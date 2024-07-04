from io import BytesIO
import streamlit as st
import docx
from pandas import DataFrame
import re
from urllib.parse import quote
import copy

st.set_page_config(layout="wide")

st.title("神秘的数据统计工具")

# data_doc_path = "/Users/lantian/Desktop/zhouji/2023年08月质控简报.docx"
uploaded_file = st.file_uploader("选择简报文件(docx)", type="docx")
if not uploaded_file:
  st.stop()

filename_match = re.match(r"(\d+)年(\d+)月\.*", uploaded_file.name)
if not filename_match:
  st.error("无法从文件名中解析年份和月份")
  st.stop()

year = int(filename_match.group(1))
month = int(filename_match.group(2))

record_doc = docx.Document(uploaded_file)
tables = record_doc.tables


def remove_blanks(s):
  return s.replace("\n", "").replace(" ", "").replace("\u3000", "")

def is_header_text(s):
  return "指标" in s or "项目" in s or "科室" in s or "科别" in s

data_tables: list[list[list[str]]] = []
for table in tables:
  data_table: list[list[str]] = []
  for row in table.rows:
    data_table.append([remove_blanks(cell.text) for cell in row.cells])
  data_tables.append(data_table)

all_keshi: set[str] = set()
for row in data_tables[1]:
  name = remove_blanks(row[0])
  if name and name != "项目" and name != "科室":
    all_keshi.add(name)

all_keshi_sorted = sorted(all_keshi)

col1, _, _ = st.columns(3)
with col1:
  index = 0
  if "keshi" in st.session_state and st.session_state.keshi in all_keshi:
    index = all_keshi_sorted.index(st.session_state.keshi)

  # def handle_keshi_change():
  keshi = st.selectbox("选择科室", all_keshi_sorted, key="keshi_selector", index=index)
  st.session_state.keshi = keshi

# 合并多列表格成一列
for i in range(len(data_tables)):
  table = data_tables[i]
  if len(table) < 1:
    continue
  headers = table[0]
  first_header = headers[0]
  same_headers = [value for value in headers if value == first_header]
  step = int(len(headers) / len(same_headers))
  if (len(headers) != len(same_headers) and
      "医院名称" not in first_header and
      len(same_headers) >= 2 and
      "".join(headers[0:step]) == "".join(headers[step:step+step])
  ):
    new_table: list[list[str]] = []
    for i, row in enumerate(table):
      new_table.append(row[0: step])
      if i > 0:
        for j in range(step, len(headers), step):
          new_table.append(row[j:j+step])
    data_tables[i] = new_table

def get_first_text(row):
  first_text = remove_blanks(row[0])
  if len(row) > 1:
    first_text += remove_blanks(row[1])
  if len(row) > 2:
    first_text += remove_blanks(row[2])
  return first_text

def filter_row(i: int, row: list[str]):
  first_text = get_first_text(row)
  is_header = is_header_text(first_text)
  is_keshi = i >= 1 and (keshi in row[0] or (len(row) > 1 and keshi in row[1]))
  return is_header or is_keshi

filtered_data: list[list[list[str]]] = []
for table in data_tables:
  filtered_data.append([row for i, row in enumerate(table) if filter_row(i, row)])

targets = [
  ["病床使用率",r"病床使用率"],
  ["病床周转率",r"病床周转率"],
  ["平均住院日",r"平均住院日"],
  ["治愈好转率",r"好转率"],
  ["入院与出院诊断符合率",r"入院与出院诊断符合率"],
  ["成分输血率",r"成份输血率"],
  ["抗菌药物使用率",r"抗菌药物使用率"],
  ["抗菌药物使用强度",r"抗菌药物使用强度"],
  ["药物构成比",r"业务收入不含耗材收入药占比"],
  ["临床路径入径率",r"临床路径入径率"],
  ["临床路径完成率",r"临床路径完成率"],
  ["临床路径覆盖率",r"临床路径覆盖率"],
  ["院内感染发生率",r"^感染率%$"],
]
results = {}
for target, search_text in targets:
  results[target] = [target, search_text, "", ""]

found_targets = set()

def find_value_v3(table, target_row_name, target_column_name):
  if len(table) == 0:
    return ["", ""]

  headers = table[0]
  column_index = next(filter(
    lambda x: re.search(target_column_name, headers[x]),
    range(len(headers))
  ), None)
  if column_index is None:
    return ["", ""]

  found_values = []
  for row in table:
    row_name = row[0] + row[1]
    cell = row[column_index]
    if target_row_name in row_name or ("科室" in row_name and "≥" in cell):
      found_values.append(cell)

  if len(found_values) == 0:
    return ["", ""]
  if len(found_values) == 2:
    expected, actual = found_values
    return [actual, expected]
  return [found_values[0], ""]

for table in filtered_data:
  for target, search_text in targets:
    if target in found_targets:
      continue

    actual, expected = find_value_v3(table, target_row_name=keshi, target_column_name=search_text)
    if actual:
      found_targets.add(target)
      results[target] = [target, search_text, actual, expected]

# 特别处理的指标
# 病种总例数＝表单内所有疾病名称下，病例数的总和。例如放疗科是
# 入径率＝入径数/病种总病例数
# 完成率＝变异完成例数/病种总例数。
def format_float(f):
  res = f"{f:.2f}"
  if res[-1] == "0":
    res = res[:-1]
  return res

def get_special_metrics():
  tables = [t for t in filtered_data if len(t) > 0 and "病种名称" in t[0]]
  if len(tables) == 0 or len(tables[0]) < 2:
    return {}
  table = tables[0]
  total_col = 2
  rujing_col = 3
  finish_col = 5
  total = 0
  rujing = 0
  finish = 0
  data_rows = table[1:]
  for row in data_rows:
    total += int(row[total_col])
    rujing += int(row[rujing_col])
    finish += int(row[finish_col])
  return {
    "临床路径入径率": ["临床路径入径率", "= 入径数/病种总病例数", format_float(rujing / total * 100), ""],
    "临床路径完成率": ["临床路径完成率", "= 变异完成例数/病种总例数", format_float(finish / total * 100), ""],
    "重点疾病例数": ["重点疾病例数", "= 病种总例数", format_float(total), ""]
  }

results.update(get_special_metrics())

results_df = DataFrame(results.values(), columns=["指标", "搜索方法", "实际值", "目标值"])
results_df.sort_values(by="指标")

# 找到诊断问题
def search_paragraph(search_text):
  results = []
  for paragraph in record_doc.paragraphs:
    text = remove_blanks(paragraph.text)
    if search_text in text:
      results.append(text)
  return results

wrong_diagnose = search_paragraph(f"{keshi}处方号")
wrong_diagnose = [f"（{i+1}）{text[4:]}" for i, text in enumerate(wrong_diagnose)]

# 找到环节病历
def get_bingli():
  tables = [t for t in filtered_data if len(t) > 0 and "患者姓名" in t[0] and "住院号" in t[0]]
  if len(tables) == 0 or len(tables[0]) < 2:
    return []
  table = tables[0]
  bingli = [row[0: 5] for row in table[1:]]
  return bingli

bingli = get_bingli()


# st.write("## 结果下载")
output_file_path = "./放射治疗科2023年03月医疗质量与安全检查记录.docx"
output_doc = docx.Document(output_file_path)
output_tables = output_doc.tables

# 写入指标
metrics_table = [table for table in output_tables if remove_blanks(table.rows[0].cells[0].text) == "指标"][0]
for row in metrics_table.rows:
  metrics_name = remove_blanks(row.cells[0].text)
  if metrics_name in results:
    _, _, actual, expected = results[metrics_name]
    row.cells[1].text = actual
    row.cells[2].text = expected
  
  metrics_name = remove_blanks(row.cells[3].text)
  if metrics_name in results:
    _, _, actual, expected = results[metrics_name]
    row.cells[4].text = actual
    row.cells[5].text = expected

# 写入甲级环节病例
jiaji_bingli = [item for item in bingli if "甲级" in item[4]]
temp_table = []
bingli_table = output_tables[1].cell(0, 1).tables[0]

def copy_cell_properties(source_cell, dest_cell):
    cell_properties = source_cell._tc.get_or_add_tcPr()
    dest_cell._tc.remove(dest_cell._tc.get_or_add_tcPr())
    cell_properties = copy.copy(cell_properties)
    dest_cell._tc.append(cell_properties)

def copy_row_properties(source_row, dest_row):
  for i in range(len(source_row.cells)):
    source_cell = source_row.cells[i]
    dest_cell = dest_row.cells[i]
    copy_cell_properties(source_cell, dest_cell)

if len(jiaji_bingli) > 3:
  first_row = bingli_table.rows[0]
  for i in range(3, len(jiaji_bingli)):
    new_row = bingli_table.add_row()
    copy_row_properties(first_row, new_row)

for i, item in enumerate(jiaji_bingli):
  _, huanzhe, zhuyuanhao, problem, level = item
  row = bingli_table.rows[i+1]
  row.cells[0].text = str(i+1)
  for j in range(1, 5):
    row.cells[j].text = item[j]

# 写入诊断问题
wenti_cell = output_tables[1].cell(0, 10)
wenti_cell.text = ""
for text in wrong_diagnose:
  if remove_blanks(wenti_cell.text) == "":
    wenti_cell.text = text
  else:
    wenti_cell.add_paragraph(text)

# 写入乙级病例
yiji_bingli = [item for item in bingli if "乙级" in item[4]]
for i, item in enumerate(yiji_bingli):
  _, huanzhe, zhuyuanhao, problem, level = item
  text = f"（{len(wrong_diagnose)+i+1}）患者{huanzhe}（住院号：{zhuyuanhao}），存在问题：{problem}"
  if remove_blanks(wenti_cell.text) == "":
    wenti_cell.text = text
  else:
    wenti_cell.add_paragraph(text)

# 替换科室和年月
for p in output_doc.paragraphs:
  p.text = p.text.replace("放射治疗科", f"{keshi}").replace(f"2023年03月", f"{year:04d}年{month:02d}月")

output = BytesIO()
output_doc.save(output)

st.write("## 下载统计结果")

download_filename = f"{keshi}{year:04d}年{month:02d}月医疗质量与安全检查记录.docx"
st.download_button(
    label=f"💾 {download_filename}",
    data=output.getvalue(),
    file_name=quote(download_filename),
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)

### 显示统计结果

# st.write("## 解析结果预览")
# st.write("**指标搜索结果**")
# col1, _ = st.columns(2)
# with col1:
#   st.table(results_df)

# wrong_diagnose_df = DataFrame(wrong_diagnose, columns=["门急诊处方点评公布"])
# st.write("**门急诊处方点评公布**")
# st.table(wrong_diagnose_df)

# bingli_df = DataFrame(bingli, columns=["科室", "患者姓名", "住院号", "存在问题", "病历等级"])
# st.write("**环节病历**")
# st.table(bingli_df)

# st.write("----")
# st.write("## 参考数据")

# for table_data in filtered_data:
#   if len(table_data) >= 1:
#     st.table(table_data)
