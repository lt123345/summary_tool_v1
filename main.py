import streamlit as st
import docx
from pandas import DataFrame
import re

st.set_page_config(layout="wide")

st.title("神秘的数据统计工具")

data_doc_path = "/Users/lantian/Desktop/zhouji/2023年08月质控简报.docx"
record_doc = docx.Document(data_doc_path)
tables = record_doc.tables

keshi = "放射治疗科"

def remove_blanks(s):
  return s.replace("\n", "").replace(" ", "")

def is_header_text(s):
  return "指标" in s or "项目" in s or "科室" in s or "科别" in s

data_tables: list[list[list[str]]] = []
for table in tables:
  data_table: list[list[str]] = []
  for row in table.rows:
    data_table.append([remove_blanks(cell.text) for cell in row.cells])
  data_tables.append(data_table)

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

st.write("## 解析结果")

results_df = DataFrame(results.values(), columns=["指标", "搜索方法", "实际值", "目标值"])
results_df.sort_values(by="指标")
col1, _ = st.columns(2)
with col1:
  st.table(results_df)

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
wrong_diagnose_df = DataFrame(wrong_diagnose, columns=["门急诊处方点评公布"])
st.table(wrong_diagnose_df)

st.write("----")
st.write("## 参考数据")

for table_data in filtered_data:
  if len(table_data) >= 1:
    st.table(table_data)
