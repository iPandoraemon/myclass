import os
import re
import pandas as pd


def is_number(string):
    return bool(re.match('^[0-9]+$', string))


def is_number_and_other(string):
    return all(char.isdigit() or not char.isalnum() for char in string)


folder_path = 'data\大学城-计算机外包'
student_list_file = 'data\移动应用开发-计算机外包20-学生名单.xlsx'
class_id = '1'
output_file = 'data\移动应用开发_实验报告提交情况表.xlsx'
sheet_name = '外包20'

exp_list = []
for root, dirs, files in os.walk(folder_path):
    for dir_name in dirs:
        exp_list.append(dir_name)

reports = pd.read_excel(student_list_file, usecols=['学号', '姓名'])
student_num = reports.shape[0]
reports['学号'] = reports['学号'].astype(str)
# for exp in exp_list:
#     reports[exp] = [0 for i in range(student_num)]

for root, dirs, files in os.walk(folder_path):
    students = {}
    for file in files:
        split_root = root.split('\\')
        exp_name = split_root[-1]
        split_file = re.split('[-.(]', file)
        if is_number(split_file[0]):
            if not is_number(split_file[1]):
                student_name = split_file[1]
                student_id = split_file[0]
        elif re.findall(r'\d+', split_file[0]):
            student_name = re.findall(r'\D+', split_file[0])[0]
            student_id = re.findall(r'\d+', split_file[0])[0]
        else:
            student_name = split_file[0]
            student_id = split_file[1]
        student_name = re.findall(r'\D+', student_name)[0]
        student_name = student_name.strip()
        student_id = re.findall(r'\d+', student_id)[0]
        student_id = student_id.strip()
        file_format = split_file[2]
        if len(student_id) < 10:
            student_id = '2000502' + class_id + student_id[-2:]
        students.setdefault(student_id, [student_id, 1])
        # reports.loc[reports['id'] == student_id, ['name']] = student_name
        # reports.loc[reports['id'] == student_id, [exp_name]] = 1
    if students:
        df = pd.DataFrame(students)
        df = df.T
        df = df.rename(columns={0: '学号', 1: exp_name})
        # df = df.rename(columns={0: '学号', 1: '姓名', 2: exp_name})
        reports = reports.merge(df, on='学号', how='left')
        print(exp_name+' 结束')

with pd.ExcelWriter(output_file) as writer:
    reports.to_excel(writer, sheet_name=sheet_name)

