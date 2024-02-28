import pandas as pd
import os
import re


def get_all_files(path):
    files = os.listdir(path)
    for file_name in files:
        print(file_name)


SHEET_NAME = '签到详情'

if __name__ == "__main__":
    filedir = input("输入你的路径")
    files = os.listdir(filedir)

    student_list_file = './data/考勤/数媒21/学生名单_移动应用开发2023-2024-1_数媒21.csv'
    checkin_report = pd.read_csv(student_list_file, header=0, index_col=0)
    checkin_report.index.astype(str, copy=False)

    for file in files:
        file_split = re.split(r'[-.]', file)
        cls_name, filetype = file_split[3], file_split[-1]

        if filetype != "xlsx":
            continue

        if "补签" in cls_name:
            cls_name = cls_name[:-2]

        df = pd.read_excel(f'{filedir}/{file}', sheet_name=SHEET_NAME)
        cols = df.iloc[4].values.tolist()
        df.columns = cols
        df = df.iloc[5:]
        df.index = df['学号/工号'].astype('int64')
        df.index.dtype

        if cls_name in checkin_report.columns.tolist():
            checkin_report[cls_name] = df['签到状态'].combine_first(checkin_report[cls_name])
        else:
            checkin_report = checkin_report.merge(right=df['签到状态'].to_frame(cls_name),
                                                  left_index=True,
                                                  right_index=True, how="left")

    checkin_report = checkin_report.sort_index(axis=1)

    checkin_types = ['已签', "教师代签", "病假", "事假"]
    counts = pd.Series([0]*checkin_report.shape[0], index=checkin_report.index)
    for t in checkin_types:
        counts = counts + checkin_report.apply(lambda x: x.value_counts().get(t, 0), axis=1)
    checkin_report['有效合计'] = counts

    checkin_report.to_excel('./data/考勤/考勤统计_移动应用开发_2023-2024-1_数媒.xlsx', index=True)
