import re
import os
import pandas as pd


if __name__ == "__main__":
    print(os.getcwd())
    path = input('实验报告文件夹路径：')
    my_class = eval(input('智医21输入0，数媒21输入1：'))
    if my_class == 0:
        my_class = "智医21"
    elif my_class == 1:
        my_class = "数媒21"
    else:
        print("输入错误！")
        exit()

    files = os.listdir(path)
    summary_file = [file for file in files if '实验报告收集记录' in file and my_class in file][0]
    checkin_files = [file for file in files if re.search(r'实验报告\d', file)]

    summary = pd.read_csv(path+'/'+summary_file, header=0, index_col=0)
    for file in checkin_files:
        print(f'Checking file: {file}...')
        checkin = pd.read_excel(path+'/'+file, header=0)
        checkin['stu_num'] = checkin['你的姓名和学号：'].apply(lambda x: re.findall(r'[\d]+', x)[0])
        checkin['name'] = checkin['你的姓名和学号：'].apply(lambda x: re.findall(r'[\u4e00-\u9fa5]+', x)[0])

        report = re.findall(r'实验报告\d', file)[0]
        print("处理{report}".format(report=report))

        if report in summary.columns.tolist():
            print(f'{report} is already existed, replace now...')
            summary.drop(columns=[report], inplace=True)

        checkin.index = checkin['stu_num']
        checkin[report] = 1
        checkin.index = checkin.index.astype('int64')
        checkin = checkin[[report]]
        summary = pd.merge(left=summary, right=checkin, left_index=True, right_index=True,
                           how='left')
        # summary.merge(right=checkin, how="right", copy=True)

    print(f'Save output file in {path}/{summary_file}')
    summary.drop_duplicates(inplace=True)
    summary.to_csv(path+'/'+summary_file, index=True, encoding='utf-8')
