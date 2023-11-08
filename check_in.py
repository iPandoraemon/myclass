import os
import re
import pandas as pd


if __name__ == '__main__':
    print(os.getcwd())
    path = input('考勤文件夹路径：')
    my_class = eval(input('智医21输入0，数媒21输入1：'))
    if my_class == 0:
        my_class = "智医21"
    elif my_class == 1:
        my_class = "数媒21"
    else:
        print("输入错误！")
        exit()

    files = os.listdir(path)
    summary_file = [file for file in files if '考勤记录' in file and my_class in file][0]
    checkin_files = [file for file in files if '课考勤' in file]

    for file in checkin_files:
        print(f'Checking file {path}/{file}...')
        xls = pd.ExcelFile(path+'/'+file)
        sheet_names = xls.sheet_names
        sheet_name = [s for s in sheet_names if "已停用" not in s][0]
        checkin = pd.read_excel(xls, sheet_name=sheet_name, header=0)
        checkin['stu_num'] = checkin['学号+姓名'].apply(lambda x: re.findall(r'[\d\u4e00-\u9fa5]+', x)[0])
        checkin['name'] = checkin['学号+姓名'].apply(lambda x: re.findall(r'[\d\u4e00-\u9fa5]+', x)[1])
        checkin['date'] = checkin['提交时间'].apply(lambda x: x.strftime('%Y%m%d%H%p'))

        # 获取每次考勤的日期
        dates = checkin['date'].drop_duplicates().tolist()

        summary = pd.read_csv(path+'/'+summary_file, header=0)

        for d in dates:
            if d in summary.columns.tolist():
                print("{date}已经记录了".format(date=d))
                continue
            chk_list = checkin[checkin['date'] == d]
            chk_list.index = chk_list['stu_num'].to_list()
            chk_list = chk_list[['name']]
            chk_list[d] = 1
            summary = summary.merge(right=chk_list, how="left")
            summary.fillna(value=0, inplace=True)

        summary.to_csv(path+'/'+summary_file, encoding='utf-8', index=False)
        print("{}/{} is updated and saved.".format(path, summary_file))
