import pandas as pd
import os
import re


def get_all_files(path):
    files = os.listdir(path)
    for file_name in files:
        print(file_name)


CHECKIN_FILE_DIR = './data/考勤/移动应用开发_计科21/'
CHECKIN_SUMMARY_FILE = './data/考勤/考勤统计_移动应用开发-理论_23-24-2_计算机科学与技术21.xlsx'
SHEET_NAME = ['1班', '2班', '3班', '4班']
CHECKIN_STATES = {0: "未参与",
                  1: "已签",
                  2: "教师代签",
                  3: "事假",
                  4: "病假",
                  5: "迟到"}
CHECKIN_FILENAME_PATTERN = r"^(.*)-(.*)-(.*)_(.*)\.xlsx$"
CHECKIN_INFO_PATTERN1 = r"第(\d+)周(.*?)课(.*?)班"
CHECKIN_INFO_PATTERN2 = r"第(\d+)周(.*?)课(.*?)班(.*?)批"


def extract_checkin_info_from_filename(filename):
    filename_pattern = re.compile(CHECKIN_FILENAME_PATTERN)
    filename_match = filename_pattern.search(filename)
    if filename_match:
        # 提取子组内容
        subject = filename_match.group(1)
        class_name = filename_match.group(2)
        check_info = filename_match.group(4)

        info_pattern1 = re.compile(CHECKIN_INFO_PATTERN1)
        info_match1 = info_pattern1.search(check_info)
        if info_match1:
            week = info_match1.group(1)
            class_type = info_match1.group(2)
            class_no = info_match1.group(3)
            group_no = ''

        else:
            print("未找到匹配的子串")

        info_pattern2 = re.compile(CHECKIN_INFO_PATTERN2)
        info_match2 = info_pattern2.search(check_info)
        if info_match2:
            week = info_match2.group(1)
            class_type = info_match2.group(2)
            class_no = info_match2.group(3)
            group_no = info_match2.group(4)

        else:
            print("未找到匹配的子串")

        print("subject:", subject)
        print("class: {name}({no}班)".format(name=class_name, no=class_no))
        print("week: ", week)
        print("type: ", class_type)
        print("group: ", group_no)
        return {"subject": subject,
                "class_type": class_type,
                "week": week,
                "class_name": class_name,
                "class_no": class_no,
                "group_no": group_no}


if __name__ == "__main__":
    print("Checkin file dir: {}".format(CHECKIN_FILE_DIR))
    print("Checkin summary file: {}".format(CHECKIN_SUMMARY_FILE))
    print("Checkin summary sheet names: {}".format(SHEET_NAME))
    filelist = os.listdir(CHECKIN_FILE_DIR)

    # Filter *.xlsx files only
    target_extension = ".xlsx"
    filtered_files = [filename for filename in filelist if filename.endswith(target_extension)]

    checkin_sheets = pd.read_excel(CHECKIN_SUMMARY_FILE, sheet_name=None, header=0)

    for filename in filtered_files:
        print("Working for file {}".format(filename))
        filedir = CHECKIN_FILE_DIR + filename
        df = pd.read_excel(filedir, header=5)
        selected_columns = df[["姓名", "学号/工号", "签到状态", "时间"]]

        checkin_info = extract_checkin_info_from_filename(filename)












