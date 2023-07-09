import os
import re
import pandas as pd

from expirment_report import combine_word_documents, get_file_names, read_reports, output_list

folder_path = 'data/云浮'
student_list_file = 'data/期末考试名单_移动应用开发_计算机科学与技术20-云浮.xlsx'
class_id = '1'
output_file = 'data/移动应用开发_实验报告提交情况表.xlsx'
sheet_name = '外包20'

# Main program start here:
if __name__ == '__main__':
    reports = read_reports(folder_path=folder_path, student_list_file=student_list_file,
                           class_id='1')
    if reports:
        output_list(reports=reports, output_file=output_file, output_sheet=sheet_name)
    print('finished')

