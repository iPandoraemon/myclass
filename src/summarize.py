'''
Filename        : summarize.py
Author          : Wen, T.
Create Date     : 2022.9.27
Descriptions    :
1. Collect homework from students and summarize their student IDs.
'''

import os
import re


def collect_id(filedir):
    '''
    Collect student IDs

    Param filedir is the folder name.

    Return a dictionary type.
    '''
    ndict = {}
    for dirpath, dirnames, filenames in os.walk(top=filedir, topdown=True):
        for name in filenames:
            name = re.split('\W', name)

            # 找出学生ID
            for item in name:
                if item.isdigit():
                    id = item

            ndict[name[0]] = int(id)
    return ndict


def copies_count(ndict):
    '''
    Count how many copies of homeworks.

    ---------------------------------------------------------------------------
    Param *ndict*: the dictionary of student names and their IDs.

    ---------------------------------------------------------------------------
    Return:
    cp_num: integer, the number of copies collected.
    '''
    return len(ndict)

# Main program start heres
if __name__ == '__main__':

    filedir = input('Give the absolute folder path: ')
    ndict = collect_id(filedir)
    print('There are %d students submitted the job.'%copies_count(ndict))