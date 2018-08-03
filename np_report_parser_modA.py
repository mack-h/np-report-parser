import glob
import re
from typing import List, Union


def get_files():
    fileList = glob.glob('R:/groups/seeley_pathology/NP Reports/Finalized/*.docx')
    for filename in fileList:
        print(filename)

def get_pNums():
    fileList = glob.glob('R:/groups/seeley_pathology/NP Reports/Finalized/*.docx')
    pNum_list= []
    for filename in fileList:
        f = filename.split('_')
        f = f[3]
        pNum = f[:5]
        pNum_list.append(pNum)
    print(sorted(pNum_list))


get_pNums()