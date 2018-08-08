import glob
import re
from typing import List, Union


def get_files():
    excluded_files = []
    fileList = sorted(glob.glob('R:/groups/seeley_pathology/NP Reports/Finalized/*.doc*'))
    for filename in fileList:
        if '~' in filename:
            excluded_files.append(filename)
            fileList.remove(filename)
        #else:
            #print(filename)
    #print(excluded_files)
    #print(len(excluded_files))
    return fileList

def get_pNums():
    fileList = get_files()
    pNum_list= []
    dup_pNums = []
    for filename in fileList:
        f = filename.split('_')
        f = f[3]
        pNum = f[:5]
        if pNum in pNum_list:
            dup_pNums.append(pNum)
        else:
            pNum_list.append(pNum)
    print(sorted(pNum_list))
    print(len(pNum_list))
    for pNum in sorted(dup_pNums):
        print(pNum)
    print(len(dup_pNums))

# This would be a good module on it's own, which calls get_pNums.
    dup_pNum_files = []
    for pNum in dup_pNums:
        for filename in fileList:
            if pNum in filename:
                dup_pNum_files.append(filename)
                # This is where a break/out of loop needs to happen.
                i = 1
                pNum_choice_numbers = [i]
                print('Which of these files would you like to keep in the queue?')
                for filename in dup_pNum_files:
                    print(f'{i}: {filename}')
                    i = i+1
                    pNum_choice_numbers.append(i)
                print(dup_pNum_files)
                usr_choice = int(input('Please type the number of file you would like to keep (on the left):'))
            #if usr_choice not in pNum_choice_numbers:

                for filename in dup_pNum_files:
                    if dup_pNum_files.index(filename) != (usr_choice - 1):
                        fileList.remove(filename)






get_pNums()
#get_files()
