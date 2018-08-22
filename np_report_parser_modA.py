import glob
import shutil
import re
from typing import List, Union

fileList_working = []

def get_files(spec = 0):

    #excluded_files = []
    if spec == 0:
        fileList_0 = sorted(glob.glob('R:/groups/seeley_pathology/NP Reports/Finalized/*.doc*'))
    else:
        fileList_0 = spec
    fileList = [f for f in fileList_0 if not '$' in f]
    #print(excluded_files)
    #print(len(excluded_files))
    #print(len(excluded_files))
    return fileList

"""
print('\nResults of get_files:')
for filename in get_files():
    print(filename)
print(len(get_files()))
"""

def get_pNums(spec=0):
    if spec == 0:
        fileList = get_files()
    else:
        fileList = spec

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

    if spec == 'duplicates':
        return dup_pNums
    else:
        return sorted(pNum_list)

"""
print('\nResults of get_pNums:')
print(get_pNums())
print(len(get_pNums()))
"""

def get_dup_pNums(spec = 0):

    pNum_list = []
    dup_pNums = []

    if spec == 0:
        fileList = get_files()
    else:
        fileList = spec

    for filename in fileList:
        f = filename.split('_')
        f = f[3]
        pNum = f[:5]

        if pNum in pNum_list:
            dup_pNums.append(pNum)
        else:
            pNum_list.append(pNum)

    return dup_pNums

"""
print('\nResults of get_dup_pNums:')
for pNum in sorted(get_dup_pNums()):
    print(pNum)
print(len(get_dup_pNums()))
"""



def remove_dup_pNums(spec = 0):

    if spec == 0:
        fileList = get_files()
    else:
        fileList = spec

    #global fileList_working
    #if not fileList_working:
    #    fileList = get_files()
    #else:
        #fileList = fileList_working

    dup_pNums = get_dup_pNums(spec)
    if not dup_pNums:
        print('\nWorking file list (no duplicate P numbers)')

    else:
        for pNum in dup_pNums:

            dup_pNum_files = []

            for filename in fileList:
                if pNum in filename:
                    dup_pNum_files.append(filename) # I think the problem is here.

            i = 1
            pNum_choice_numbers = [i]

            print('Which of these files would you like to keep in the queue?')

            if len(dup_pNum_files) > 1:
                for filename in dup_pNum_files:

                    print(f'{i}: {filename}')

                    i = i+1
                    pNum_choice_numbers.append(i)

                usr_choice = int(input('Please type the number of file you would like to keep (on the left):'))

                for filename in dup_pNum_files:
                    if dup_pNum_files.index(filename) != (usr_choice - 1):
                        fileList.remove(filename)

    fileList_working = fileList

    return fileList_working

fileList_working_0 = remove_dup_pNums()
fileList_working = remove_dup_pNums(fileList_working_0)

print('\nResults of remove_dup_pNums (fileList_working_0):')
for filename in fileList_working_0:
    print(filename)
print(len(fileList_working_0))

"""
print('\nResults of remove_dup_pNums (fileList_working):')
for filename in fileList_working:
    print(filename)
print(len(fileList_working))

## Okay, this seems chill...

## Let's try and actually get in the files...
"""


#print('Result of get_pNums:')
#get_pNums()
#get_dup_pNums()
#print('Result of remove_dup_pNums:')
#remove_dup_pNums()
#print('Result of ')
#get_files()

