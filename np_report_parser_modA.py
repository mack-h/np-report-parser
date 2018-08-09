import glob
import re
from typing import List, Union

fileList_working = []

def get_files():

    #excluded_files = []
    fileList_0 = sorted(glob.glob('R:/groups/seeley_pathology/NP Reports/Finalized/*.doc*'))
    fileList = [f for f in fileList_0 if not '$' in f]
    #print(excluded_files)
    #print(len(excluded_files))
    #print(len(excluded_files))
    return fileList

"""
    for filename in fileList:

        if 'NP' not in filename:
            excluded_files.append(filename)
            fileList.remove(filename)
"""
        #else:
            #print(filename)

for filename in get_files():
    print(filename)
print(len(get_files()))



def get_pNums(spec=0):

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

    if spec == 'duplicates':
        return dup_pNums
    else:
        return sorted(pNum_list)

print(get_pNums())
print(len(get_pNums()))

def get_dup_pNums():

    dup_pNums = get_pNums('duplicates')

    return dup_pNums



for pNum in sorted(get_dup_pNums()):
    print(pNum)
print(len(get_dup_pNums()))




def remove_dup_pNums():

    global fileList_working
    dup_pNums = get_dup_pNums()

    if not fileList_working:
        fileList = get_files()
    else:
        fileList = fileList_working

    for pNum in dup_pNums:

        dup_pNum_files = []

        for filename in fileList:
            if pNum in filename:
                dup_pNum_files.append(filename)

        i = 1
        pNum_choice_numbers = [i]

        print('Which of these files would you like to keep in the queue?')

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




fileList_working = remove_dup_pNums()
for filename in fileList_working:
    print(filename)
print(len(fileList_working))

#get_pNums()
#get_dup_pNums()
#remove_dup_pNums()
#get_files()

