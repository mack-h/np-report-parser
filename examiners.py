import parser_functions
import re

## Given how similar these all are, examiner could be a class

def get_wmr(pNum, lines):

    pNums = []
    for item in lines:
        if 'pallor' in item:
            #if any(region in item.upper() for region in (cortical_regions or deep_gray)):
             print(f'{pNum}: {item}')
             pNums.append(pNum)
    for item in lines:
          if 'rarefaction' in item:
            #if any(region in item.upper() for region in (cortical_regions or deep_gray)):
            print(f'{pNum}: {item}')
            pNums.append(pNum)
    pNums = set(pNums)
    print(len(pNums))

def examine_microbleeds(pNum, lines):
    pNums = []
    for item in lines:
        if 'bleed' in item:
            print(f'{pNum}: {item}')
            pNums.append(pNum)

    pNums = set(pNums)
    print(len(pNums))

def examine_CTE(pNum, lines):
    pNums = []

    for item in lines:
        if 'hronic traumatic encephalopathy' in item:
            print(f'{pNum}: {item}')
            pNums.append(pNum)

    pNums = set(pNums)
    print(len(pNums))

def examine_ARTAG(pNum, lines):
    pNums = []

    for item in lines:
        if 'ARTAG' in item:
            print(f'{pNum}: {item}')
            pNums.append(pNum)

    pNums = set(pNums)
    print(len(pNums))

def examine_ATAC(pNum, lines):
    pNums = []

    for item in lines:
        if 'ATAC' in item:
            print(f'{pNum}: {item}')
            pNums.append(pNum)

    pNums = set(pNums)
    print(len(pNums))

def examine_HS(pNum, lines):
    pNums = []

    for item in lines:
        if 'HIPPOCAMPAL SCLEROSIS' in item:
            print(f'{pNum}: {item}')
            pNums.append(pNum)

    if not pNums:
        print(f'No hippocampal sclerosis mentioned for {pNum}')
   #pNums = set(pNums)
   # print(len(pNums))

def examine_AGD(pNum, lines):
    pNums = []

    for item in lines:
        if 'rgyrophilic grain' in item:
            print(f'{pNum}: {item}')
            pNums.append(pNum)

    pNums = set(pNums)
    print(len(pNums))

def examine_LBD(pNum, dx_sxn):
    pNums = []

    for item in dx_sxn:
        if 'LEWY' in item:
            if 'BRAAK' not in item:
            #if any(region in item.upper() for region in (cortical_regions or deep_gray)):
                print(f'{pNum}: {item}')
                pNums.append(pNum)
    pNums = set(pNums)
    print(len(pNums))

def examine_TDP_nonFTLD(pNum, dx_sxn):
    pNums = []

    for item in dx_sxn:
        if 'TDP-43 PROTEINOPATHY' in item:
            #if any(region in item.upper() for region in (cortical_regions or deep_gray)):
            print(f'{pNum}: {item}')
            pNums.append(pNum)
    pNums = set(pNums)
    print(len(pNums))

def examine_VBI(pNum, dx_sxn):
    pNums = []

    for item in dx_sxn:
        if 'VASCULAR BRAIN INJURY' in item:
            print(f'{pNum}: {item}')
            pNums.append(pNum)
    pNums = set(pNums)
    for pNum in pNums:
        print(f'{pNum}')
    print(len(pNums))
