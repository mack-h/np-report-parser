import glob
import re
import csv
import pandas
import xlrd
import xlsxwriter
import win32com.client as win32
import os
import docx
#from examiners import examine_HS
from docx import Document


all_txt_files = sorted(glob.glob('R:/groups/seeley/Mack/NP report parser/TXT files/*.txt'))
all_word_files = sorted(glob.glob('R:/groups/seeley/Mack/NP report parser/Word files/*'))

specific_file = ''

# Specify P number pattern to search filename for: P followed by 4 digits
pNum_pattern = re.compile('P\d\d\d\d')

# File paths to NP Workflow (to get author info), NP Slide Review Sheet (for H&E tables)
SlideReviewSheet = 'R:/groups/seeley/Mack/NP report parser/databases/Copy of ADRC_NPSlideReviewSheet.xls'
NPWorkflow = 'R:/groups/seeley_pathology/databases/Neuropathology Diagnostic Workflow_20180710.xlsx'
HE_datasheet =  'R:/groups/seeley/Mack/NP report parser/databases/HE_SlideReviewData.xlsx'

def line_after(index):
    index = index + 1
    return index

def get_files(pNum_list=['all']):
    working_file_dict = {}
    if 'all' in pNum_list:
        for filename in all_txt_files:
            pNum = re.findall(pNum_pattern, filename)[0]

            #print(pNum)
            working_file_dict[pNum] = filename

    elif isinstance(pNum_list, list):
        for filename in all_txt_files:
            for pNum in pNum_list:
                if pNum in filename:
                    working_file_dict[pNum] = filename
    else:
        pNum_list = get_pNums()
        if 'all' in pNum_list:
            get_files(pNum_list)
        else:
            for pNum in pNum_list:
                for filename in all_txt_files:
                    if pNum in filename:
                        working_file_dict[pNum] = filename
    print(working_file_dict)
    return working_file_dict

def get_Word_files(pNum_list=['all']):
    working_file_dict = {}
    if 'all' in pNum_list:
        for filename in all_word_files:
            pNum = re.findall(pNum_pattern, filename)[0]

            #print(pNum)
            working_file_dict[pNum] = filename

    else:
        pNum_list = get_pNums()
        if 'all' in pNum_list:
            get_files[pNum_list]
        else:
            for pNum in pNum_list:
                for filename in all_txt_files:
                    if pNum in filename:
                        working_file_dict[pNum] = filename
    #print(working_file_dict)
    return working_file_dict


def get_pNums():

    input1 = input('To type in P numbers manually, enter 1.\nTo use working_pNums file, enter 2.\nTo loop through all reports, enter "all".\nEnter:\n')
        ## PLACEHOLDER FOR SYNERGIZING WITH CASE FINDER

    if input1 == '1':
        pNum_list = input('Type in P numbers, separated by commas (no spaces):')
        pNum_list = pNum_list.split(',')
        print(f'P numbers: {pNum_list}')
        return pNum_list

    if input1 == '2':
        print('Pulling from working pNum list file... (under construction)')
        ## NEED TO MAKE SURE THIS SYNERGIZES WITH GET FILES
        return 'USING FILE'

    if input1 == 'all':
        return ['all']

## OPEN FILE

def open_file(working_file):
    with open(working_file, 'rt', encoding='utf-8', errors='ignore') as f:

        lines = []

        # Adds all lines of a file (demarcated by \n) to a list. Each line will be stored as a string in the list.
        for line in f:
            line = line.strip()
            if not line:
                continue
            lines.append(line)
    return lines

## GET PIDN -- USE ADRC inventory

## GET AUTHOR
def get_author(pNum):
    df = pandas.read_excel(NPWorkflow, sheet='Main', usecols=[0,1])
    try:
        author = df.loc[df['P #'] == pNum, 'Assignmt'].values[0]
    except IndexError:
        author = 'Check report'
    if author == 'Bill':
        author = 'Seeley'
    if author == 'Salvo':
        author = 'Spina'
    if author == 'Lea':
        author = 'Grinberg'
    return author

## GET DIAGNOSIS SECTION

def get_dx_sxn(filename):

    lines = open_file(filename)

    i_1 = 0
    i_2 = 0

    for line in lines:
        if any (x in line for x in ['Final Neuropathological Diagnosis', 'Final Neuropathological Diagnoses']):
            i_1 = lines.index(line)
        if 'Clinical History' in line:
            i_2 = lines.index(line) - 1

    if not i_1:
        print(f'Error finding diagnostic section for {filename}\n')
    dx_line_sxn = lines[i_1:i_2+1]
    dx_line_sxn = [item.upper() for item in dx_line_sxn]

    return dx_line_sxn

#print(len(get_dx_sxn()))
#print(type(get_dx_sxn()))

## Testing get_dx_sxn
#for line in get_dx_sxn():
#    print(line)
#print(get_dx_sxn())

## Gross obs section
def get_grossObs(filename):
    lines = open_file(filename)

    grossObs = []

    i_1 = 0
    i_2 = 0

    for line in lines:
        line2 = line.upper()

        if 'GROSS OBSERVATIONS' in line2:
            i_1 = lines.index(line)
        if 'OTHER GROSS FINDINGS' in line2:
            i_2 = lines.index(line) + 1

    if not i_1:
        print(f'No gross observations found for {filename}')
    else:
        grossObs = lines[i_1:i_2]
        grossObs = [item.upper() for item in grossObs]
        return grossObs

## Testing get_grossObs
#for line in get_grossObs(filename):
#    print(line)

def get_PrimDx(dx_sxn):
    primDx_list = []
    for line in dx_sxn:
        if 'PRIMARY DIAGNOSIS' in line:
            if ':' in line:
                dx = line[line.index(':')+2:]
                primDx_list.append(dx)
            elif '#' in line:
                dx = line[line.index('#') + 2:]
                primDx_list.append(dx)

    return primDx_list

#primDx_list = get_PrimDx()

#print(f'\nNumber of primary diagnoses: {len(primDx_list)}')
#for counter, dx in enumerate(primDx_list, 1):
#    print(f'{counter}: {dx}')

def get_ContributingDx(dx_sxn):
    contributingDx_list = []
    for line in dx_sxn:
        if 'CONTRIBUTING DIAGNOSIS' in line:
            if ':' in line:
                dx = line[line.index(':')+2:]
                contributingDx_list.append(dx)

            elif '#' in line:
                dx = line[line.index('#') + 2:]
                contributingDx_list.append(dx)

    return contributingDx_list

#contributingDx_list = get_ContributingDx()

#print(f'\nNumber of contributing diagnoses: {len(contributingDx_list)}')
#for counter, dx in enumerate(contributingDx_list, 1):
#    print(f'{counter}: {dx}')

def get_IncidentalDx(dx_sxn):

    IncidentalDx_list = []

    for line in dx_sxn:
        if 'INCIDENTAL DIAGNOSIS' in line:
            if ':' in line:
                dx = line[line.index(':')+2:]
                IncidentalDx_list.append(dx)
            elif '#' in line:
                dx = line[line.index('#') + 3:]
# This is to account for typos where there is no colon following a number.
                IncidentalDx_list.append(dx)

    return IncidentalDx_list

#incidentalDx_list = get_IncidentalDx()

#print(f'\nNumber of incidental diagnoses: {len(incidentalDx_list)}')
#for counter, dx in enumerate(incidentalDx_list, 1):
#    print(f'{counter}: {dx}')

def get_ADNC(lines):

    ADNC_raw_data = []
    ADNC_values = []
    ADNC_keys = ['ADNC level', 'ADNC_level', 'Thal Phase', 'Braak Stage', 'CERAD NP Score', 'CERAD DP Score', 'NIA-Reagan']
    ADNC_dict = dict(zip(ADNC_keys, ADNC_values))
    i = 0
    for line in lines:
        if 'ALZHEIMER' in line:
            i += 1
            if i > 1:
                pass
            else:
                ADNC_index = lines.index(line)
                i = ADNC_index #This is the line number of where the 'ADNC' header appears in the title.
                a, b, c, d, e = lines[i+1], lines[i+2], lines[i+3], lines[i+4], lines[i+5] # This is just so the code below aligns with what we see in the NP report.

                ADNC_desig_index = lines.index(a)
                Thal_index = lines.index(b)
                Braak_index = lines.index(c)
                CERAD_index = lines.index(d)
                NIAReag_index = lines.index(e)

                ADNC_raw_data = [a, b, c, d, e]
                #ADNC_raw_data.append(lines[a:e])
                #ADNC_raw_data = ADNC_raw_data[0] # Doing this because the above makes it an array (list of a list)

    for line in ADNC_raw_data:
        line.upper()
    print(ADNC_raw_data)
## ADNC level
## Calculate it based on the formula


    ADNC_level = ADNC_raw_data[0]
    ADNC_level = ADNC_level.split()

    ADNC_level = [word.upper() for word in ADNC_level]

    if all(x in ADNC_level for x in ['NOT', 'LOW']): # Not to low
        ADNC_update_value = 4
    elif all(x in ADNC_level for x in ['INTERMEDIATE', 'LOW']): # Low to intermediate
        ADNC_update_value = 5
    elif all(x in ADNC_level for x in ['INTERMEDIATE', 'HIGH']): # Intermediate to high
        ADNC_update_value = 6
    elif re.search('NOT\s+ADNC', ' '.join(ADNC_level)):
        ADNC_update_value = 0
    elif ('LOW') in ADNC_level:
        ADNC_update_value = 1
    elif ('INTERMEDIATE') in ADNC_level:
        ADNC_update_value = 2
    elif ('HIGH') in ADNC_level:
        ADNC_update_value = 3


    elif any(x in ' '.join(ADNC_level) for x in ['NOT CONFORMING','NON-CONFORMING','NOT APPLICABLE']):
        ADNC_update_value = (-4)
    else:
        ADNC_update_value = 'Unknown, check NP report'

    ADNC_dict["ADNC level"] = ADNC_update_value

## Thal Phase

    Thal_data = ADNC_raw_data[1]

    for c in Thal_data.split():
        try:
            Thal_value = (int(c))
        except ValueError:
            pass

    ADNC_dict["Thal Phase"] = Thal_value

## Braak Stage

    Braak_data = ADNC_raw_data[2]

    for c in Braak_data.split():
        try:
            Braak_value = (int(c))
        except ValueError:
            Braak_value = int('-3')
            # i.e. 2748: Braak value could not be determined

    ADNC_dict["Braak Stage"] = Braak_value

## CERAD Scores

    CERAD_data = ADNC_raw_data[3].split(',') # The two scores are split by a comma in the NP report
    CERAD_values = ['NULL', 'NULL']

    CERAD_NP_data = CERAD_data[0].split()
    CERAD_DP_data = CERAD_data[1].split()
    CERAD_NP_value = int
    CERAD_DP_value = int

## Neuritic Plaques

    CERAD_NP_data = [word.upper() for word in CERAD_NP_data]

    if CERAD_NP_data[-1] == ('ABSENT'):
        CERAD_NP_value = 0
    elif CERAD_NP_data[-1] == ('SPARSE'):
        CERAD_NP_value = 1
    elif CERAD_NP_data[-1] == ('MODERATE'):
        CERAD_NP_value = 2
    elif CERAD_NP_data[-1]  == ('FREQUENT'):
        CERAD_NP_value = 3
    else:
        print('Atypical CERAD Neuritic Plaque value, check NP report')

    CERAD_values[0] = CERAD_NP_value
    #print(CERAD_NP_data[-1], CERAD_NP_value)

## Diffuse Plaques

    CERAD_DP_data = [word.upper() for word in CERAD_DP_data]

    if CERAD_DP_data[-1] == ('ABSENT'):
        CERAD_DP_value = 0
    elif CERAD_DP_data[-1] == ('SPARSE'):
        CERAD_DP_value = 1
    elif CERAD_DP_data[-1] == ('MODERATE'):
        CERAD_DP_value = 2
    elif CERAD_DP_data[-1]  == ('FREQUENT'):
        CERAD_DP_value = 3
    else:
        print('Atypical CERAD Neuritic Plaque value, check NP report')

    CERAD_values[1] = (CERAD_DP_value)
    #print(CERAD_DP_data[-1], CERAD_DP_value)

    ADNC_dict["CERAD NP Score"] = CERAD_NP_value
    ADNC_dict["CERAD DP Score"] = CERAD_DP_value

## Calculate A B C score

    ABC_score = {'A': '', 'B': '', 'C': ''}

    if Thal_value == 0:
        ABC_score['A'] = 0
    elif Thal_value in (1,2):
        ABC_score['A'] = 1
    elif Thal_value == 3:
        ABC_score['A'] = 2
    elif Thal_value in (4,5):
        ABC_score['A'] = 3

    if Braak_value in (1,2):
        ABC_score['B'] = 1
    elif Braak_value in (3,4):
        ABC_score['B'] = 2
    elif Braak_value in (5,6):
        ABC_score['B'] = 3

    if CERAD_NP_value in (0,1,2,3):
        ABC_score['C'] = CERAD_NP_value

    if ABC_score['A'] == 0 and ABC_score['C'] == 0:
        ADNC_dict['ADNC_level'] = 0
    elif ABC_score['A'] in (1,2,3) and ABC_score['B'] in (0,1):
        ADNC_dict['ADNC_level'] = 1
    elif ABC_score['A'] == 1 and ABC_score['C'] in (0,1):
        ADNC_dict['ADNC level'] = 1
    elif ABC_score['A'] == 1 and ABC_score['C'] in (2,3) and ABC_score['B'] in (2,3):
        ADNC_dict['ADNC_level'] = 2
    elif ABC_score['A'] == 2 and ABC_score['B'] in (2,3):
        ADNC_dict['ADNC_level'] = 2
    elif ABC_score['A'] == 3 and ABC_score['C'] in (0,1):
        ADNC_dict['ADNC_level'] = 2
    elif ABC_score['A'] == 3 and ABC_score['C'] in (2,3) and ABC_score['B'] == 2:
        ADNC_dict['ADNC_level'] = 2
    elif ABC_score['A'] == 3 and ABC_score['C'] in (2,3) and ABC_score['B'] == 3:
        ADNC_dict['ADNC_level'] = 3

## NIA-Reagan Criteria
## Convert to calculating NIA-Reagan


    NIAReag_data = ADNC_raw_data[4]
    NIAReag_data = NIAReag_data.split()
    NIAReag_data = [word.upper() for word in NIAReag_data]
    NIAReag_value = int

    if NIAReag_data[-1] in ('CONFORMING', 'NONCONFORMING', 'NON-CONFORMING', 'APPLICABLE'): # Come up with a better system for this
        NIAReag_value = (-4)
    elif NIAReag_data[-1] == ('MET'):
        NIAReag_value = 0
    elif NIAReag_data[-2] == 'LOW':
        NIAReag_value = 1
    elif NIAReag_data[-2] == 'INTERMEDIATE':
        NIAReag_value = 2
    elif NIAReag_data[-2] == 'HIGH':
        NIAReag_value = 3
    else:
        print('See NP report for NIA-Reagan Criteria.')

    ADNC_dict["NIA-Reagan"] = NIAReag_value

    print(ADNC_dict)

    return ADNC_dict

# Cerebral amyloid angiopathy

def get_CAA(dx_sxn):

    has_CAA = False
    CAA_value = 0
    CAA_data = []

    for item in dx_sxn:
        if 'CEREBRAL AMYLOID ANGIOPATHY' in item:
            has_CAA = True
            CAA_data.append(item)
            CAA_line = item.split(',')

            if 'MILD' in CAA_line[1]:
                CAA_value = 1
            elif 'MODERATE' in CAA_line[1]:
                CAA_value = 2
            elif 'SEVERE' in CAA_line[1]:
                CAA_value = 3

    #print(dx_line_sxn)
    print(f'CAA value: {CAA_value}')

    return CAA_value

def get_lbd_stage(dx_sxn):

    lines = dx_sxn[1:-1]
    has_lbd = False
    LBDStage_value = int
    possible_lbd_values = []

    for line in lines:
        if 'LEWY' in line:
            has_lbd = True

            if 'BRAINSTEM' in line:
                LBDStage_value = 1
            elif 'LIMBIC' in line:
                LBDStage_value = 2
            elif 'NEOCORTICAL' in line:
                LBDStage_value = 3
            elif 'AMYGDALA' in line:
                LBDStage_value = 4
            elif 'SUBSTANTIA NIGRA' in line:
                LBDStage_value = 5
            else:
                LBDStage_value = 'atypical'
                print('Atypical LBD stage value, check NP report')


    if has_lbd == False:
        LBDStage_value = 0

    print(f'LBD Stage: {LBDStage_value}')
    return LBDStage_value

def get_PDBraak(dx_sxn):

    PDBraak_value = 0
    roman_nums = dict(zip(['I','II', 'III', 'IV', 'V', 'VI'],[1, 2, 3, 4, 5, 6]))

    for line in dx_sxn:
        if 'LEWY' in line:
            if 'BRAAK' not in line:
                PDBraak_value = (-4)
        if all(x in line for x in ['LEWY', 'BRAAK']):
            lbd_index = dx_sxn.index(line)
            next_line = dx_sxn[lbd_index + 1]

            lbd_dx = line.split()
            if re.search('NON-*\s*CONFORMING', lbd_dx[-1]):
                    PDBraak_value = (-4)
            #elif isinstance((lbd_dx[-1]), int) == True:
            else:
               try:
                   PDBraak_value = int(lbd_dx[-1])
                   if PDBraak_value > 6:
                       num_tooBig = lbd_dx[-1]  ## This is for when the parser includes the citation superscript
                       PDBraak_value = int(num_tooBig[0])
               except ValueError:
                   if lbd_dx[-1] in roman_nums.keys():
                       PDBraak_value = roman_nums[(lbd_dx[-1])]
                   elif bool(re.search('\s?STAGE', next_line)):
                       PDBraak_value = int(next_line[-1])
                   elif bool(re.search('\s?NON-*\s*[CONFORMING|APPLICABLE]', next_line)):
                       PDBraak_value = (-4)
                   else:
                        PDBraak_value = 'atypical'
                        print('Atypical PD Braak value, check NP report')

    print(f'PD Braak Stage: {PDBraak_value}')
    return PDBraak_value

def get_ATAC(dx_sxn):

    ATAC_value = 0

    for line in dx_sxn:
        if 'ATAC' in line:
            ATAC_value = 1
        elif 'ARGYROPHILIC THORNY' in line:
            ATAC_value = 1

    print(f'ATAC_value: {ATAC_value}')
    return ATAC_value

## Working on CTE -- varies between cases. ASK BILL about what value we should put when no stage is given (ex. P2502)

def get_CTE(dx_sxn):

    CTE_value = 0
    CTE_stages = {'STAGE 0': 0, 'STAGE I': 1, 'STAGE II': 2, 'STAGE III': 3, 'STAGE IV': 4, 'STAGE V': 5, 'STAGE VI': 6, 'NOT CONFORMING': -4, 'NON-CONFORMING': -4, 'NONCONFORMING': -4, 'NOT APPLICABLE': -4}
    #CTE_stages_nums =
    for line in dx_sxn:
        if 'CHRONIC TRAUMATIC' in line:
            if 'STAGE' in line:
                for stage in CTE_stages.keys():
                    if stage in line:
                        CTE_value = CTE_stages[stage]
# Even though this will always initially change CTE_value to 1, it will update if needed as it goes down the list
            else:
                CTE_value = -7
# Check with Bill about this
    print(f'CTE value: {CTE_value}')
    return CTE_value


# Hippocampal sclerosis, in progress (need win32com)
"""
def get_HS():

    HS_value = 0
    HS_stages = [0,1,2,3,4,5]
    HS_laterality = 0

    for line in get_dx_sxn():
        if 'HIPPOCAMPAL SCLEROSIS' in line:
            has_HS = True

    if has_HS:
"""

## ARTERIOLOSCLEROSIS
def get_arterio_sxn(filename):
    lines = open_file(filename)
    arteriolo_sxn = ''
    for line in lines:
        line = line.upper()
        if 'ARTERIOLOSCLEROSIS:' in line:
            arteriolo_sxn = line
    return arteriolo_sxn

def get_arterio(dx_sxn):
# 5/29/19 fixed so that 'moderate to severe --> severe

# Arteriolosclerosis value defaults to 0 unless mentioned
    arterio_value = 0

# Review Dx line items for arteriolosclerosis
    for line in dx_sxn:
        #if re.match(re.compile('ARTERIOLOSCLEROSIS\b?:'), line):
        if 'ARTERIOLOSCLEROSIS' in line:
            if 'MILD' in line:
                arterio_value = 1
            if 'MODERATE' in line:
                arterio_value = 2
            if 'SEVERE' in line:
                arterio_value = 3
            else:
                print('Atypical arteriolosclerosis value in Dx section, check H&E section of NP report')


    print(f'Arteriolosclerosis value: {arterio_value}')
    return arterio_value


## ATHEROSCLEROSIS

def get_athero(grossObs):

    athero_value = 0

    for line in grossObs:
        if re.match(re.compile('ATHEROSCLEROSIS\b?:'), line):
            if 'ABSENT' in line:
                athero_value = 0
            elif 'MILD' in line:
                athero_value = 1
            elif 'MODERATE' in line:
                athero_value = 2
                print(line)
            elif 'SEVERE' in line:
                athero_value = 3
            else:
                print('Atypical atherosclerosis value, check NP report')

    print(f'Atherosclerosis value: {athero_value}')
    return athero_value

## ARGYROPHILIC GRAIN DISEASE (AGD)

def get_AGD(dx_sxn):

    AGD_value = 0

    AGD_stages = {'LIMBIC': 1, 'DIFFUSE': 2}

    for line in dx_sxn:
        if 'GRAIN DISEASE' in line:
            for stage in AGD_stages.keys():
                if stage in line:
                    AGD_value = AGD_stages[stage]

    print(f'AGD value: {AGD_value}')
    return AGD_value

def get_huntington(dx_sxn):

    HD_value = 0
    pat = r'GRADE [0-4]'

    for line in dx_sxn:
        if 'HUNTINGTON' in line:
            grade = re.findall(pat, line)
            #print(f'\nHD grade: {grade}\n')
            if grade:
                grade = grade[0]
                HD_value = int(grade[-1:])
            else:
                print('Atypical HD value, check NP report')

    print(f'Huntington value: {HD_value}')
    return HD_value

## WHITE MATTER RAREFACTION

def get_wmr():
    wmr_value = 'NA'
    return wmr_value

## MICROBLEEDS (IN PROGRESS)

def get_microbleed():
    old_cortex = 0
    old_wm = 0
    old_deepgray = 0
    old_reptilian = 0
    acute_cortex = 0
    acute_wm = 0
    acute_deepgray = 0
    acute_reptilian = 0

    for line in get_dx_sxn():
        if 'MICROBLEED' in line:
            print(line)
            usr_input = input('Select region and age: 1 = cortex, old')

def make_new_rows():
    file_dict = get_files()
    for pNum in file_dict.keys():
        print(f' --- {pNum} ---')
        filename = file_dict[pNum]
        all_contents = open_file(filename)
        dx_sxn = get_dx_sxn(filename)
        grossObs = get_grossObs(filename)
        ADNC_dict = get_ADNC(dx_sxn)
        Thal_phase = ADNC_dict['Thal Phase']
        AD_Braak = ADNC_dict['Braak Stage']
        AD_CERAD_NP = ADNC_dict['CERAD NP Score']
        AD_CERAD_DP = ADNC_dict['CERAD DP Score']
        NIAReag = ADNC_dict['NIA-Reagan']
        CAA = get_CAA(dx_sxn)
        ADNC_level = ADNC_dict['ADNC_level']
        LBD = get_lbd_stage(dx_sxn)
        PD_Braak = get_PDBraak(dx_sxn)
        ATAC = get_ATAC(dx_sxn)
        CTE = get_CTE(dx_sxn)
        HS = 'NA'
        HS_laterality = 'NA'
        Arterio = get_arterio(dx_sxn)
        Athero = get_athero(grossObs)
        TDP_proteinopathy = 'NA'
        AGD = get_AGD(dx_sxn)
        HD = get_huntington(dx_sxn)

        field_names = 'pNum, Thal phase, AD Braak, AD CERAD NP, AD CERAD DP, NIA-Reagan, CAA, ADNC_level, LBD, PD_Braak, ATAC, CTE, HS, HS Laterality, Aterio, Athero, TDP Proteinopathy, AGD, HD'
        new_row_fields = field_names.split(',')
        new_row_values = [pNum, Thal_phase, AD_Braak, AD_CERAD_NP, AD_CERAD_DP, NIAReag, CAA, ADNC_level, LBD, PD_Braak, ATAC, CTE, HS, HS_laterality, Arterio, Athero, TDP_proteinopathy, AGD, HD]

        primDx_list = get_PrimDx(dx_sxn)
        print(f'\nNumber of primary diagnoses: {len(primDx_list)}')
        for counter, dx in enumerate(primDx_list, 1):
            print(f'{counter}: {dx}')

        contributingDx_list = get_ContributingDx(dx_sxn)
        print(f'\nNumber of contributing diagnoses: {len(contributingDx_list)}')
        for counter, dx in enumerate(contributingDx_list, 1):
            print(f'{counter}: {dx}')

        incidentalDx_list = get_IncidentalDx(dx_sxn)
        print(f'\nNumber of incidental diagnoses: {len(incidentalDx_list)}')
        for counter, dx in enumerate(incidentalDx_list, 1):
            print(f'{counter}: {dx}')


        print(dict(zip(new_row_fields, new_row_values)))

        with open('R:/groups/seeley/Mack/NP report parser/test.csv', 'a', encoding = 'utf-8') as fd:
            writer = csv.writer(fd)
            writer.writerow(new_row_values)

#if __name__ == '__main__':
#    make_new_rows()

#get_files()

# Get H&E table. From darrienwiens.wordpress.com
"""
def get_HE_table(pNum, filename):

    # Open Excel document
    XL = win32.DispatchEx('Excel.Application')
    XL.Visible = 0

    # load pre-made workbook
    XLBook = XL.Workbooks.Open('R:/groups/seeley/Mack/NP report parser/H&E tables.xlsx')

    # add worksheet with pNum name (from StackOverflow, 'Adding Excel Sheets to End of Workbook')
    newSheet = XLBook.Sheets.Add(Before=None, After=XLBook.Sheets(XLBook.Sheets.count))
    newSheet.Name = (f'{pNum}')

    # counter to keep track of Excel row
    XLrow = 2
    XLcol = 1
    HErow = 1
    HEcol = 1

    # open invisible Word app
    word = win32.DispatchEx('Word.Application')
    word.visible = 0

    # open Word document
    word.Documents.Open(filename)
    doc = word.ActiveDocument

    # access first table in Word doc (H&E table)
    HE_table = doc.Tables(1)

    # get (Word) and set (Excel) some data
    for row in HE_table.rows:
        for col in HE_table.columns:
            newSheet.Cells(row, col).Value = HE_table.Cell(Row=row, Column=col).Range.Text


    # close word doc
    doc.Close()

    # save and close Excel app
    XLBook.Close(True)
"""

test_files = get_Word_files()

# Pull values from word document

def get_HE_table_word(pNum, filename):

    document = Document(filename)
    table = document.tables[0]

    data = []

    for i, row in enumerate(table.rows):
        text = (str(cell.text) for cell in row.cells)

        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys,text))
        data.append(row_data)
    for item in data:
        print(item)
    #print(data)
    df = pandas.DataFrame(data)
    return df

def get_HE_table_sheet(pNum): # Add 'pNum' to argument
    HE_template = pandas.read_excel(HE_datasheet, sheet_name='Template')
    #print(HE_template)
    HE_Region_List = HE_template.Region.tolist()
    HE_Block_List = HE_template.Block.tolist()
    #print(HE_Region_List)
    #print(HE_Block_List)
    #print(list(HE_template))
    ## Get H&E data for a given p# from the slide review sheet
    worksheet = pandas.read_excel(SlideReviewSheet, sheet_name=pNum, header=3, usecols=[0,1,2,3,4,5,6,7])
    #worksheet = pandas.DataFrame(init_worksheet.values[2:])
    try:
        worksheet_regions = dict(zip(list(worksheet['Block'] + ' ' + worksheet['Region']), list(worksheet['Neuronal loss'])))
    except KeyError:
        try:
            worksheet_regions = dict(zip(list(worksheet['Block'] + ' ' + worksheet['Region']), list(worksheet['Neuron loss'])))
        except KeyError:
            print(f'Key error for {pNum}\n')
        # will need way to valudate this
    #print(worksheet_regions)
    #print(worksheet)
    try:
        return worksheet_regions
    except UnboundLocalError:
        print(f'No slide review sheet for {pNum}')

working_files = get_files()

"""
def HS_dx_vs_HE():

    pNum_dict = {'pNum': [], 'Hippocampal sclerosis mentioned?': [], 'C12 Region': [], 'Neuronal loss': []}
    pNum_list = []
    HS_mentioned_list = []
    C12_Region_list = []
    Neuronal_loss_list = []
    #NP_Report = []

    C12_pat = 'C12\s.*'

    i = 0
    for pNum in working_files.keys():

        filename = working_files[pNum]

        pNum_int = int(pNum[1:])
        if pNum_int in range(2302, 2330):
            pNum = pNum + '.10'
        if pNum_int in range(2330, 2391):
            pNum = pNum + '.11'
        if pNum_int in range(2391, 2447):
            pNum = pNum + '.12'
        if pNum_int in range(2447, 2468):
            pNum = pNum + '.13'
        print(f'Working on {pNum}...')
        i += 1
        print(f'Percent finished: {100*i/434}')
        ## Add pNum to dict
        #pNum_list.append(pNum)


        lines = open_file(filename)
        lines = [item.upper() for item in lines]

        has_HS = False
        for item in lines:
            if 'HIPPOCAMPAL SCLEROSIS' in item:
                if 'HIPPOCAMPAL SCLEROSIS WAS NOT IDENTIFIED' not in item:
                    #print(f'{pNum}: {item}')
                    has_HS = True

        # Add HS value to pNum_dict
        #HS_mentioned_list.append(has_HS)


        try:
            NL_dict = get_HE_table_sheet(pNum)
            if not NL_dict:
                continue

            else:
                for block, severity in NL_dict.items():
                    if isinstance(block, str):
                        if re.search(C12_pat, block):

                            #Add C12 region to pNum_dict
                            C12_Region_list.append(block)
                            Neuronal_loss_list.append(severity)
                            pNum_list.append(pNum)
                            HS_mentioned_list.append(has_HS)
                            #NP_Report.append(filename)
                            #if severity == 3.0:
                                #print(f'{pNum} -- {block}: {severity}')
                            #else:
                            #    continue
                    else:
                        continue
        except ValueError:
            print(f'Value error: {pNum} is not in list')
        #    continue
        except xlrd.biffh.XLRDError:
            print(f'XLRDError: No sheet named <"{pNum}">')
            continue
    pNum_dict['pNum'] = pNum_list
    pNum_dict['Hippocampal sclerosis mentioned?'] = HS_mentioned_list
    pNum_dict['C12 Region'] = C12_Region_list
    pNum_dict['Neuronal loss'] = Neuronal_loss_list
    for key, value in pNum_dict.items():
        print(key, value)
    df = pandas.DataFrame(data=pNum_dict)

    df.to_excel('R:/groups/seeley/Mack/NP report parser/HS Discrepancy Data.xlsx')
"""
"""
for pNum in working_files.keys():
    HS_dx = dict()
    HE_value = dict()

    has_HS = False

    filename = working_files[pNum]

    pNum_int = int(pNum[1:])
    if pNum_int in range(2302, 2330):
        pNum = pNum + '.10'
    if pNum_int in range(2330, 2391):
        pNum = pNum + '.11'
    if pNum_int in range(2391, 2447):
        pNum = pNum + '.12'
    if pNum_int in range(2447, 2468):
        pNum = pNum + '.13'
    print(f'Working on {pNum}...')

    lines = open_file(filename)
    lines = [item.upper() for item in lines]

    for item in lines:
        if 'HIPPOCAMPAL SCLEROSIS' in item:
            if 'HIPPOCAMPAL SCLEROSIS WAS NOT IDENTIFIED' not in item:
                print(f'{pNum}: {item}')
                has_HS = True
            else:
                continue

    try:
        NL_dict = get_HE_table_sheet(pNum)
        for severity, block in NL_dict.items():
            if block == 'C12':
                print(type(block))
                if severity == 3.0:
                    print(f'{pNum} -- {block}: {severity}')
                else:
                    continue
    except ValueError:
        print(f'Value error: {pNum} is not in list')
    #    continue
    except xlrd.biffh.XLRDError:
        print(f'XLRDError: No sheet named <"{pNum}">')
        continue


"""


"""
i = 0
pNum = 'P2307'
filename = test_files[pNum]
print(filename)
get_HE_table(pNum, filename)


get_HE_table_sheet('P2558')
"""

#HS_dx_vs_HE()

## Generate list of pNums with issues (false positive, false negative)