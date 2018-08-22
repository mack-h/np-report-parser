import re
import shutil
# import win32compat.client as win32

#from np_report_parser_modA import get_files, get_pNums, get_dup_pNums, remove_dup_pNums

#import np_report_parser_modA

import sqlite3


#fileList = remove_dup_pNums() # save for later


working_file = 'R:\groups\seeley\Mack/NP report parser/_NP_Report_P2679_final.txt'
#table_file = 'R:\groups\seeley\Mack/NP report parser/NP_Report_P2722_final.docx'


## Copy files as plain text, store in NP report parser folder




# Okay. First, let's try and generate something resembling the Excel spreadsheet from the NP report.

# Okay actually, let's try and generate a SQLite table that reconfigures the data from the Microscopic findings.

"""
def create_connection(db_file):
    conn = sqlite3.connect(db_file)
    return conn

def create_table(conn, create_table_sql):
    c = conn.cursor()
    c.execute(create_table_sql)
"""

"""
pNum = np_report_parser_modA.get_pNums(working_file)
print(f'\n{pNum[0]}')
"""

#def get_PrimDx(pNum): ## Use this once you figure all of the above stuff out


## Note!! When converting files to .txt, be sure not to include line breaks.

def open_file():
    with open(working_file, 'rt') as f:

        lines = []

        # Adds all lines of a file (demarcated by \n) to a list. Each line will be stored as a string in the list.
        for line in f:
            line = line.strip()
            if not line:
                continue
            lines.append(line)
    return lines




def get_dx_sxn():

    lines = open_file()

    for line in lines:
        if 'Final Neuropathological Diagnoses' in line:
            i_1 = lines.index(line)
        if 'Clinical History' in line:
            i_2 = lines.index(line) - 1

    dx_line_sxn = lines[i_1:i_2+1]
    dx_line_sxn = [item.upper() for item in dx_line_sxn]


    return dx_line_sxn

#print(len(get_dx_sxn()))
#print(type(get_dx_sxn()))

## Testing get_dx_sxn
for line in get_dx_sxn():
    print(line)
print(get_dx_sxn())

## Gross obs section
def get_grossObs():
    lines = open_file()

    for line in lines:
        if 'Gross Observations' in line:
            i_1 = lines.index(line)
        if 'Other gross findings' in line:
            i_2 = lines.index(line) + 1

    grossObs = lines[i_1:i_2]
    grossObs = [item.upper() for item in grossObs]

    return grossObs

## Testing get_grossObs
for line in get_grossObs():
    print(line)

def get_PrimDx():
    lines = get_dx_sxn()

    primDx_list = []
    for line in lines:
        if 'PRIMARY DIAGNOSIS' in line:
            if ':' in line:
                dx = line[line.index(':')+2:]
                primDx_list.append(dx)
            elif '#' in line:
                dx = line[line.index('#') + 2:]
                primDx_list.append(dx)

    return primDx_list

primDx_list = get_PrimDx()

print(f'\nNumber of primary diagnoses: {len(primDx_list)}')
for counter, dx in enumerate(primDx_list, 1):
    print(f'{counter}: {dx}')

def get_ContributingDx():

    lines = get_dx_sxn()

    contributingDx_list = []
    for line in lines:
        if 'CONTRIBUTING DIAGNOSIS' in line:
            if ':' in line:
                dx = line[line.index(':')+2:]
                contributingDx_list.append(dx)

            elif '#' in line:
                dx = line[line.index('#') + 2:]
                contributingDx_list.append(dx)

    return contributingDx_list

contributingDx_list = get_ContributingDx()

print(f'\nNumber of contributing diagnoses: {len(contributingDx_list)}')
for counter, dx in enumerate(contributingDx_list, 1):
    print(f'{counter}: {dx}')

def get_IncidentalDx():

    lines = get_dx_sxn()

    IncidentalDx_list = []

    for line in lines:
        if 'INCIDENTAL DIAGNOSIS' in line:
            if ':' in line:
                dx = line[line.index(':')+2:]
                IncidentalDx_list.append(dx)
            elif '#' in line:
                dx = line[line.index('#') + 3:]
# This is to account for typos where there is no colon following a number.
                IncidentalDx_list.append(dx)

    return IncidentalDx_list

incidentalDx_list = get_IncidentalDx()

print(f'\nNumber of incidental diagnoses: {len(incidentalDx_list)}')
for counter, dx in enumerate(incidentalDx_list, 1):
    print(f'{counter}: {dx}')

def get_ADNC():

    lines = get_dx_sxn()

    ADNC_raw_data = []
    ADNC_values = []
    ADNC_keys = ['ADNC level', 'Thal Phase', 'Braak Stage', 'CERAD NP Score', 'CERAD DP Score', 'NIA-Reagan']
    ADNC_dict = dict(zip(ADNC_keys, ADNC_values))
    for line in lines:
        if 'ALZHEIMER' in line:

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


## ADNC level

    ADNC_level = ADNC_raw_data[0]
    ADNC_level = ADNC_level.split()

    ADNC_level = [word.upper() for word in ADNC_level]

    if all(x in ADNC_level for x in ['NOT', 'LOW']): # Not to low
        ADNC_update_value = 4
    elif all(x in ADNC_level for x in ['INTERMEDIATE', 'LOW']): # Low to intermediate
        ADNC_update_value = 5
    elif all(x in ADNC_level for x in ['INTERMEDIATE', 'HIGH']): # Intermediate to high
        ADNC_update_value = 6
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
            pass

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


## NIA-Reagan Criteria

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

def get_CAA():

    lines = open_file()

    dx_line_sxn = get_dx_sxn()

    has_CAA = False
    CAA_value = 0
    CAA_data = []

    for item in dx_line_sxn:
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

def get_lbd_stage():

    lines = get_dx_sxn()[1:-1]
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

def get_PDBraak():

    lines = get_dx_sxn()
    PDBraak_value = 0

    for line in lines:

        if all(x in line for x in ['LEWY', 'BRAAK']):
            lbd_dx = line.split()

            if isinstance(lbd_dx[-1], int) == True:
                PDBraak_value = lbd_dx[-1]
            elif lbd_dx[-1] == 'NON-CONFORMING':
                    PDBraak_value = (-4)
            else:
                PDBraak_value = 'atypical'
                print('Atypical PD Braak value, check NP report')

    print(f'PD Braak Stage: {PDBraak_value}')
    return PDBraak_value

def get_ATAC():

    ATAC_value = 0

    for line in get_dx_sxn():
        if 'ARTAG' in line:
            ATAC_value = 1

    print(f'ATAC_value: {ATAC_value}')
    return ATAC_value

## Working on CTE -- varies between cases. ASK BILL about what value we should put when no stage is given (ex. P2502)

def get_CTE():

    CTE_value = 0
    CTE_stages = {'STAGE 0': 0, 'STAGE I': 1, 'STAGE II': 2, 'STAGE III': 3, 'STAGE IV': 4, 'STAGE V': 5, 'STAGE VI': 6, 'NOT CONFORMING': -4, 'NON-CONFORMING': -4, 'NONCONFORMING': -4, 'NOT APPLICABLE': -4}

    for line in get_dx_sxn():
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

def get_arterio():

    arterio_value = 0

    for line in get_dx_sxn():
        if 'ARTERIOLOSCLEROSIS' in line:
            if 'MILD' in line:
                arterio_value = 1
            elif 'MODERATE' in line:
                arterio_value = 2
            elif 'SEVERE' in line:
                arterio_value = 3
            else:
                print('Atypical arteriolosclerosis value, check NP report')

    print(f'Arteriolosclerosis value: {arterio_value}')
    return arterio_value

def get_athero():

    athero_value = 0

    for line in get_grossObs():
        if 'ATHEROSCLEROSIS' in line:
            if 'ABSENT' in line:
                athero_value = 0
            elif 'MILD' in line:
                athero_value = 1
            elif 'MODERATE' in line:
                athero_value = 2
            elif 'SEVERE' in line:
                athero_value = 3
            else:
                print('Atypical atherosclerosis value, check NP report')

    print(f'Atherosclerosis value: {athero_value}')
    return athero_value

def get_AGD():

    AGD_value = 0

    AGD_stages = {'LIMBIC': 1, 'DIFFUSE': 2}

    for line in get_dx_sxn():
        if 'GRAIN DISEASE' in line:
            for stage in AGD_stages.keys():
                if stage in line:
                    AGD_value = AGD_stages[stage]

    print(f'AGD value: {AGD_value}')
    return AGD_value

def get_huntington():

    HD_value = 0
    pat = r'GRADE [0-4]'

    for line in get_dx_sxn():
        if 'HUNTINGTON' in line:
            grade = re.findall(pat, line)
            #print(f'\nHD grade: {grade}\n')
            if grade:
                grade = grade[0]
                HD_value = grade[-1:]
            else:
                print('Atypical HD value, check NP report')

    print(f'Huntington value: {HD_value}')
    return HD_value

print('\n')
get_ADNC()
get_CAA()
get_lbd_stage()
get_PDBraak()
get_ATAC()
get_CTE()
get_arterio()
get_athero()
get_AGD()
get_huntington()