import re

#from np_report_parser_modA import get_files, get_pNums, get_dup_pNums, remove_dup_pNums

#import np_report_parser_modA

import sqlite3

#fileList = remove_dup_pNums() # save for later


working_file = 'R:\groups\seeley\Mack/NP report parser/_NP_Report_P2817_final.txt'

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
        if '*Designations' in line:
            i_2 = lines.index(line)

    dx_line_sxn = lines[i_1:i_2+1]
    dx_line_sxn = [item.upper() for item in dx_line_sxn]


    return dx_line_sxn

#print(len(get_dx_sxn()))
#print(type(get_dx_sxn()))
for line in get_dx_sxn():
    print(line)
print(get_dx_sxn())

def get_PrimDx():
    lines = get_dx_sxn()

    primDx_list = []
    for line in lines:
        if 'PRIMARY DIAGNOSIS' in line:
            dx = line[line.index(':')+2:]
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
            dx = line[line.index(':')+2:]
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
            dx = line[line.index(':')+2:]
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


print('\n')
get_ADNC()
get_CAA()
get_lbd_stage()
get_PDBraak()

