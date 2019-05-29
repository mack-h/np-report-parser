import parser_functions
import examiners
import csv
import xlsxwriter
import time
import sys
import re

date_today = time.strftime('%Y.%m.%d')

pNum_list = []

section_names = ['']

def pNum_input():
    with open('R:/groups/seeley/Mack/NP report parser/working_pNums.csv', 'rt') as pNum_file:
        pNums = csv.reader(pNum_file)
        for pNum in pNums:
            pNum_list.append(pNum[0]) #The [0] is added because pNums is a list of list, the way the code is written

# pNum list test
    print(f'pNum list:\n{pNum_list}\n')

# pNum files test
pNum_files = parser_functions.get_files()
print(f'pNum files:\n{pNum_files}\n')

## Make dictionary with pNum + filepath

filepath_dict = {value: key for key, value in pNum_files.items()}

## Print the pNums and filenames (test)

for pNum in pNum_files:
    print(f'{pNum}: {pNum_files[pNum]}')

## Main loop

def get_all_dx():
    for f in pNum_files.values():
        pNum = filepath_dict[f]

        dx_sxn = parser_functions.get_dx_sxn(f)
        all_content = parser_functions.open_file(f)

        print(f'\n --- {pNum} --- \n')

        primDx_list = parser_functions.get_PrimDx(dx_sxn)
        print(f'Number of primary diagnoses: {len(primDx_list)}')
        for counter, dx in enumerate(primDx_list, 1):
            print(f'{counter}: {dx}')

        contributingDx_list = parser_functions.get_ContributingDx(dx_sxn)
        print(f'\nNumber of contributing diagnoses: {len(contributingDx_list)}')
        for counter, dx in enumerate(contributingDx_list, 1):
            print(f'{counter}: {dx}')

        incidentalDx_list = parser_functions.get_IncidentalDx(dx_sxn)
        print(f'\nNumber of incidental diagnoses: {len(incidentalDx_list)}')
        for counter, dx in enumerate(incidentalDx_list, 1):
            print(f'{counter}: {dx}')

def examine_CTE():
    file_list = parser_functions.get_files()
    for pNum, filename in file_list.items():
        lines = parser_functions.open_file(filename)
        for line in lines:
            if any(x in line for x in ('hronic traumatic', 'CTE')):
                print(f'{pNum}: {line}')

def examine_lbd():
    file_list = parser_functions.get_files()
    lbd_specs = {}

    for pNum, filename in file_list.items():
        lines = parser_functions.get_dx_sxn(filename)
        for line in lines:
            if 'LEWY' in line:
                lewy_index = lines.index(line)
                line = line.split(':')[1]
                line = line.strip()
                if 'DEMENTIA WITH' in line:
                    print(f'{pNum}: {line}')
                if line in lbd_specs:
                    continue
                else:
                    lbd_specs[line] = pNum
    for item in lbd_specs.items():
        print(item)




def all_dx_to_file():

    orig_stdout = sys.stdout
    newfile = open(f'R:/groups/seeley/Mack/NP report parser/NACC backlog dx sections_{date_today}.txt', 'w')
    sys.stdout = newfile

    for f in pNum_files.values():
        pNum = filepath_dict[f]

        dx_sxn = parser_functions.get_dx_sxn(f)
        all_content = parser_functions.open_file(f)

        print(f'\n --- {pNum} --- \n')

        primDx_list = parser_functions.get_PrimDx(dx_sxn)
        print(f'Number of primary diagnoses: {len(primDx_list)}')
        for counter, dx in enumerate(primDx_list, 1):
            print(f'{counter}: {dx}')

        contributingDx_list = parser_functions.get_ContributingDx(dx_sxn)
        print(f'\nNumber of contributing diagnoses: {len(contributingDx_list)}')
        for counter, dx in enumerate(contributingDx_list, 1):
            print(f'{counter}: {dx}')

        incidentalDx_list = parser_functions.get_IncidentalDx(dx_sxn)
        print(f'\nNumber of incidental diagnoses: {len(incidentalDx_list)}')
        for counter, dx in enumerate(incidentalDx_list, 1):
            print(f'{counter}: {dx}')

    sys.stdout = orig_stdout
    newfile.close()

def vbi_dx_to_file():
    current_time = time.strftime('%d-%m-%Y %H:%M:%S')

    orig_stdout = sys.stdout
    newfile = open(f'R:/groups/seeley/Mack/NP report parser/VBI sxns_{date_today}.txt', 'w')
    sys.stdout = newfile

    print(f'| Cases with VBI |\nRun from all files {current_time}\n\n')
    vbi_list = []

    for f in pNum_files.values():
        print(f'Working on {f}...')
        pNum = filepath_dict[f]

        dx_sxn = parser_functions.get_dx_sxn(f)
        all_content = parser_functions.open_file(f)

        hx_sxn = {}

        try:
            gross_obs = parser_functions.get_grossObs(f)
            gross_obs = [item.upper() for item in gross_obs]
        except TypeError:
            print(f'No gross obs found for {pNum}')

        for line in all_content:
            if '(H & E)' in line:
                if 'Other significant pathology' not in line:
                    hx_sxn[line] = all_content[all_content.index(line) + 2]
                else:
                    hx_sxn[line] = all_content[all_content.index(line) + 1]

        gross_infarcts = {}

        if gross_obs:
            for line in gross_obs:
                if any(x in line for x in ['TERRITORIAL INFARCTS', 'LACUNAR INFARCTS']):
                    if any(x in line for x in ['ABSENT', 'NONE']):
                        continue
                    else:
                        line = line.split(':')
                        gross_infarcts[line[0]] = line[1]

        vbi_lineItem = ''

        for line in dx_sxn:
            if 'VASCULAR BRAIN INJURY' in line:
                vbi_lineItem = line
                vbi_list.append(pNum)

        if vbi_lineItem != '':
            print(f'--- {pNum} ---\n')
            for key in hx_sxn:
                if 'Microinfarcts' in key:
                    print(f'{key}:\n{hx_sxn[key]}\n')
                if 'Other significant pathology' in key:
                    print(f'{key}:\n{hx_sxn[key]}\n')

            if gross_infarcts:
                for key, value in gross_infarcts.items():
                    print(f'{key}: {value}')

    print('P# List:')
    for pNum in vbi_list:
        print(f'{pNum}')
    print(len(vbi_list))

    sys.stdout = orig_stdout
    newfile.close()

def weird_vbi_to_file():
        current_time = time.strftime('%d-%m-%Y %H:%M:%S')

        orig_stdout = sys.stdout
        newfile = open(f'R:/groups/seeley/Mack/NP report parser/Abnormal VBI sxns_{date_today}.txt', 'w')
        sys.stdout = newfile

        print(f'| Cases with Abnormal VBI |\nRun from working_pNums file on {current_time}\n\n')
        vbi_list = []
        gross_infarct_list = []

        check_grossObs_list = []
        check_dx_list = []


        for f in pNum_files.values():
            pNum = filepath_dict[f]

            try:
                dx_sxn = parser_functions.get_dx_sxn(f)
                all_content = parser_functions.open_file(f)
            except UnboundLocalError:
                print(f'{pNum}: Diagnostic sxn index referenced before assignment')

            hx_sxn = {}
            try:
                gross_obs = parser_functions.get_grossObs(f)
                gross_obs = [item.upper() for item in gross_obs]
            except TypeError:
                print(f'{pNum}: No gross obs')

            for line in all_content:
                if '(H & E)' in line:
                    if 'Other significant pathology' not in line:
                        hx_sxn[line] = all_content[all_content.index(line) + 2]
                    else:
                        hx_sxn[line] = all_content[all_content.index(line) + 1]


            gross_infarcts = {}

            if gross_obs:
                for line in gross_obs:
                    if any(x in line for x in ['TERRITORIAL INFARCTS', 'LACUNAR INFARCTS']):
                        if any(x in line for x in ['ABSENT', 'NONE']):
                            continue
                        else:
                            line = line.split(':')
                            gross_infarcts[line[0]] = line[1]

            if gross_infarcts:
                vbi_list.append(pNum)
                gross_infarct_list.append(pNum)

                print(f'--- {pNum} ---\n')
                for key, value in gross_infarcts.items():
                    print(f'{key}: {value}\n')

            vbi_lineItem = ''

            for line in dx_sxn:
                if 'VASCULAR BRAIN INJURY' in line:
                    vbi_lineItem = line

            if vbi_lineItem != '':

                for key in hx_sxn:
                    if 'Microinfarcts' in key:
                        if all(x not in hx_sxn[key].upper() for x in ['SCARCE', 'MODERATE', 'SEVERE']):
                            #print(f'{key}:\n{hx_sxn[key]}\n')
                            vbi_list.append(pNum)

        vbi_list = sorted(set(vbi_list))
        gross_infarct_list = sorted(set(gross_infarct_list))

        print('Weird VBI P# List:')
        for pNum in vbi_list:
            print(f'{pNum}')
        print(len(vbi_list))

        print('Gross Infarct List:')
        for pNum in gross_infarct_list:
            print(f'{pNum}')
        print(len(gross_infarct_list))

        sys.stdout = orig_stdout
        newfile.close()



def examine_grossObs():

    cases_without_grossObs = []
    for pNum in pNum_files.keys():
        filename = pNum_files[pNum]
        all_content = parser_functions.open_file(filename)

        i = 0
        for line in all_content:
            line = line.upper()
            if 'GROSS OBSERVATIONS' in line:
                i += 1
        if i == 0:
            cases_without_grossObs.append(pNum)

    for pNum in cases_without_grossObs:
        print(pNum)


def examine_vbi():
    pNum_list = []

    for f in pNum_files.values():
        print(f'Working on {f}...')
        pNum = filepath_dict[f]

        dx_sxn = parser_functions.get_dx_sxn(f)
        all_content = parser_functions.open_file(f)

        for line in all_content:
            line = line.upper()
            ## Note: needs to account for 'MICROINFARCTION(S), AGE, WAS/WERE FOUND'
            if 'MICROINFARCTION WAS FOUND' in line:
                pNum_list.append(pNum)
            if 'MICROINFARCT WAS FOUND' in line:
                pNum_list.append(pNum)
            if 'MICROINFARCTS WERE FOUND' in line:
                pNum_list.append(pNum)
            if 'MICROINFARCTIONS WERE FOUND' in line:
                pNum_list.append(pNum)

    pNum_list = sorted(set(pNum_list))
    return pNum_list

microinfarct_phrases = ['MICROINFARCTION WAS FOUND', 'MICROINFARCT WAS FOUND', 'MICROINFARCTS WERE FOUND']

def get_microinfarct_section(lines):
    start0 = 0
    end0 = 0
    lines = (line.upper() for line in lines)
    lines = list(lines)
    for line in lines:
        if 'MICROINFARCTS (H & E)' in line:
            start0 = lines.index(line)
        for line in lines[start0:]:
            if any(x in line for x in ['OTHER SIGNIFICANT PATHOLOGY', 'OTHER FINDINGS', 'LEWY']):
                end0 = lines.index(line)
                break
            elif any(x in line for x in ['IMMUNOHISTOCHEMISTRY']):
                end0 = lines.index(line)
                break
    if start0 == 0:
        print('No start!')
    if end0 == 0:
        print('No end!')
    #print(lines[start0:end0])

    microinfarct_sxn = lines[start0:end0]
    ''.join(microinfarct_sxn)
    return microinfarct_sxn

def count_occurrences(word, sentence):
    return sentence.split().count(word)

def examine_vbi_regions():

    all_vbi_terms = []
    distinct_sentences_list = []

    for pNum in examine_vbi():
        dict_entry = parser_functions.get_files([pNum])
        filename = dict_entry[pNum]
        print(f'Filename: {filename}')
        lines = parser_functions.open_file(filename)

        lines = get_microinfarct_section(lines)
        print(lines)

        for line in lines:
            line = line.upper()
            if any(x in line for x in microinfarct_phrases):
                print(line)
                distinct_sentences = line.split('.')
                for sentence in distinct_sentences:
                    distinct_sentences_list.append(sentence)
                    print(f'Sentence: {sentence}')
                    distinct_clauses = sentence.split('AND')
                    for clause in distinct_clauses:
                        print(f'Clause: {clause}')
                        distinct_terms = clause.split(',')
                        for term in distinct_terms:
                            print(f'Term: {term}')
                            all_vbi_terms.append(term)

    clean_vbi_terms = []
    before_vbi_terms = []
    gray_matter_terms = []
    white_matter_terms = []
    before_region_terms = []
    after_region_terms = []

    for term in all_vbi_terms:
        term = term.strip()
        if any(x in term for x in ['GRAY', 'CORTEX OF', '(CORTEX)', '(GRAY)', '(GRAY']):
            gray_matter_terms.append(term)
        if any(x in term for x in ['SUBCORTICAL', '(SUBCORTICAL', '(SUBCORTICAL)''WHITE MATTER', 'MATTER)']):
            white_matter_terms.append(term)
        if 'NO MICROINFARCTION' in term:
            continue
        elif r'FINDINGS:\t' in term:
            terms = term.split(r'FINDINGS:\t')
            if 'FOUND IN' in terms[1]:
                terms = term[1].split('FOUND IN')
                before_region_terms.append(terms[0])
                for article in ['THE ', 'A ']:
                    if article in terms[1]:
                        terms = terms[1].split(article)
                clean_vbi_terms.append(terms[1].strip())
            else:
                clean_vbi_terms.append(terms[1].strip())
        elif 'FOUND IN' in term:
            terms = term.split('FOUND IN')
            before_vbi_terms.append(terms[0])
            has_article = 0
            for article in ['THE ', 'A ']:
                if article in terms[1]:
                    has_article = 1
                    terms = terms[1].split(article)
                    clean_vbi_terms.append(term.strip())
            if has_article == 0:
                clean_vbi_terms.append(terms[1].strip())
        elif term == '':
            continue
        else:
            clean_vbi_terms.append(term.strip())
    clean_vbi_terms = set(clean_vbi_terms)
    before_vbi_terms = set(before_vbi_terms)
    gray_matter_terms = sorted(set(gray_matter_terms))
    white_matter_terms = sorted(set(white_matter_terms))
    #for term in sorted(clean_vbi_terms):
    #    print(term)
    #print(len(clean_vbi_terms))
    #for term in sorted(before_vbi_terms):
    #    print(term)
    #print(len(before_vbi_terms))
    print(f'Gray matter terms:')
    for term in gray_matter_terms:
        print(term)
    print(f'White matter terms:')
    for term in white_matter_terms:
        print(term)
    for sentence in set(distinct_sentences_list):
        print(sentence)
    print(len(set(distinct_sentences_list)))


def percent_authorship():
    files = parser_functions.get_files()
    pNums = files.keys()
    author_dict = {}
    for pNum in pNums:
        try:
            author_dict[pNum] = parser_functions.get_author(pNum)
        except IndexError:
            continue
    bill_count = 0
    lea_count = 0
    salvo_count = 0
    total_count = 0
    for author in author_dict.values():
        if 'Seeley' in author:
            bill_count += 1
        if 'Grinberg' in author:
            lea_count += 1
        if "Spina" in author:
            salvo_count += 1
        total_count += 1
    bill_percent = (bill_count/total_count)*100
    salvo_percent = (salvo_count/total_count)*100
    lea_percent = (lea_count/total_count)*100
    print(f'Total cases: {total_count}\nBill {bill_count}, Salvo {salvo_count}, Lea {lea_count}\n{bill_percent}% Bill, {salvo_percent}% Salvo, {lea_percent}% Lea\n')


"""
        
    print('P# List:')
    for pNum in vbi_list:
        print(f'{pNum}')
    print(len(vbi_list))

    sys.stdout = orig_stdout
    newfile.close()

def examine_diagnoses():
    pat1 = r'(?<=[PRIMARY|CONTRIBUTING|INCIDENTAL] DIAGNOSIS #\d:).*'
    pat2 = r'(?<=[PRIMARY|CONTRIBUTING|INCIDENTAL] DIAGNOSIS:).*'
    pat3 = r'(?<=[PRIMARY|CONTRIBUTING|INCIDENTAL]#.:).*'
    pat4 = r'(?<=\d\.).*'
    pat5 = r'[ABCDEF]\. '

    patterns_list = [pat1, pat2, pat3, pat4]

    #diagnosis_patterns_list = []
    diagnosis_list = []
    dx_pNum_dict  = {}

    for pNum in pNum_files.keys():
        filename = pNum_files[pNum]
        dx_sxn = parser_functions.get_dx_sxn(filename)

        for item in dx_sxn:
            for pat in patterns_list:
                if re.compile(pat).search(item):

                    # Add diagnostic item to list of diagnostic items
                    return_item = re.search(pat, item)
                    return_item = return_item.group(0)
                    return_item = return_item.strip()
                    diagnosis_list.append(return_item)

                    # Associate pNumber with this particular diagnostic item
                    if return_item in dx_pNum_dict.keys():
                        dx_pNum_dict[return_item] = (f'{dx_pNum_dict[return_item]}, {pNum}')
                    else:
                        dx_pNum_dict[return_item] = (f'{pNum}')

                else:
                    continue

    #return sorted(set(diagnosis_list))
    return dx_pNum_dict

    #    for item in dx_sxn:
     #       if pat.search(item):
      #          continue
       #     else:
        #        diagnosis_patterns_list.append(item)
    # return set(diagnosis_patterns_list)


def diagnosis_to_pNum():
    dx_pNum_dict = {}
    for dx in examine_diagnoses():
        dx_pNum_dict[dx] = ''
        for pNum in pNum_files.keys():
            filename = pNum_files[pNum]
            dx_sxn = parser_functions.get_dx_sxn(filename)
            for item in dx_sxn:
                if dx in item:
                    if dx_pNum_dict[dx] != '':
                        dx_pNum_dict[dx] = (f'{dx_pNum_dict[dx]}, {pNum}')
                    else:
                        dx_pNum_dict[dx] = pNum
    return dx_pNum_dict
"""
"""
for item in diagnosis_to_pNum():
    print(item)
"""
"""
for (key, value) in sorted(examine_diagnoses().items()):
    print(f'{key}: {value}')

"""

#examine_vbi_regions()
#examine_CTE()
examine_lbd()