from collections import OrderedDict
from datetime import datetime

from pyexcel_ods import get_data
from pyexcel_ods import save_data


def iteratelist(datatype, databank):
    list_nm = []
    list_st = []

    div = False
    for rowInst in databank['Sheet1']:
        if len(rowInst) == 0:
            continue
        if rowInst[0] == '':
            div = True
            continue
        if div:
            list_nm.append(rowInst[0])
            if datatype == 'data':
                sub_hp = rowInst[1] + rowInst[2] + rowInst[5]
                sub_lp = rowInst[3] + rowInst[4] + rowInst[6] + rowInst[7]
                list_st.append((2 * sub_hp + sub_lp) / 1000)
            elif datatype == 'ielts':
                sub_ep = rowInst[1] + rowInst[2] + rowInst[3] + rowInst[4]
                list_st.append(sub_ep / 40)
            elif datatype == 'interview':
                sub_ep = rowInst[1] + rowInst[2] + rowInst[3] + rowInst[4] + rowInst[5]
                list_st.append(sub_ep / 50)
            else:
                print('!You made a mistake here!')

    return [list_nm, list_st]


# Lists Inits #
DataX_list = []
Ielts_list = []
Inter_list = []
###############

# read'n files #
for i in range(10):
    try:
        file_dat = get_data('Data' + str(i) + '.ods')
        DataX_list.append(iteratelist('data', file_dat))

        print('Data', i, '.ods file read.', sep='')
    except:
        pass

try:
    file_iel = get_data('IELTS.ods')
    Ielts_list.append(iteratelist('ielts', file_iel))

    print('IELTS.ods file read.')
except:
    print('IELTS.ods file not read.')

try:
    file_iel = get_data('Interview.ods')
    Inter_list.append(iteratelist('interview', file_iel))

    print('Interview.ods file read.')
except:
    print('Interview.ods file not read.')
###############

# buble sort'n #
# Sort of not required...
###############

# processing #
names_list = DataX_list[0][0]
subTotal = [0] * len(names_list)

for row in DataX_list:
    kay = row[1]
    for i in range(len(names_list)):
        subTotal[i] += row[1][i] / (len(DataX_list) + 2)

for i in range(len(names_list)):
    subTotal[i] += (Ielts_list[0][1][i] + Inter_list[0][1][i]) / (len(DataX_list) + 2)

for i in range(len(names_list)):
    subTotal[i] = round(subTotal[i], 5)
###############

# buble sort'n #
for i in range(len(names_list)):
    for j in range(len(names_list) - i - 1):
        if subTotal[j] < subTotal[j + 1]:
            names_list[j], names_list[j + 1] = names_list[j + 1], names_list[j]
            subTotal[j], subTotal[j + 1] = subTotal[j + 1], subTotal[j]
###############

# writ'n data #
row = []
for i in range(len(names_list)):
    row.append([names_list[i], subTotal[i]])

column = []
for cell in row:
    column.append(cell)

data = OrderedDict()
data.update({'Sheet 1': column})

name = 'result_' + str(datetime.now()) + '.ods'
name = name.replace(':', ',')
save_data(name, data)

print('Result saved as', name)
###############
