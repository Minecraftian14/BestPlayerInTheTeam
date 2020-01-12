from collections import OrderedDict

import pyexcel_ods
from pyexcel_ods import save_data

#
#
#
data = pyexcel_ods.get_data("Data1.ods")

dat_names = []
dat_strength = []
driver = False
for row in data['Sheet1']:
    if len(row) == 0:
        continue
    if row[0] == '':
        driver = True
        continue
    if driver:
        dat_names.append(row[0])
        wt_MP = row[1] + row[2] + row[5]
        wt_RA = row[3] + row[4] + row[6] + row[7]
        dat_strength.append((2 * wt_MP + wt_RA) / 1000)
#
#
#
data = pyexcel_ods.get_data("IELTS.ods")

iel_names = []
iel_strength = []
driver = False
for row in data['Sheet1']:
    if len(row) == 0:
        continue
    if row[0] == '':
        driver = True
        continue
    if driver:
        iel_names.append(row[0])
        wt_AS = row[1] + row[2] + row[3] + row[4]
        iel_strength.append(wt_AS / 40)
#
#
#
data = pyexcel_ods.get_data("Interview.ods")

int_names = []
int_strength = []
driver = False
for row in data['Sheet1']:
    if len(row) == 0:
        continue
    if row[0] == '':
        driver = True
        continue
    if driver:
        int_names.append(row[0])
        wt_AT = row[1] + row[2] + row[3] + row[4] + row[5]
        int_strength.append(wt_AT / 50)
#
#
#
for i in range(len(dat_strength)):
    for j in range(len(dat_strength) - i - 1):
        if dat_names[j] > dat_names[j + 1]:
            dat_names[j], dat_names[j + 1] = dat_names[j + 1], dat_names[j]
            dat_strength[j], dat_strength[j + 1] = dat_strength[j + 1], dat_strength[j]
#
for i in range(len(iel_strength)):
    for j in range(len(iel_strength) - i - 1):
        if iel_names[j] > iel_names[j + 1]:
            iel_names[j], iel_names[j + 1] = iel_names[j + 1], iel_names[j]
            iel_strength[j], iel_strength[j + 1] = iel_strength[j + 1], iel_strength[j]
#
for i in range(len(int_strength)):
    for j in range(len(int_strength) - i - 1):
        if int_names[j] > int_names[j + 1]:
            int_names[j], int_names[j + 1] = int_names[j + 1], int_names[j]
            int_strength[j], int_strength[j + 1] = int_strength[j + 1], int_strength[j]
#
#
#
final = []
for i in range(len(dat_names)):
    final.append(round(0.4 * dat_strength[i] + 0.3 * iel_strength[i] + 0.3 * int_strength[i], 5))
# 
# 
# 
for i in range(len(dat_strength)):
    for j in range(len(dat_strength) - i - 1):
        if final[j] < final[j + 1]:
            dat_names[j], dat_names[j + 1] = dat_names[j + 1], dat_names[j]
            final[j], final[j + 1] = final[j + 1], final[j]
#
#
#

row = []
for i in range(len(dat_names)):
    row.append([dat_names[i], final[i]])

column = []
for i in range(len(dat_names)):
    column.append(row[i])

data = OrderedDict()
data.update({'Sheet 1': column})
save_data("result.ods", data)
