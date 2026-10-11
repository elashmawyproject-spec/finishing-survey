# -*- coding: utf-8 -*-
import sys, json, base64, io
from openpyxl import load_workbook
import build_boq as B, patch_v2 as P
from build_boq import *
P.patch('orig.xlsx', '_tmp2.xlsx')
out = sys.argv[1]
wb = load_workbook('template.xlsx')
ls = wb[S_LS]
lists = {}
for nm, col, ucol in (('FloorMat','A','B'),('WallMat','C','D'),('LinMat','E','F'),('CeilMat','G','H'),('PlMat','I','J'),('OpenTypes','K',None),('Levels','N',None),('CeilCond','O',None),('RoomNames','P',None)):
    items = [ls[f'{col}{r}'].value for r in range(LS_R0, LS_RN + 1)]
    lists[nm] = {'col': col, 'ucol': ucol, 'items': [i for i in items if i is not None]}
layout = {
    'r0': R0, 'listR0': LS_R0, 'listRN': LS_RN,
    'sheets': {'PD': S_PD, 'SP': S_SP, 'FL': S_FL, 'WL': S_WL, 'OP': S_OP, 'LN': S_LN, 'CL': S_CL, 'PL': S_PL, 'LS': S_LS, 'SM': S_SM},
    'cap': {'SP': N_SP, 'FL': N_FL, 'WL': N_WL, 'OP': N_OP, 'LN': N_LN, 'CL': N_CL, 'PL': N_PL},
    'cols': {k: B.COLS[k] for k in ('SP', 'FL', 'WL', 'OP', 'LN', 'CL', 'PL')},
    'lists': lists,
    'pd': {'name': 'B3', 'owner': 'B4', 'location': 'B5', 'date': 'B6', 'surveyor': 'B7', 'method': 'B9', 'notes': 'B12'},
}
io.open(out + '/js/layout.js', 'w', encoding='utf-8').write('globalThis.BOQ_LAYOUT = ' + json.dumps(layout, ensure_ascii=False, indent=1) + ';\n')
b64 = base64.b64encode(open('template.xlsx', 'rb').read()).decode()
io.open(out + '/js/template.js', 'w', encoding='utf-8').write('globalThis.BOQ_TEMPLATE_B64 = "' + b64 + '";\n')
print('ok', len(b64))
