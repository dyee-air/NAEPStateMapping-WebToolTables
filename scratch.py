# %%
import openpyxl as xl
from openpyxl.cell.cell import Cell, MergedCell
from copy import copy

WB = xl.load_workbook('template_state.xlsx')
SHEETS = list(WB)
WS3 = SHEETS[2]
# %%
stname = 'TX'

for sht in WB:
    sht.title = sht.title.replace('_ST_', stname)

# %%


note_rows = list(WS1.rows)[13:]

rowinfo = []
for row in note_rows:
    cell = row[0]
    val = cell.value
    fnt = copy(cell.font)
    aln = copy(cell.alignment)
    ht = WS1.row_dimensions[cell.row].height
    rowinfo.append({'value': val, 'height': ht, 'font': fnt, 'alignment': aln})


# %%
startrow = 20
for i, note in enumerate(rowinfo[1::2]):
    rownum = startrow+i

    tgt_cell = WS1[f'A{rownum}']
    tgt_cell.value = note['value']
    tgt_cell.font = note['font']
    tgt_cell.alignment = note['alignment']
    WS1.row_dimensions[rownum].height = note['height']
    WS1.merge_cells(f'A{rownum}:I{rownum}')

# %%