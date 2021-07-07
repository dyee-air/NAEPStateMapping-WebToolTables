# %%
from common import StateTableGenerator
# %%


DATAPATH = "U:/ESSIN Task 14/Mapping Report/2019/Standard Method/02_Mapping Tool Design/Data Source/SourceTable_States/SAS/Updated data/Long Files/Updated 0413"
DATAFILE = "long file.xlsx"
OUTPATH = 'output/state'

# %%
TBL = StateTableGenerator(f'{DATAPATH}/{DATAFILE}')
# %%

for state in TBL.states:
    TBL.generate(state).save(f'{OUTPATH}/{state}.xlsx')

# %%
