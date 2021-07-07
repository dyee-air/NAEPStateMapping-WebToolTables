# %%
from common import SnakeTableGenerator

# %%

DATAPATH = "U:/ESSIN Task 14/Mapping Report/2019/Standard Method/02_Mapping Tool Design/Data Source/SourceTable_States/SAS/Updated data/Long Files/Updated 0413"
DATAFILE = "long file.xlsx"
OUTPATH = 'output/snakechart'


# %%
TBL = SnakeTableGenerator(f'{DATAPATH}/{DATAFILE}')

# %%

for year in set(TBL.state_data['year']):
    for sg in 'R4', 'M4', 'R8', 'M8':
        TBL.generate(year, sg).save(
            f"{OUTPATH}/snake_chart_table_{year}{TBL.letters[sg]}.xlsx")

# %%
