# %%
import json
from itertools import product
from typing import Sequence, Union
import pandas as pd

DATAPATH = "U:/ESSIN Task 14/Mapping Report/2019/Standard Method/06_Analysis and Results/Darrick/py_webtables"
DATAFILE = "StateMappingResults.csv"

# %%
SUBJGRADES = ('M4', 'R4', 'M8', 'R8')


def get_cells(data: pd.DataFrame, yrs: Sequence[Union[int, str]]):

    cells = dict()
    for sg in SUBJGRADES:
        out = dict()
        for included in product((True, False), repeat=2):
            key = '_'.join(
                f'{"Included" if included[i] else "Excluded"}{yrs[i]}' for i in range(2))

            states = data[(data[f'{sg}_{yrs[0]}'] == included[0])
                          & (data[f'{sg}_{yrs[1]}'] == included[1])]['state_abbr']

            out[key] = ', '.join(states) + \
                f' ({len(states)} state{"s" if len(states) > 1 else ""})' if len(
                    states) else ''

        cells[sg] = out

    return cells


# %%
dta = pd.read_csv(f'{DATAPATH}/{DATAFILE}')

df: pd.DataFrame = dta[(dta.is_consortium == False)][[
    'fips', 'state_abbr', 'state']].drop_duplicates()
df['fips'] = df.fips.astype(int)

for year in 2017, 2019:
    for sg in SUBJGRADES:
        sgstates = set(dta[(dta.year == year) & (
            dta.subjgrade == sg) & (dta.exclude == False) & (pd.notna(dta.nse))]['state'])
        df[f'{sg}_{year}'] = df['state'].isin(sgstates)

# %%
YEARS = (2017, 2019)

df.to_csv('StateCrosstab_{0}_{1}_Data.csv'.format(*YEARS))

with open('StateCrosstab_{0}_{1}_Cells.json'.format(*YEARS), 'w') as outfile:
    outfile.write(json.dumps(get_cells(df, YEARS),
                             sort_keys=True, indent=4))

# %%
# Update - Alt vs Std
XLSFILE = 'StateCrosstab_Alt_Std.xlsx'

asdta = pd.read_excel(f'{DATAPATH}/{XLSFILE}')
# %%
YEARS = ('ALT', 'STD')
with open('StateCrosstab_{0}_{1}_Cells.json'.format(*YEARS), 'w') as outfile:
    outfile.write(json.dumps(get_cells(asdta, YEARS),
                             sort_keys=True, indent=4))
# %%
