# %%
import csv
import pandas as pd
import us
from numpy import nan

DATAPATH = "U:/ESSIN Task 14/Mapping Report/2019/Standard Method/02_Mapping Tool Design/Data Source/SourceTable_States/SAS/Updated data/Long Files/Updated 0413"
DATAFILE = "long file.xlsx"
# %%
dta = pd.read_excel(f'{DATAPATH}/{DATAFILE}')
years = set(dta.year)

# %%

dta.rename({'Consortia': 'consortium'}, axis=1, inplace=True)

# %% Add NECAP rows

for year in years:
    datasets = {sg: dta[(dta.subjgrade == sg) & (dta.year == year)]
                for sg in ('M4', 'M8', 'R4', 'R8')}
    for sg, results in datasets.items():
        if 'NECAP' in set(results.consortium):
            necap_row = results[results.consortium == 'NECAP'].iloc[0]
            necap_row.state = 'NECAP'
            necap_row.consortium = nan
            dta.append(necap_row)

# %% Replace mark
dta['nse_re_mark'] = dta['nse_re'].apply(lambda val: val >= 0.5)

# %% Exclusion indicator
dta['exclude'] = dta['IN_SNAKECHART_FILE'].apply(lambda s: s == 'NO')

# %% Consortium indicator
dta['is_consortium'] = dta['state'].isin(set(dta.consortium))

# %% State abbreviations and fips codes
st_abbr = us.states.mapping('name', 'abbr')
st_fips = us.states.mapping('name', 'fips')
dta['fips'] = dta['state'].apply(lambda st: st_fips.get(st, nan))
dta['state_abbr'] = dta['state'].apply(lambda st: st_abbr.get(st, nan))


# %% Achievement levels
NAEP_LABELS = ('Below NAEP Basic',
               'NAEP Basic',
               'NAEP Proficient',
               'NAEP Advanced')

NAEP_CUTS = {
    'M4': (214, 249, 282),
    'M8': (262, 299, 333),
    'R4': (208, 238, 268),
    'R8': (243, 281, 323)
}


def get_level(nse: float, nse_se: float, subjgrade: str):

    if subjgrade in NAEP_CUTS and pd.notna(nse):
        upper = nse + 1.96*nse_se
        for i, cut in enumerate(NAEP_CUTS[subjgrade]):
            if upper < cut:
                return NAEP_LABELS[i]

        return NAEP_LABELS[-1]

    return None

dta['level'] = dta[['nse', 'nse_se', 'subjgrade']].apply(lambda row: get_level(*row), axis=1)

# %% Export required columns
cols = ['year',
        'subjgrade',
        'fips',
        'state_abbr',
        'state',
        'consortium',
        'nse',
        'nse_se',
        'nse_re',
        'level',
        'nse_re_mark',
        'exclude',
        'is_consortium']

dta[cols].sort_values(cols).to_csv(
    'StateMappingResults.csv', columns=cols, index=False, quoting=csv.QUOTE_NONNUMERIC)
# %%
