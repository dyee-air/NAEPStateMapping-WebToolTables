# %%
from typing import Optional, Union
import openpyxl as xl


class Row:  # pylint: disable=inherit-non-class
    """Simple class for NSE row display values
    """

    def __init__(self, year: int,
                 nse: Optional[float] = None,
                 nse_se: Optional[float] = None,
                 nse_re: Optional[float] = None) -> None:
        self.year = year
        self.nse: Union[float, str] = nse or '-'
        self.nse_se: Union[float, str] = nse_se or (
            '†' if nse is None else '-')
        self.nse_re: Union[float, str] = nse_re or (
            '†' if nse is None else '-')

    def astuple(self) -> tuple[int, float, str]:
        return tuple((self.year, self.nse, self.nse_se, self.nse_re))

    def __getitem__(self, idx):
        return self.astuple()[idx]

    def __iter__(self):
        return iter(self.astuple())

    def __repr__(self) -> str:
        return f"Row(year={self.year}, nse={self.nse}, nse_se={self.nse_se}, nse_re={self.nse_re})"

# %%


wb = xl.load_workbook('source.xlsx')

sheets = {sg: wb[sg] for sg in ('G4', 'G8')}

data = {
    'state': 'Vermont',
    'R4': [Row(2005, None, None, None),
           Row(2007, 263, 1.4, None),
           Row(2009, 253.502772375101, 0.638233566056301, 0.2665152040405)]
}

for sg, ws in sheets.items():
    ws['B5'].value = ws['B5'].value.replace('[STATE_NAME]', data['state'])
    datarows = data['R4']

    for row in range(len(data['R4'])):
        for col in range(4):
            cell = ws.cell(row+9, col+2)
            cell.value = datarows[row][col]


wb.save('mywb.xlsx')
# %%
# –	†	†
# 263 	1.4 	–
# 254 	0.6 	0.3
# 248 	0.6 	0.2
# 245 	0.5 	0.2
# 266 	0.8 	0.2
# 272 	1.5 	0.4
