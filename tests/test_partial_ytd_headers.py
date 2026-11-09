import openpyxl
from ir_parse import Item,extract,CUM,POINT


def test_partial_ytd_cannot_replace_fy_or_be_summed_as_quarter(tmp_path):
    w=openpyxl.Workbook();s=w.active;s.title='CSM'
    s.cell(2,2).value=2025;s.merge_cells('B2:E2')
    for col,label in enumerate(['1Q','2Q','3Q','4Q'],2):s.cell(3,col).value=label
    for col,label in [(7,2025),(8,'3Q25'),(9,'1H25')]:
        s.cell(2,col).value=label;s.cell(3,col).value='YTD'
    s.cell(4,1).value='신계약'
    for col,value in [(2,100),(3,200),(4,300),(5,400),(7,1000),(8,600),(9,300)]:s.cell(4,col).value=value
    path=tmp_path/'factsheet.xlsx';w.save(path)
    item=Item('CSM','신계약',76,CUM)
    annual,_=extract(path,'2512',[item],unit=1)
    quarter,_=extract(path,'2509',[item],unit=1)
    half,_=extract(path,'2506',[item],unit=1)
    assert annual[76]==1000 and quarter[76]==600 and half[76]==300


def test_annual_missing_ytd_sums_four_quarters_without_partial_ytd(tmp_path):
    w=openpyxl.Workbook();s=w.active;s.title='CSM'
    s.cell(2,2).value=2025;s.merge_cells('B2:E2')
    for col,label in enumerate(['1Q','2Q','3Q','4Q'],2):s.cell(3,col).value=label
    s.cell(2,7).value='1H25';s.cell(3,7).value='YTD';s.cell(4,1).value='신계약'
    for col,value in [(2,100),(3,200),(4,300),(5,400),(7,300)]:s.cell(4,col).value=value
    path=tmp_path/'factsheet.xlsx';w.save(path)
    value,_=extract(path,'2512',[Item('CSM','신계약',76,CUM)],unit=1)
    assert value[76]==1000
