"""Do not count Excel's blank-dependent zero caches as extracted facts."""
from __future__ import annotations
import re
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import range_boundaries
from .core import numeric


def formula_available(book, sheet, coordinate, seen=None):
    seen=set() if seen is None else set(seen)
    key=(sheet,coordinate)
    if key in seen:return False
    seen.add(key)
    cell=book[sheet][coordinate]
    if cell.data_type!='f':return numeric(cell.value)
    try:
        for token in Tokenizer(cell.value).items:
            if token.subtype!='RANGE':continue
            text=token.value
            if '!' in text:
                name,text=text.rsplit('!',1);name=name.strip("'")
            else:name=sheet
            if name not in book.sheetnames or not re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?',text):return False
            c1,r1,c2,r2=range_boundaries(text)
            if (r2-r1+1)*(c2-c1+1)>5000:return False
            for row in book[name].iter_rows(min_row=r1,max_row=r2,min_col=c1,max_col=c2):
                for ref in row:
                    if not formula_available(book,name,ref.coordinate,seen):return False
        return True
    except Exception:return False
