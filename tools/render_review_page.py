"""Render/register a selected page before the current assistant inspects it."""
from pathlib import Path
import argparse
import sys
import pdfplumber
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from quarterly.core import load,dump,digest,COMPANIES

def render(run,source_prefix,pages):
    source=next(s for s in load(run/'source_manifest.json') if s['sha256'].startswith(source_prefix))
    assets=load(run/'review_assets.json');folder=run/'selected_pages'/source['sha256'][:16];folder.mkdir(parents=True,exist_ok=True)
    with pdfplumber.open(source['path']) as pdf:
        for number in pages:
            page=pdf.pages[number-1];image=folder/f'{number:04d}.png';text=folder/f'{number:04d}.txt'
            page.to_image(resolution=110).save(image,format='PNG');text.write_text(page.extract_text() or '',encoding='utf-8')
            assets=[a for a in assets if not (a['source_id']==source['id'] and a['page']==number)]
            assets.append(dict(source_id=source['id'],company=source['company'],short=dict(COMPANIES)[source['company']],
                page=number,image=str(image.resolve()),image_sha256=digest(image),text=str(text.resolve()),text_sha256=digest(text),
                pdf_path=source['path'],pdf_sha256=source['sha256']))
            print(image,flush=True)
    dump(run/'review_assets.json',assets)

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('run',type=Path);p.add_argument('source');p.add_argument('pages',nargs='+',type=int);a=p.parse_args();render(a.run,a.source,a.pages)
