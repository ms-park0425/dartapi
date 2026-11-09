"""Public IR archive adapters. Select exact reporting titles, never the latest item."""
import re
import hashlib
from pathlib import Path
import subprocess
from urllib.parse import urljoin

from .sources import period_hint


def matching(title, period):
    if period.quarter == 4 and re.search(r'분기|상반기|[123]Q|Q[123]', title, re.I):
        return False
    return period_hint(title, period)


def public_post(collector, url, **kwargs):
    response = collector.session.post(url, timeout=(10, 40), **kwargs)
    response.raise_for_status()
    return response.json()


def discover(collector, company):
    p = collector.period
    found = []
    if company == '메리츠화재':
        data = public_post(collector, 'https://www.meritzgroup.com/web/ir1_1_search.do',
                           data={'bd_div_cd': 'IR011', 'acrsl_anun_yy': str(p.year)})
        for item in data.get('result', []):
            title = item.get('title_ko', '')
            if not matching(title, p):
                continue
            for file in item.get('fileMap', []):
                path = file.get('file_url', '')
                if file.get('lang_div_cd') in ('P', 'F') and re.search(r'\.(pdf|xlsx)$', path, re.I):
                    found.append((urljoin('https://www.meritzgroup.com/commfiles/', path.lstrip('/')), title))
    elif company == '삼성생명':
        endpoint = 'https://www.samsunglife.com/gw/api/display/board/content/'
        for year in ([p.year, p.year+1] if p.quarter == 4 else [p.year]):
            data = public_post(collector, endpoint+'list', data=dict(apiCgoNm='cm0024', pageNo=1,
                offset=100, orderByIdx=19, order='DESC', bltnItmCntnt04='Y', bltnItmCntnt06='KOR', bltnYmd=year))
            for item in data.get('response', {}).get('list', []):
                title = item.get('bltnTitlCntnt', '')
                if not matching(title, p):
                    continue
                for key in ('bltnItmCntnt01', 'bltnItmCntnt02'):
                    if not item.get(key):
                        continue
                    file = public_post(collector, endpoint+'file', data={'cntsApndFileId': item[key]})
                    response = file.get('response', {})
                    for record in (response if isinstance(response,list) else [response]):
                        if isinstance(record, dict) and record.get('fileDownUrl'):
                            found.append((urljoin('https://www.samsunglife.com', record['fileDownUrl']), title))
    return list(dict.fromkeys(found))


def collect_hanwha(collector, company):
    p=collector.period
    endpoint='https://company.hanwhalife.com/api/comp/v1/get/'
    data=public_post(collector,endpoint+'earnings-releases',json={'year':str(p.year)})
    for item in data.get('items',[]):
        title=item.get('title','')
        if not matching(title,p):continue
        for key in ('pdfAttachSeq','factSheetAttachSeq'):
            if not item.get(key):continue
            params={'fileType':'FACTSHEET','fileSeq':item[key]}
            response=collector.session.post(endpoint+'file',json=params,timeout=(10,40))
            response.raise_for_status();body=response.content
            if body.startswith(b'%PDF'):ext,kind='.pdf','pdf'
            elif body.startswith(b'PK'):ext,kind='.xlsx','factsheet'
            elif body.startswith(bytes.fromhex('d0cf11e0')):ext,kind='.xls','factsheet_xls'
            else:raise ValueError('Unsupported official IR file response')
            path=collector.root/'data'/company/'IR'/p.code/('download_'+hashlib.sha256(body).hexdigest()[:16]+ext)
            path.parent.mkdir(parents=True,exist_ok=True);path.write_bytes(body)
            original=collector.add(path,company,kind,endpoint+'file',title=title,request_method='POST',request_json=params,
                period_verification='exact_archive_title_requires_content_review')
            if ext=='.xls':
                converted=path.with_suffix('.xlsx')
                if not converted.exists():
                    subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-File',
                        str(Path(__file__).resolve().parents[1]/'convert_legacy_factsheet.ps1'),'-Path',str(path.resolve())],
                        check=True,capture_output=True,timeout=120)
                collector.add(converted,company,'factsheet',endpoint+'file',title=title,
                    request_method='POST',request_json=params,derived_from=original['id'],conversion='Excel XLS to XLSX')
