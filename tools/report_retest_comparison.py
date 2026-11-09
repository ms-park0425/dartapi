"""Compare a new frozen annual run with its preserved baseline on identical targets."""
from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from quarterly.core import load,dump,digest,COMPANIES

def compare(base):
    baseline=load(base/'retest_baseline.json');old=Path(baseline['path'])
    assert digest(old/'run/final.xlsx')==baseline['final_sha256']
    assert digest(base/'run/final.xlsx')==load(base/'frozen_outputs.json')['final_sha256']
    one=load(old/'accuracy_summary.json');two=load(base/'accuracy_summary.json')
    a={(r['company'],r['field']):r for r in load(old/'stage2_scores.json')}
    b={(r['company'],r['field']):r for r in load(base/'stage2_scores.json')}
    assert a.keys()==b.keys(),'Comparison target sets differ; do not compare headline percentages silently'
    improved=[dict(company=k[0],field=k[1],before=a[k]['status'],after=r['status']) for k,r in b.items() if a[k]['status']!='OK' and r['status']=='OK']
    regressed=[dict(company=k[0],field=k[1],before=a[k]['status'],after=r['status']) for k,r in b.items() if a[k]['status']=='OK' and r['status']!='OK']
    dump(base/'retest_comparison.json',dict(baseline=str(old),current=str(base),target_count=len(b),improved=improved,regressed=regressed,
        before=one['stage2']['ALL'],after=two['stage2']['ALL'],baseline_unchanged=True))
    lines=['# 2512 지침 보완 후 재추출 비교','','동일한 정답본·직접 입력 대상에서 비교했다. 기존 실행과 정답본은 보존했고 새 결과를 확정한 뒤 채점했다. 이미 진단한 2025년 자료를 사용하는 개발 재검증으로, 독립 미학습 성능이 아니다.','',
        '| 항목 | 기존 최종 결과 | 이번 최종 결과 |','|---|---:|---:|']
    for label,key in [('비교 대상','total'),('일치','ok'),('불일치','wrong'),('빈칸','missing')]:
        lines.append(f"| {label} | {one['stage2']['ALL'][key]} | {two['stage2']['ALL'][key]} |")
    for label,key in [('전체 대상 일치율','accuracy_all'),('채워진 값의 일치율','accuracy_filled')]:
        lines.append(f"| {label} | {one['stage2']['ALL'][key]:.2%} | {two['stage2']['ALL'][key]:.2%} |")
    lines += ['',f'기존 불일치/빈칸 → 일치: {len(improved)}개. 기존 일치 → 불일치/빈칸: {len(regressed)}개.','',
        '## 반영한 보완','',
        '- 2025년 네 분기에서 얻은 단위·범위·기간·부호·잠정/확정 지침을 사용했다.',
        '- 연간 XML을 수집기가 저장하는 이름과 기존 파서가 읽는 이름을 같은 원본 스냅샷으로 연결했다.',
        '- PDF 대조 중 메리츠 FY 옆의 3Q/1H YTD를 연간으로 오인하는 오류를 발견했다. 채점 전 수정하고 정밀 팩트시트 값으로 재입력했다. 변경 전 출력도 `before_partial_ytd_fix/`에 보존했다.',
        '- 38개 페이지 이미지를 다시 확인했다. 실제 검증 항목·보류 항목은 실행 요약과 근거 파일에 기록했다.',
        '- PDF의 다른 숫자만으로 기존값을 교체하지 않았다. 원수/순보유 등 범위가 확정되지 않은 항목은 보류했다. FISIS와 CALC·수식 셀은 제외했다.','',
        '## 회사별 최종 결과','','| 회사 | 대상 | 기존 일치 | 이번 일치 | 이번 불일치 | 빈칸 |','|---|---:|---:|---:|---:|---:|']
    for _,short in COMPANIES:
        x,y=one['stage2'][short],two['stage2'][short]
        lines.append(f"| {short} | {y['total']} | {x['ok']} | {y['ok']} | {y['wrong']} | {y['missing']} |")
    lines += ['','## 기존에 맞았으나 이번에 달라진 항목','']
    reasons={('현대해상',92):'IR 이미지에서 잠정 표시를 확인하지 못해 빈칸으로 보류했다. 이전의 단순 IR→잠정 분류를 그대로 적용하지 않았다.',
             ('DB손보',97):'사업보고서 주석이 외부검증 전 잠정치라고 명시했지만 기존 파서가 최종 칸에도 복사했다. 최종 경영공시와 기준이 달랐다. 원인 확인 후 잠정치의 최종 칸 복사를 막는 코드를 추가했으며 이 확정 실행에는 소급 적용하지 않았다.'}
    lines += [f"- {r['company']} Col{r['field']}: {r['before']} → {r['after']}. {reasons.get((r['company'],r['field']),'원문·기준 차이 추가 진단 필요.')} 이번 확정값·점수는 소급 변경하지 않는다." for r in regressed] or ['- 없음']
    lines += ['','[이번 엑셀](2512_추출결과.xlsx) · [이번 1차/2차 상세 검증](검증결과.md) · [셀별 비교](accuracy_detail.csv)','']
    (base/'재추출_비교.md').write_text('\n'.join(lines),encoding='utf-8')
    print('Retest:',two['stage2']['ALL'],'improved',len(improved),'regressed',len(regressed))

if __name__=='__main__':
    sys.stdout.reconfigure(encoding='utf-8');compare(Path(sys.argv[1]))
