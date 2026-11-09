"""Company-specific PDF navigation; never a lookup of historical answer values."""
from pathlib import Path
from .core import load

RULEBOOK = Path(__file__).resolve().parent.parent / 'docs' / 'company_field_sources.json'
LESSONS = RULEBOOK.parent / 'quarterly_lessons.json'
NAVIGATION = RULEBOOK.parent / 'review_navigation_2025.json'


def company_catalog(catalog, company, config):
    """Run-specific semantic definitions; never mutate other companies' fields."""
    result = {str(k):dict(v) for k,v in catalog.items()}
    for field, meta in config.get('company_fields', {}).get(company, {}).items():
        if str(field) not in result:
            raise ValueError(f'회사별 정의에 없는 Col: {field}')
        result[str(field)].update(meta)
    return result


def reviewer_guide(company, catalog):
    book = load(RULEBOOK, {})
    entry = book.get('companies', {}).get(company, {})
    fields = {int(f) for f in catalog}
    groups = []
    for rule in book.get('field_groups', []):
        selected = sorted(fields.intersection(rule['fields']))
        if selected:
            groups.append({**rule, 'fields': selected})
    lessons = load(LESSONS, {})
    navigation = load(NAVIGATION, {})
    return dict(policy=book.get('policy', []), company_notes=entry.get('notes', []),
                cross_quarter_lessons=[r for r in lessons.get('rules', []) if
                    (not r.get('companies') or company in r['companies']) and fields.intersection(r.get('fields', fields))],
                reviewed_2025_navigation=[r for r in navigation.get('examples', []) if r['company']==company and r['field'] in fields],
                field_groups=groups,
                factsheet_locations=[e for e in entry.get('factsheet_locations', []) if e['field'] in fields],
                historical_navigation_examples=[e for e in entry.get('examples', []) if e['field'] in fields],
                example_policy='2025년 직접 확인한 탐색 예시이며 현재 분기의 수치·페이지·표 구조는 다시 확인. 정답 숫자는 포함하지 않음.')
