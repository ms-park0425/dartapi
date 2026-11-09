"""Prepare all 2025 quarter reconstructions; no reference-based filling or AI delegation."""
from pathlib import Path
import argparse
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from quarterly.core import Period, load, dump
from backtest_2512 import initialize, run_prepare


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--periods', nargs='+', default=['2503', '2506', '2509', '2512'])
    parser.add_argument('--reuse', action='store_true', help='Reuse a completed preparation, never overwrite frozen outputs')
    args = parser.parse_args()
    for code in args.periods:
        period = Period.parse(code)
        if period.year != 2025:
            parser.error('This development reconstruction is authorized for 2025 only')
        pointer = ROOT/'runs'/('active_backtest.txt' if code == '2512' else f'active_backtest_{code}.txt')
        if args.reuse and pointer.exists():
            base = Path(pointer.read_text(encoding='utf-8').strip())
            if (base/'run/state.json').exists():
                print(code, 'existing preparation retained:', base, flush=True)
                continue
        base = initialize(period)
        config = load(base/'run_config.json')
        config.update(ai_source_discovery=False, use_fisis=False,
                      review_keywords=['지급여력', '보험계약마진', '보험부채현황', '보험계약부채', 'CSM'])
        dump(base/'run_config.json', config)
        run_prepare(base)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
