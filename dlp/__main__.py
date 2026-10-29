import argparse
import json

from .detector import analyze_xlsx


def main():
    parser = argparse.ArgumentParser(description='Inspect XLSX formatting for CellVeilXL streams')
    parser.add_argument('file')
    parser.add_argument('--range', dest='host_range')
    args = parser.parse_args()
    result = analyze_xlsx(args.file, host_range=args.host_range)
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 2 if result['status'] != 'ok' else 1 if result['risk_score'] >= result['policy']['alert_score'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
