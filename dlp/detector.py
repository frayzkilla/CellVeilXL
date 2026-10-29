import hashlib
import io
import zipfile
from dataclasses import asdict
from pathlib import Path
from time import perf_counter

from defusedxml import ElementTree as SafeET
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

from .models import Limits, Policy, ScanLimit
from .analyzers.cellveilxl import analyze_sheet
from .analyzers.package import analyze_package
from .aggregation import aggregate
from .features import analyze_features


def _preflight(data, limits):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        entries = archive.infolist()
        if len(entries) > limits.zip_entries or sum(e.file_size for e in entries) > limits.expanded_bytes:
            raise ScanLimit('archive_size_limit')
        names = [e.filename for e in entries]
        if len(names) != len(set(names)):
            raise ValueError('duplicate_archive_members')
        if '[Content_Types].xml' not in names or 'xl/workbook.xml' not in names:
            raise ValueError('not_an_xlsx_workbook')
        cells = merged_area = sheets = 0
        for entry in entries:
            if entry.flag_bits & 1:
                raise ValueError('encrypted_archive')
            if not entry.filename.endswith(('.xml', '.rels')):
                continue
            with archive.open(entry) as stream:
                for _, node in SafeET.iterparse(stream, events=('end',), forbid_dtd=True):
                    tag = node.tag.rsplit('}', 1)[-1]
                    if tag == 'c':
                        cells += 1
                        if cells > limits.cells:
                            raise ScanLimit('cell_limit')
                    if entry.filename == 'xl/workbook.xml' and tag == 'sheet':
                        sheets += 1
                        if sheets > limits.sheets:
                            raise ScanLimit('sheet_limit')
                    if tag in ('dimension', 'mergeCell') and node.get('ref'):
                        left, top, right, bottom = range_boundaries(node.get('ref'))
                        if None in (left, top, right, bottom) or (right - left + 1) * (bottom - top + 1) > limits.sheet_area:
                            raise ScanLimit('sheet_area_limit')
                        if tag == 'mergeCell':
                            merged_area += (right - left + 1) * (bottom - top + 1)
                            if merged_area > limits.cells:
                                raise ScanLimit('merged_cell_limit')
                    node.clear()


def analyze_xlsx(path, *, host_range=None, limits=None, policy=None):
    started = perf_counter()
    limits = limits or Limits()
    policy = policy or Policy()
    report = {'schema_version': 1, 'detector': 'xlsx_formatting', 'detector_version': '1.0',
              'status': 'ok', 'sha256': None, 'risk_score': 0, 'verdict': 'no_indicators',
              'findings': [], 'errors': [], 'sheets_scanned': 0, 'host_range': host_range,
              'coverage': [], 'features': {'sheets': [], 'package': {}}, 'profiles': {},
              'policy': asdict(policy), 'limits': asdict(limits)}
    workbook = None
    try:
        with Path(path).open('rb') as stream:
            data = stream.read(limits.file_bytes + 1)
        if len(data) > limits.file_bytes:
            raise ScanLimit('file_size_limit')
        report['sha256'] = hashlib.sha256(data).hexdigest()
        _preflight(data, limits)
        workbook = load_workbook(io.BytesIO(data), rich_text=True, data_only=False, keep_links=False)
        if len(workbook.worksheets) > limits.sheets:
            raise ScanLimit('sheet_limit')
        budget = [0]
        feature_budget = [0]
        for ws in workbook.worksheets:
            features, generic_findings = analyze_features(ws, limits, policy, feature_budget)
            report['features']['sheets'].append(features)
            report['findings'].extend(generic_findings)
            findings, coverage = analyze_sheet(ws, limits, host_range, budget)
            for finding in findings:
                finding['family'] = 'cellveilxl'
            report['findings'].extend(findings)
            report['coverage'].append(coverage)
            report['sheets_scanned'] += 1
            if len(report['findings']) > limits.findings:
                raise ScanLimit('finding_limit')
        sheets = report['features']['sheets']
        metrics, findings = analyze_package(data,
            sum(s['visible_value_utf8_bytes_estimate'] for s in sheets),
            sum(s['text_utf8_bytes'] for s in sheets),
            sum(s['serialized_value_utf8_bytes'] for s in sheets), policy)
        report['features']['package'] = metrics
        report['findings'].extend(findings)
        if len(report['findings']) > limits.findings:
            raise ScanLimit('finding_limit')
    except ScanLimit as exc:
        report['status'] = 'limit_exceeded'
        report['errors'].append(str(exc))
    except Exception as exc:
        report['status'] = 'error'
        report['errors'].append(type(exc).__name__)
    finally:
        if workbook is not None:
            workbook.close()
    report['profiles'] = aggregate(report['findings'], policy)
    report['risk_score'] = max(p['risk_score'] for p in report['profiles'].values())
    score = report['risk_score']
    report['verdict'] = 'high_risk' if score >= 85 else 'suspicious' if score >= policy.alert_score else 'review' if score else 'no_indicators'
    if report['status'] != 'ok':
        report['verdict'] = 'inconclusive'
    report['duration_ms'] = round((perf_counter() - started) * 1000, 2)
    return report
