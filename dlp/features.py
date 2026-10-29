from .analyzers.color import ColorAnalyzer
from .analyzers.formatting import FormattingAnalyzer
from .analyzers.rotation import RotationAnalyzer
from .analyzers.visibility import VisibilityAnalyzer
from .models import ScanLimit
from .observations import observe


def analyze_features(ws, limits, policy, budget):
    from openpyxl.utils.cell import column_index_from_string
    analyzers = {'formatting': FormattingAnalyzer(policy), 'rotation': RotationAnalyzer(policy),
                 'color': ColorAnalyzer(policy), 'visibility': VisibilityAnalyzer()}
    hidden_columns = set()
    for name, dimension in ws.column_dimensions.items():
        if dimension.hidden:
            first = dimension.min or column_index_from_string(name)
            last = dimension.max or first
            hidden_columns.update(range(max(1, first), min(last, 16384) + 1))
    text_bytes = value_bytes = visible_bytes = formula_cells = 0
    for _, cell in sorted(ws._cells.items()):
        if cell.value is None:
            continue
        if cell.data_type == 'f':
            formula_cells += 1
            continue
        value_size = len(str(cell.value).encode('utf-8', errors='surrogatepass'))
        value_bytes += value_size
        hidden = ws.sheet_state != 'visible' or bool(ws.row_dimensions[cell.row].hidden) or cell.column in hidden_columns
        if not hidden and cell.number_format != ';;;':
            visible_bytes += value_size
        observation = observe(cell, hidden=hidden)
        if observation is None:
            continue
        budget[0] += observation.units
        if budget[0] > limits.characters:
            raise ScanLimit('character_limit')
        text_bytes += observation.text_bytes
        for analyzer in analyzers.values():
            analyzer.observe(observation)
    metrics, findings = {}, []
    for name, analyzer in analyzers.items():
        metrics[name], signals = analyzer.result(ws.title)
        findings.extend(signals)
    return {'sheet': ws.title, 'state': ws.sheet_state, 'analyzers': metrics,
            'text_utf8_bytes': text_bytes, 'serialized_value_utf8_bytes': value_bytes,
            'visible_value_utf8_bytes_estimate': visible_bytes,
            'formula_cells_unmeasured': formula_cells}, findings
