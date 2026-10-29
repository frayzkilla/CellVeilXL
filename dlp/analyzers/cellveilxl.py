from openpyxl.cell.rich_text import CellRichText
from openpyxl.utils.cell import range_boundaries

from ..models import ScanLimit
from ..observations import _runs


FONTS = ('arial', 'calibri', 'segoe ui', 'tahoma')
MAX_PACKET = 1048576


def _cell_bits(cell):
    rotation = cell.alignment.textRotation or 0
    if rotation == 255:
        return None
    orientation = 90 - rotation if rotation > 90 else rotation
    runs = list(_runs(cell))
    names = {str(name).lower() for text, name, _, _ in runs if text}
    if len(names) != 1 or next(iter(names)) not in FONTS:
        return None
    bits = [f'{orientation & 3:02b}{FONTS.index(next(iter(names))):02b}']
    for text, _, color, _ in runs:
        if color is None:
            return ''.join(bits)
        symbol = ((color >> 16) & 3) * 16 + ((color >> 8) & 3) * 4 + (color & 3)
        units = len(text.encode('utf-16-le', errors='surrogatepass')) // 2
        bits.append(f'{symbol:06b}' * units)
    return ''.join(bits)


def _packet_findings(bits, boundaries, sheet, limits):
    findings = []
    for offset, address in boundaries:
        if len(findings) >= limits.findings:
            raise ScanLimit('finding_limit')
        if len(bits) - offset < 80:
            continue
        header = bytes(int(bits[i:i + 8], 2) for i in range(offset, offset + 80, 8))
        known = header[:4] == b'CVX3'
        mode, version = header[4:6]
        length = int.from_bytes(header[6:10], 'little')
        minimum = {1: 29, 2: 50}.get(mode, MAX_PACKET + 1)
        key_size = 11 if mode == 1 else 32
        plausible = version == 1 and minimum <= length <= MAX_PACKET and (length - key_size - 16) % 2 == 0
        if not plausible:
            if known:
                findings.append({'rule': 'cvx3_signature', 'score': 70, 'sheet': sheet,
                                 'cell': address, 'complete': False})
            continue
        complete = offset + (length + 10) * 8 <= len(bits)
        if not complete and not known:
            continue
        key_start = offset + 80
        key_bits = bits[key_start:key_start + key_size * 8]
        if complete:
            public = int.from_bytes(bytes(int(key_bits[i:i + 8], 2)
                                         for i in range(0, len(key_bits), 8)), 'little')
            valid_public = (1 < public < 154742504910672534362390399 and
                            pow(public, 77371252455336267181195199, 154742504910672534362390399) == 1) if mode == 1 else 0 < public < 2 ** 255 - 19
            if not valid_public:
                if known:
                    findings.append({'rule': 'cvx3_signature', 'score': 70, 'sheet': sheet,
                                     'cell': address, 'complete': False})
                continue
        findings.append({'rule': 'cellveilxl_packet' if known else 'cellveilxl_packet_custom_magic',
                         'score': (95 if known else 85) if complete else 70,
                         'sheet': sheet, 'cell': address, 'magic_hex': header[:4].hex(),
                         'mode': 'demo' if mode == 1 else 'production', 'version': version,
                         'payload_bytes': length, 'complete': complete,
                         'checks': ['mode', 'version', 'length_bounds', 'utf16_length_parity'] +
                         (['available_bits', 'public_key_constraints'] if complete else [])})
        if len(findings) >= limits.findings:
            raise ScanLimit('finding_limit')
    return findings


def analyze_sheet(ws, limits, host_range, budget):
    findings, chunks, boundaries = [], [], []
    position = 0
    examined, decodable = 0, 0
    bounds = range_boundaries(host_range) if host_range else None
    if bounds and (None in bounds or bounds[0] > bounds[2] or bounds[1] > bounds[3]):
        raise ValueError('invalid_host_range')
    excluded = list(ws.merged_cells.ranges)
    excluded.extend(r for cf in ws.conditional_formatting for r in cf.sqref.ranges)
    if len(excluded) > 1000:
        raise ScanLimit('formatting_range_limit')

    def flush():
        nonlocal position
        findings.extend(_packet_findings(''.join(chunks), boundaries, ws.title, limits))
        chunks.clear()
        boundaries.clear()
        position = 0

    for _, cell in sorted(ws._cells.items()):
        if bounds and not (bounds[0] <= cell.column <= bounds[2] and bounds[1] <= cell.row <= bounds[3]):
            continue
        if cell.data_type in ('f', 'e') or not isinstance(cell.value, (str, CellRichText)) or not str(cell.value):
            continue
        if any(cell.coordinate in r for r in excluded):
            continue
        runs = list(_runs(cell))
        bold = {b for text, _, _, b in runs if text}
        if bold == {True}:
            continue
        if bold == {False, True}:
            flush()
            continue
        budget[0] += len(str(cell.value).encode('utf-16-le', errors='surrogatepass')) // 2
        if budget[0] > limits.characters:
            raise ScanLimit('character_limit')
        bits = _cell_bits(cell)
        examined += 1
        if bits is None:
            flush()
            continue
        decodable += 1
        boundaries.append((position, cell.coordinate))
        chunks.append(bits)
        position += len(bits)
        if len(bits) < 4 + 6 * len(str(cell.value).encode('utf-16-le', errors='surrogatepass')) // 2:
            flush()
    flush()
    if len(findings) > limits.findings:
        raise ScanLimit('finding_limit')
    return findings, {'sheet': ws.title, 'text_cells_examined': examined,
                      'decodable_cells': decodable}


