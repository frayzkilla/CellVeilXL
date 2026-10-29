from collections import Counter

from ..observations import histogram_entropy


class ColorAnalyzer:
    def __init__(self, policy):
        self.policy = policy
        self.known = self.total = 0
        self.models = {(depth, channel): {'histogram': Counter(), 'matching_cells': 0, 'examples': []}
                       for depth in (1, 2, 3) for channel in ('rgb', 'r', 'g', 'b')}

    def observe(self, cell):
        self.total += cell.units
        known = sum(cell.colors.values())
        self.known += known
        for (depth, channel), state in self.models.items():
            mask = (1 << depth) - 1
            shifts = (16, 8, 0) if channel == 'rgb' else ({'r': 16, 'g': 8, 'b': 0}[channel],)
            packed_mask = sum(mask << shift for shift in shifts)
            bases = Counter()
            for color, count in cell.colors.items():
                bases[color & (0xFFFFFF ^ packed_mask)] += count
                low = 0
                for shift in shifts:
                    low = low << depth | ((color >> shift) & mask)
                state['histogram'][low] += count
            if known < self.policy.min_color_units:
                continue
            base, count = bases.most_common(1)[0]
            lows = Counter({c & packed_mask: n for c, n in cell.colors.items()
                            if c & (0xFFFFFF ^ packed_mask) == base})
            total_bits = len(shifts) * depth
            min_distinct = min(16, max(2, (1 << total_bits) // 2))
            if (count / cell.units >= self.policy.color_base_share and len(lows) >= min_distinct
                    and histogram_entropy(lows) / total_bits >= self.policy.color_entropy_fraction):
                state['matching_cells'] += 1
                if len(state['examples']) < 5:
                    state['examples'].append(cell.address)

    def result(self, sheet):
        metrics = {'rgb_units': self.known, 'text_units': self.total,
                   'rgb_coverage': self.known / max(self.total, 1), 'bit_depths': {}}
        for (depth, channel), state in self.models.items():
            metrics['bit_depths'].setdefault(str(depth), {})[channel] = {
                'histogram': {str(k): n for k, n in sorted(state['histogram'].items())},
                'entropy': histogram_entropy(state['histogram']),
                'matching_cells': state['matching_cells']}
        (depth, channel), best = max(self.models.items(), key=lambda item: item[1]['matching_cells'])
        findings = []
        if best['matching_cells']:
            findings.append({'rule': 'rgb_low_bits_pattern', 'family': 'color',
                             'score': 50 if best['matching_cells'] >= 2 else 30,
                             'sheet': sheet, 'cells': best['examples'], 'bit_depth': depth, 'channel': channel,
                             'matching_cells': best['matching_cells']})
        return metrics, findings
