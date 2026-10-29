from collections import Counter

from ..observations import histogram_entropy


class RotationAnalyzer:
    def __init__(self, policy):
        self.policy = policy
        self.rotations = Counter()

    def observe(self, cell):
        self.rotations[cell.rotation] += 1

    def result(self, sheet):
        metrics = {'angle_histogram': {str(a): n for a, n in sorted(self.rotations.items())}, 'bit_depths': {}}
        findings = []
        for depth in (1, 2, 3):
            bands = Counter()
            for angle, count in self.rotations.items():
                if angle != 255:
                    bands[angle // (1 << depth)] += count
            total = sum(bands.values())
            band, count = bands.most_common(1)[0] if bands else (0, 0)
            local = Counter({angle & ((1 << depth) - 1): n for angle, n in self.rotations.items()
                             if angle != 255 and angle // (1 << depth) == band})
            entropy = histogram_entropy(local)
            metrics['bit_depths'][str(depth)] = {'dominant_band_share': count / max(total, 1),
                                                'dominant_band_low_bits_entropy': entropy}
            if (total >= self.policy.min_rotation_cells and count / total >= self.policy.rotation_band_share
                    and len(local) >= 2 and entropy / depth >= self.policy.rotation_entropy_fraction and not findings):
                findings.append({'rule': 'rotation_low_bits_pattern', 'family': 'rotation',
                                 'score': 35, 'sheet': sheet, 'matching_cells': count, 'bit_depth': depth})
        return metrics, findings
