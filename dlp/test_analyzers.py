import io
import random
import unittest
import zipfile
from collections import Counter

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

from dlp import Policy
from dlp.aggregation import aggregate
from dlp.analyzers.color import ColorAnalyzer
from dlp.analyzers.formatting import FormattingAnalyzer
from dlp.analyzers.package import analyze_package
from dlp.analyzers.rotation import RotationAnalyzer
from dlp.analyzers.visibility import VisibilityAnalyzer
from dlp.observations import CellObservation, observe


def sample(**changes):
    values = dict(address='A1', units=128, text_bytes=128, run_lengths=(),
                  colors=Counter({0x3C5068: 128}), fonts=Counter({'unrelated font': 128}),
                  rotation=0, hidden=False, hidden_number_format=False)
    values.update(changes)
    return CellObservation(**values)


class AnalyzerTests(unittest.TestCase):
    def test_color_depths_and_unknown_fonts(self):
        for depth in (1, 2, 3):
            colors = Counter()
            mask = (1 << depth) - 1
            for n in range(min(64, 1 << (3 * depth))):
                red, green, blue = (n >> (depth * 2)) & mask, (n >> depth) & mask, n & mask
                colors[(0x386078 & ~(mask * 0x010101)) | red << 16 | green << 8 | blue] += 8
            analyzer = ColorAnalyzer(Policy())
            for address in ('A1', 'A2'):
                analyzer.observe(sample(address=address, colors=colors, units=sum(colors.values())))
            metrics, findings = analyzer.result('Sheet')
            self.assertEqual(findings[0]['bit_depth'], depth)
            self.assertEqual(findings[0]['score'], 50)
            self.assertEqual(metrics['rgb_coverage'], 1)

    def test_low_entropy_color_is_not_enough(self):
        analyzer = ColorAnalyzer(Policy())
        analyzer.observe(sample(colors=Counter({0x3C5068: 127, 0x3C5069: 1})))
        self.assertEqual(analyzer.result('Sheet')[1], [])

    def test_single_channel_one_bit(self):
        analyzer = ColorAnalyzer(Policy())
        for address in ('A1', 'A2'):
            analyzer.observe(sample(address=address, colors=Counter({0x3C5068: 64, 0x3D5068: 64})))
        metrics, findings = analyzer.result('Sheet')
        self.assertEqual(findings[0]['channel'], 'r')
        self.assertEqual(findings[0]['bit_depth'], 1)
        self.assertEqual(metrics['bit_depths']['1']['r']['entropy'], 1)

    def test_unknown_colors_have_explicit_coverage(self):
        analyzer = ColorAnalyzer(Policy())
        analyzer.observe(sample(colors=Counter()))
        metrics, findings = analyzer.result('Sheet')
        self.assertEqual(metrics['rgb_coverage'], 0)
        self.assertEqual(findings, [])

    def test_formatting_short_runs(self):
        analyzer = FormattingAnalyzer(Policy())
        analyzer.observe(sample(run_lengths=(1,) * 128))
        analyzer.observe(sample(run_lengths=(2,) * 64))
        metrics, findings = analyzer.result('Sheet')
        self.assertEqual(metrics['rich_text_cell_share'], 1)
        self.assertEqual(metrics['short_run_unit_share'], 1)
        self.assertEqual(findings[0]['score'], 45)

    def test_normal_highlighting_and_small_samples(self):
        analyzer = FormattingAnalyzer(Policy())
        analyzer.observe(sample(run_lengths=(64, 64)))
        analyzer.observe(sample(units=10, run_lengths=(1,) * 10))
        self.assertEqual(analyzer.result('Sheet')[1], [])

    def test_rotation_distribution(self):
        for depth in (1, 2, 3):
            analyzer = RotationAnalyzer(Policy())
            for i in range(64):
                analyzer.observe(sample(rotation=-8 + i % (1 << depth)))
            metrics, findings = analyzer.result('Sheet')
            self.assertTrue(findings)
            self.assertEqual(findings[0]['bit_depth'], depth)
            self.assertEqual(len(metrics['bit_depths']), 3)

    def test_rotation_constant_or_too_small(self):
        for count in (8, 64):
            analyzer = RotationAnalyzer(Policy())
            for i in range(count):
                analyzer.observe(sample(rotation=i % 4 if count == 8 else 0))
            self.assertEqual(analyzer.result('Sheet')[1], [])

    def test_visibility_is_weak_evidence(self):
        analyzer = VisibilityAnalyzer()
        analyzer.observe(sample(hidden=True))
        analyzer.observe(sample(hidden_number_format=True))
        metrics, findings = analyzer.result('Sheet')
        self.assertEqual(metrics['hidden_text_cells'], 2)
        self.assertEqual(findings[0]['score'], 15)

    def test_package_media_and_extra_data(self):
        for prefix, expected in [('xl/media/', 0), ('xl/embeddings/', 1)]:
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as archive:
                archive.writestr(prefix + 'data.bin', random.Random(1).randbytes(300000))
            metrics, findings = analyze_package(buffer.getvalue(), 100, 100, 100, Policy())
            self.assertEqual(len(findings), expected)
            if findings:
                self.assertLess(findings[0]['score'], Policy().alert_score)

    def test_aggregation_does_not_sum_repeated_signals(self):
        finding = {'family': 'color', 'score': 50}
        self.assertEqual(aggregate([finding] * 100, Policy())['formatting_steganography']['risk_score'], 50)
        result = aggregate([finding, {'family': 'fragmentation', 'score': 45}], Policy())
        self.assertEqual(result['formatting_steganography']['risk_score'], 60)
        self.assertEqual(result['cellveilxl']['risk_score'], 0)

    def test_policy_changes_thresholds(self):
        analyzer = FormattingAnalyzer(Policy(min_fragmented_units=1000))
        analyzer.observe(sample(run_lengths=(1,) * 128))
        self.assertEqual(analyzer.result('Sheet')[1], [])

    def test_utf16_observation_and_no_plaintext(self):
        wb = Workbook()
        cell = wb.active['A1']
        cell.value = CellRichText([TextBlock(InlineFont(color='FF123456'), '\U0001f600'), 'secret'])
        result = observe(cell)
        self.assertEqual(result.units, 8)
        self.assertEqual(result.run_lengths, (2, 6))
        self.assertNotIn('secret', repr(result))


if __name__ == '__main__':
    unittest.main()
