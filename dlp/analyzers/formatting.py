from collections import Counter

from ..observations import histogram_entropy


class FormattingAnalyzer:
    def __init__(self, policy):
        self.policy = policy
        self.cells = self.rich_cells = self.fragmented_cells = 0
        self.units = self.fragmented_units = 0
        self.examples = []
        self.fonts = Counter()

    def observe(self, cell):
        self.cells += 1
        self.fonts.update(cell.fonts)
        self.units += cell.units
        self.rich_cells += bool(cell.run_lengths)
        short_units = sum(n for n in cell.run_lengths if n <= 2)
        self.fragmented_units += short_units
        if cell.units >= self.policy.min_fragmented_units and short_units / cell.units >= self.policy.short_run_share:
            self.fragmented_cells += 1
            if len(self.examples) < 5:
                self.examples.append(cell.address)

    def result(self, sheet):
        metrics = {'text_cells': self.cells, 'rich_text_cells': self.rich_cells,
                   'rich_text_cell_share': self.rich_cells / max(self.cells, 1),
                   'short_run_unit_share': self.fragmented_units / max(self.units, 1),
                   'short_run_definition': 'UTF-16 run length <= 2',
                   'font_unit_histogram': dict(self.fonts), 'font_entropy': histogram_entropy(self.fonts),
                   'fragmented_cells': self.fragmented_cells}
        findings = []
        if self.fragmented_cells:
            findings.append({'rule': 'character_formatting', 'family': 'fragmentation',
                             'score': 45 if self.fragmented_cells >= 2 else 25,
                             'sheet': sheet, 'cells': self.examples,
                             'matching_cells': self.fragmented_cells})
        return metrics, findings
