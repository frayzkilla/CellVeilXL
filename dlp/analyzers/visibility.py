class VisibilityAnalyzer:
    def __init__(self):
        self.cells = self.hidden_cells = self.hidden_units = self.hidden_format_cells = 0
        self.examples = []

    def observe(self, cell):
        self.cells += 1
        if cell.hidden or cell.hidden_number_format:
            self.hidden_cells += 1
            self.hidden_units += cell.units
            self.hidden_format_cells += cell.hidden_number_format
            if len(self.examples) < 5:
                self.examples.append(cell.address)

    def result(self, sheet):
        metrics = {'hidden_text_cells': self.hidden_cells, 'hidden_text_units': self.hidden_units,
                   'hidden_text_cell_share': self.hidden_cells / max(self.cells, 1),
                   'hidden_number_format_cells': self.hidden_format_cells}
        findings = []
        if self.hidden_cells:
            findings.append({'rule': 'hidden_text_content', 'family': 'visibility', 'score': 15,
                             'sheet': sheet, 'cells': self.examples, 'matching_cells': self.hidden_cells})
        return metrics, findings
