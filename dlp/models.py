from dataclasses import dataclass


@dataclass(frozen=True)
class Limits:
    file_bytes: int = 20 * 1024 * 1024
    expanded_bytes: int = 64 * 1024 * 1024
    zip_entries: int = 2000
    cells: int = 200000
    sheet_area: int = 500000
    sheets: int = 100
    characters: int = 2000000
    findings: int = 100

    def __post_init__(self):
        if any(type(value) is not int or value < 0 for value in self.__dict__.values()):
            raise ValueError('Limits must be nonnegative integers')


class ScanLimit(Exception):
    pass


@dataclass(frozen=True)
class Policy:
    alert_score: int = 60
    generic_cap: int = 80
    corroboration_bonus: int = 10
    min_color_units: int = 32
    color_base_share: float = .9
    color_entropy_fraction: float = .58
    min_rotation_cells: int = 16
    rotation_band_share: float = .9
    rotation_entropy_fraction: float = .7
    min_fragmented_units: int = 32
    short_run_share: float = .8
    size_min_bytes: int = 262144
    size_ratio: float = 50

    def __post_init__(self):
        if not 0 < self.alert_score <= self.generic_cap < 85:
            raise ValueError('Expected 0 < alert_score <= generic_cap < 85')
        if not 0 <= self.corroboration_bonus <= 20:
            raise ValueError('Invalid corroboration bonus')
        if min(self.min_color_units, self.min_rotation_cells, self.min_fragmented_units, self.size_ratio) <= 0 or self.size_min_bytes < 0:
            raise ValueError('Invalid sample-size or size threshold')
        fractions = (self.color_base_share, self.color_entropy_fraction,
                     self.rotation_band_share, self.rotation_entropy_fraction, self.short_run_share)
        if any(not 0 < fraction <= 1 for fraction in fractions):
            raise ValueError('Expected fractional thresholds in (0, 1]')


