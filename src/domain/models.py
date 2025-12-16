"""Domain models for form filler application."""
from dataclasses import dataclass
from typing import Dict, Optional


@dataclass
class FormField:
    """Represents a detected form field."""
    label: str
    target_coordinate: str
    is_merged: bool = False

    def __post_init__(self):
        self.is_merged = ":" in self.target_coordinate


@dataclass
class FieldDetectionResult:
    """Result of field detection operation."""
    fields: Dict[str, FormField]
    merged_map: Dict[str, str]
    
    @property
    def count(self) -> int:
        return len(self.fields)


@dataclass
class FillResult:
    """Result of filling Excel with data."""
    filled_count: int
    skipped_count: int
    errors: list[str]
    
    @property
    def success(self) -> bool:
        return len(self.errors) == 0

