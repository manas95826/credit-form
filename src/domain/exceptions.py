"""Domain-specific exceptions."""


class FormFillerError(Exception):
    """Base exception for form filler operations."""
    pass


class FieldDetectionError(FormFillerError):
    """Raised when field detection fails."""
    pass


class AIGenerationError(FormFillerError):
    """Raised when AI data generation fails."""
    pass


class ExcelProcessingError(FormFillerError):
    """Raised when Excel processing fails."""
    pass


class ConfigurationError(FormFillerError):
    """Raised when configuration is invalid."""
    pass

