"""Application configuration."""
import os
from typing import List
from dataclasses import dataclass


@dataclass
class AIConfig:
    """AI service configuration."""
    api_key: str
    model: str = "gpt-4o-mini"
    temperature: float = 0.2
    max_retries: int = 3
    base_tokens: int = 2000
    tokens_per_field: int = 80
    token_buffer: int = 1000


@dataclass
class LabelDetectionConfig:
    """Label detection configuration."""
    keywords: List[str]
    min_label_length: int = 2
    max_label_length_with_colon: int = 100
    max_label_length_without_colon: int = 40
    min_left_cell_text_length: int = 10


class Config:
    """Application configuration."""
    
    def __init__(self):
        self._ai_config = None
        self._label_config = None
    
    @property
    def ai(self) -> AIConfig:
        """Get AI configuration."""
        if self._ai_config is None:
            api_key = os.getenv("OPENAI_API_KEY")
            if not api_key:
                raise ValueError("OPENAI_API_KEY environment variable is not set")
            self._ai_config = AIConfig(api_key=api_key)
        return self._ai_config
    
    @property
    def label_detection(self) -> LabelDetectionConfig:
        """Get label detection configuration."""
        if self._label_config is None:
            keywords = [
                'nombre', 'dirección', 'teléfono', 'telefono', 'email', 'correo', 'fecha',
                'código', 'codigo', 'actividad', 'representante', 'razón', 'razon', 'social',
                'nit', 'rfc', 'curp', 'ciudad', 'estado', 'país', 'pais', 'cp', 'código postal',
                'banco', 'cuenta', 'clabe', 'swift', 'iban', 'moneda', 'monto', 'importe',
                'apellido', 'paterno', 'materno', 'nacimiento', 'edad', 'género', 'genero',
                'ocupación', 'ocupacion', 'profesión', 'profesion', 'empresa', 'puesto',
                'documento', 'identificación', 'identificacion', 'pasaporte', 'licencia',
                'contacto', 'emergencia', 'parentesco', 'beneficiario', 'titular',
                'firma', 'fecha de', 'lugar de', 'hora', 'folio', 'referencia', 'número', 'numero'
            ]
            self._label_config = LabelDetectionConfig(keywords=keywords)
        return self._label_config


# Global configuration instance
config = Config()

