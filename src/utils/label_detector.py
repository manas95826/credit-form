"""Label detection utilities for identifying form field labels."""
import re
from typing import Optional
from src.config import config


def looks_like_label(text: Optional[str]) -> bool:
    """Determine if text looks like a form label rather than a field value."""
    if not text or not isinstance(text, str):
        return False
    
    text = text.strip()
    cfg = config.label_detection
    
    if len(text) < cfg.min_label_length:
        return False
    
    text_lower = text.lower()
    ends_with_colon = text.rstrip().endswith(':')
    has_keyword = any(keyword in text_lower for keyword in cfg.keywords)
    
    if not (ends_with_colon or has_keyword):
        return False
    
    if ends_with_colon:
        if len(text) > cfg.max_label_length_with_colon:
            return False
        if '@' in text and not any(kw in text_lower for kw in ['email', 'correo', 'mail']):
            return False
        if text_lower.startswith(('http', 'www')):
            return False
        return True
    
    if len(text) > cfg.max_label_length_without_colon:
        return False
    
    if re.search(r'\d+', text) and not text[0].isdigit() and not text[-1].isdigit():
        return False
    
    if '@' in text and 'email' not in text_lower and 'correo' not in text_lower:
        return False
    
    if text_lower.startswith(('http', 'www')):
        return False
    
    if re.search(r'\d{3,}', text):
        return False
    
    return True

