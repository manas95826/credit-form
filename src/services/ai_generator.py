"""AI service for generating form field data."""
import json
from typing import Dict, List
from openai import OpenAI
from src.domain.exceptions import AIGenerationError
from src.config import config


class AIGenerator:
    """Service for generating form field data using AI."""
    
    def __init__(self, ai_config=None):
        self.config = ai_config or config.ai
        self.client = OpenAI(api_key=self.config.api_key)
    
    def _build_prompt(self, field_labels: List[str]) -> str:
        """Build the prompt for AI data generation."""
        field_list = ", ".join(field_labels)
        return f"""Generate realistic sample data for the following form fields: {field_list}

IMPORTANT: Return ONLY a complete, valid JSON object. The JSON must:
1. Have each field name as a key (exactly as listed above)
2. Have appropriate sample data as the value for each key
3. Be complete and properly closed (all brackets and braces must be closed)
4. Be valid JSON that can be parsed

Return ONLY the JSON object, no explanations, no markdown, no code blocks.

Example format:
{{"Full Name": "John Doe", "Address": "123 Main St, City, Country", "DOB": "01-15-1990", "Gender": "Male"}}
"""
    
    def _estimate_tokens(self, field_count: int) -> int:
        """Estimate required tokens based on field count."""
        return max(
            self.config.base_tokens,
            field_count * self.config.tokens_per_field + self.config.token_buffer
        )
    
    def _clean_response(self, response: str) -> str:
        """Clean AI response to extract JSON."""
        ai_response = response.strip()
        
        if ai_response.startswith("```"):
            parts = ai_response.split("```")
            for part in parts:
                part = part.strip()
                if part.startswith("json"):
                    part = part[4:].strip()
                if part.startswith("{") or part.startswith("["):
                    ai_response = part
                    break
            else:
                ai_response = next(
                    (p.strip() for p in parts if p.strip() and not p.strip().startswith("json")),
                    ai_response
                )
        
        ai_response = ai_response.strip()
        if not ai_response.endswith("}") and not ai_response.endswith("]"):
            open_braces = ai_response.count("{") - ai_response.count("}")
            open_brackets = ai_response.count("[") - ai_response.count("]")
            
            if open_braces > 0 or open_brackets > 0:
                ai_response += "}" * open_braces + "]" * open_brackets
        
        return ai_response
    
    def _call_ai(self, prompt: str, max_tokens: int) -> str:
        """Call AI API with retry logic."""
        messages = [
            {
                "role": "system",
                "content": (
                    "You are a helpful assistant that generates realistic sample data. "
                    "Always return complete, valid JSON objects. Never return incomplete or "
                    "truncated JSON. Ensure all field names from the user's list are included as keys."
                )
            },
            {"role": "user", "content": prompt}
        ]
        
        try:
            return self.client.chat.completions.create(
                model=self.config.model,
                messages=messages,
                temperature=self.config.temperature,
                max_tokens=max_tokens,
                response_format={"type": "json_object"}
            ).choices[0].message.content.strip()
        except Exception:
            return self.client.chat.completions.create(
                model=self.config.model,
                messages=messages,
                temperature=self.config.temperature,
                max_tokens=max_tokens
            ).choices[0].message.content.strip()
    
    def generate_data(self, field_labels: List[str]) -> Dict[str, str]:
        """Generate data for form fields using AI."""
        if not field_labels:
            return {}
        
        prompt = self._build_prompt(field_labels)
        max_tokens = self._estimate_tokens(len(field_labels))
        
        for attempt in range(self.config.max_retries):
            try:
                response = self._call_ai(prompt, max_tokens)
                cleaned_response = self._clean_response(response)
                data = json.loads(cleaned_response)
                return data
            except json.JSONDecodeError as e:
                if attempt < self.config.max_retries - 1:
                    max_tokens = int(max_tokens * 1.5)
                    continue
                raise AIGenerationError(
                    f"Failed to parse AI response as JSON after {self.config.max_retries} attempts: {e}"
                )
            except Exception as e:
                if attempt < self.config.max_retries - 1:
                    continue
                raise AIGenerationError(f"Error calling AI API: {e}")
        
        raise AIGenerationError("Failed to generate data after all retries")

