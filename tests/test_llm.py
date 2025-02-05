import pytest
from llm_processor import validate_response

SAMPLE_RESPONSE = """
{
    "action": "create",
    "project": {
        "name": "Projet Test",
        "timeline": {
            "start": "2025-01-01",
            "end": "2025-03-15"
        },
        "appearance": {
            "color": {
                "input": "bleu",
                "rgb": [0,0,255]
            }
        }
    }
}
"""

def test_validation():
    result = validate_response(SAMPLE_RESPONSE)
    assert result['action'] == 'create'
    assert result['project']['name'] == 'Projet Test'


def test_invalid_response():
    with pytest.raises(Exception):
        validate_response("{invalid_json}")
