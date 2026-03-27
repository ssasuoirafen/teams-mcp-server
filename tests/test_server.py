import json

import pytest

from teams_mcp.server import _extract_adaptive_card_text


def test_textblock():
    card = {
        "type": "AdaptiveCard",
        "body": [
            {"type": "TextBlock", "text": "Hello world"},
            {"type": "TextBlock", "text": "Second line"},
        ],
    }
    assert _extract_adaptive_card_text(card) == "Hello world\nSecond line"


def test_factset():
    card = {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "FactSet",
                "facts": [
                    {"title": "Status", "value": "Open"},
                    {"title": "Priority", "value": "High"},
                ],
            }
        ],
    }
    assert _extract_adaptive_card_text(card) == "Status: Open\nPriority: High"


def test_richtextblock():
    card = {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "RichTextBlock",
                "inlines": [
                    {"type": "TextRun", "text": "Bold text"},
                    {"type": "TextRun", "text": " and more"},
                ],
            }
        ],
    }
    assert _extract_adaptive_card_text(card) == "Bold text and more"


def test_empty_card():
    assert _extract_adaptive_card_text({"type": "AdaptiveCard", "body": []}) == ""
    assert _extract_adaptive_card_text({}) == ""


def test_unknown_elements_skipped():
    card = {
        "type": "AdaptiveCard",
        "body": [
            {"type": "Input.Text", "id": "name"},
            {"type": "TextBlock", "text": "Visible"},
        ],
    }
    assert _extract_adaptive_card_text(card) == "Visible"
