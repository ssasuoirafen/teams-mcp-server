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


def test_container_nested():
    card = {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "Container",
                "items": [
                    {"type": "TextBlock", "text": "Inside container"},
                    {
                        "type": "Container",
                        "items": [
                            {"type": "TextBlock", "text": "Deeply nested"},
                        ],
                    },
                ],
            }
        ],
    }
    assert _extract_adaptive_card_text(card) == "Inside container\nDeeply nested"


def test_columnset():
    card = {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "ColumnSet",
                "columns": [
                    {
                        "type": "Column",
                        "items": [{"type": "TextBlock", "text": "Col A"}],
                    },
                    {
                        "type": "Column",
                        "items": [{"type": "TextBlock", "text": "Col B"}],
                    },
                ],
            }
        ],
    }
    assert _extract_adaptive_card_text(card) == "Col A\nCol B"


def test_table():
    card = {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "Table",
                "rows": [
                    {
                        "type": "TableRow",
                        "cells": [
                            {
                                "type": "TableCell",
                                "items": [{"type": "TextBlock", "text": "R1C1"}],
                            },
                            {
                                "type": "TableCell",
                                "items": [{"type": "TextBlock", "text": "R1C2"}],
                            },
                        ],
                    }
                ],
            }
        ],
    }
    assert _extract_adaptive_card_text(card) == "R1C1\nR1C2"


def test_actions():
    card = {
        "type": "AdaptiveCard",
        "body": [{"type": "TextBlock", "text": "Click below"}],
        "actions": [
            {"type": "Action.OpenUrl", "title": "Open Jira", "url": "https://jira.example.com/DWH-383"},
            {"type": "Action.Submit", "title": "Approve"},
        ],
    }
    result = _extract_adaptive_card_text(card)
    assert result == "Click below\nOpen Jira (https://jira.example.com/DWH-383)\nApprove"
