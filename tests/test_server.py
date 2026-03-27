import json

import pytest

from teams_mcp.server import _extract_adaptive_card_text, _extract_attachments_text, _format_message


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


def test_extract_attachments_text_adaptive_card():
    attachments = [
        {
            "id": "att-1",
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": json.dumps({
                "type": "AdaptiveCard",
                "body": [{"type": "TextBlock", "text": "Task DWH-383 created"}],
            }),
        }
    ]
    assert _extract_attachments_text(attachments) == "Task DWH-383 created"


def test_extract_attachments_text_non_adaptive_skipped():
    attachments = [
        {
            "id": "att-1",
            "contentType": "application/vnd.microsoft.card.hero",
            "content": json.dumps({"title": "Hero card"}),
        }
    ]
    assert _extract_attachments_text(attachments) == ""


def test_extract_attachments_text_invalid_json():
    attachments = [
        {
            "id": "att-1",
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": "not valid json {{{",
        }
    ]
    assert _extract_attachments_text(attachments) == ""


def test_extract_attachments_text_multiple_cards():
    attachments = [
        {
            "id": "att-1",
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": json.dumps({
                "type": "AdaptiveCard",
                "body": [{"type": "TextBlock", "text": "Card one"}],
            }),
        },
        {
            "id": "att-2",
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": json.dumps({
                "type": "AdaptiveCard",
                "body": [{"type": "TextBlock", "text": "Card two"}],
            }),
        },
    ]
    assert _extract_attachments_text(attachments) == "Card one\nCard two"


def test_format_message_with_adaptive_card():
    msg = {
        "id": "msg-1",
        "from": {"user": {"displayName": "DWH Bot"}},
        "createdDateTime": "2026-03-26T15:21:29Z",
        "body": {"content": "", "contentType": "html"},
        "attachments": [
            {
                "id": "att-1",
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": json.dumps({
                    "type": "AdaptiveCard",
                    "body": [
                        {"type": "TextBlock", "text": "Task DWH-383 created"},
                        {
                            "type": "FactSet",
                            "facts": [
                                {"title": "Type", "value": "Airflow Access"},
                                {"title": "Status", "value": "Open"},
                            ],
                        },
                    ],
                }),
            }
        ],
    }
    result = _format_message(msg)
    assert result["content"] == "Task DWH-383 created\nType: Airflow Access\nStatus: Open"
    assert result["sender"] == "DWH Bot"


def test_format_message_body_and_card_combined():
    msg = {
        "id": "msg-2",
        "from": {"user": {"displayName": "Alice"}},
        "createdDateTime": "2026-03-26T10:00:00Z",
        "body": {"content": "<p>Check this out</p>", "contentType": "html"},
        "attachments": [
            {
                "id": "att-1",
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": json.dumps({
                    "type": "AdaptiveCard",
                    "body": [{"type": "TextBlock", "text": "Card content"}],
                }),
            }
        ],
    }
    result = _format_message(msg)
    assert result["content"] == "Check this out\nCard content"


def test_format_message_no_attachments_unchanged():
    msg = {
        "id": "msg-3",
        "from": {"user": {"displayName": "Bob"}},
        "createdDateTime": "2026-03-26T10:00:00Z",
        "body": {"content": "Plain message", "contentType": "text"},
    }
    result = _format_message(msg)
    assert result["content"] == "Plain message"
