import json

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


def test_extract_attachments_text_content_as_dict():
    """Graph API sometimes returns content as pre-parsed dict."""
    attachments = [
        {
            "id": "att-1",
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "type": "AdaptiveCard",
                "body": [{"type": "TextBlock", "text": "Already parsed"}],
            },
        }
    ]
    assert _extract_attachments_text(attachments) == "Already parsed"


def test_action_showcard():
    card = {
        "type": "AdaptiveCard",
        "body": [{"type": "TextBlock", "text": "Main text"}],
        "actions": [
            {
                "type": "Action.ShowCard",
                "title": "Show details",
                "card": {
                    "type": "AdaptiveCard",
                    "body": [
                        {"type": "TextBlock", "text": "Hidden detail"},
                        {"type": "FactSet", "facts": [{"title": "Key", "value": "Val"}]},
                    ],
                },
            }
        ],
    }
    result = _extract_adaptive_card_text(card)
    assert result == "Main text\nShow details\nHidden detail\nKey: Val"


# --- Forwarded message tests ---


def test_forwarded_message_real_api_format():
    """Forwarded message with real Graph API field names (originalMessage*)."""
    attachments = [
        {
            "id": "1727881360458",
            "contentType": "forwardedMessageReference",
            "content": json.dumps({
                "originalMessageId": "1727881360458",
                "originalMessageContent": "\n<p>Here are the credentials</p>\n",
                "originalConversationId": "19:abc123@thread.v2",
                "originalSentDateTime": "2026-04-01T10:00:00+00:00",
                "originalMessageSender": {
                    "user": {
                        "userIdentityType": "aadUser",
                        "id": "user-id-123",
                        "displayName": "Alice",
                    }
                },
            }),
        }
    ]
    result = _extract_attachments_text(attachments)
    assert "[Forwarded from Alice]" in result
    assert "Here are the credentials" in result


def test_forwarded_message_dict_content():
    """Forwarded message with pre-parsed dict content."""
    attachments = [
        {
            "id": "ref-2",
            "contentType": "forwardedMessageReference",
            "content": {
                "originalMessageSender": {"user": {"displayName": "Bob"}},
                "originalMessageContent": "<p>Check this out</p>",
            },
        }
    ]
    result = _extract_attachments_text(attachments)
    assert "[Forwarded from Bob]" in result
    assert "Check this out" in result


def test_forwarded_message_null_displayname():
    """Forwarded message where sender displayName is null (Graph API quirk)."""
    attachments = [
        {
            "id": "ref-3",
            "contentType": "forwardedMessageReference",
            "content": json.dumps({
                "originalMessageSender": {"user": {"displayName": None, "id": "uid"}},
                "originalMessageContent": "<p>Important info</p>",
            }),
        }
    ]
    result = _extract_attachments_text(attachments)
    assert "Important info" in result
    assert "[Forwarded from" not in result


def test_quoted_reply_messageref_fallback():
    """messageReference (quoted reply) uses messageSender/messagePreview fields."""
    attachments = [
        {
            "id": "ref-4",
            "contentType": "messageReference",
            "content": json.dumps({
                "messageSender": {"user": {"displayName": "Carol"}},
                "messagePreview": "Original quoted text",
            }),
        }
    ]
    result = _extract_attachments_text(attachments)
    assert "[Forwarded from Carol]" in result
    assert "Original quoted text" in result


def test_forwarded_message_empty_content():
    """Forwarded message with no content returns empty."""
    attachments = [
        {
            "id": "ref-5",
            "contentType": "forwardedMessageReference",
            "content": "",
        }
    ]
    assert _extract_attachments_text(attachments) == ""


def test_forwarded_message_not_in_format_attachments():
    """Forwarded messages should be excluded from _format_attachments."""
    from teams_mcp.server import _format_attachments

    attachments = [
        {"id": "ref-6", "contentType": "forwardedMessageReference", "content": "{}"},
        {"id": "file-1", "contentType": "application/pdf", "name": "doc.pdf"},
    ]
    result = _format_attachments(attachments)
    assert len(result) == 1
    assert result[0]["name"] == "doc.pdf"


def test_format_message_with_forwarded():
    """Full _format_message pipeline includes forwarded content."""
    msg = {
        "id": "msg-1",
        "from": {"user": {"displayName": "Dave"}},
        "createdDateTime": "2026-04-06T08:00:00Z",
        "body": {"content": "", "contentType": "html"},
        "attachments": [
            {
                "id": "ref-7",
                "contentType": "forwardedMessageReference",
                "content": json.dumps({
                    "originalMessageSender": {"user": {"displayName": "Eve"}},
                    "originalMessageContent": "<p>Secret data</p>",
                }),
            }
        ],
    }
    result = _format_message(msg)
    assert "Eve" in result["content"]
    assert "Secret data" in result["content"]


def test_forwarded_message_body_fallback():
    """Falls back to body field when no originalMessageContent or messagePreview."""
    attachments = [
        {
            "id": "ref-8",
            "contentType": "messageReference",
            "content": json.dumps({
                "messageSender": {"user": {"displayName": "Frank"}},
                "body": {"content": "<p>Body <b>fallback</b></p>"},
            }),
        }
    ]
    result = _extract_attachments_text(attachments)
    assert "[Forwarded from Frank]" in result
    assert "Body fallback" in result
