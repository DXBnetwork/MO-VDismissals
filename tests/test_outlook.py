# tests/test_outlook.py
import pytest
from unittest.mock import AsyncMock, MagicMock, patch
from outlook import case_number, notify_n8n, process_matching_email, renew_subscription


def test_case_number_extracts_from_standard_subject():
    subject = "New filing in Hamed et al v. Rubio et al: NOTICE of Voluntary Dismissal by (vaed-1:2025-cv-02484)"
    assert case_number(subject) == "vaed-1:2025-cv-02484"


def test_case_number_returns_none_when_missing():
    subject = "New filing: some other document"
    assert case_number(subject) is None


def test_case_number_returns_none_for_empty_or_none():
    assert case_number("") is None
    assert case_number(None) is None


@pytest.mark.asyncio
async def test_process_matching_email_ignores_non_dismissal():
    email = MagicMock()
    email.subject = "ORDER re: Motion to Compel (vaed-1:2025-cv-00111)"
    email.has_attachments = True

    with patch("outlook.get_email", new=AsyncMock(return_value=email)):
        result = await process_matching_email("msg-123")

    assert result == {"status": "ignored", "message_id": "msg-123"}

@pytest.mark.asyncio
async def test_process_matching_email_ignores_email_without_attachments():
    email = MagicMock()
    email.subject = "NOTICE of Voluntary Dismissal by (vaed-1:2025-cv-02484)"
    email.has_attachments = False

    with patch("outlook.get_email", new=AsyncMock(return_value=email)):
        result = await process_matching_email("msg-123")

    assert result == {"status": "ignored", "message_id": "msg-123"}


@pytest.mark.asyncio
async def test_process_matching_email_returns_missing_case_number():
    email = MagicMock()
    email.subject = "NOTICE of Voluntary Dismissal by no parens here"
    email.has_attachments = True

    with patch("outlook.get_email", new=AsyncMock(return_value=email)):
        result = await process_matching_email("msg-123")

    assert result == {"status": "missing_case_number", "message_id": "msg-123"}


@pytest.mark.asyncio
async def test_process_matching_email_returns_missing_attachment():
    email = MagicMock()
    email.subject = "NOTICE of Voluntary Dismissal by (vaed-1:2025-cv-02484)"
    email.has_attachments = True

    with patch("outlook.get_email", new=AsyncMock(return_value=email)), \
         patch("outlook.download_email_attachment", new=AsyncMock(return_value=(None, None))):
        result = await process_matching_email("msg-123")

    assert result == {"status": "missing_attachment", "message_id": "msg-123"}


@pytest.mark.asyncio
async def test_process_matching_email_returns_folder_not_found():
    email = MagicMock()
    email.subject = "NOTICE of Voluntary Dismissal by (vaed-1:2025-cv-02484)"
    email.has_attachments = True

    with patch("outlook.get_email", new=AsyncMock(return_value=email)), \
         patch("outlook.download_email_attachment", new=AsyncMock(return_value=(b"data", "file.pdf"))), \
         patch("outlook.get_access_token", new=AsyncMock(return_value="token")), \
         patch("outlook.search_folder", new=AsyncMock(return_value=None)):
        result = await process_matching_email("msg-123")

    assert result == {
        "status": "folder_not_found",
        "message_id": "msg-123",
        "case_number": "vaed-1:2025-cv-02484",
    }


@pytest.mark.asyncio
async def test_process_matching_email_uploads_successfully():
    email = MagicMock()
    email.subject = "NOTICE of Voluntary Dismissal by (vaed-1:2025-cv-02484)"
    email.has_attachments = True

    folder_item = {"id": "folder-456", "name": "vaed-1:2025-cv-02484"}
    upload_result = {"id": "upload-789"}

    with patch("outlook.get_email", new=AsyncMock(return_value=email)), \
         patch("outlook.download_email_attachment", new=AsyncMock(return_value=(b"data", "dismissal.pdf"))), \
         patch("outlook.get_access_token", new=AsyncMock(return_value="token")), \
         patch("outlook.search_folder", new=AsyncMock(return_value=folder_item)), \
         patch("outlook.upload_file_to_sharepoint", new=AsyncMock(return_value=upload_result)):
        result = await process_matching_email("msg-123")

    assert result == {
        "status": "uploaded",
        "message_id": "msg-123",
        "case_number": "vaed-1:2025-cv-02484",
        "upload_id": "upload-789",
        "destination_name": "vaed-1:2025-cv-02484",
    }


@pytest.mark.asyncio
async def test_renew_subscription_patches_expiration():
    renewed = {"id": "sub-abc", "expirationDateTime": "2026-05-03T00:00:00Z"}

    mock_response = MagicMock()
    mock_response.json.return_value = renewed

    mock_client = AsyncMock()
    mock_client.__aenter__.return_value.patch = AsyncMock(return_value=mock_response)

    with patch("outlook.get_access_token", new=AsyncMock(return_value="token")), \
         patch("httpx.AsyncClient", return_value=mock_client):
        result = await renew_subscription("sub-abc")

    assert result == renewed
    patch_call = mock_client.__aenter__.return_value.patch
    called_url = patch_call.call_args[0][0]
    assert "sub-abc" in called_url
    body = patch_call.call_args[1]["json"]
    assert "expirationDateTime" in body


@pytest.mark.asyncio
async def test_notify_n8n_posts_correct_payload():
    mock_response = MagicMock()
    mock_client = AsyncMock()
    mock_client.__aenter__.return_value.post = AsyncMock(return_value=mock_response)

    with patch("outlook.N8N_WEBHOOK_URL", "https://n8n.example.com/webhook/test"), \
         patch("httpx.AsyncClient", return_value=mock_client):
        await notify_n8n("91342164756", "dismissal.pdf", "https://sharepoint.example.com/file")

    post_call = mock_client.__aenter__.return_value.post
    assert post_call.called
    payload = post_call.call_args[1]["json"]
    assert payload["matter_number"] == "91342164756"
    assert payload["file_name"] == "dismissal.pdf"
    assert payload["sharepoint_url"] == "https://sharepoint.example.com/file"


@pytest.mark.asyncio
async def test_process_matching_email_notifies_n8n_after_upload():
    email = MagicMock()
    email.subject = "NOTICE of Voluntary Dismissal by (vaed-1:2025-cv-02484)"
    email.has_attachments = True

    folder_item = {"id": "folder-456", "name": "Gabibli, Aitan (91342164756)", "parentReference": {}}
    upload_result = {"id": "upload-789", "webUrl": "https://sharepoint.example.com/file"}

    mock_notify = AsyncMock()

    with patch("outlook.get_email", new=AsyncMock(return_value=email)), \
         patch("outlook.download_email_attachment", new=AsyncMock(return_value=(b"data", "dismissal.pdf"))), \
         patch("outlook.get_access_token", new=AsyncMock(return_value="token")), \
         patch("outlook.search_folder", new=AsyncMock(return_value=folder_item)), \
         patch("outlook.upload_file_to_sharepoint", new=AsyncMock(return_value=upload_result)), \
         patch("outlook.notify_n8n", mock_notify):
        result = await process_matching_email("msg-123")

    assert result["status"] == "uploaded"
    mock_notify.assert_awaited_once_with(
        matter_number="91342164756",
        file_name="dismissal.pdf",
        sharepoint_url="https://sharepoint.example.com/file",
    )


@pytest.mark.asyncio
async def test_process_matching_email_does_not_notify_n8n_when_folder_not_found():
    email = MagicMock()
    email.subject = "NOTICE of Voluntary Dismissal by (vaed-1:2025-cv-02484)"
    email.has_attachments = True

    mock_notify = AsyncMock()

    with patch("outlook.get_email", new=AsyncMock(return_value=email)), \
         patch("outlook.download_email_attachment", new=AsyncMock(return_value=(b"data", "dismissal.pdf"))), \
         patch("outlook.get_access_token", new=AsyncMock(return_value="token")), \
         patch("outlook.search_folder", new=AsyncMock(return_value=None)), \
         patch("outlook.notify_n8n", mock_notify):
        result = await process_matching_email("msg-123")

    assert result["status"] == "folder_not_found"
    mock_notify.assert_not_awaited()
