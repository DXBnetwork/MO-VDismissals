import base64
import os
import re
from datetime import datetime, timedelta, timezone

import httpx
from azure.identity import ClientSecretCredential
from dotenv import load_dotenv
from fastapi import FastAPI, Query, Request
from fastapi.responses import PlainTextResponse
from msgraph import GraphServiceClient

from sharepoint import search_folder, upload_file_to_sharepoint

load_dotenv()

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
MAILBOX_USER_ID = os.getenv("mailbox_user_id", "litigationfilings@murrayosorio.com")
WATCH_FOLDER_NAME = os.getenv("watch_folder_name", "DocketBird")
SUBSCRIPTION_SECRET = os.getenv("subscription_secret")
NOTIFICATION_URL = os.getenv("notification_url")

app = FastAPI()
 
credential = ClientSecretCredential(
    os.getenv("tenant_id"),
    os.getenv("client_id"),
    os.getenv("client_secret"),
)
graph_client = GraphServiceClient(credential)


async def get_folder(user_id: str, folder_name: str):
    folders = await graph_client.users.by_user_id(user_id).mail_folders.get()
    for folder in folders.value:
        if folder.display_name == folder_name:
            return folder.id
    return None


async def get_email(user_id: str, message_id: str):
    return await graph_client.users.by_user_id(user_id).messages.by_message_id(message_id).get()


async def download_email_attachment(user_id: str, message_id: str):
    attachments = await (
        graph_client.users.by_user_id(user_id)
        .messages.by_message_id(message_id)
        .attachments.get()
    )
    for attachment in attachments.value:
        content = await (
            graph_client.users.by_user_id(user_id)
            .messages.by_message_id(message_id)
            .attachments.by_attachment_id(attachment.id)
            .get()
        )
        if getattr(content, "content_bytes", None):
            return base64.b64decode(content.content_bytes), attachment.name
    return None, None


def case_number(subject: str):
    match = re.search(r"\(([^)]+)\)", subject or "")
    return match.group(1) if match else None


async def get_access_token():
    token = credential.get_token(GRAPH_SCOPE)
    return token.token


async def process_matching_email(message_id: str):
    email = await get_email(MAILBOX_USER_ID, message_id)
    if not email or "Voluntary Dismissal" not in (email.subject or "") or not email.has_attachments:
        return {"status": "ignored", "message_id": message_id}

    matter_number = case_number(email.subject)
    if not matter_number:
        return {"status": "missing_case_number", "message_id": message_id}

    file_bytes, file_name = await download_email_attachment(MAILBOX_USER_ID, message_id)
    if not file_bytes or not file_name:
        return {"status": "missing_attachment", "message_id": message_id}

    access_token = await get_access_token()
    folder_item = await search_folder(access_token, matter_number)
    if not folder_item:
        return {
            "status": "folder_not_found",
            "message_id": message_id,
            "case_number": matter_number,
        }

    upload_result = await upload_file_to_sharepoint(
        access_token=access_token,
        folder_item=folder_item,
        file_name=file_name,
        file_bytes=file_bytes,
    )
    return {
        "status": "uploaded",
        "message_id": message_id,
        "case_number": matter_number,
        "upload_id": upload_result.get("id"),
        "destination_name": folder_item.get("name"),
    }


@app.post("/webhook")
async def webhook_handler(request: Request, validationToken: str = Query(None)):
    if validationToken:
        return PlainTextResponse(validationToken)

    payload = await request.json()
    for notification in payload.get("value", []):
        resource = notification.get("resource", "")
        if not resource:
            continue
        message_id = resource.split("/")[-1]
        result = await process_matching_email(message_id)
        print(result)

    return PlainTextResponse("OK", status_code=202)


async def create_subscription():
    if not SUBSCRIPTION_SECRET:
        raise RuntimeError("Missing required env var: subscription_secret")
    if not NOTIFICATION_URL:
        raise RuntimeError("Missing required env var: notification_url")

    access_token = await get_access_token()
    folder_id = await get_folder(MAILBOX_USER_ID, WATCH_FOLDER_NAME)
    if not folder_id:
        raise RuntimeError(f"Mail folder '{WATCH_FOLDER_NAME}' was not found for {MAILBOX_USER_ID}.")

    expiration = (datetime.now(timezone.utc) + timedelta(minutes=4230)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )
    payload = {
        "changeType": "created",
        "notificationUrl": NOTIFICATION_URL,
        "resource": f"users/{MAILBOX_USER_ID}/mailFolders/{folder_id}/messages",
        "expirationDateTime": expiration,
        "clientState": SUBSCRIPTION_SECRET,
    }
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    async with httpx.AsyncClient() as client:
        response = await client.post(
            "https://graph.microsoft.com/v1.0/subscriptions",
            headers=headers,
            json=payload,
        )
        response.raise_for_status()
        return response.json()
