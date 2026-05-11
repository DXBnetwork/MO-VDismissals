import asyncio
import base64
import hmac
import logging
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

logger = logging.getLogger(__name__)

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
MAILBOX_USER_ID = os.getenv("mailbox_user_id", "litigationfilings@murrayosorio.com")
WATCH_FOLDER_NAME = os.getenv("watch_folder_name", "Test Cases")
SUBSCRIPTION_SECRET = os.getenv("subscription_secret")
NOTIFICATION_URL = os.getenv("notification_url")
N8N_WEBHOOK_URL = os.getenv("n8n_webhook_url")

app = FastAPI()
 
credential = ClientSecretCredential(
    os.getenv("tenant_id"),
    os.getenv("client_id"),
    os.getenv("client_secret"),
)
graph_client = GraphServiceClient(credential)


async def get_folder(user_id: str, folder_name: str):
    access_token = await get_access_token()
    headers = {"Authorization": f"Bearer {access_token}"}
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"https://graph.microsoft.com/v1.0/users/{user_id}/mailFolders",
            headers=headers,
            params={"$top": "100"},
        )
        response.raise_for_status()
        data = response.json()
    for folder in data.get("value", []):
        if folder["displayName"] == folder_name:
            return folder["id"]
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


async def notify_n8n(matter_number: str, file_name: str, sharepoint_url: str):
    if not N8N_WEBHOOK_URL:
        logger.warning("n8n_webhook_url is not set — skipping notification.")
        return
    payload = {
        "matter_number": matter_number,
        "file_name": file_name,
        "sharepoint_url": sharepoint_url,
    }
    async with httpx.AsyncClient() as client:
        response = await client.post(N8N_WEBHOOK_URL, json=payload)
        response.raise_for_status()


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

    folder_name = folder_item.get("name", "")
    salesforce_matter_number = case_number(folder_name)
    try:
        await notify_n8n(
            matter_number=salesforce_matter_number,
            file_name=file_name,
            sharepoint_url=upload_result.get("webUrl", ""),
        )
    except Exception:
        logger.exception("Failed to notify n8n for message %s.", message_id)

    return {
        "status": "uploaded",
        "message_id": message_id,
        "case_number": matter_number,
        "upload_id": upload_result.get("id"),
        "destination_name": folder_name,
    }


@app.post("/webhook")
async def webhook_handler(request: Request, validationToken: str = Query(None)):
    if validationToken:
        return PlainTextResponse(validationToken)

    payload = await request.json()
    notifications = payload.get("value", [])

    for notification in notifications:
        client_state = notification.get("clientState")
        if not SUBSCRIPTION_SECRET or not client_state:
            logger.warning("Rejected webhook notification with missing clientState.")
            continue
        if not hmac.compare_digest(client_state, SUBSCRIPTION_SECRET):
            logger.warning("Rejected webhook notification with invalid clientState.")
            continue
        resource = notification.get("resource", "")
        if not resource:
            logger.warning("Skipped webhook notification with missing resource.")
            continue
        message_id = resource.split("/")[-1]
        try:
            result = await process_matching_email(message_id)
            logger.info("Processed webhook notification: %s", result)
        except Exception:
            logger.exception("Failed to process webhook notification for message %s.", message_id)

    return PlainTextResponse("OK", status_code=202)


async def renew_subscription(subscription_id: str):
    access_token = await get_access_token()
    expiration = (datetime.now(timezone.utc) + timedelta(minutes=4230)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    async with httpx.AsyncClient() as client:
        response = await client.patch(
            f"https://graph.microsoft.com/v1.0/subscriptions/{subscription_id}",
            headers=headers,
            json={"expirationDateTime": expiration},
        )
        response.raise_for_status()
        return response.json()


async def subscription_renewal_loop(subscription_id: str):
    while True:
        await asyncio.sleep(24 * 60 * 60)
        try:
            await renew_subscription(subscription_id)
            logger.info("Subscription %s renewed.", subscription_id)
        except Exception:
            logger.exception("Failed to renew subscription %s.", subscription_id)


async def get_existing_subscription(access_token: str, resource: str):
    headers = {"Authorization": f"Bearer {access_token}"}
    async with httpx.AsyncClient() as client:
        response = await client.get(
            "https://graph.microsoft.com/v1.0/subscriptions",
            headers=headers,
        )
        response.raise_for_status()
        data = response.json()
    for sub in data.get("value", []):
        if sub.get("resource") == resource:
            return sub
    return None


async def create_subscription():
    if not SUBSCRIPTION_SECRET:
        raise RuntimeError("Missing required env var: subscription_secret")
    if not NOTIFICATION_URL:
        raise RuntimeError("Missing required env var: notification_url")

    access_token = await get_access_token()
    folder_id = await get_folder(MAILBOX_USER_ID, WATCH_FOLDER_NAME)
    if not folder_id:
        raise RuntimeError(f"Mail folder '{WATCH_FOLDER_NAME}' was not found for {MAILBOX_USER_ID}.")
    
    resource = f"users/{MAILBOX_USER_ID}/mailFolders/{folder_id}/messages"
    existing = await get_existing_subscription(access_token, resource)
    if existing:
        logger.info("Reusing existing subscription %s.", existing.get("id"))
        return existing

    expiration = (datetime.now(timezone.utc) + timedelta(minutes=4230)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )
    payload = {
        "changeType": "created",
        "notificationUrl": NOTIFICATION_URL,
        "resource": resource,
        "expirationDateTime": expiration,
        "clientState": SUBSCRIPTION_SECRET,
    }
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    
    async with httpx.AsyncClient(timeout=30.0) as client:
        response = await client.post(
            "https://graph.microsoft.com/v1.0/subscriptions",
            headers=headers,
            json=payload,
        )
        response.raise_for_status()
        return response.json()
