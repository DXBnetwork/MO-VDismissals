import httpx


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


async def search_folder(access_token: str, query_string: str):
    if not query_string:
        return None

    payload = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {"queryString": query_string},
            }
        ]
    }
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE_URL}/search/query",
            headers=headers,
            json=payload,
        )
        response.raise_for_status()
        data = response.json()

    for search_request in data.get("value", []):
        for hit_container in search_request.get("hitsContainers", []):
            for hit in hit_container.get("hits", []):
                resource = hit.get("resource", {})
                parent_reference = resource.get("parentReference", {})
                if resource.get("folder") and parent_reference.get("driveId") and resource.get("id"):
                    return resource
    return None


async def upload_file_to_sharepoint(
    access_token: str,
    folder_item: dict,
    file_name: str,
    file_bytes: bytes,
):
    if not folder_item:
        raise ValueError("A destination folder result is required for upload.")

    drive_id = folder_item.get("parentReference", {}).get("driveId")
    item_id = folder_item.get("id")
    if not drive_id or not item_id:
        raise ValueError("Search result does not include the drive/item identifiers needed for upload.")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/octet-stream",
    }
    upload_url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}:/{file_name}:/content"

    async with httpx.AsyncClient() as client:
        response = await client.put(upload_url, headers=headers, content=file_bytes)
        response.raise_for_status()
        return response.json()
