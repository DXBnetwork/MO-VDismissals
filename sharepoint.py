from urllib.parse import quote

import httpx


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
SHAREPOINT_SITE_ID = "69ddd908-8766-42d0-a501-367166fe5883"


def is_in_configured_site(resource: dict):
    parent_reference = resource.get("parentReference", {})
    site_id = parent_reference.get("siteId", "")
    return (
        parent_reference.get("driveId")
        and resource.get("id")
        and SHAREPOINT_SITE_ID.lower() in site_id.lower()
    )


async def search_folder(access_token: str, query_string: str):
    if not query_string:
        return None

    search_query = query_string.replace("-", " ").replace(":", " ")
    payload = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {"queryString": search_query},
                "from": 0,
                "size": 25,
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
                if not is_in_configured_site(resource):
                    continue
                if resource.get("folder"):
                    continue
                parent_ref = resource.get("parentReference", {})
                drive_id = parent_ref.get("driveId")
                folder_id = parent_ref.get("id")
                path = parent_ref.get("path", "")
                if not drive_id or not folder_id:
                    continue
                # path looks like: /drives/xxx/root:/A/Lastname, Firstname (12345)/District Court
                after_root = path.split("root:")[-1] if "root:" in path else path
                segments = [s for s in after_root.split("/") if s]
                matter_name = next(
                    (s for s in reversed(segments) if "(" in s and ")" in s),
                    "",
                )
                return {
                    "id": folder_id,
                    "name": matter_name,
                    "parentReference": {"driveId": drive_id},
                }
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
    upload_url = f"{GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}:/{quote(file_name, safe='')}:/content"

    async with httpx.AsyncClient() as client:
        response = await client.put(upload_url, headers=headers, content=file_bytes)
        response.raise_for_status()
        return response.json()
