import asyncio
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse

from outlook import create_subscription, subscription_renewal_loop, webhook_handler


@asynccontextmanager
async def lifespan(app: FastAPI):
    renewal_task = None
    try:
        subscription = await create_subscription()
        sub_id = subscription.get("id")
        print({"subscription_status": "created", "subscription_id": sub_id})
        if sub_id:
            renewal_task = asyncio.create_task(subscription_renewal_loop(sub_id))
    except Exception as exc:
        import traceback
        print({"subscription_status": "failed", "error": str(exc), "type": type(exc).__name__, "trace": traceback.format_exc()})

    yield

    if renewal_task:
        renewal_task.cancel()
        try:
            await renewal_task
        except asyncio.CancelledError:
            pass


app = FastAPI(lifespan=lifespan)

app.add_api_route("/webhook", webhook_handler, methods=["POST"])


async def webhook_validate(request: Request):
    token = request.query_params.get("validationToken")
    if token:
        return PlainTextResponse(token)
    return PlainTextResponse("ok")


app.add_api_route("/webhook", webhook_validate, methods=["GET"])
