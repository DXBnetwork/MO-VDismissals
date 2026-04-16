from contextlib import asynccontextmanager

from fastapi import FastAPI, Request

from outlook import create_subscription, webhook_handler

from fastapi.responses import PlainTextResponse


@asynccontextmanager
async def lifespan(app: FastAPI):
    try:
        subscription = await create_subscription()
        print({"subscription_status": "created", "subscription_id": subscription.get("id")})
    except Exception as exc:
        print(
            {
                "subscription_status": "failed",
                "error": str(exc),
            }
        )
    yield


app = FastAPI(lifespan=lifespan)
app.add_api_route("/webhook", webhook_handler, methods=["POST"])


async def webhook_validate(request: Request):
    token=request.query_params.get("validationToken")
    if token:
        return PlainTextResponse(token)
    return PlainTextResponse("ok")
app.add_api_route("/webhook", webhook_validate, methods=["GET"])
