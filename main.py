from contextlib import asynccontextmanager

from fastapi import FastAPI

from outlook import create_subscription, webhook_handler


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
