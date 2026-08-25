from fastapi import FastAPI, Request
from datetime import datetime
import json

app = FastAPI(
    title="Lynk.id Automation Bot",
    version="1.0.0"
)


# =========================
# HOME / ROOT
# =========================
@app.get("/")
async def root():
    return {
        "status": "online",
        "service": "Lynk.id Automation Bot"
    }


# =========================
# HEALTH CHECK
# =========================
@app.get("/health")
async def health():
    return {
        "status": "healthy",
        "time": datetime.now().isoformat()
    }


# =========================
# LYNK.ID WEBHOOK
# =========================
@app.post("/webhook/lynk")
async def lynk_webhook(request: Request):

    try:
        # Membaca data yang dikirim oleh Lynk.id
        body = await request.body()

        print("\n========== LYNK WEBHOOK ==========")
        print(body.decode("utf-8"))
        print("==================================\n")

        # Coba membaca data sebagai JSON
        try:
            data = json.loads(body)
        except json.JSONDecodeError:
            data = {
                "raw_data": body.decode("utf-8")
            }

        print("Data webhook:")
        print(data)

        # Untuk sementara hanya menerima data.
        # Database dan Telegram akan kita tambahkan
        # setelah format webhook Lynk.id diketahui.

        return {
            "success": True,
            "message": "Webhook received"
        }

    except Exception as e:

        print("Webhook error:", str(e))

        return {
            "success": False,
            "message": str(e)
        }