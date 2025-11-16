# server.py
"""Unified backend server.

Runs:
- FastAPI HTTP API on http://127.0.0.1:8000 (for FaceAuth)
- Existing WebSocket bridge on ws://127.0.0.1:8765
- Main assistant loop from main.py

Entry point: python server.py
"""

from __future__ import annotations

import logging
import os
import sys
import threading
from pathlib import Path
from typing import Tuple

import uvicorn
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
AUTH_DIR = os.path.join(BASE_DIR, "auth")
if AUTH_DIR not in sys.path:
    sys.path.append(AUTH_DIR)

from face_auth import FaceAuthSystem  # type: ignore  # noqa: E402
from main import (  # type: ignore  # noqa: E402
    main as assistant_main,
    push_face_auth_status,
    set_face_auth_granted,
    start_websocket_server,
)


logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

app = FastAPI(title="FYP Backend", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5174",
        "http://127.0.0.1:5174",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

FACE_DATASET_DIR = Path(AUTH_DIR) / "dataset"
FACE_MODEL_DIR = Path(AUTH_DIR) / "models"
face_auth = FaceAuthSystem(dataset_path=str(FACE_DATASET_DIR), model_path=str(FACE_MODEL_DIR))


def start_backend_threads() -> Tuple[threading.Thread, threading.Thread]:
    """Start WebSocket server and main assistant in background threads."""
    ws_thread = threading.Thread(target=start_websocket_server, daemon=True)
    ws_thread.start()

    assistant_thread = threading.Thread(target=assistant_main, kwargs={"start_ws": False}, daemon=True)
    assistant_thread.start()

    logging.info("Backend threads started (WebSocket + main loop).")
    return ws_thread, assistant_thread


@app.post("/api/face-auth")
async def face_auth_endpoint(image: UploadFile = File(...)):
    """Authenticate a single uploaded face image.

    Expects multipart/form-data with field name 'image'.
    Returns JSON: { authorized: bool, person: str, score: float, status: str, message: str }.
    """
    if image.content_type not in ("image/jpeg", "image/png"):
        raise HTTPException(status_code=400, detail="Image must be JPEG or PNG")

    data = await image.read()
    authorized, person, score, status, message = face_auth.authenticate_image(data)

    if authorized:
        set_face_auth_granted(person)
    else:
        push_face_auth_status(status, message, person if person != "Unknown" else None)

    if authorized:
        push_face_auth_status("granted", message or "Access granted", person)

    return {
        "authorized": bool(authorized),
        "person": person,
        "score": float(score),
        "status": status,
        "message": message,
    }


if __name__ == "__main__":
    start_backend_threads()
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")