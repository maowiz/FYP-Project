# FYP-Project-main/server.py
import asyncio
import logging
import json
import queue
import threading
import uvicorn
import cv2
import numpy as np
import os
import sys
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import websockets

# Ensure auth directory is in path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
AUTH_DIR = os.path.join(BASE_DIR, "auth")
if AUTH_DIR not in sys.path:
    sys.path.append(AUTH_DIR)

# Import existing logic
from auth.face_auth import FaceAuthSystem
from main import (
    ui_message_queue, 
    python_command_queue, 
    websocket_handler, 
    send_ui_updates,
    main as run_assistant_logic,
    set_face_auth_granted,
    push_face_auth_status
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
log = logging.getLogger(__name__)

app = FastAPI()

# Setup Face Auth
FACE_DATASET_DIR = Path(AUTH_DIR) / "dataset"
FACE_MODEL_DIR = Path(AUTH_DIR) / "models"
face_auth = FaceAuthSystem(dataset_path=str(FACE_DATASET_DIR), model_path=str(FACE_MODEL_DIR))

# CORS (Still good to have)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- API ROUTES ---
@app.post("/api/face-auth")
async def api_face_auth(image: UploadFile = File(...)):
    try:
        image_data = await image.read()
        # Use the new authentication method
        authorized, person, score, status, message = face_auth.authenticate_image(image_data)
        
        if authorized:
            log.info(f"API: Face auth SUCCESS for {person}")
            set_face_auth_granted(person) # Notify main.py
            return {"authorized": True, "person": person, "score": score}
        else:
            log.warning(f"API: Face auth FAILED. Status: {status}")
            return {"authorized": False, "person": "Unknown", "score": score}
            
    except Exception as e:
        log.error(f"API: Face auth error: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail="Error processing image")

# --- SERVE FRONTEND ---
# This serves the React app from the 'dist' folder
if os.path.exists("dist"):
    # 1. Serve the assets folder (JS/CSS)
    app.mount("/assets", StaticFiles(directory="dist/assets"), name="assets")
    
    # 2. Serve the audio folder if it exists
    if os.path.exists("dist/audio"):
        app.mount("/audio", StaticFiles(directory="dist/audio"), name="audio")

    # 3. Catch-all route for the main HTML file
    @app.get("/{catchall:path}")
    async def read_index(catchall: str):
        # Check if file exists (e.g. favicon.ico), otherwise serve index.html
        file_path = os.path.join("dist", catchall)
        if os.path.exists(file_path) and os.path.isfile(file_path):
            return FileResponse(file_path)
        return FileResponse("dist/index.html")
else:
    log.warning("⚠️ 'dist' folder not found! Frontend will not load.")

# --- SERVER STARTUP ---
async def start_websocket_server():
    server = await websockets.serve(websocket_handler, "127.0.0.1", 8765)
    log.info("WebSocket server started on ws://127.0.0.1:8765")
    asyncio.create_task(send_ui_updates())
    await server.wait_closed()

@app.on_event("startup")
async def startup_event():
    log.info("Starting background tasks...")
    ws_loop = asyncio.new_event_loop()
    threading.Thread(target=lambda: (asyncio.set_event_loop(ws_loop), ws_loop.run_until_complete(start_websocket_server())), daemon=True).start()
    
    # Pass start_ws=False so main.py doesn't try to start another WS server
    threading.Thread(target=run_assistant_logic, kwargs={"start_ws": False}, daemon=True).start()
    log.info("Background tasks started.")

if __name__ == "__main__":
    log.info("Starting Unified Server on http://127.0.0.1:8000")
    uvicorn.run(app, host="127.0.0.1", port=8000)