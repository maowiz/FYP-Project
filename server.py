# server.py
import asyncio
import websockets
import json
import logging

logging.basicConfig(level=logging.INFO)

async def handler(websocket, path):
    logging.info("Frontend connected.")
    try:
        # Keep the connection alive and listen for messages
        async for message in websocket:
            data = json.loads(message)
            logging.info(f"Received from UI: {data}")

            # Example of echoing a message back
            if data.get("action") == "request_state":
                response = {
                    "type": "state",
                    "payload": {"phase": "listening"}
                }
                await websocket.send(json.dumps(response))
    except websockets.exceptions.ConnectionClosed:
        logging.info("Frontend disconnected.")
    finally:
        logging.info("Connection handler finished.")

async def main():
    # Start the WebSocket server on localhost, port 8765
    async with websockets.serve(handler, "localhost", 8765):
        logging.info("WebSocket server started on ws://localhost:8765")
        await asyncio.Future()  # run forever

if __name__ == "__main__":
    asyncio.run(main())