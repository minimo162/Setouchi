#!/usr/bin/env python3
"""Bridge script that runs CopilotWorker once and streams responses as JSON."""

import argparse
import json
import queue
import sys
import threading
from typing import Any, Dict

from excel_copilot.ui.worker import CopilotWorker
from excel_copilot.ui.messages import RequestMessage, RequestType, ResponseMessage, ResponseType


def _emit(event: Dict[str, Any]) -> None:
    json.dump(event, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    sys.stdout.flush()


def _response_to_dict(message: ResponseMessage) -> Dict[str, Any]:
    return {
        "event": "response",
        "type": message.type.value,
        "content": message.content,
        "metadata": message.metadata,
    }


def run_worker(envelope: Dict[str, Any]) -> int:
    worker_payload = envelope.get("worker_payload")
    if not isinstance(worker_payload, dict):
        _emit({"event": "error", "message": "worker_payload missing"})
        return 1
    request_q: "queue.Queue[Any]" = queue.Queue()
    response_q: "queue.Queue[Any]" = queue.Queue()

    worker = CopilotWorker(
        request_q,
        response_q,
        sheet_name=envelope.get("sheet_name"),
        workbook_name=envelope.get("workbook_name"),
    )
    thread = threading.Thread(target=worker.run, name="CopilotWorkerThread", daemon=True)
    thread.start()

    request_q.put(RequestMessage(RequestType.USER_INPUT, worker_payload))

    stop_types = {ResponseType.END_OF_TASK, ResponseType.ERROR, ResponseType.FINAL_ANSWER}
    try:
        while True:
            try:
                message = response_q.get(timeout=300)
            except queue.Empty:
                _emit({"event": "timeout", "message": "response queue empty"})
                break
            if isinstance(message, ResponseMessage):
                _emit(_response_to_dict(message))
                if message.type in stop_types:
                    break
            else:
                _emit({"event": "unknown", "payload": str(message)})
    finally:
        request_q.put(RequestMessage(RequestType.QUIT))
        thread.join(timeout=5)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Copilot worker bridge")
    parser.add_argument("command", choices=["run"], help="Action to execute")
    args = parser.parse_args()
    data = json.load(sys.stdin)
    if args.command == "run":
        return run_worker(data)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
