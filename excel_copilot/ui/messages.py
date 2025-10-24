"""Messaging primitives shared by the desktop UI and worker thread."""

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Any, Dict, Optional, Union


class AppState(Enum):
    """High-level lifecycle states for the desktop application."""

    INITIALIZING = auto()
    READY = auto()
    TASK_IN_PROGRESS = auto()
    STOPPING = auto()
    ERROR = auto()


class RequestType(str, Enum):
    """Types of requests the UI can post to the worker thread."""

    USER_INPUT = "USER_INPUT"
    STOP = "STOP"
    QUIT = "QUIT"
    UPDATE_CONTEXT = "UPDATE_CONTEXT"
    RESET_BROWSER = "RESET_BROWSER"


class ResponseType(str, Enum):
    """Categories of responses emitted by the worker thread."""

    STATUS = "status"
    ERROR = "error"
    INFO = "info"
    END_OF_TASK = "end_of_task"
    INITIALIZATION_COMPLETE = "initialization_complete"
    THOUGHT = "thought"
    ACTION = "action"
    OBSERVATION = "observation"
    FINAL_ANSWER = "final_answer"
    CHAT_PROMPT = "chat_prompt"
    CHAT_RESPONSE = "chat_response"
    SHUTDOWN_COMPLETE = "shutdown_complete"


@dataclass(frozen=True)
class RequestMessage:
    """Serializable request envelope that flows from UI to worker."""

    type: RequestType
    payload: Optional[Any] = None

    @classmethod
    def from_raw(cls, raw: Union["RequestMessage", Dict[str, Any]]) -> "RequestMessage":
        """Coerce dictionaries or message instances into the canonical form."""
        if isinstance(raw, cls):
            return raw
        if not isinstance(raw, dict):
            raise ValueError(f"Unsupported request payload type: {type(raw)}")

        raw_type = raw.get("type")
        if isinstance(raw_type, RequestType):
            request_type = raw_type
        else:
            try:
                request_type = RequestType(str(raw_type))
            except ValueError as exc:  # pragma: no cover - defensive guard
                raise ValueError(f"Unsupported request type: {raw_type}") from exc

        return cls(type=request_type, payload=raw.get("payload"))


@dataclass(frozen=True)
class ResponseMessage:
    """Serializable response envelope that flows from worker to UI."""

    type: ResponseType
    content: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_raw(cls, raw: Union["ResponseMessage", Dict[str, Any]]) -> "ResponseMessage":
        """Coerce dictionaries or message instances into the canonical form."""
        if isinstance(raw, cls):
            return raw
        if not isinstance(raw, dict):
            raise ValueError(f"Unsupported response payload type: {type(raw)}")

        raw_type = raw.get("type")
        if isinstance(raw_type, ResponseType):
            response_type = raw_type
        else:
            try:
                response_type = ResponseType(str(raw_type))
            except ValueError:
                response_type = ResponseType.INFO

        content = raw.get("content", "")
        metadata = {k: v for k, v in raw.items() if k not in {"type", "content"}}
        if raw_type and (not isinstance(raw_type, ResponseType)) and raw_type != response_type.value:
            metadata.setdefault("source_type", raw_type)
        return cls(type=response_type, content=content, metadata=metadata)


__all__ = [
    "AppState",
    "RequestType",
    "ResponseType",
    "RequestMessage",
    "ResponseMessage",
]
