import pathlib
import sys
import unittest

sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))

from excel_copilot.ui.messages import (
    RequestMessage,
    RequestType,
    ResponseMessage,
    ResponseType,
)


class RequestMessageFromRawTests(unittest.TestCase):
    def test_passthrough_when_instance_given(self) -> None:
        message = RequestMessage(type=RequestType.USER_INPUT, payload={"text": "hello"})

        result = RequestMessage.from_raw(message)

        self.assertIs(result, message)

    def test_string_type_dict_is_normalized(self) -> None:
        result = RequestMessage.from_raw({"type": "STOP", "payload": {"reason": "user"}})

        self.assertEqual(result.type, RequestType.STOP)
        self.assertEqual(result.payload, {"reason": "user"})

    def test_unsupported_type_raises_value_error(self) -> None:
        with self.assertRaises(ValueError):
            RequestMessage.from_raw({"type": "UNKNOWN"})

    def test_non_dict_payload_raises_value_error(self) -> None:
        with self.assertRaises(ValueError):
            RequestMessage.from_raw(123)  # type: ignore[arg-type]


class ResponseMessageFromRawTests(unittest.TestCase):
    def test_passthrough_when_instance_given(self) -> None:
        message = ResponseMessage(type=ResponseType.STATUS, content="ok")

        result = ResponseMessage.from_raw(message)

        self.assertIs(result, message)

    def test_string_type_and_extra_metadata_are_preserved(self) -> None:
        result = ResponseMessage.from_raw(
            {
                "type": "error",
                "content": "failed",
                "payload": {"message": "boom"},
                "attempt": 3,
            }
        )

        self.assertEqual(result.type, ResponseType.ERROR)
        self.assertEqual(result.content, "failed")
        self.assertEqual(result.metadata, {"payload": {"message": "boom"}, "attempt": 3})

    def test_invalid_type_defaults_to_info_and_marks_source(self) -> None:
        result = ResponseMessage.from_raw({"type": "mystery", "content": "???"})

        self.assertEqual(result.type, ResponseType.INFO)
        self.assertEqual(result.metadata, {"source_type": "mystery"})

    def test_non_dict_payload_raises_value_error(self) -> None:
        with self.assertRaises(ValueError):
            ResponseMessage.from_raw("status")  # type: ignore[arg-type]


if __name__ == "__main__":
    unittest.main()
