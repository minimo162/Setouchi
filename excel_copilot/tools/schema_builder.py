# excel_copilot/tools/schema_builder.py

import inspect
import typing
from docstring_parser import parse
from ..core.exceptions import SchemaGenerationError

def _get_type_mapping(type_hint: typing.Any) -> dict:
    """Pythonの型ヒントをJSON Schemaの型定義に変換する"""
    if hasattr(type_hint, '__origin__'):
        origin = typing.get_origin(type_hint)
        args = typing.get_args(type_hint)

        if origin is list or origin is list:
            if args:
                # list[str] -> {"type": "array", "items": {"type": "string"}}
                return {"type": "array", "items": _get_type_mapping(args[0])}
            return {"type": "array", "items": {}} # list -> {"type": "array", "items": {}}

        if origin is dict or origin is dict:
            if args and len(args) == 2:
                # dict[str, int] -> {"type": "object", "additionalProperties": {"type": "integer"}}
                return {"type": "object", "additionalProperties": _get_type_mapping(args[1])}
            return {"type": "object"} # dict -> {"type": "object"}

        if origin is typing.Union:
            # Optional[str] は Union[str, None] と等価
            types = [arg for arg in args if arg is not type(None)]
            if len(types) == 1:
                # Optional[str] のようなケース
                return _get_type_mapping(types[0])
            # Union[str, int] のようなケース
            return {"oneOf": [_get_type_mapping(t) for t in types]}

    if type_hint is str:
        return {"type": "string"}
    if type_hint is int:
        return {"type": "integer"}
    if type_hint is float:
        return {"type": "number"}
    if type_hint is bool:
        return {"type": "boolean"}
    if type_hint is typing.Any:
        return {}  # Anyはスキーマ制限なし
    if type_hint is type(None):
        return {"type": "null"}

    # 不明な型はstringとして扱うか、エラーを出すか選択
    # ここでは警告を出しつつstringとして扱う
    # print(f"Warning: Unknown type hint '{type_hint}' treated as string.")
    return {"type": "string"}

def create_tool_schema(func: callable) -> dict:
    """
    関数オブジェクトから、AI (LLM) に渡すためのJSONスキーマを生成する。
    型ヒントとdocstringを解析し、関数の能力を構造化された形式で表現する。
    """
    try:
        sig = inspect.signature(func)
        docstring = parse(func.__doc__)

        param_descriptions = {param.arg_name: param.description for param in docstring.params}
        
        properties = {}
        required_params = []

        for name, param in sig.parameters.items():
            # 内部でのみ使用する引数 (例: actions, browser_manager) はスキーマに含めない
            if name in ["actions", "excel_manager", "browser_manager"]:
                continue

            if param.annotation is inspect.Parameter.empty:
                raise SchemaGenerationError(f"関数 '{func.__name__}' の引数 '{name}' に型ヒントがありません。")

            schema_type = _get_type_mapping(param.annotation)
            schema_type["description"] = param_descriptions.get(name, "")
            properties[name] = schema_type

            if param.default is inspect.Parameter.empty:
                required_params.append(name)

        return {
            "name": func.__name__,
            "description": docstring.short_description or f"関数 '{func.__name__}' の説明がありません。",
            "parameters": {
                "type": "object",
                "properties": properties,
                "required": required_params
            }
        }
    except Exception as e:
        # エラーにもとの例外をチェインすることで、デバッグしやすくなる
        raise SchemaGenerationError(f"関数 '{func.__name__}' のスキーマ生成に失敗しました。") from e
