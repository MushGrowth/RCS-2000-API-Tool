import json
import sys
from pathlib import Path


VERSION_FILES = {
    "RCS 3.x": "rcs-3.3.json",
    "RCS 4.x": "rcs-4.x.json",
}


def resource_path(relative_path: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative_path


def load_interfaces(version: str) -> list[dict]:
    file_name = VERSION_FILES.get(version)
    if not file_name:
        raise ValueError(f"不支持的 RCS 版本：{version}")
    path = resource_path(f"interfaces/{file_name}")
    with path.open("r", encoding="utf-8") as stream:
        document = json.load(stream)
    interfaces = document.get("interfaces")
    if not isinstance(interfaces, list):
        raise ValueError(f"接口配置格式错误：{path}")
    result = []
    for item in interfaces:
        converted = dict(item)
        converted["fields"] = [
            (field["name"], field.get("requirement", "选填"))
            for field in item.get("fields", [])
        ]
        result.append(converted)
    return result

