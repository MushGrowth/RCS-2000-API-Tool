import json
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Callable
from urllib.parse import urlsplit


V3_CALLBACK_PATHS = {
    "/rcs/callback/task": "任务执行回调",
    "/rcs/callback/box": "料箱取放申请",
    "/rcs/callback/alarm": "告警推送",
    "/rcs/callback/storage": "申请回库仓位",
    "/rcs/callback/equipment": "外设交互",
}

V4_CALLBACK_PATHS = {
    "/api/robot/reporter/task": "任务执行过程回馈",
    "/api/robot/reporter/resource": "请求资源",
    "/api/robot/reporter/equipment": "请求外设",
    "/api/robot/reporter/homing": "机器人归巢完成",
    "/api/robot/reporter/banish": "机器人驱离完成",
    "/api/robot/reporter/bind": "绑定解绑通知",
    "/api/robot/reporter/ctu-check": "CTU 入库校验",
    "/api/robot/reporter/traffic": "交管区申请回调",
}


@dataclass
class ResponseProfile:
    status: int = 200
    delay_seconds: float = 0
    body: dict = field(default_factory=lambda: {"code": "0", "message": "成功"})


class CallbackServer:
    def __init__(self, host: str = "0.0.0.0", port: int = 8090, on_record: Callable | None = None):
        self.host = host
        self.port = port
        self.on_record = on_record
        self.profile = ResponseProfile()
        self.records: list[dict] = []
        self._server: ThreadingHTTPServer | None = None
        self._thread: threading.Thread | None = None

    @property
    def running(self) -> bool:
        return self._server is not None

    def start(self) -> None:
        if self.running:
            return
        owner = self

        class Handler(BaseHTTPRequestHandler):
            def do_POST(self):
                length = int(self.headers.get("Content-Length", "0"))
                raw = self.rfile.read(length)
                try:
                    payload = json.loads(raw.decode("utf-8")) if raw else {}
                except (UnicodeDecodeError, json.JSONDecodeError):
                    payload = {"_raw": raw.decode("utf-8", errors="replace")}
                task_id = next(
                    (payload.get(key) for key in ("robotTaskCode", "taskCode", "reqCode") if payload.get(key)),
                    "",
                )
                path = urlsplit(self.path).path
                record = {
                    "time": datetime.now().isoformat(timespec="seconds"),
                    "path": path,
                    "name": V3_CALLBACK_PATHS.get(path) or V4_CALLBACK_PATHS.get(path) or "自定义回调",
                    "task_id": task_id,
                    "headers": dict(self.headers),
                    "payload": payload,
                }
                owner.records.append(record)
                if owner.on_record:
                    owner.on_record(record)
                profile = owner.profile
                if profile.delay_seconds > 0:
                    time.sleep(profile.delay_seconds)
                body = json.dumps(profile.body, ensure_ascii=False).encode("utf-8")
                self.send_response(profile.status)
                self.send_header("Content-Type", "application/json;charset=UTF-8")
                request_id = self.headers.get("X-lr-request-id")
                if request_id:
                    self.send_header("X-lr-request-id", request_id)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, _format, *_args):
                return

        self._server = ThreadingHTTPServer((self.host, self.port), Handler)
        self.port = self._server.server_port
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        server, self._server = self._server, None
        if server:
            server.shutdown()
            server.server_close()
        if self._thread:
            self._thread.join(timeout=2)
            self._thread = None

