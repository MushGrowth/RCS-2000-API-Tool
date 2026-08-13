import hashlib
import hmac
import json
import secrets
import uuid
from datetime import datetime
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import requests


def iso_timestamp() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def compact_json(payload: dict) -> str:
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def build_v4_headers(
    request_id: str,
    app_key: str = "",
    source: str = "",
    api_version: str = "v1.0",
    trace_id: str = "",
    algorithm: str = "HMAC-SHA256",
    nonce: str = "",
    timestamp: str = "",
) -> dict[str, str]:
    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "X-lr-request-id": request_id,
    }
    if not app_key:
        return headers
    nonce = nonce or secrets.token_hex(4)
    timestamp = timestamp or iso_timestamp()
    headers.update(
        {
            "Authorization": f'nonce="{nonce}",method="{algorithm}",timestamp="{timestamp}"',
            "X-lr-appkey": app_key,
            "X-lr-version": api_version,
            "X-lr-trace-id": trace_id or uuid.uuid4().hex,
        }
    )
    if source:
        headers["X-lr-source"] = source
    return headers


def signing_text(method: str, url: str, headers: dict[str, str], body: str) -> str:
    parsed = urlsplit(url)
    path = parsed.path or "/"
    if parsed.query:
        path += f"?{parsed.query}"
    normalized = {key.upper(): value for key, value in headers.items()}
    lines = [f"{method.upper()} {path} HTTP/1.1"]
    for name in (
        "AUTHORIZATION",
        "HOST",
        "X-LR-APPKEY",
        "X-LR-REQUEST-ID",
        "X-LR-SOURCE",
        "X-LR-TRACE-ID",
        "X-LR-VERSION",
    ):
        value = parsed.netloc if name == "HOST" else normalized.get(name)
        if value:
            lines.append(f"{name}: {value}")
    return "\n".join(lines) + "\n\n" + body


def sign_v4_request(
    method: str, url: str, headers: dict[str, str], body: str, app_secret: str
) -> tuple[str, str]:
    algorithm = "HMAC-SHA512" if "HMAC-SHA512" in headers.get("Authorization", "") else "HMAC-SHA256"
    digest = hashlib.sha512 if algorithm == "HMAC-SHA512" else hashlib.sha256
    raw = signing_text(method, url, headers, body)
    salted = hmac.new(app_secret.encode(), raw.encode(), digest).hexdigest()
    # 协议示例使用 MD5 32 位结果的中间 16 位作为最终 sign。
    signature = hashlib.md5(salted.encode()).hexdigest()[8:24]
    parsed = urlsplit(url)
    query = parse_qsl(parsed.query, keep_blank_values=True)
    query.append(("sign", signature))
    signed_url = urlunsplit((parsed.scheme, parsed.netloc, parsed.path, urlencode(query), parsed.fragment))
    return signed_url, raw


def post_json(
    url: str,
    payload: dict,
    headers: dict | None = None,
    connect_timeout: float = 30,
    read_timeout: float = 60,
    verify_tls: bool = True,
) -> requests.Response:
    body = compact_json(payload)
    return requests.post(
        url,
        data=body.encode("utf-8"),
        headers=headers or {"Content-Type": "application/json;charset=UTF-8"},
        timeout=(connect_timeout, read_timeout),
        verify=verify_tls,
    )
