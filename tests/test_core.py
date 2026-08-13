import json
import unittest
import urllib.request

from callback_server import CallbackServer, ResponseProfile
from interface_catalog import load_interfaces
from rcs_http import build_v4_headers, compact_json, sign_v4_request, signing_text


class CoreTests(unittest.TestCase):
    def test_catalog_counts_and_new_fields(self):
        v3 = load_interfaces("RCS 3.x")
        v4 = load_interfaces("RCS 4.x")
        self.assertEqual(len(v3), 23)
        self.assertEqual(len(v4), 23)
        submit = next(item for item in v4 if item["suffix"].endswith("/task/submit"))
        self.assertIn(("carrierSpeedScale", "选填-载具速度比例"), submit["fields"])
        self.assertTrue(any(item["suffix"] == "wcs/api/outer/rest/trafficTask" for item in v4))

    def test_v4_signature_is_stable_and_sign_is_last(self):
        url = "https://example.test:443/api/test"
        headers = build_v4_headers(
            "1234567890abcdef",
            "app-key",
            "wms",
            trace_id="0123456789abcdef0123456789abcdef",
            nonce="12345678",
            timestamp="2026-07-03T12:00:00+08:00",
        )
        body = compact_json({"taskCode": "T001"})
        raw = signing_text("POST", url, headers, body)
        signed_url, returned_raw = sign_v4_request("POST", url, headers, body, "secret")
        self.assertEqual(raw, returned_raw)
        self.assertEqual(signed_url.count("?sign="), 1)
        self.assertEqual(signed_url.split("sign=", 1)[1], "e9c58318e509c783")

    def test_callback_server_records_and_responds(self):
        server = CallbackServer("127.0.0.1", 0)
        server.profile = ResponseProfile(body={"code": "0", "message": "ok"})
        server.start()
        try:
            request = urllib.request.Request(
                f"http://127.0.0.1:{server.port}/api/robot/reporter/task",
                data=json.dumps({"robotTaskCode": "T001"}).encode(),
                headers={"Content-Type": "application/json", "X-lr-request-id": "req-1"},
                method="POST",
            )
            with urllib.request.urlopen(request, timeout=2) as response:
                self.assertEqual(json.loads(response.read())["code"], "0")
                self.assertEqual(response.headers["X-lr-request-id"], "req-1")
            self.assertEqual(server.records[0]["task_id"], "T001")
        finally:
            server.stop()


if __name__ == "__main__":
    unittest.main()
