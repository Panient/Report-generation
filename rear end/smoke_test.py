from __future__ import annotations

import json
import urllib.request


BASE = "http://127.0.0.1:8000"


def request(method: str, path: str, data=None, token: str | None = None, raw: bool = False):
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = "Bearer " + token
    body = json.dumps(data, ensure_ascii=False).encode("utf-8") if data is not None else None
    req = urllib.request.Request(BASE + path, data=body, method=method, headers=headers)
    with urllib.request.urlopen(req, timeout=30) as resp:
        content = resp.read()
        if raw:
            return resp.status, resp.headers, content
        obj = json.loads(content.decode("utf-8"))
        if obj["code"] != 0:
            raise RuntimeError(obj)
        return obj["data"]


def main() -> None:
    token = request("POST", "/api/auth/login", {"username": "student", "password": "123456"})["accessToken"]
    report = request(
        "POST",
        "/api/reports",
        {
            "reportName": "测试电厂迎峰度夏检查报告",
            "reportType": "summerCheck",
            "topic": "迎峰度夏安全检查",
            "major": "电气",
            "plant": "测试电厂",
            "year": 2026,
        },
        token,
    )
    report_id = report["reportId"]
    outline = request(
        "POST",
        f"/api/reports/{report_id}/outline/generate",
        {"reportType": "summerCheck", "topic": "迎峰度夏安全检查", "templateId": "tpl_001"},
        token,
    )
    request(
        "POST",
        f"/api/reports/{report_id}/content/generate",
        {"chapterIds": [], "regenerate": False, "forceOverwrite": False},
        token,
        raw=True,
    )
    detail = request("GET", f"/api/reports/{report_id}", token=token)
    export = request("POST", f"/api/reports/{report_id}/exports", {"fileFormat": "docx"}, token)
    request("GET", f"/api/reports/{report_id}/exports/{export['exportId']}/download", token=token, raw=True)
    admin = request("POST", "/api/auth/login", {"username": "admin", "password": "admin123"})["accessToken"]
    request("GET", "/api/templates?page=1&pageSize=10", token=admin)
    request("GET", "/api/admin/model-config", token=admin)
    super_token = request("POST", "/api/auth/login", {"username": "super", "password": "super123"})["accessToken"]
    request("GET", "/api/admin/users?page=1&pageSize=10", token=super_token)
    print("OK", report_id, len(outline["outline"]), len(detail["contents"]), export["fileSize"])


if __name__ == "__main__":
    main()
