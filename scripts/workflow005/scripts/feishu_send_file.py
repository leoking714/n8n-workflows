#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ROLE: 把文件（xlsx/pdf/png…）上传到飞书并发送到群聊
ENTRY: python3 feishu_send_file.py --path <文件> [--chat oc_xxx] [--name 自定义文件名]
CALLED-BY: workflow005（发计算表）/ workflow007（发 PDF）
DEPS: feishu_newbase.json（飞书凭据）

流程：im/v1/files 上传拿 file_key → im/v1/messages 发 file 消息
"""
import argparse
import json
import mimetypes
import os
import sys
import urllib.request
import uuid

CRED = "/data/data/workflow005/config/feishu_newbase.json"
DEFAULT_CHAT = "oc_76b442ae3f9e004ad15fc88783e24bba"


def token_of(cfg):
    req = urllib.request.Request(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        data=json.dumps({"app_id": cfg["app_id"], "app_secret": cfg["app_secret"]}).encode(),
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode())["tenant_access_token"]


def file_type_of(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return "pdf"
    if ext in (".xls", ".xlsx"):
        return "xls"
    if ext in (".doc", ".docx"):
        return "doc"
    if ext in (".ppt", ".pptx"):
        return "ppt"
    if ext == ".mp4":
        return "mp4"
    return "stream"


def upload(token, path, name):
    """上传文件到飞书 im/v1/files：优先 requests（multipart 可靠），无则用 urllib 兜底"""
    fname = name or os.path.basename(path)
    try:
        import requests
        with open(path, "rb") as f:
            r = requests.post(
                "https://open.feishu.cn/open-apis/im/v1/files",
                headers={"Authorization": "Bearer " + token},
                data={"file_type": file_type_of(path), "file_name": fname},
                files={"file": (fname, f, mimetypes.guess_type(fname)[0] or "application/octet-stream")},
                timeout=120)
        return r.json()
    except ImportError:
        pass
    boundary = "----dsh" + uuid.uuid4().hex
    with open(path, "rb") as f:
        content = f.read()
    parts = []

    def field(k, v):
        parts.append(("--" + boundary + "\r\n").encode())
        parts.append((f'Content-Disposition: form-data; name="{k}"\r\n\r\n').encode())
        parts.append(v.encode() + b"\r\n")

    field("file_type", file_type_of(path))
    field("file_name", fname)
    parts.append(("--" + boundary + "\r\n").encode())
    parts.append((f'Content-Disposition: form-data; name="file"; filename="{fname}"\r\n').encode())
    parts.append((mimetypes.guess_type(fname)[0] or "application/octet-stream").encode() + b"\r\n\r\n")
    parts.append(content + b"\r\n")
    parts.append(("--" + boundary + "--\r\n").encode())
    body = b"".join(parts)
    req = urllib.request.Request(
        "https://open.feishu.cn/open-apis/im/v1/files", data=body,
        headers={"Authorization": "Bearer " + token,
                 "Content-Type": "multipart/form-data; boundary=" + boundary})
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read().decode())


def send_file_message(token, chat, file_key):
    body = json.dumps({"receive_id": chat, "msg_type": "file",
                       "content": json.dumps({"file_key": file_key})}).encode()
    req = urllib.request.Request(
        "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
        data=body, headers={"Authorization": "Bearer " + token,
                            "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--path", required=True)
    ap.add_argument("--chat", default=DEFAULT_CHAT)
    ap.add_argument("--name", default="")
    a = ap.parse_args()
    if not os.path.exists(a.path):
        raise SystemExit("文件不存在: " + a.path)
    cfg = json.load(open(CRED, encoding="utf-8"))
    token = token_of(cfg)
    up = upload(token, a.path, a.name)
    if up.get("code") != 0:
        print(json.dumps({"ok": False, "step": "upload", "code": up.get("code"),
                          "msg": up.get("msg")}, ensure_ascii=False))
        return 1
    fk = (up.get("data") or {}).get("file_key")
    msg = send_file_message(token, a.chat, fk)
    ok = msg.get("code") == 0
    print(json.dumps({"ok": ok, "file": os.path.basename(a.path),
                      "size": os.path.getsize(a.path), "file_key": fk,
                      "message_id": (msg.get("data") or {}).get("message_id"),
                      "code": msg.get("code"), "msg": msg.get("msg")}, ensure_ascii=False))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
