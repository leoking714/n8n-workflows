# -*- coding: utf-8 -*-
"""
路径解析：代码在 /scripts（只读），数据/输出要写 /data（可写）。

本地开发时（代码目录本身可写）自动用 工程目录/data 和 工程目录/输出；
在 n8n python-runner 里（/scripts 只读）自动切到 /data 下。

可用环境变量强制指定：
    BILL_DATA  配置 JSON 目录
    BILL_OUT   计算表输出目录
"""
import os

BASE = os.path.dirname(os.path.abspath(__file__))

# n8n 里 /data 挂的是宿主 /Users/zuo/n8n/files
RUN_DATA = "/data/data/workflow003/60103飞书/data"
RUN_OUT = "/data/data/workflow003/60103飞书/输出"


def _writable(path):
    """目录能写就返回 True（不存在则看父级）"""
    p = path
    while p and not os.path.exists(p):
        p = os.path.dirname(p)
    return os.access(p or "/", os.W_OK)


def data_dir():
    env = os.environ.get("BILL_DATA")
    if env:
        os.makedirs(env, exist_ok=True)
        return env
    local = os.path.join(BASE, "data")
    if _writable(local):
        os.makedirs(local, exist_ok=True)
        return local
    os.makedirs(RUN_DATA, exist_ok=True)
    return RUN_DATA


def out_dir():
    env = os.environ.get("BILL_OUT")
    if env:
        os.makedirs(env, exist_ok=True)
        return env
    local = os.path.join(BASE, "输出")
    if _writable(local):
        os.makedirs(local, exist_ok=True)
        return local
    os.makedirs(RUN_OUT, exist_ok=True)
    return RUN_OUT


def template():
    """计算表模板：优先工程目录，其次 /data 下的备份"""
    env = os.environ.get("BILL_TEMPLATE")
    if env:
        return env
    for p in (os.path.join(BASE, "模板", "计算表模板.xlsx"),
              "/data/data/workflow003/60103飞书/模板/计算表模板.xlsx"):
        if os.path.exists(p):
            return p
    raise FileNotFoundError("找不到计算表模板 模板/计算表模板.xlsx")
