# -*- coding: utf-8 -*-
"""
飞书多维表格 读取层（新工程，独立于老脚本 generate_table.py）

只做三件事：
    1. 拿 tenant_access_token
    2. 按「表名」找到 table_id（也可以直接用环境变量指定）
    3. 分页拉全部记录，并把飞书那套嵌套字段展平成普通 dict

不依赖 requests，只用标准库 urllib。
"""
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request

API = "https://open.feishu.cn/open-apis"
PAGE_SIZE = 500          # 飞书单次最多 500 条

# 表名 → 环境变量名（用来覆盖 table_id）
TABLE_ENV = {
    "商品配置": "TBL_PRODUCTS",
    "运价表": "TBL_RATES",
    "澳洲分区": "TBL_AU",
    "全局参数": "TBL_PARAMS",
    "国家规则": "TBL_COUNTRY_RULES",
}


def _request(url, headers=None, body=None, method="GET", timeout=60):
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    if data is not None:
        req.add_header("Content-Type", "application/json; charset=utf-8")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "ignore")
        raise RuntimeError(f"HTTP {e.code} {url}\n{detail[:400]}") from None


class Feishu:
    """飞书多维表格客户端（tenant_access_token 自动续期）"""

    def __init__(self, app_id=None, app_secret=None, app_token=None):
        self.app_id = app_id or os.environ.get("FEISHU_APP_ID")
        self.app_secret = app_secret or os.environ.get("FEISHU_APP_SECRET")
        self.app_token = app_token or os.environ.get("FEISHU_APP_TOKEN")
        if not (self.app_id and self.app_secret):
            raise RuntimeError("缺少 FEISHU_APP_ID / FEISHU_APP_SECRET")
        if not self.app_token:
            raise RuntimeError("缺少 FEISHU_APP_TOKEN（多维表格地址里 /base/ 后面那串）")
        self._token = None
        self._exp = 0

    # ---------- token ----------
    @property
    def token(self):
        if self._token and time.time() < self._exp - 60:
            return self._token
        j = _request(f"{API}/auth/v3/tenant_access_token/internal",
                     body={"app_id": self.app_id, "app_secret": self.app_secret},
                     method="POST")
        if j.get("code") != 0:
            raise RuntimeError(f"取 token 失败：{j}")
        self._token = j["tenant_access_token"]
        self._exp = time.time() + j.get("expire", 7200)
        return self._token

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}"}

    # ---------- 表 ----------
    def list_tables(self):
        """返回 [(table_id, name), ...]"""
        out, page_token = [], None
        while True:
            q = {"page_size": 100}
            if page_token:
                q["page_token"] = page_token
            j = _request(f"{API}/bitable/v1/apps/{self.app_token}/tables?"
                         + urllib.parse.urlencode(q), self._headers())
            if j.get("code") != 0:
                raise RuntimeError(f"列数据表失败：{j}")
            data = j.get("data", {})
            out += [(t["table_id"], t.get("name", "")) for t in data.get("items", [])]
            if not data.get("has_more"):
                return out
            page_token = data.get("page_token")

    def table_ids(self):
        """表名 → table_id；环境变量优先（TBL_PRODUCTS 等）"""
        mapping = {}
        for tid, name in self.list_tables():
            mapping.setdefault(name.strip(), tid)
        for name, env in TABLE_ENV.items():
            v = os.environ.get(env)
            if v:
                mapping[name] = v
        return mapping

    # ---------- 记录 ----------
    def list_records(self, table_id, view_id=None):
        """分页拉全表，返回扁平化后的 dict 列表"""
        rows, page_token = [], None
        while True:
            q = {"page_size": PAGE_SIZE, "field_names": "[]"}
            q.pop("field_names")
            if view_id:
                q["view_id"] = view_id
            if page_token:
                q["page_token"] = page_token
            url = (f"{API}/bitable/v1/apps/{self.app_token}/tables/{table_id}/records?"
                   + urllib.parse.urlencode(q))
            j = _request(url, self._headers())
            if j.get("code") != 0:
                raise RuntimeError(f"拉记录失败 table={table_id}：{j}")
            data = j.get("data", {})
            for item in data.get("items", []):
                rows.append(flatten(item.get("fields", {})))
            if not data.get("has_more"):
                return rows
            page_token = data.get("page_token")


# ============================================================================
# 飞书字段 → 普通值
# ============================================================================
def flatten(fields):
    """
    飞书字段值有多种形态，统一成 Python 基本类型：
        文本      → str
        数字      → float / int
        单选      → str（取 name）
        多选/人员 → ["a", "b"]
        日期      → 毫秒时间戳转 "YYYY-MM-DD"
        公式/查找 → 递归取 value
        附件      → ["url", ...]
    """
    out = {}
    for k, v in fields.items():
        key = str(k).strip()
        val = _norm(v)
        if isinstance(val, (int, float)) and key.endswith(("日期", "时间")) and val > 10 ** 11:
            import datetime
            val = datetime.datetime.fromtimestamp(val / 1000).strftime("%Y-%m-%d")
        out[key] = val
    return out


def _norm(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        return v.strip() or None
    if isinstance(v, list):
        vals = [_norm(x) for x in v]
        vals = [x for x in vals if x is not None]
        if len(vals) == 1:
            return vals[0]
        return vals
    if isinstance(v, dict):
        for key in ("text", "name", "value", "link", "url", "file_token"):
            if key in v:
                return _norm(v[key])
        if "values" in v:
            return _norm(v["values"])
        return {k: _norm(x) for k, x in v.items()}
    return str(v)


def fetch_all(feishu=None):
    """一次拉齐五张表，返回 {表名: [记录...]}"""
    fs = feishu or Feishu()
    ids = fs.table_ids()
    need = ["商品配置", "运价表", "澳洲分区", "全局参数", "国家规则"]
    missing = [n for n in need if n not in ids]
    if missing:
        raise RuntimeError(
            f"多维表格里找不到这些数据表：{missing}\n"
            f"现有：{sorted(ids)}\n"
            f"（表名必须完全一致，或用 TBL_PRODUCTS 等环境变量指定 table_id）")
    return {n: fs.list_records(ids[n]) for n in need}


if __name__ == "__main__":
    data = fetch_all()
    for name, rows in data.items():
        print(f"{name}: {len(rows)} 行")
        if rows:
            print("   首行：", json.dumps(rows[0], ensure_ascii=False)[:200])
