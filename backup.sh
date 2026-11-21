#!/bin/bash
# 全量备份脚本：n8n 已发布工作流 + 脚本 + 模板 + 配置快照 → git 提交 → 推送 GitHub
set -euo pipefail

REPO="$(cd "$(dirname "$0")" && pwd)"
DATE="$(date '+%Y-%m-%d %H:%M:%S')"
SRC_SCRIPTS="/Users/zuo/n8n/files/automation/scripts"
SRC_CONFIG="/Users/zuo/n8n/files/data/workflow005/config"
SRC_LOGS="/Users/zuo/n8n/files/data/workflow005/logs"

say() { printf '[%s] %s\n' "$1" "$2"; }

# ---------------------------------------------------------------- 1. 工作流
say 1/6 "导出 n8n 已发布工作流"
docker exec n8n-n8n-1 n8n export:workflow --all --published --output=/tmp/n8n_all.json >/dev/null
docker cp n8n-n8n-1:/tmp/n8n_all.json "$REPO/.export.json" >/dev/null

/usr/bin/python3 - "$REPO" <<'PY'
import json, os, re, sys
repo = sys.argv[1]
data = json.load(open(os.path.join(repo, ".export.json"), encoding="utf-8"))
outdir = os.path.join(repo, "workflows")
os.makedirs(outdir, exist_ok=True)
for f in os.listdir(outdir):
    if f.endswith(".json"):
        os.remove(os.path.join(outdir, f))

SECRET_PATTERNS = [
    (re.compile(r'(\\?"app_secret\\?"\s*:\s*\\?")([^"\\]+)'), r'\1***REDACTED***'),
    (re.compile(r'(\\?"appSecret\\?"\s*:\s*\\?")([^"\\]+)'), r'\1***REDACTED***'),
    (re.compile(r'(Bearer\s+)[A-Za-z0-9._\-]{12,}'), r'\1***REDACTED***'),
    (re.compile(r'(\\?"(?:password|passwd|token|api[_-]?key)\\?"\s*:\s*\\?")([^"\\]{6,})', re.I), r'\1***REDACTED***'),
]

def slug(name, wid):
    s = re.sub(r'[^0-9A-Za-z\u4e00-\u9fa5]+', '-', name).strip('-')[:60]
    return f"{s or 'workflow'}-{wid[:8]}"

index = []
for w in data:
    wid, name = w.get("id", ""), w.get("name", "")
    text = json.dumps(json.loads(json.dumps(w, ensure_ascii=False)), ensure_ascii=False, indent=1)
    redacted = 0
    for pat, rep in SECRET_PATTERNS:
        text, n = pat.subn(rep, text); redacted += n
    fn = slug(name, wid) + ".json"
    open(os.path.join(outdir, fn), "w", encoding="utf-8").write(text)
    index.append({"id": wid, "name": name, "active": bool(w.get("active")),
                  "nodes": len(w.get("nodes", [])), "file": fn, "redacted": redacted})
index.sort(key=lambda x: (not x["active"], x["name"]))
json.dump(index, open(os.path.join(repo, "workflows.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print(f"      工作流 {len(index)} 个")
PY

# ---------------------------------------------------------------- 2. 脚本
say 2/6 "同步自动化脚本"
mkdir -p "$REPO/scripts"
rsync -a --delete \
  --exclude '.DS_Store' --exclude '__pycache__' --exclude '*.pyc' \
  --exclude '*.log' --exclude '*.tmp' \
  "$SRC_SCRIPTS/" "$REPO/scripts/" 2>/dev/null || cp -R "$SRC_SCRIPTS/." "$REPO/scripts/"

# ---------------------------------------------------------------- 3. 模板
say 3/6 "同步模板文件"
mkdir -p "$REPO/templates"
for f in 计算表模板.xlsx bill_template.docx bill_header.png 计算表模板_旧_20261120.xlsx; do
  [ -f "$SRC_CONFIG/$f" ] && cp -f "$SRC_CONFIG/$f" "$REPO/templates/"
done

# ---------------------------------------------------------------- 4. 配置快照
say 4/6 "同步配置快照（脱敏）"
mkdir -p "$REPO/config"
/usr/bin/python3 - "$REPO" "$SRC_CONFIG" <<'PY'
import json, os, re, sys
repo, cfgdir = sys.argv[1], sys.argv[2]
src = os.path.join(cfgdir, "feishu_newbase.json")
if os.path.exists(src):
    c = json.load(open(src, encoding="utf-8"))
    for k in list(c):
        if re.search(r'secret|token|password', k, re.I):
            c[k] = "***REDACTED***"
    json.dump(c, open(os.path.join(repo, "config", "feishu_base.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("      feishu_base.json 已脱敏")
PY

# 全局参数快照
docker exec n8n-n8n-1 node -e '
(async()=>{
  const fs=require("fs");
  const cred=require("/tmp/cred.json").find(x=>x.type==="larkApi").data;
  const tr=await fetch("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({app_id:cred.appid,app_secret:cred.appsecret})}).then(r=>r.json());
  const H={Authorization:"Bearer "+tr.tenant_access_token,"Content-Type":"application/json"};
  const APP="RvrMbDYgJaZ8mts6vPOcoCb8nMb", TID="tblDPHP5ZgioWinV";
  const flat=v=>v==null?"":(Array.isArray(v)?v.map(x=>(x&&typeof x==="object")?(x.text??""):x).join(""):(typeof v==="object"?(v.text??v.value??""):String(v)));
  const q=await fetch(`https://open.feishu.cn/open-apis/bitable/v1/apps/${APP}/tables/${TID}/records/search?page_size=100`,{method:"POST",headers:H,body:JSON.stringify({filter:{conjunction:"and",conditions:[]}})}).then(r=>r.json());
  const out={};
  for (const i of (q.data?.items||[])) {
    const k=flat(i.fields["参数键"]).trim();
    if (k) out[k]=flat(i.fields["参数值"]).trim();
  }
  fs.writeFileSync("/tmp/params_snapshot.json", JSON.stringify(out,null,1));
})();' >/dev/null 2>&1
docker cp n8n-n8n-1:/tmp/params_snapshot.json "$REPO/config/全局参数.json" >/dev/null 2>&1 && say ok "全局参数快照已保存"

# ---------------------------------------------------------------- 5. 日志
say 5/6 "同步最近明细日志"
mkdir -p "$REPO/logs"
rsync -a --delete --include 'bill_detail_*.json' --exclude '*' "$SRC_LOGS/" "$REPO/logs/" 2>/dev/null || true
ls -t "$REPO/logs"/bill_detail_*.json 2>/dev/null | tail -n +11 | xargs -r rm -f

# ---------------------------------------------------------------- 6. README + git
say 6/6 "生成清单并提交"
{
  echo "# 文件备份"
  echo
  echo "最后备份：$DATE"
  echo
  echo "## 目录"
  echo
  echo "| 目录 | 内容 |"
  echo "|---|---|"
  echo "| \`workflows/\` | n8n 已发布工作流 JSON（敏感字段已脱敏）|"
  echo "| \`scripts/\` | 自动化 Python 脚本 |"
  echo "| \`templates/\` | 计算表模板 / Bill 模板 / 页头图 |"
  echo "| \`config/\` | 全局参数快照、飞书 Base 配置（脱敏）|"
  echo "| \`logs/\` | 最近 10 次账单明细 JSON |"
  echo
  echo "## 已发布工作流"
  echo
  echo "| 名称 | 节点数 | 文件 |"
  echo "|---|---|---|"
  /usr/bin/python3 - "$REPO" <<'PY'
import json, sys, os
repo = sys.argv[1]
for x in json.load(open(os.path.join(repo, "workflows.json"), encoding="utf-8")):
    print(f"| {x['name']} | {x['nodes']} | `workflows/{x['file']}` |")
PY
  echo
  echo "## 恢复"
  echo
  echo '```bash'
  echo "# 工作流"
  echo 'docker exec n8n-n8n-1 n8n import:workflow --input=/data/.../workflows/xxx.json'
  echo "# 脚本（挂载目录直接覆盖）"
  echo 'cp -R scripts/workflow005 /Users/zuo/n8n/files/automation/scripts/'
  echo "# 模板"
  echo 'cp templates/计算表模板.xlsx /Users/zuo/n8n/files/data/workflow005/config/'
  echo '```'
  echo
  echo "> 密钥类字段在导出时替换为 \`***REDACTED***\`，恢复后需重新填写。"
} > "$REPO/README.md"

cd "$REPO"
git add -A
if git diff --cached --quiet; then
  say ok "无变化，跳过提交"
else
  git commit -q -m "backup: $DATE"
  say ok "已提交 $(git log -1 --oneline)"
fi

if git remote get-url origin >/dev/null 2>&1; then
  say push "推送到 GitHub"
  git push -q origin HEAD 2>&1 | tail -2 || say warn "推送失败（检查 Token）"
fi
du -sh "$REPO" | awk '{print "仓库大小: "$1}'
