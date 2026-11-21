#!/bin/bash
# n8n 工作流备份脚本：导出全部工作流 → 脱敏 → git 提交（可选推送 GitHub）
set -euo pipefail

REPO="$(cd "$(dirname "$0")" && pwd)"
DATE="$(date '+%Y-%m-%d %H:%M:%S')"
WORKDIR="$REPO/workflows"

echo "[1/4] 导出 n8n 已发布工作流"
# 只导出「已发布/已激活」的流程
docker exec n8n-n8n-1 n8n export:workflow --all --published --output=/tmp/n8n_all.json >/dev/null
docker cp n8n-n8n-1:/tmp/n8n_all.json "$REPO/.export.json" >/dev/null

echo "[2/4] 拆分为单文件 + 脱敏"
/usr/bin/python3 - "$REPO" <<'PY'
import json, os, re, sys, hashlib
repo = sys.argv[1]
data = json.load(open(os.path.join(repo, ".export.json"), encoding="utf-8"))
outdir = os.path.join(repo, "workflows")
os.makedirs(outdir, exist_ok=True)

SECRET_PATTERNS = [
    # 支持被转义的 JSON 字符串（如 jsonBody 里的 \"app_secret\": \"...\"）
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
    wid = w.get("id", "")
    name = w.get("name", "")
    clean = json.loads(json.dumps(w, ensure_ascii=False))
    text = json.dumps(clean, ensure_ascii=False, indent=1)
    redacted = 0
    for pat, rep in SECRET_PATTERNS:
        text, n = pat.subn(rep, text)
        redacted += n
    path = os.path.join(outdir, slug(name, wid) + ".json")
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    index.append({
        "id": wid, "name": name, "active": bool(w.get("active")),
        "nodes": len(w.get("nodes", [])), "file": os.path.basename(path),
        "redacted": redacted,
    })

index.sort(key=lambda x: (not x["active"], x["name"]))
with open(os.path.join(repo, "workflows.json"), "w", encoding="utf-8") as f:
    json.dump(index, f, ensure_ascii=False, indent=1)
print(f"      工作流 {len(index)} 个，脱敏 {sum(i['redacted'] for i in index)} 处")
PY

echo "[3/4] 生成本次备份说明"
{
  echo "# n8n 工作流备份"
  echo
  echo "最后备份：$DATE"
  echo
  echo "| 状态 | 名称 | 节点数 | 文件 |"
  echo "|---|---|---|---|"
  /usr/bin/python3 - "$REPO" <<'PY'
import json, sys, os
repo = sys.argv[1]
for x in json.load(open(os.path.join(repo, "workflows.json"), encoding="utf-8")):
    flag = "✅" if x["active"] else "⏸"
    print(f"| {flag} | {x['name']} | {x['nodes']} | `{x['file']}` |")
PY
  echo
  echo "> 敏感字段（app_secret / Bearer token）已自动脱敏为 \`***REDACTED***\`，恢复时需重新填写。"
} > "$REPO/README.md"

echo "[4/4] 提交 git"
cd "$REPO"
if [ ! -d .git ]; then
  git init -q
  git config user.name "n8n-backup"
  git config user.email "backup@local"
fi
git add -A
if git diff --cached --quiet; then
  echo "      无变化，跳过提交"
else
  git commit -q -m "backup: $DATE"
  echo "      已提交：$(git log -1 --oneline)"
fi

if git remote get-url origin >/dev/null 2>&1; then
  echo "[push] 推送到 GitHub"
  git push -q origin HEAD 2>&1 | tail -3 || echo "      推送失败（检查凭证）"
else
  echo "[push] 未配置 origin，跳过（配置方法见 README）"
fi
