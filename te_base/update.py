"""按"真产品 + 款式"重构：更新 1311 条流水的 产品名/款式/关联产品（就地更新，不重建）。"""
import json
import os
import subprocess
import sys

BASE = "TdJfbyPA9a2vITsJ3oWczTqvnKe"
T_PRODUCT = "tblPgeXmiaTQ8tna"
T_FLOW = "tblmiTg77d4bYnT1"
CWD = "/workspace/te_base"
OUT = os.path.join(CWD, "out")
tables = json.load(open(os.path.join(OUT, "tables.json"), encoding="utf-8"))


def cli(args):
    r = subprocess.run(["lark-cli"] + args, capture_output=True, text=True, cwd=CWD)
    txt = r.stdout or r.stderr
    return json.loads(txt)


def fetch(table, fields):
    base = ["base", "+record-list", "--base-token", BASE, "--table-id", table, "--format", "json"]
    for f in fields:
        base += ["--field-id", f]
    base += ["--as", "user"]
    all_ids, all_rows, cols, offset = [], [], None, 0
    while True:
        d = cli(base + ["--limit", "200", "--offset", str(offset)])["data"]
        cols = d["fields"]
        all_ids += d["record_id_list"]
        all_rows += d["data"]
        if not d.get("has_more"):
            break
        offset += 200
    return all_ids, all_rows, cols


# 1. 产品表：真产品名 -> record_id（真产品 = tables.json 里的 14 个）
parent_names = {p["产品名"] for p in tables["产品"]}
ids, rows, cols = fetch(T_PRODUCT, ["产品名"])
name2id = {}
for rid, r in zip(ids, rows):
    if r[cols.index("产品名")] in parent_names:
        name2id[r[cols.index("产品名")]] = rid
print("真产品映射:", len(name2id))

# 2. 流水表：记录顺序 == 源数据顺序（前面有 3 条你保留的测试记录，产品名为空）
ids, rows, cols = fetch(T_FLOW, ["产品名"])
pi = cols.index("产品名")
test = [rid for rid, r in zip(ids, rows) if r[pi] is None]
real = [rid for rid, r in zip(ids, rows) if r[pi] is not None]
print(f"流水记录: 共{len(ids)} 条, 测试残留{len(test)} 条, 正式{len(real)} 条")
assert len(real) == len(tables["库存流水"]), "记录数与源数据不一致，中止"

# 3. 生成更新：按顺序一一对应
updates = {}
for rid, f in zip(real, tables["库存流水"]):
    rec = {"产品名": f["产品名"], "关联产品": [{"id": name2id[f["产品名"]]}]}
    rec["款式"] = f["款式"] if f["款式"] else None
    updates[rid] = rec

items = list(updates.items())
ok = 0
for i in range(0, len(items), 200):
    chunk = dict(items[i:i + 200])
    json.dump({"update_records": chunk}, open(os.path.join(OUT, "_upd.json"), "w", encoding="utf-8"), ensure_ascii=False)
    res = cli(["base", "+record-batch-update", "--base-token", BASE, "--table-id", T_FLOW,
               "--json", "@out/_upd.json", "--as", "user"])
    if not res.get("ok"):
        print("[FAIL]", i, json.dumps(res, ensure_ascii=False)[:400]); sys.exit(1)
    ok += len(chunk)
    print(f"[OK] 更新 {ok}/{len(items)}")
print("完成，更新", ok, "条")
