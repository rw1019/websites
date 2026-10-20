"""把清洗后的 tables.json 灌入飞书多维表格：建表已就绪，这里只做记录写入、关联与图片上传。"""
import json
import os
import subprocess
import sys

BASE = "TdJfbyPA9a2vITsJ3oWczTqvnKe"
TABLE = {
    "产品": "tblPgeXmiaTQ8tna",
    "库存流水": "tblmiTg77d4bYnT1",
    "Y2核算": "tbltnInh519yFukz",
    "物料采购": "tblgxsU4AIDJj3BD",
    "春节备货": "tbleHUo6DYjFB0uT",
    "固定成本": "tblUCSe1M9rYXyGg",
}
CWD = "/workspace/te_base"
OUT = os.path.join(CWD, "out")

tables = json.load(open(os.path.join(OUT, "tables.json"), encoding="utf-8"))


def cli(args):
    r = subprocess.run(["lark-cli"] + args, capture_output=True, text=True, cwd=CWD)
    txt = r.stdout or r.stderr
    try:
        return json.loads(txt)
    except Exception:
        print("非JSON输出:", txt[:800])
        raise


def clean_rec(rec, drop=()):
    return {k: v for k, v in rec.items() if v is not None and v != "" and k not in drop}


def insert(table, records, label):
    """分批（<=200）写入，返回 record_id 列表（保持顺序）。"""
    ids, batch = [], 200
    for i in range(0, len(records), batch):
        chunk = records[i:i + batch]
        payload = {"create_records": chunk}
        p = os.path.join(OUT, "_payload.json")
        json.dump(payload, open(p, "w", encoding="utf-8"), ensure_ascii=False)
        res = cli(["base", "+record-batch-create", "--base-token", BASE,
                   "--table-id", TABLE[table], "--json", "@out/_payload.json", "--as", "user"])
        if not res.get("ok"):
            print(f"[FAIL] {label} 第{i//batch+1}批:", json.dumps(res, ensure_ascii=False)[:600])
            sys.exit(1)
        d = res["data"]
        got = d.get("record_id_list") or d.get("records") or []
        got = [g if isinstance(g, str) else g.get("record_id") for g in got]
        ids += got
        print(f"[OK] {label} {i+len(chunk)}/{len(records)}")
    return ids


# ---- 1. 产品：复用已写入的记录，构建 产品名 -> record_id 映射 ----
d = cli(["base", "+record-list", "--base-token", BASE, "--table-id", TABLE["产品"],
         "--limit", "200", "--format", "json", "--as", "user"])["data"]
ids, rows, cols = d["record_id_list"], d["data"], d["fields"]
name2id = {r[cols.index("产品名")]: rid for rid, r in zip(ids, rows)}
print("产品记录(已有):", len(name2id))
prod_ids = [name2id[p["产品名"]] for p in tables["产品"]]

# ---- 2. 库存流水（带 关联产品）----
flow_recs = []
for f in tables["库存流水"]:
    rec = clean_rec(f)
    rid = name2id.get(f["产品名"])
    if rid:
        rec["关联产品"] = [{"id": rid}]
    flow_recs.append(rec)
insert("库存流水", flow_recs, "库存流水")

# ---- 3. 其余表 ----
for t in ["Y2核算", "物料采购", "春节备货", "固定成本"]:
    insert(t, [clean_rec(r) for r in tables[t]], t)

# ---- 4. 上传缩略图 ----
uploaded = 0
for p, rid in zip(tables["产品"], prod_ids):
    if not p["缩略图"]:
        continue
    rel = "images/" + p["缩略图"]
    res = cli(["base", "+record-upload-attachment", "--base-token", BASE,
               "--table-id", TABLE["产品"], "--record-id", rid,
               "--field-id", "缩略图", "--file", rel, "--as", "user"])
    if res.get("ok"):
        uploaded += 1
    else:
        print("[IMG FAIL]", p["产品名"], json.dumps(res, ensure_ascii=False)[:300])
print(f"图片上传完成 {uploaded} 张")
json.dump(name2id, open(os.path.join(OUT, "name2id.json"), "w", encoding="utf-8"), ensure_ascii=False)
