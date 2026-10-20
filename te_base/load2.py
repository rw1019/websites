"""新建多维表格并写入 5 张（6 个）表。"""
import json
import os
import subprocess
import sys

OUT = "/workspace/te_base/out2"
tables = json.load(open(f"{OUT}/tables.json", encoding="utf-8"))


def cli(args):
    r = subprocess.run(["lark-cli"] + args, capture_output=True, text=True, cwd="/workspace/te_base")
    txt = r.stdout or r.stderr
    try:
        return json.loads(txt)
    except Exception:
        print("BAD:", txt[:500]); sys.exit(1)


def num(name, p=2):
    return {"name": name, "type": "number", "style": {"type": "plain", "precision": p}}


def txt(name):
    return {"name": name, "type": "text"}


def dt(name):
    return {"name": name, "type": "datetime", "style": {"format": "yyyy-MM-dd"}}


def sel(name, opts):
    return {"name": name, "type": "select", "multiple": False,
            "options": [{"name": o, "hue": h} for o, h in opts]}


T1 = "采购发货流水"
T2 = "SKU档案库"
T3 = "物料采购流水"
T4 = "月度资金对账"
T5a = "Y2测算"
T5b = "备货预警"

f1 = [txt("主分类"), txt("子规格"), txt("产品完整名称"), dt("采购日期"), txt("采购年月"),
      dt("发货日期"), sel("发货状态", [("已发货", "Green"), ("待发货", "Orange")]),
      num("采购单价"), num("采购数量"), num("国内运费"), num("TEMU运费"), num("单个综合成本"),
      num("本次总支出"), num("申报价"), num("单件利润"), num("净利率", 4),
      num("预计销售总额"), num("预计毛利总额"), txt("SKC"), txt("备注")]

f2 = [txt("产品完整名称"), txt("所属大类"), txt("SKC编码"), num("最新/常用采购价"),
      num("最新申报价"), num("累计采购批次", 0), num("累计采购总量"), num("累计采购总金额"),
      num("累计理论毛利贡献"), txt("图片关联ID"), {"name": "缩略图", "type": "attachment"}]

f3 = [txt("耗材品名"), dt("采购日期"), txt("采购月份"), num("支出金额")]

f4 = [dt("对账月份"), num("平台结算回款"), num("其他收入/补贴"), num("当月总支出"),
      num("当月现金净结余"), num("累计滚存资金")]

f5a = [txt("产品名"), dt("订单日期"), num("采购价"), num("申报价"), num("活动价"), num("重量kg"),
       num("国内快递"), num("Y2头程"), num("贴单费"), num("尾程运费"), num("汇率"),
       num("成本价"), num("每单赚"), num("毛利率", 4)]

f5b = [txt("产品名称"), num("最近7天销量", 0), num("最近30天销量", 0), num("7天日均销速", 4),
       num("30天日均销速", 4), num("仓内可用"), num("已发库存"), num("已采购"), num("在途可用库存"),
       num("卖家发货天数", 0), num("物流天数", 0), num("入库时长", 0), dt("预计可销售到"),
       num("安全采购量(7天)"), num("安全采购量(30天)"), num("单价"), num("建议采购总金额"),
       txt("采购状态")]


def field_names(fields):
    return [f["name"] for f in fields]


def clean_rec(rec, fields):
    names = field_names(fields)
    out = {}
    for n in names:
        v = rec.get(n)
        if v is None:
            continue
        if n == "发货状态":
            out[n] = [v]
        else:
            out[n] = v
    return out


def insert(base, tid, recs, fields):
    items = [clean_rec(r, fields) for r in recs]
    for i in range(0, len(items), 200):
        chunk = items[i:i + 200]
        json.dump({"create_records": chunk}, open(f"{OUT}/_c.json", "w", encoding="utf-8"), ensure_ascii=False)
        res = cli(["base", "+record-batch-create", "--base-token", base, "--table-id", tid,
                   "--json", "@out2/_c.json", "--as", "user"])
        if not res.get("ok"):
            print("[FAIL]", i, json.dumps(res, ensure_ascii=False)[:500]); sys.exit(1)
        print(f"  {tid} 写入 {min(i+200,len(items))}/{len(items)}")


# 1. 创建（或复用）Base + 第一个表
base_json = f"{OUT}/base.json"
if os.path.exists(base_json):
    BASE = json.load(open(base_json, encoding="utf-8"))["base"]
    print("复用已有 Base:", BASE)
else:
    res = cli(["base", "+base-create", "--name", "Te 数据表（重构版）", "--time-zone", "Asia/Shanghai",
               "--table-name", T1, "--fields", json.dumps(f1, ensure_ascii=False), "--as", "user"])
    if not res.get("ok"):
        print("[FAIL base-create]", json.dumps(res, ensure_ascii=False)[:600]); sys.exit(1)
    BASE = res["data"]["base"]["base_token"]
    json.dump({"base": BASE}, open(base_json, "w", encoding="utf-8"))
    print("Base:", BASE, res["data"]["base"].get("url"))


def table_ids(base):
    tl = cli(["base", "+table-list", "--base-token", base, "--as", "user", "--format", "json"])
    return {t["name"]: t["id"] for t in tl["data"]["tables"]}


ids = table_ids(BASE)
for name, fields in ((T2, f2), (T3, f3), (T4, f4), (T5a, f5a), (T5b, f5b)):
    if name in ids:
        print("已存在，跳过:", name); continue
    r = cli(["base", "+table-create", "--base-token", BASE, "--name", name,
             "--fields", json.dumps(fields, ensure_ascii=False), "--as", "user"])
    if not r.get("ok"):
        print(f"[FAIL {name}]", json.dumps(r, ensure_ascii=False)[:500]); sys.exit(1)
    ids = table_ids(BASE)
    print("建表:", name, ids[name])
for k in (T1, T2, T3, T4, T5a, T5b):
    print("  表:", k, ids[k])

# 3. 写数据
print("写数据:")
insert(BASE, ids[T1], tables["采购发货流水"], f1)
insert(BASE, ids[T2], tables["SKU档案库"], f2)
insert(BASE, ids[T3], tables["物料采购流水"], f3)
insert(BASE, ids[T4], tables["月度资金对账"], f4)
insert(BASE, ids[T5a], tables["Y2测算"], f5a)
insert(BASE, ids[T5b], tables["备货预警"], f5b)

json.dump({"base": BASE, "tables": ids}, open(f"{OUT}/base.json", "w", encoding="utf-8"))
print("完成 BASE=", BASE)
