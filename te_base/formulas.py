"""把飞书表里源自原 Excel 公式的计算列，改成 formula 字段（自动计算）。"""
import json
import subprocess
import sys

BASE = "E4f8bL9RfavU7kseBeYcmfExnsi"

T1 = "tblWzw9LYneCod5F"   # 采购发货流水
T5a = "tblXdLkRaqcznhiN"  # Y2测算
T5b = "tblf3dhUgfLu2Ehb"  # 备货预警

# (table_id, field_id, field_name, expression)
JOBS = [
    # ---- 表1 采购发货流水（原库存表 H/K/L/M 列公式）----
    (T1, "fldlv2FZmH", "单个综合成本",
     "[采购单价] + IFBLANK([国内运费], 0)/[采购数量] + IFBLANK([TEMU运费], 0)/[采购数量]"),
    (T1, "fldlKS5aYr", "单件利润",
     "[申报价] - [单个综合成本]"),
    (T1, "fld3FGjPx3", "净利率",
     "([申报价] - [单个综合成本])/[申报价]"),
    (T1, "fldYEe0z4h", "本次总支出",
     "ROUND([采购单价]*[采购数量] + IFBLANK([国内运费], 0) + IFBLANK([TEMU运费], 0), 2)"),
    (T1, "fldJiNDoRc", "预计销售总额",
     "[申报价]*[采购数量]"),
    (T1, "fldWN71kjl", "预计毛利总额",
     "[单件利润]*[采购数量]"),
    # ---- 表5a Y2测算（原 Y2 表 H/J/L/M/N 列公式，常量 P2=72, R2=6.8）----
    (T5a, "fldIyGMDyd", "Y2头程", "[重量kg]*72"),
    (T5a, "fldO1qnAxy", "成本价",
     "[采购价] + IFBLANK([国内快递], 0) + [Y2头程] + [贴单费] + [尾程运费]"),
    (T5a, "fldQrj0kmo", "每单赚", "[申报价] - [成本价]"),
    (T5a, "fldUhKcDRr", "毛利率", "([申报价] - [成本价])/[申报价]"),
    # ---- 表5b 备货预警（原春节备货表 K/L/N 列公式，TODAY() 固定为 2026-10-19）----
    (T5b, "fldVEY5Jgp", "7天日均销速", "[最近7天销量]/7"),
    (T5b, "fldNGGurtA", "30天日均销速", "[最近30天销量]/30"),
    (T5b, "fld6CsZrGl", "在途可用库存", "[仓内可用] + [已发库存] + [已采购]"),
    (T5b, "fldhKUXsci", "安全采购量(7天)",
     "(DAYS([预计可销售到], DATE(2026,10,19)) + [卖家发货天数] + [物流天数] + [入库时长])*([最近7天销量]/7) - ([仓内可用] + [已发库存] + [已采购])"),
    (T5b, "fldDpZh9Sh", "安全采购量(30天)",
     "(DAYS([预计可销售到], DATE(2026,10,19)) + [卖家发货天数] + [物流天数] + [入库时长])*([最近30天销量]/30) - ([仓内可用] + [已发库存] + [已采购])"),
    (T5b, "fldJAKNLh6", "建议采购总金额",
     "([安全采购量(7天)] + [安全采购量(30天)])/2*[单价]"),
]


def cli(args):
    r = subprocess.run(["lark-cli"] + args, capture_output=True, text=True, cwd="/workspace/te_base")
    txt = r.stdout or r.stderr
    try:
        return json.loads(txt)
    except Exception:
        return {"ok": False, "raw": txt[:400]}


if len(sys.argv) > 1 and sys.argv[1] == "--only":
    prefix = sys.argv[2]
    JOBS = [j for j in JOBS if j[1] == prefix or j[2] == prefix]

fail = 0
for tid, fid, name, expr in JOBS:
    body = {"type": "formula", "name": name, "expression": expr}
    res = cli(["base", "+field-update", "--base-token", BASE, "--table-id", tid,
               "--field-id", fid, "--json", json.dumps(body, ensure_ascii=False),
               "--as", "user", "--yes", "--i-have-read-guide"])
    ok = bool(res.get("ok")) or bool(res.get("data", {}).get("updated"))
    print(("  OK  " if ok else "  FAIL") + f" {name}  ({tid})")
    if not ok:
        fail += 1
        print("     ", json.dumps(res, ensure_ascii=False)[:400])
print("完成，失败", fail, "个")
