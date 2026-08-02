#!/usr/bin/env python3
"""Generate per-account and per-asset-class net-worth history for each snapshot.
Method: compute the CURRENT net (USD) of each slice, choose a START allocation
that differs from today's, interpolate each slice's weight from start->current
across the snapshot dates, normalize so slices sum to that date's total, and
store by_account / by_class on every snapshot. Distinct paths, exact totals."""
import json, datetime
P="/home/claude/invest/dataset.json"
d=json.load(open(P))
fx=d["fx"]; ar=fx["ars_rate_selected"]
def per(ccy): return {"USD":1.0,"EUR":fx["EUR_per_USD"],"GBP":fx["GBP_per_USD"],"ARS":fx["ars_rates"][ar]}.get(ccy,1.0)
def tu(a,ccy): return a/per(ccy)
today=datetime.date.fromisoformat(d["meta"]["as_of"])
def dsince(s):
    y,m,dd=map(int,s.split("-")); return (today-datetime.date(y,m,dd)).days

def acct_of(x): return x.get("account","")
# ---- current net USD per class and per account ----
CLASSES=["Global Equities","Argentine Equities","Funds","International Bonds","Argentine Bonds",
         "Structured Products","Options","Futures","Cash","Loans"]
def class_vals(filter_acct=None):
    def ok(x): return filter_acct is None or x.get("account")==filter_acct
    v={c:0.0 for c in CLASSES}
    for p in d.get("funds",[]):
        if ok(p): v["Funds"]+=tu(p["units"]*p["nav"],p["ccy"])
    for p in d["equities_global"]:
        if ok(p): v["Global Equities"]+=tu(p["qty"]*p["price"],p["ccy"])
    for p in d["equities_arg"]:
        if ok(p): v["Argentine Equities"]+=tu(p["qty"]*p["price"],p["ccy"])
    for b in d["bonds_intl"]:
        if ok(b): v["International Bonds"]+=tu(b["face"]*b["price_pct"]/100 + b["face"]*(b["coupon_pct"]/100)*dsince(b["last_coupon"])/365,b["ccy"])
    for b in d["bonds_arg"]:
        if ok(b): v["Argentine Bonds"]+=tu(b["face"]*b["price_pct"]/100 + b["face"]*(b["coupon_pct"]/100)*dsince(b["last_coupon"])/365,b["ccy"])
    for s in d["structured"]:
        if ok(s): v["Structured Products"]+=tu(s["notional"]*s.get("mark_pct",100)/100 + s["notional"]*(s["accrual_pct"]/100)*dsince(s["last_coupon"])/365,s["ccy"])
    for o in d["options"]:
        if ok(o): v["Options"]+=o["contracts"]*o["multiplier"]*o["price"]/per(o["ccy"])
    for f in d["futures"]:
        if ok(f): v["Futures"]+=(f["price"]-f["entry_price"])*f["qty"]*f["multiplier"]/per(f["ccy"])
    for t in d["ledger"]:
        if filter_acct is None or t.get("account")==filter_acct:
            v["Cash"]+=tu(t["cash_flow"],t["ccy"])
    for l in d["loans"]:
        if ok(l): v["Loans"]+= -tu(l["principal"]+l["principal"]*(l["rate_pct"]/100)*dsince(l["start_date"])/365,l["ccy"])
    return v

cur_class=class_vals()
accounts=[a["id"] for a in d["accounts"]]
cur_acct={a:sum(class_vals(a).values()) for a in accounts}
net_now=sum(cur_class.values())

# ---- start-of-year allocation tweaks (relative multipliers vs current weight) ----
# >1 => bigger share at start (declined since); <1 => smaller share at start (grew since)
CLASS_START_MULT={"Global Equities":0.80,"Argentine Equities":0.72,"Funds":0.95,"International Bonds":1.12,
  "Argentine Bonds":0.90,"Structured Products":1.05,"Options":0.30,"Futures":0.35,
  "Cash":1.55,"Loans":0.25}
ACCT_START_MULT={"IBKR":0.85,"Balanz":0.78,"Cocos":0.75,"JPM":1.10,"Santander":1.4,"Galicia":1.6}

snaps=sorted(d["snapshots"], key=lambda s:s["date"])
d0=datetime.date.fromisoformat(snaps[0]["date"]); span=(today-d0).days or 1
def frac(dt): return (datetime.date.fromisoformat(dt)-d0).days/span

def build_breakdown(cur, mult):
    keys=list(cur.keys())
    # start value per slice = current * mult (un-normalized); we lerp value then scale to total
    start={k:cur[k]*mult.get(k,1.0) for k in keys}
    for sp in snaps:
        f=frac(sp["date"]); tot=sp["net_worth_usd"]
        raw={k:start[k]+(cur[k]-start[k])*f for k in keys}
        s=sum(raw.values()) or 1.0
        scale=tot/s
        sp.setdefault("_bd",{})
        for k in keys: raw[k]=round(raw[k]*scale,2)
        yield sp, raw

for sp,raw in build_breakdown(cur_class, CLASS_START_MULT):
    sp["by_class"]=raw
for sp,raw in build_breakdown(cur_acct, ACCT_START_MULT):
    sp["by_account"]=raw
for sp in snaps:
    sp.pop("_bd",None)

json.dump(d, open(P,"w"), indent=2, ensure_ascii=False)
print("current net:", round(net_now))
print("by_class YTD (start->now):")
for k in CLASSES:
    s0=snaps[0]["by_class"][k]; print(f"  {k:22} {s0:>12,.0f} -> {cur_class[k]:>12,.0f}  ({(cur_class[k]/s0-1)*100 if s0 else 0:+.1f}%)")
print("by_account YTD:")
for a in accounts:
    s0=snaps[0]["by_account"][a]; print(f"  {a:10} {s0:>12,.0f} -> {cur_acct[a]:>12,.0f}  ({(cur_acct[a]/s0-1)*100 if s0 else 0:+.1f}%)")
