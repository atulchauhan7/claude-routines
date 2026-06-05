"""
Daily Store & Ads Performance Dashboard Generator
Date: 2026-06-04 (yesterday in Asia/Kolkata)
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
import os

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
REPORT_DATE = "2026-06-04"
BRAND = "Dhirai"
DARK_NAVY   = "1A2744"
MID_NAVY    = "2E4272"
ACCENT_GOLD = "C9A84C"
LIGHT_BG    = "F5F7FA"
GREEN_LIGHT = "D6F5D6"
RED_LIGHT   = "FFD6D6"
YELLOW_BG   = "FFF9D6"
WHITE       = "FFFFFF"

def fill(c): return PatternFill(fill_type="solid", fgColor=c)
def font(bold=False, color="000000", size=11): return Font(bold=bold, color=color, size=size, name="Calibri")
def border():
    s = Side(border_style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)
def align(h="left", v="center", wrap=False): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def col_w(ws, col, w): ws.column_dimensions[get_column_letter(col)].width = w
def row_h(ws, row, h): ws.row_dimensions[row].height = h

def style_hdr(c, bg=DARK_NAVY, sz=11):
    c.fill = fill(bg); c.font = font(True, "FFFFFF", sz)
    c.alignment = align("center", "center", True); c.border = border()

def style_sub(c, bg=MID_NAVY):
    c.fill = fill(bg); c.font = font(True, "FFFFFF", 10)
    c.alignment = align("center", "center"); c.border = border()

def style_data(c, bg=WHITE, center=False, bold=False, sz=10):
    c.fill = fill(bg); c.font = font(bold, "000000", sz)
    c.alignment = align("center" if center else "left", "center"); c.border = border()

def merge_title(ws, rng, text, bg=DARK_NAVY, sz=14, fc="FFFFFF", bold=True):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = text; c.fill = fill(bg)
    c.font = font(bold, fc, sz)
    c.alignment = align("center", "center")

def section(ws, row, c1, c2, text, bg="C9A84C"):
    cl = get_column_letter
    merge_title(ws, f"{cl(c1)}{row}:{cl(c2)}{row}", text, bg=bg, sz=12, fc=DARK_NAVY)
    row_h(ws, row, 22)

def pct_chg(val, avg):
    return (val - avg) / avg * 100 if avg else 0

# ─────────────────────────────────────────────
# DATA
# ─────────────────────────────────────────────
# Shopify Yesterday
SY = dict(gross=114081.80, net=105301.05, orders=48, aov=2312.46,
          returns=5697.00, discounts=3083.75, taxes=1947.44,
          total_sales=107248.49, sessions=4413, cart=220,
          checkout_reach=22, checkout_done=5, conv_rate=0.00113)

# Shopify 7-day history (28 May – 3 Jun)
SHOP_HIST = [
    ("2026-05-28", 160630.00, 140368.98, 59, 2682.39, 17892.00, 2369.02),
    ("2026-05-29", 166625.00, 138446.33, 67, 2452.75, 25888.00, 2290.67),
    ("2026-05-30", 129240.00, 106520.96, 55, 2320.20, 21090.00, 1629.04),
    ("2026-05-31", 143728.07, 125599.95, 63, 2242.75, 15693.00, 2435.12),
    ("2026-06-01", 117347.00, 114533.40, 50, 2330.65,  1999.00,  814.60),
    ("2026-06-02", 155330.00, 149053.16, 62, 2437.94,  2099.00, 4177.84),
    ("2026-06-03", 102051.00,  97708.08, 45, 2171.29,     0.00, 4342.92),
]
def avg7(col): return sum(r[col] for r in SHOP_HIST) / 7
S7 = dict(gross=avg7(1), net=avg7(2), orders=avg7(3), aov=avg7(4),
          returns=avg7(5), discounts=avg7(6))

# Meta campaigns yesterday
CAMPS = [
    dict(name="Dhirai Scale -ASC", status="ACTIVE", spend=20665.56,
         impr=230572, reach=209640, clicks=3509, ctr=1.52, cpc=5.89,
         cpm=89.63, purch=28, cpa=738.06, roas=2.97),
    dict(name="Black DHURANDHAR", status="ACTIVE", spend=10023.10,
         impr=51350, reach=42395, clicks=1826, ctr=3.56, cpc=5.49,
         cpm=195.19, purch=9, cpa=1113.68, roas=2.42),
    dict(name="Retargeting - New", status="ACTIVE", spend=6929.13,
         impr=22497, reach=15596, clicks=798, ctr=3.55, cpc=8.68,
         cpm=308.00, purch=9, cpa=769.90, roas=3.30),
    dict(name="Summer Special Campaign", status="PAUSED", spend=210.81,
         impr=1306, reach=1173, clicks=24, ctr=1.84, cpc=8.78,
         cpm=161.42, purch=0, cpa=None, roas=None),
]
META_7D = dict(spend=(155932.75+71654.51+53567.49)/7, purch=(195+83+88)/7)

TOT_SPEND   = sum(c["spend"] for c in CAMPS)
TOT_PURCH   = sum(c["purch"] for c in CAMPS)
TOT_IMPR    = sum(c["impr"]  for c in CAMPS)
TOT_REACH   = sum(c["reach"] for c in CAMPS)
TOT_CLICKS  = sum(c["clicks"] for c in CAMPS)
BLEND_CTR   = TOT_CLICKS / TOT_IMPR * 100 if TOT_IMPR else 0
BLEND_CPC   = TOT_SPEND / TOT_CLICKS if TOT_CLICKS else 0
BLEND_CPA   = TOT_SPEND / TOT_PURCH if TOT_PURCH else 0
REV_ATTR    = sum(c["spend"]*c["roas"] for c in CAMPS if c["roas"])
BLEND_ROAS  = REV_ATTR / (TOT_SPEND - 210.81) if (TOT_SPEND-210.81) else 0

TOP_PRODS = [
    ("The Essential – White Cotton Shirt & Pant Co-ord Set",45951.88,45851.93,21),
    ("Embroidered Rayon Kurti with Farshi Salwar Set",      11296.00,10411.52, 4),
    ("Black Cotton Kurta Pant Set",                          7691.07, 7691.07, 4),
    ("Ira Rayon Co-ord Set",                                 4998.00, 2499.00, 2),
    ("VAIRA Airy Linen Everyday Co-ord Set",                 3998.00, 3998.00, 2),
    ("The Essential Co-ord Set",                             3998.00, 3998.00, 2),
    ("Inayat — Power Pastel Cotton Shirt & Trouser Co-ord", 3998.00, 3898.05, 2),
    ("Noor Blush Pink Cotton Lined Kurta & Palazzo Set",     3499.00, 3499.00, 1),
    ("Navya Cotton Anarkali Suit Set",                       2999.00, 2999.00, 1),
    ("Nidra – Jet Black Premium Georgette Suit Set",         2999.00, 2999.00, 1),
]

RECS = [
    ("HIGH",   "Meta Ads – Scale",     "Retargeting - New ROAS 3.30: increase budget by 20–30% today.",       "Media Buyer"),
    ("HIGH",   "Meta Ads – Monitor",   "ASC campaign ROAS 2.97: if CPP stays <₹800 add ₹5K budget.",          "Media Buyer"),
    ("MEDIUM", "Meta Ads – Review",    "Black DHURANDHAR CPA ₹1,114 (high): refresh creatives urgently.",     "Creative Team"),
    ("HIGH",   "Shopify – Checkout",   "Only 22.7% checkout completion (5/22): audit payment & UX flow.",      "Tech / CRO"),
    ("HIGH",   "Shopify – Hero SKU",   "'The Essential' drove 40% of revenue; verify full size inventory.",    "Operations"),
    ("MEDIUM", "Shopify – Sales",      "Gross sales ₹1,14,082 is 18% below 7-day avg; check for issues.",     "Operations"),
    ("MEDIUM", "Shopify – Conv Rate",  "Conv rate 0.11% — add urgency banners, low-stock alerts on PDPs.",     "CRO / Marketing"),
    ("LOW",    "Meta Ads – Paused",    "Summer Special Campaign (PAUSED) spent ₹211 — retire or reactivate?", "Media Buyer"),
    ("INFO",   "Data Gap",             "Abandoned checkouts & failed payments unavailable via ShopifyQL API.", "Tech"),
    ("INFO",   "Attribution Note",     "Meta pixel purchases (46) match Shopify orders (48) — tracking OK.",   "Media Buyer"),
]

# ─────────────────────────────────────────────
# BUILD WORKBOOK
# ─────────────────────────────────────────────
wb = Workbook()
default = wb.active; wb.remove(default)

ws_dash = wb.create_sheet("Dashboard")
ws_shop = wb.create_sheet("Shopify Daily Data")
ws_meta = wb.create_sheet("Meta Ads Daily Data")
ws_prod = wb.create_sheet("Product Performance")
ws_camp = wb.create_sheet("Campaign Performance")
ws_rec  = wb.create_sheet("Recommendations & Notes")

for ws, tc in [(ws_dash, DARK_NAVY), (ws_shop, "0070C0"),
               (ws_meta, "3B5998"), (ws_prod, "70AD47"),
               (ws_camp, "ED7D31"), (ws_rec,  "C9A84C")]:
    ws.sheet_properties.tabColor = tc

# ── TAB: Shopify Daily Data ──────────────────
ws = ws_shop; ws.freeze_panes = "A3"
merge_title(ws, "A1:N1", f"Shopify Daily Data — {BRAND}", sz=14); row_h(ws, 1, 28)
hdrs = ["Date","Gross Sales (₹)","Net Sales (₹)","Orders","AOV (₹)",
        "Returns (₹)","Discounts (₹)","Taxes (₹)","Total Sales (₹)",
        "Sessions","Cart Adds","Checkout Reached","Checkout Completed","Conv Rate (%)"]
for ci, h in enumerate(hdrs, 1):
    style_hdr(ws.cell(row=2, column=ci, value=h)); row_h(ws, 2, 34)

for ri, row in enumerate(SHOP_HIST, 3):
    bg = LIGHT_BG if ri % 2 == 0 else WHITE
    vals = [row[0], row[1], row[2], row[3], row[4], row[5], row[6], 0, row[2],
            None, None, None, None, None]
    fmts = [None,"₹#,##0.00","₹#,##0.00","#,##0","₹#,##0.00","₹#,##0.00","₹#,##0.00","₹#,##0.00","₹#,##0.00",
            "#,##0","#,##0","#,##0","#,##0","0.00%"]
    for ci, (v, f) in enumerate(zip(vals, fmts), 1):
        c = ws.cell(row=ri, column=ci, value=v)
        style_data(c, bg=bg, center=True)
        if f: c.number_format = f
    row_h(ws, ri, 18)

ri = len(SHOP_HIST) + 3
yvals = [REPORT_DATE+" ★", SY["gross"], SY["net"], SY["orders"], SY["aov"],
         SY["returns"], SY["discounts"], SY["taxes"], SY["total_sales"],
         SY["sessions"], SY["cart"], SY["checkout_reach"], SY["checkout_done"],
         SY["conv_rate"]]
yfmts = [None,"₹#,##0.00","₹#,##0.00","#,##0","₹#,##0.00","₹#,##0.00","₹#,##0.00","₹#,##0.00","₹#,##0.00",
         "#,##0","#,##0","#,##0","#,##0","0.00%"]
for ci, (v, f) in enumerate(zip(yvals, yfmts), 1):
    c = ws.cell(row=ri, column=ci, value=v)
    c.fill = fill("FFF2CC"); c.font = font(True, "000000", 10)
    c.alignment = align("center","center"); c.border = border()
    if f: c.number_format = f
row_h(ws, ri, 20)
ws.auto_filter.ref = f"A2:N{ri}"
for ci, w in enumerate([14,16,14,10,12,14,14,12,14,10,10,18,20,16], 1): col_w(ws, ci, w)

# ── TAB: Meta Ads Daily Data ─────────────────
ws = ws_meta; ws.freeze_panes = "A3"
merge_title(ws, "A1:N1", f"Meta Ads Daily Data — {BRAND}", sz=14); row_h(ws, 1, 28)
mhdrs = ["Date","Campaign","Status","Spend (₹)","Impressions","Reach",
         "Clicks","CTR (%)","CPC (₹)","CPM (₹)","Purchases","CPA (₹)","ROAS","7D Avg ROAS"]
for ci, h in enumerate(mhdrs, 1):
    style_hdr(ws.cell(row=2, column=ci, value=h)); row_h(ws, 2, 34)

camp_7d_roas = {"Dhirai Scale -ASC":2.98,"Black DHURANDHAR":2.76,
                "Retargeting - New":3.92,"Summer Special Campaign":3.87}
ri = 3
for i, camp in enumerate(CAMPS):
    bg = GREEN_LIGHT if camp["status"]=="ACTIVE" else LIGHT_BG
    vals = [REPORT_DATE, camp["name"], camp["status"], camp["spend"],
            camp["impr"], camp["reach"], camp["clicks"],
            camp["ctr"]/100, camp["cpc"], camp["cpm"],
            camp["purch"], camp["cpa"], camp["roas"],
            camp_7d_roas.get(camp["name"])]
    fmts = [None,None,None,"₹#,##0.00","#,##0","#,##0","#,##0",
            "0.00%","₹#,##0.00","₹#,##0.00","#,##0","₹#,##0.00","0.00","0.00"]
    for ci, (v, f) in enumerate(zip(vals, fmts), 1):
        c = ws.cell(row=ri, column=ci, value=v)
        style_data(c, bg=bg, center=(ci!=2))
        if f and v is not None: c.number_format = f
    ri += 1

# Totals row
totals = [REPORT_DATE,"TOTAL / BLENDED","",TOT_SPEND,TOT_IMPR,TOT_REACH,TOT_CLICKS,
          BLEND_CTR/100,BLEND_CPC,None,TOT_PURCH,BLEND_CPA,BLEND_ROAS,None]
tfmts = [None,None,None,"₹#,##0.00","#,##0","#,##0","#,##0","0.00%","₹#,##0.00",None,"#,##0","₹#,##0.00","0.00",None]
for ci, (v, f) in enumerate(zip(totals, tfmts), 1):
    c = ws.cell(row=ri, column=ci, value=v)
    c.fill = fill(DARK_NAVY); c.font = font(True,"FFFFFF",10)
    c.alignment = align("center","center"); c.border = border()
    if f and v is not None: c.number_format = f

ws.auto_filter.ref = f"A2:N{ri-1}"
for ci, w in enumerate([12,28,10,13,12,12,10,9,10,10,12,12,9,12], 1): col_w(ws, ci, w)
# Conditional: ROAS vs 7D
ws.conditional_formatting.add(f"M3:M{ri-1}",
    CellIsRule("greaterThanOrEqual",["3"],fill=fill(GREEN_LIGHT)))
ws.conditional_formatting.add(f"M3:M{ri-1}",
    CellIsRule("lessThan",["2.5"],fill=fill(RED_LIGHT)))

# ── TAB: Product Performance ─────────────────
ws = ws_prod; ws.freeze_panes = "A3"
merge_title(ws, "A1:H1", f"Product Performance — {BRAND}", sz=14); row_h(ws, 1, 28)
phdrs = ["Date","Product Name","Gross Sales (₹)","Net Sales (₹)","Orders","AOV (₹)",
         "7D Avg Daily Gross (₹)","% of Day Gross"]
for ci, h in enumerate(phdrs, 1):
    style_hdr(ws.cell(row=2, column=ci, value=h)); row_h(ws, 2, 34)

prod_7d = {
    "The Essential – White Cotton Shirt & Pant Co-ord Set": 337533/7,
    "Embroidered Rayon Kurti with Farshi Salwar Set":       140950/7,
    "Black Cotton Kurta Pant Set":                           32084/7,
    "Ira Rayon Co-ord Set":                                  39984/7,
}
for ri, (name, gross, net, orders) in enumerate(TOP_PRODS, 3):
    bg = LIGHT_BG if ri%2==0 else WHITE
    aov_p = gross/orders if orders else 0
    d7g = prod_7d.get(name)
    vals = [REPORT_DATE, name, gross, net, orders, aov_p, d7g, gross/SY["gross"]]
    fmts = [None,None,"₹#,##0.00","₹#,##0.00","#,##0","₹#,##0.00","₹#,##0.00","0.0%"]
    for ci, (v, f) in enumerate(zip(vals, fmts), 1):
        c = ws.cell(row=ri, column=ci, value=v)
        style_data(c, bg=bg, center=(ci!=2))
        if f and v is not None: c.number_format = f
    row_h(ws, ri, 18)

ws.auto_filter.ref = f"A2:H{ri}"
ws.conditional_formatting.add(f"G3:G{ri}",
    CellIsRule("greaterThan",["0"],fill=fill(GREEN_LIGHT)))
for ci, w in enumerate([12,52,16,14,10,12,20,16], 1): col_w(ws, ci, w)

# ── TAB: Campaign Performance ────────────────
ws = ws_camp; ws.freeze_panes = "A3"
merge_title(ws, "A1:R1", f"Campaign Performance — {BRAND}", sz=14); row_h(ws, 1, 28)
cht = ["Date","Campaign","Status","Spend (₹)","Impressions","Reach","Clicks",
       "CTR (%)","CPC (₹)","CPM (₹)","Purchases","CPA (₹)","ROAS",
       "7D Avg Spend/Day","7D Avg Purch/Day","7D Avg ROAS",
       "Spend vs 7D (%)","ROAS vs 7D (%)"]
for ci, h in enumerate(cht, 1):
    style_hdr(ws.cell(row=2, column=ci, value=h)); row_h(ws, 2, 44)

camp_7d_data = {
    "Dhirai Scale -ASC":      dict(spend=155932.75/7, purch=195/7, roas=2.98),
    "Black DHURANDHAR":       dict(spend=71654.51/7,  purch=83/7,  roas=2.76),
    "Retargeting - New":      dict(spend=53567.49/7,  purch=88/7,  roas=3.92),
    "Summer Special Campaign":dict(spend=7464.74/7,   purch=13/7,  roas=3.87),
}
ri = 3
for i, camp in enumerate(CAMPS):
    bg = GREEN_LIGHT if camp["status"]=="ACTIVE" else LIGHT_BG
    d7 = camp_7d_data.get(camp["name"], {})
    s7 = d7.get("spend",0); p7 = d7.get("purch",0); r7 = d7.get("roas")
    spct = pct_chg(camp["spend"],s7)/100 if s7 else None
    rpct = pct_chg(camp["roas"] or 0, r7)/100 if r7 and camp["roas"] else None
    vals = [REPORT_DATE,camp["name"],camp["status"],camp["spend"],camp["impr"],
            camp["reach"],camp["clicks"],camp["ctr"]/100,camp["cpc"],camp["cpm"],
            camp["purch"],camp["cpa"],camp["roas"],s7,p7,r7,spct,rpct]
    fmts = [None,None,None,"₹#,##0.00","#,##0","#,##0","#,##0","0.00%","₹#,##0.00","₹#,##0.00",
            "#,##0","₹#,##0.00","0.00","₹#,##0.00","0.0","0.00","0.0%","0.0%"]
    for ci, (v, f) in enumerate(zip(vals, fmts), 1):
        c = ws.cell(row=ri, column=ci, value=v)
        style_data(c, bg=bg, center=(ci!=2))
        if f and v is not None: c.number_format = f
    ri += 1

ws.auto_filter.ref = f"A2:R{ri-1}"
ws.conditional_formatting.add(f"R3:R{ri-1}",
    CellIsRule("greaterThan",["0"],fill=fill(GREEN_LIGHT)))
ws.conditional_formatting.add(f"R3:R{ri-1}",
    CellIsRule("lessThan",["0"],fill=fill(RED_LIGHT)))
ws.conditional_formatting.add(f"L3:L{ri-1}",
    CellIsRule("greaterThan",["1000"],fill=fill(RED_LIGHT)))
ws.conditional_formatting.add(f"L3:L{ri-1}",
    CellIsRule("lessThan",["800"],fill=fill(GREEN_LIGHT)))
for ci, w in enumerate([12,28,10,13,12,12,10,9,10,10,10,12,9,16,14,12,16,14], 1): col_w(ws, ci, w)

# ── TAB: Recommendations & Notes ─────────────
ws = ws_rec; ws.freeze_panes = "A3"
merge_title(ws, "A1:E1", f"Recommendations & Notes — {BRAND} — {REPORT_DATE}", sz=14); row_h(ws, 1, 28)
rhdrs = ["#","Category","Priority","Recommendation / Note","Action Owner"]
for ci, h in enumerate(rhdrs, 1):
    style_hdr(ws.cell(row=2, column=ci, value=h)); row_h(ws, 2, 28)

pcol = {"HIGH":"FFD6D6","MEDIUM":"FFF9D6","LOW":"D6F5D6","INFO":"E8F4FD"}
for i, (pri, cat, note, owner) in enumerate(RECS):
    ri = i + 3
    bg = pcol.get(pri, WHITE)
    vals = [i+1, cat, pri, note, owner]
    for ci, v in enumerate(vals, 1):
        c = ws.cell(row=ri, column=ci, value=v)
        c.fill = fill(bg); c.font = font(ci==3, "000000", 10)
        c.alignment = align("center" if ci!=4 else "left","top",ci==4)
        c.border = border()
    row_h(ws, ri, 48)

ws.auto_filter.ref = f"A2:E{len(RECS)+2}"
for ci, w in enumerate([5,22,12,80,20], 1): col_w(ws, ci, w)

# ── TAB: Dashboard ────────────────────────────
ws = ws_dash; ws.sheet_view.showGridLines = False

merge_title(ws,"A1:L1",f"Daily Store & Ads Performance — {BRAND}",sz=18); row_h(ws,1,38)
merge_title(ws,"A2:L2",
    f"Report Date: {REPORT_DATE}  |  Timezone: Asia/Kolkata  |  7-Day Avg: 28 May – 3 Jun 2026",
    bg=MID_NAVY, sz=11); row_h(ws,2,20)

def kpi(ws, row, col, label, value, chg=None, fmt="₹#,##0"):
    c1 = ws.cell(row=row, column=col, value=label)
    c1.fill=fill(LIGHT_BG); c1.font=font(True,"555555",9)
    c1.alignment=align("left","center"); c1.border=border()
    c2 = ws.cell(row=row+1, column=col, value=value)
    c2.fill=fill(WHITE); c2.font=font(True,DARK_NAVY,14)
    c2.alignment=align("left","center"); c2.border=border()
    if fmt: c2.number_format=fmt
    if chg is not None:
        c3 = ws.cell(row=row+2, column=col, value=chg/100)
        c3.number_format="+0.0%;-0.0%;0.0%"
        fc = "375623" if chg>=0 else "C00000"
        c3.fill=fill(GREEN_LIGHT if chg>=0 else RED_LIGHT)
        c3.font=font(True,fc,9)
        c3.alignment=align("left","center"); c3.border=border()
    else:
        c3 = ws.cell(row=row+2, column=col)
        c3.fill=fill(LIGHT_BG); c3.border=border()

# ── Shopify KPIs ──
section(ws, 4, 1, 6, "SHOPIFY — Yesterday vs 7-Day Average", bg=ACCENT_GOLD)
kpi_shop = [
    ("Gross Sales",SY["gross"],pct_chg(SY["gross"],S7["gross"]),"₹#,##0"),
    ("Net Sales",SY["net"],pct_chg(SY["net"],S7["net"]),"₹#,##0"),
    ("Orders",SY["orders"],pct_chg(SY["orders"],S7["orders"]),"#,##0"),
    ("Avg Order Value",SY["aov"],pct_chg(SY["aov"],S7["aov"]),"₹#,##0.00"),
    ("Returns",SY["returns"],pct_chg(SY["returns"],S7["returns"]),"₹#,##0"),
    ("Sessions",SY["sessions"],None,"#,##0"),
]
for i, (lbl, val, chg, fmt) in enumerate(kpi_shop):
    kpi(ws, 5, i+1, lbl, val, chg, fmt)
for r in range(5,8): row_h(ws,r,22)

# ── Checkout Funnel ──
section(ws, 9, 1, 4, "CHECKOUT FUNNEL — Yesterday", bg="0070C0")
ws.cell(row=9,column=1).font=Font(bold=True,color="FFFFFF",size=12)
funnel_data = [("Sessions",SY["sessions"]),("Added to Cart",SY["cart"]),
               ("Checkout Reached",SY["checkout_reach"]),("Checkout Done",SY["checkout_done"])]
for ci, (lbl, val) in enumerate(funnel_data, 1):
    c = ws.cell(row=10, column=ci, value=lbl)
    c.fill=fill("D6E4FF"); c.font=font(True,"000000",9); c.alignment=align("center","center"); c.border=border()
    c2 = ws.cell(row=11, column=ci, value=val)
    c2.fill=fill(WHITE); c2.font=Font(bold=True,color=MID_NAVY,size=16)
    c2.alignment=align("center","center"); c2.number_format="#,##0"; c2.border=border()
row_h(ws,10,20); row_h(ws,11,32)
ws.merge_cells("A12:D12")
cr = ws.cell(row=12, column=1,
    value=f"Conv Rate: {SY['conv_rate']*100:.2f}%   |   Checkout Completion: {5/22*100:.1f}%")
cr.fill=fill(YELLOW_BG); cr.font=font(True,DARK_NAVY,10)
cr.alignment=align("center","center"); cr.border=border(); row_h(ws,12,18)

# ── Meta KPIs ──
section(ws, 14, 1, 6, "META ADS — Blended Yesterday vs 7-Day Avg", bg="3B5998")
ws.cell(row=14,column=1).font=Font(bold=True,color="FFFFFF",size=12)
kpi_meta = [
    ("Total Spend",TOT_SPEND,pct_chg(TOT_SPEND,META_7D["spend"]),"₹#,##0"),
    ("Total Purchases",TOT_PURCH,pct_chg(TOT_PURCH,META_7D["purch"]),"#,##0"),
    ("Blended ROAS",BLEND_ROAS,None,"0.00"),
    ("Blended CPA",BLEND_CPA,None,"₹#,##0.00"),
    ("Impressions",TOT_IMPR,None,"#,##0"),
    ("Total Clicks",TOT_CLICKS,None,"#,##0"),
]
for i, (lbl, val, chg, fmt) in enumerate(kpi_meta):
    kpi(ws, 15, i+1, lbl, val, chg, fmt)
for r in range(15,18): row_h(ws,r,22)

# ── Top 5 Products ──
section(ws, 19, 1, 4, "TOP 5 PRODUCTS — Yesterday", bg="70AD47")
for ci, h in enumerate(["Product","Gross Sales","Orders","% of Day"],1):
    style_sub(ws.cell(row=20,column=ci,value=h)); row_h(ws,20,22)
tot_g = SY["gross"]
for i, (nm, gr, net, ord_) in enumerate(TOP_PRODS[:5]):
    ri = 21+i; bg = LIGHT_BG if i%2==0 else WHITE
    for ci, (v, f) in enumerate(zip([nm,gr,ord_,gr/tot_g],
        [None,"₹#,##0","#,##0","0.0%"]),1):
        c = ws.cell(row=ri, column=ci, value=v)
        style_data(c, bg=bg, center=(ci!=1))
        if f: c.number_format=f
    row_h(ws,ri,18)

# ── Active Campaigns ──
section(ws, 27, 1, 6, "ACTIVE CAMPAIGN SUMMARY — Yesterday", bg="ED7D31")
for ci, h in enumerate(["Campaign","Spend (₹)","Purchases","ROAS","CPA (₹)","Status"],1):
    style_sub(ws.cell(row=28,column=ci,value=h)); row_h(ws,28,22)
for i, camp in enumerate(CAMPS):
    ri=29+i; bg=GREEN_LIGHT if camp["status"]=="ACTIVE" else LIGHT_BG
    for ci, (v, f) in enumerate(zip(
        [camp["name"],camp["spend"],camp["purch"],camp["roas"],camp["cpa"],camp["status"]],
        [None,"₹#,##0.00","#,##0","0.00","₹#,##0.00",None]),1):
        c = ws.cell(row=ri, column=ci, value=v)
        style_data(c, bg=bg, center=(ci!=1))
        if f and v is not None: c.number_format=f
    row_h(ws,ri,18)

# ── Key Wins ──
section(ws, 34, 1, 6, "KEY WINS", bg="375623")
ws.cell(row=34,column=1).font=Font(bold=True,color="FFFFFF",size=12)
wins = [
    "Retargeting - New ROAS 3.30 — best campaign of the day",
    "'The Essential' White Co-ord drove 40% of gross revenue (₹45,952)",
    "ASC campaign ROAS 2.97 — holding above 2.5x benchmark",
    "Meta purchase count (46) closely matches Shopify orders (48) — tracking healthy",
]
for i, w in enumerate(wins):
    ri=35+i; ws.merge_cells(f"A{ri}:F{ri}")
    c=ws.cell(row=ri,column=1,value=f"✓  {w}")
    c.fill=fill(GREEN_LIGHT); c.font=font(False,"375623",10)
    c.alignment=align("left","center",True); c.border=border(); row_h(ws,ri,18)

# ── Key Issues ──
section(ws, 40, 1, 6, "KEY ISSUES", bg="C00000")
ws.cell(row=40,column=1).font=Font(bold=True,color="FFFFFF",size=12)
issues = [
    "Gross Sales ₹1,14,082 — 18.1% below 7-day avg ₹1,39,279 (investigate root cause)",
    "Black DHURANDHAR CPA ₹1,114 — highest of all active campaigns (refresh creatives)",
    "Checkout completion 22.7% (5/22) — severe checkout drop-off (review payment UX)",
    "Store conversion rate 0.11% — below e-comm benchmark (add urgency / CRO elements)",
]
for i, iss in enumerate(issues):
    ri=41+i; ws.merge_cells(f"A{ri}:F{ri}")
    c=ws.cell(row=ri,column=1,value=f"⚠  {iss}")
    c.fill=fill(RED_LIGHT); c.font=font(False,"C00000",10)
    c.alignment=align("left","center",True); c.border=border(); row_h(ws,ri,18)

# ── Recommended Actions ──
section(ws, 46, 1, 6, "RECOMMENDED ACTIONS FOR TODAY", bg=DARK_NAVY)
ws.cell(row=46,column=1).font=Font(bold=True,color="FFFFFF",size=12)
actions = [
    "1. SCALE 'Retargeting - New' budget by 20–30% — ROAS 3.30 warrants immediate action",
    "2. REFRESH creatives for 'Black DHURANDHAR' — CPA ₹1,114 is 29% above best campaign",
    "3. AUDIT checkout flow — investigate why only 5/22 shoppers complete purchase",
    "4. VERIFY inventory for 'The Essential' White Co-ord (23 units sold; check all sizes)",
    "5. MONITOR ASC hourly CPP — if <₹800, add ₹5K extra budget by 12 PM IST",
]
for i, act in enumerate(actions):
    ri=47+i; ws.merge_cells(f"A{ri}:F{ri}")
    c=ws.cell(row=ri,column=1,value=act)
    c.fill=fill("EBF3FB"); c.font=font(i==0,DARK_NAVY,10)
    c.alignment=align("left","center",True,); c.border=border(); row_h(ws,ri,22)

for ci, w in enumerate([26,16,14,14,14,16,14,14,14,14,14,14], 1): col_w(ws, ci, w)

# ─────────────────────────────────────────────
# SAVE
# ─────────────────────────────────────────────
out = "/home/user/claude-routines/Daily_Store_Ads_Performance_Sheet_2026-06-04.xlsx"
wb.save(out)
print(f"Saved: {out}")
import os; print("Size:", os.path.getsize(out), "bytes")
