"""Derive every number the deck uses from the Project 1 and Project 2 inputs.

Run:  python src/derive.py   ->  data/derived.json
Nothing in the deck is typed by hand: build_deck.js reads only data/derived.json.
"""
import csv
import json
from pathlib import Path
from statistics import median

ROOT = Path(__file__).resolve().parents[1]
IN = ROOT / "data" / "inputs"


def rows(name):
    with open(IN / name, encoding="utf8") as f:
        return list(csv.DictReader(f))


def by_company(name):
    return {r["company"]: r for r in rows(name)}


share = by_company("01_market_share.csv")
pool = by_company("02_ai_pool.csv")
core = by_company("03_core_vs_ai.csv")
prod = by_company("04_productivity.csv")
mom = by_company("05_momentum_vs_money.csv")
score = by_company("06_scorecard.csv")
stats = json.loads((IN / "stats.json").read_text())
scen = rows("scenarios.csv")
p1 = rows("p1_ci_events.csv")
tcs_ai = rows("tcs_ai_disclosures.csv")

FIRMS = ["Accenture", "TCS", "Capgemini", "Cognizant", "Infosys", "HCLTech", "Wipro", "Tech Mahindra"]
f = float

# --- TCS base (calendar Q3 2026 = TCS Q2 FY27; run-rate = quarter x 4) ---
c = core["TCS"]
tcs_total = f(c["revenue_end_usd_m"]) * 4 / 1000          # $bn
tcs_ai_rr = f(pool["TCS"]["ai_runrate_mid_usd_bn"])        # $bn, company-reported annualised
tcs_core = f(c["core_end_usd_m"]) * 4 / 1000               # $bn
core_g = f(c["core_growth_annualised_pct"])                # %/yr

# AI growth needed just to hold total revenue flat if the core keeps shrinking at today's pace
ai_needed_flat = -core_g / 100 * tcs_core / tcs_ai_rr * 100

# Sensitivity: one-year total growth for combinations of core growth and AI growth
CORE_CASES = [core_g, -1.5, 0.0, 2.0]
AI_CASES = [30, 60, 100]
sens = [
    {
        "core_growth_pct": cg,
        "cells": [round((tcs_ai_rr * ag / 100 + tcs_core * cg / 100) / tcs_total * 100, 1) for ag in AI_CASES],
    }
    for cg in CORE_CASES
]
# AI growth needed for 8% total growth (management's "high-single-digit" ambition, N4) at each core case
ai_needed_8 = [{"core_growth_pct": cg, "ai_growth_pct": round((0.08 * tcs_total - tcs_core * cg / 100) / tcs_ai_rr * 100)} for cg in CORE_CASES]
# Elasticities: total-growth points per 1pt of core growth and per 10pts of AI growth
per_core_pt = tcs_core / tcs_total
per_ai_10pt = tcs_ai_rr * 0.10 / tcs_total * 100

ai_series = [{"period": r["period_label"], "usd_bn": f(r["value"]) / 1000} for r in tcs_ai]
ai_3q_growth = (ai_series[-1]["usd_bn"] / ai_series[0]["usd_bn"] - 1) * 100

rpe = {k: f(prod[k]["rev_per_employee_k_usd"]) for k in FIRMS}
p1_tcs = [r for r in p1 if r["company"] == "TCS" and r["validation_status"] == "Verified"]
p1_verified = [r for r in p1 if r["validation_status"] == "Verified"]

# Largest disclosed acquisition values in Project 1 (USD m)
def deal(company, counterparty):
    for r in p1:
        if r["company"] == company and counterparty in r["counterparty"] and r["deal_value_usd_m"]:
            return f(r["deal_value_usd_m"])
    return None

out = {
    "basis": "Latest quarter = calendar Q3 2026 (TCS Q2 FY27); run-rate = quarter x 4; USD",
    "tcs": {
        "revenue_runrate_bn": round(tcs_total, 2),
        "ai_runrate_bn": tcs_ai_rr,
        "ai_share_pct": f(pool["TCS"]["ai_share_of_revenue_pct"]),
        "core_runrate_bn": round(tcs_core, 2),
        "core_growth_pct": core_g,
        "ai_growth_annualised_pct": f(c["ai_growth_annualised_pct"]),
        "ai_3q_growth_pct": round(ai_3q_growth),
        "ai_series": ai_series,
        "yoy_cc_pct": f(share["TCS"]["yoy_cc_pct"]),
        "yoy_usd_pct": f(share["TCS"]["yoy_usd_pct"]),
        "book_to_bill": f(score["TCS"]["book_to_bill"]),
        "tcv_q_bn": 9.6,  # TCS Q2 FY27 press release (TCS_Q2FY27_QUARTR); bookings_q_usd_m = 9600 in Project 2
        "rev_per_employee_k": rpe["TCS"],
        "headcount": int(prod["TCS"]["headcount"]),
        "headcount_yoy_pct": f(prod["TCS"]["headcount_yoy_pct"]),
        "share_of_gartner_pct": f(share["TCS"]["share_of_gartner_it_services_pct"]),
        "share_of_eight_pct": f(share["TCS"]["share_of_eight_pct"]),
        "p1_events": int(mom["TCS"]["p1_events"]),
        "p1_partnerships": int(mom["TCS"]["partnerships"]),
        "p1_frontier_labs": mom["TCS"]["frontier_labs"],
        "p1_frontier_lab_events": int(mom["TCS"]["frontier_lab_events"]),
        "largest_acq_usd_m": deal("TCS", "Coastal Cloud"),
    },
    "ai_needed_flat_pct": round(ai_needed_flat),
    "ai_needed_for_8pct": ai_needed_8,
    "ceo": {
        # Value of each $1k of revenue per employee at TCS's headcount (USD bn)
        "rpe_1k_value_bn": round(int(prod["TCS"]["headcount"]) * 1000 / 1e9, 1),
        # Fixing the core (from today's decline to 0%): total-growth points and revenue per year
        "core_fix_total_pts": round(-core_g * per_core_pt, 1),
        "core_fix_value_bn": round(-core_g / 100 * tcs_core, 1),
        "core_fix_equiv_ai_pts": round(-core_g * per_core_pt / per_ai_10pt * 10),
        "core_vs_ai_ratio": round(per_core_pt / (per_ai_10pt / 10)),
        "growth_rank_cc": int(mom["TCS"]["rank_growth"]),
        "btb_rank": sorted([f(score[k]["book_to_bill"]) for k in FIRMS], reverse=True).index(f(score["TCS"]["book_to_bill"])) + 1,
    },
    "sensitivity": {"ai_cases_pct": AI_CASES, "rows": sens},
    "elasticity": {"total_pts_per_core_pt": round(per_core_pt, 2), "total_pts_per_10pt_ai": round(per_ai_10pt, 2)},
    "peers": {
        k: {
            "core_growth_pct": f(core[k]["core_growth_annualised_pct"]) if k in core else None,
            "ai_growth_pct": f(core[k]["ai_growth_annualised_pct"]) if k in core else None,
            "yoy_cc_pct": f(share[k]["yoy_cc_pct"]),
            "p1_events": int(mom[k]["p1_events"]),
            "book_to_bill": f(score[k]["book_to_bill"]),
            "rev_per_employee_k": rpe[k],
        }
        for k in FIRMS
    },
    "rpe_peer_median_k": round(median(v for k, v in rpe.items() if k != "TCS"), 1),  # the seven peers, excluding TCS
    "market": {
        "eight_yoy_usd_pct": round(stats["eight_firm_yoy_usd_pct"], 1),
        "gartner_it_services_growth_pct": round(stats["gartner_it_services_growth_2026_pct"], 1),
        "spearman_events_vs_growth": stats["spearman_p1_events_vs_growth"],
        "spearman_ai_vs_growth": stats["spearman_ai_intensity_vs_growth_all8"],
        "pool_mid_bn": f(pool["TCS"]["pool_mid_usd_bn"]),
        "eight_revenue_bn": stats["eight_firm_revenue_runrate_usd_bn"],
    },
    "scenarios_2030_bn": {s: round(f(r["total_usd_bn"])) for r in scen if r["year"] == "2030" for s in [r["scenario"]]},
    "deals": {
        "capgemini_wns_usd_m": deal("Capgemini", "WNS"),
        "accenture_fy26_acq_usd_m": next(f(r["deal_value_usd_m"]) for r in p1 if r["event_id"] == "ACN-005"),
        "cognizant_openai_staff": 25000,  # Project 1 event CTS-009 headline: 25,000+ staff
        "cognizant_gemini_staff": 18000,  # Project 1 event CTS-008 headline: 18,000-strong group
    },
    "p1": {"events_total": len(p1), "events_verified": len(p1_verified), "tcs_verified": len(p1_tcs)},
}

(ROOT / "data" / "derived.json").write_text(json.dumps(out, indent=2))
print(json.dumps(out, indent=2))
