"""Pins the headline numbers in the deck. Run: python -m pytest -q"""
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
subprocess.run([sys.executable, str(ROOT / "src" / "derive.py")], check=True, capture_output=True)
D = json.loads((ROOT / "data" / "derived.json").read_text())


def test_tcs_base():
    t = D["tcs"]
    assert t["ai_runrate_bn"] == 3.1 and t["ai_share_pct"] == 10.1
    assert t["core_growth_pct"] == -3.6 and t["book_to_bill"] == 1.26
    assert abs(t["ai_runrate_bn"] + t["core_runrate_bn"] - t["revenue_runrate_bn"]) < 0.02


def test_arithmetic():
    assert D["ai_needed_flat_pct"] == 32
    assert D["sensitivity"]["rows"][0]["cells"] == [-0.2, 2.8, 6.9]
    assert [r["ai_growth_pct"] for r in D["ai_needed_for_8pct"]] == [111, 92, 79, 61]
    assert D["elasticity"] == {"total_pts_per_core_pt": 0.9, "total_pts_per_10pt_ai": 1.01}


def test_project_links():
    assert D["p1"] == {"events_total": 80, "events_verified": 80, "tcs_verified": 11}
    assert D["market"]["spearman_events_vs_growth"] == -0.57
    assert D["rpe_peer_median_k"] == 63.0


def test_ceo_page():
    c = D["ceo"]
    assert c == {"rpe_1k_value_bn": 0.6, "core_fix_total_pts": 3.2, "core_fix_value_bn": 1.0,
                 "core_fix_equiv_ai_pts": 32, "core_vs_ai_ratio": 9, "growth_rank_cc": 5, "btb_rank": 2}
