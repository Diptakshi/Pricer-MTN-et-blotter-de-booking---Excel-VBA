"""Independent Python re-implementation of the workbook, checked against the values Excel
(LibreOffice) computed and cached in MTN_Pricer.xlsx."""
from math import exp, log
from pathlib import Path

import pytest
from openpyxl import load_workbook

WB = load_workbook(Path(__file__).parent.parent / "MTN_Pricer.xlsx", data_only=True)
M, P, C = WB["Market"], WB["Pricer"], WB["Curve"]


def market():
    rows = [(M.cell(r, 1).value, M.cell(r, 2).value / 100, M.cell(r, 3).value / 1e4)
            for r in range(7, 7 + 30) if M.cell(r, 1).value is not None]
    return [x[0] for x in rows], [x[1] for x in rows], [x[2] for x in rows]


def interp(xs, ys, t):
    if t <= xs[0]:
        return ys[0]
    if t >= xs[-1]:
        return ys[-1]
    for i in range(len(xs) - 1):
        if xs[i] <= t <= xs[i + 1]:
            return ys[i] + (t - xs[i]) / (xs[i + 1] - xs[i]) * (ys[i + 1] - ys[i])


def curve(shift=0.0):
    ten, par, spr = market()
    df, dfr, fwd, ann = [1.0], [], [], 0.0
    for t in range(1, 31):
        s = interp(ten, par, t) + shift
        d = (1 - s * ann) / (1 + s)
        ann += d
        fwd.append(df[-1] / d - 1)
        df.append(d)
        dfr.append(d * exp(-interp(ten, spr, t) * t))
    return df[1:], dfr, fwd


DF, DFR, FWD = curve()
MAT, CPN, NOM = P["B7"].value, P["B8"].value, P["B6"].value


def test_bootstrap_reprices_every_par_swap():
    ten, par, _ = market()
    for t in range(1, 31):
        s = interp(ten, par, t)
        assert s * sum(DF[:t]) + DF[t - 1] == pytest.approx(1.0, abs=1e-12)


def test_excel_curve_matches_python():
    for t in range(1, 31):
        assert C.cell(3 + t, 4).value == pytest.approx(DF[t - 1], abs=1e-12)
        assert C.cell(3 + t, 7).value == pytest.approx(DFR[t - 1], abs=1e-12)


def test_par_coupon_prices_at_100():
    par = (1 - DFR[MAT - 1]) / sum(DFR[:MAT]) * 100
    assert P["B12"].value == pytest.approx(par, abs=1e-9)
    price = (par / 100 * sum(DFR[:MAT]) + DFR[MAT - 1]) * 100
    assert price == pytest.approx(100, abs=1e-9)


def test_price_at_proposed_coupon():
    assert P["B13"].value == pytest.approx((CPN / 100 * sum(DFR[:MAT]) + DFR[MAT - 1]) * 100, abs=1e-9)


def test_frn_margin_is_zero_without_funding_spread():
    df, _, fwd = curve()
    m = (1 - df[MAT - 1] - sum(f * d for f, d in zip(fwd[:MAT], df[:MAT]))) / sum(df[:MAT])
    assert m == pytest.approx(0.0, abs=1e-12)        # a risk-free FRN at Euribor flat is worth par


def test_frn_margin_matches_excel():
    m = (1 - DFR[MAT - 1] - sum(f * d for f, d in zip(FWD[:MAT], DFR[:MAT]))) / sum(DFR[:MAT]) * 1e4
    assert P["B16"].value == pytest.approx(m, abs=1e-6)


def test_all_in_cost():
    ten, par, _ = market()
    assert P["B15"].value == pytest.approx((CPN / 100 - interp(ten, par, MAT)) * 1e4, abs=1e-9)


def test_dv01_duration_close_to_curve_bump():
    _, up, _ = curve(+1e-4)
    _, dn, _ = curve(-1e-4)
    price = lambda d: (CPN / 100 * sum(d[:MAT]) + d[MAT - 1]) * 100
    bump = (price(dn) - price(up)) / 2 / 100 * NOM
    assert P["B18"].value == pytest.approx(bump, rel=0.05)     # duration vs par-curve bump
