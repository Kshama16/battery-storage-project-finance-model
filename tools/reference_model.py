"""Independent Python implementation of the Excel model's logic.

Used to cross-check the workbook (tools/verify_workbook.py), to solve for the
capacity price that hits the target equity IRR, and to pre-compute nothing that
the workbook does not also calculate live.

    python tools/reference_model.py            # base case summary + required price
"""
from dataclasses import dataclass, field

import numpy as np
import numpy_financial as npf

from assumptions import ASSETS, FINANCING, GLOBAL, MACRS, MODEL_START, MODEL_YEARS, SCENARIOS

N_MONTHS = MODEL_YEARS * 12
YEARS = np.arange(MODEL_START.year, MODEL_START.year + MODEL_YEARS)


def month_index(d):
    return (d.year - MODEL_START.year) * 12 + d.month - MODEL_START.month


@dataclass
class Case:
    capex_adj: float = 0.0     # e.g. 0.10 = +10% capex
    price_adj: float = 0.0     # applied to the capacity price only
    delay: int = 0             # months added to every COD (capex timing unchanged)
    rate_bps: float = 0.0      # added to senior and mezz rates
    om_adj: float = 0.0        # applied to fixed O&M
    capacity_price: float = field(default_factory=lambda: GLOBAL["capacity_price"])

    @classmethod
    def scenario(cls, name, **overrides):
        s = SCENARIOS[name]
        c = cls(capex_adj=s["capex"], price_adj=s["price"], delay=s["delay"], rate_bps=s["rate"], om_adj=s["om"])
        for k, v in overrides.items():
            setattr(c, k, v)
        return c


def asset_capex(a, case):
    unit_cost_per_kw = a["cost"] * (a["hours"] if a["unit"] == "$/kWh" else 1)
    return a["mw"] * unit_cost_per_kw * (1 + GLOBAL["contingency"]) * (1 + case.capex_adj)  # $k


def run(case=None):
    case = case or Case()
    g, f = GLOBAL, FINANCING
    m = np.arange(N_MONTHS)
    yi = m // 12  # years since model start

    rev_cap = np.zeros(N_MONTHS)
    rev_energy = np.zeros(N_MONTHS)
    opex = np.zeros(N_MONTHS)
    capex = np.zeros(N_MONTHS)
    any_op = np.zeros(N_MONTHS, dtype=bool)
    itc_cash = np.zeros(MODEL_YEARS)
    tax_dep = np.zeros(MODEL_YEARS)
    cod_years = []

    price = case.capacity_price * (1 + case.price_adj) * (1 + g["price_escalation"]) ** yi
    om_idx = (1 + g["om_escalation"]) ** yi

    for a in ASSETS:
        cx = asset_capex(a, case)
        cod = month_index(a["cod"])
        cod_eff = cod + case.delay
        cod_year_idx = cod_eff // 12
        cod_years.append(MODEL_START.year + cod_year_idx)

        ms = m - cod_eff
        op = (ms >= 0) & (ms < a["life"] * 12)
        any_op |= op
        op_year = np.where(op, ms // 12, 0)
        acc_mw = op * a["mw"] * a["accred"] * (1 - a["degr"]) ** op_year

        rev_cap += acc_mw * price
        rev_energy += acc_mw * a["margin"] / 12 * (1 + g["price_escalation"]) ** yi
        opex += op * (a["mw"] * a["fom"] / 12 * om_idx * (1 + case.om_adj)
                      + cx * (g["insurance"] + g["property_tax"]) / 12
                      + a["lease"] / 12 * om_idx)
        capex += ((m >= cod - a["build"]) & (m < cod)) * cx / a["build"]

        if a["itc"]:
            if cod_year_idx + 1 < MODEL_YEARS:
                itc_cash[cod_year_idx + 1] += cx * g["itc_rate"] * g["itc_transfer_price"]
        basis = cx * (1 - g["itc_basis_reduction"] * g["itc_rate"] * a["itc"])
        for k, rate in enumerate(MACRS[a["macrs"]]):
            if cod_year_idx + k < MODEL_YEARS:
                tax_dep[cod_year_idx + k] += basis * rate

    opex += any_op * g["ga_cost"] / 12 * om_idx

    def annual(x):
        return x.reshape(MODEL_YEARS, 12).sum(axis=1)

    rev = annual(rev_cap) + annual(rev_energy)
    ebitda = rev - annual(opex)
    capex_y = annual(capex)
    total_capex = capex_y.sum()

    # Debt: sculpted to target DSCR on pre-tax CFADS (= EBITDA), drawn at term conversion
    conv = max(cod_years)
    t = YEARS - conv
    dflag = (t >= 1) & (t <= f["tenor"])
    rs = f["senior_rate"] + case.rate_bps / 1e4
    rm = f["mezz_rate"] + case.rate_bps / 1e4

    ds_s_unc = dflag * ebitda / f["senior_dscr"]
    pv_s = (ds_s_unc / (1 + rs) ** t).sum()
    senior = min(pv_s, f["senior_max_gearing"] * total_capex)
    ds_s = ds_s_unc * senior / pv_s

    ds_m_unc = dflag * ebitda / f["combined_dscr"] - ds_s
    pv_m = (ds_m_unc / (1 + rm) ** t).sum()
    mezz = max(0.0, min(pv_m, f["total_max_gearing"] * total_capex - senior))
    ds_m = ds_m_unc * mezz / pv_m

    def schedule(amount, ds, r):
        bal, interest, closing = 0.0, np.zeros(MODEL_YEARS), np.zeros(MODEL_YEARS)
        for i, y in enumerate(YEARS):
            opening = bal
            draw = amount if y == conv else 0.0
            interest[i] = opening * r
            bal = opening + draw - (ds[i] - interest[i])
            closing[i] = bal
        return interest, closing

    int_s, bal_s = schedule(senior, ds_s, rs)
    int_m, bal_m = schedule(mezz, ds_m, rm)

    draws = (YEARS == conv) * (senior + mezz)
    fees = (YEARS == conv) * (senior * f["senior_fee"] + mezz * f["mezz_fee"])
    tax = g["tax_rate"] * (ebitda - int_s - int_m - tax_dep)
    unlev_tax = g["tax_rate"] * (ebitda - tax_dep)

    project_cf = -capex_y + ebitda - unlev_tax + itc_cash
    equity_cf = -capex_y + draws - fees + ebitda - ds_s - ds_m - tax + itc_cash

    equity = total_capex - senior - mezz
    t_ = g["tax_rate"]
    wacc = (senior * rs * (1 - t_) + mezz * rm * (1 - t_) + equity * f["cost_of_equity"]) / total_capex

    def npv(rate, cf):  # Excel NPV(): first value discounted one period
        return float((cf / (1 + rate) ** np.arange(1, len(cf) + 1)).sum())

    dscr_s = ebitda[dflag] / ds_s[dflag]
    dscr_c = ebitda[dflag] / (ds_s + ds_m)[dflag]
    cum = np.cumsum(equity_cf)
    payback = int(YEARS[np.argmax(cum >= 0)]) if (cum >= 0).any() else None

    return dict(
        total_capex=total_capex, senior=senior, mezz=mezz, equity=equity,
        gearing=(senior + mezz) / total_capex, wacc=wacc,
        project_irr=float(npf.irr(project_cf)), equity_irr=float(npf.irr(equity_cf)),
        project_npv=npv(wacc, project_cf), equity_npv=npv(f["cost_of_equity"], equity_cf),
        min_dscr_senior=float(dscr_s.min()), min_dscr_combined=float(dscr_c.min()),
        avg_dscr_senior=float(dscr_s.mean()), payback_year=payback, conversion_year=conv,
        itc_cash=float(itc_cash.sum()), revenue=rev, ebitda=ebitda, capex=capex_y,
        project_cf=project_cf, equity_cf=equity_cf, senior_balance=bal_s, mezz_balance=bal_m,
    )


def required_capacity_price(case=None, target=None, lo=1.0, hi=60.0):
    """Capacity price ($/kW-month, 2027 $) at which equity IRR equals the target (bisection)."""
    case = case or Case()
    target = FINANCING["cost_of_equity"] if target is None else target
    for _ in range(100):
        mid = (lo + hi) / 2
        case.capacity_price = mid
        if run(case)["equity_irr"] < target:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


if __name__ == "__main__":
    for name in SCENARIOS:
        r = run(Case.scenario(name))
        print(f"{name:9s} capex {r['total_capex']:>9,.0f}  senior {r['senior']:>9,.0f}  mezz {r['mezz']:>8,.0f}  "
              f"gearing {r['gearing']:.1%}  WACC {r['wacc']:.2%}  projIRR {r['project_irr']:.2%}  "
              f"eqIRR {r['equity_irr']:.2%}  NPV {r['project_npv']:>9,.0f}  minDSCR {r['min_dscr_senior']:.2f}/"
              f"{r['min_dscr_combined']:.2f}  payback {r['payback_year']}")
    print(f"Required capacity price (base, eq IRR = {FINANCING['cost_of_equity']:.0%}): "
          f"${required_capacity_price():.2f}/kW-month")
    b = run()
    print("EBITDA by year:", np.round(b["ebitda"]).astype(int).tolist())
    print("Senior closing balance last debt year:", round(b["senior_balance"][-1], 4))
