"""Base-case assumptions shared by the Excel builder and the Python reference model.

All figures are illustrative portfolio assumptions, not data from a real project.
Currency is nominal USD; model outputs are in $ thousands ($k).
"""
from datetime import date

MODEL_START = date(2027, 1, 1)
MODEL_YEARS = 23  # Jan 2027 - Dec 2049 covers a 20-year life for the latest COD plus a 12-month delay

GLOBAL = {
    "capacity_price": 11.00,        # $/kW-month of accredited capacity (2027 $)
    "price_escalation": 0.020,      # capacity price and energy margin escalation, per year
    "om_escalation": 0.025,         # O&M, lease and G&A escalation, per year
    "contingency": 0.05,            # owner's costs and contingency, % of EPC capex
    "insurance": 0.004,             # % of total capex per year
    "property_tax": 0.005,          # % of total capex per year
    "ga_cost": 1000.0,              # portfolio G&A and asset management, $k/yr (2027 $)
    "tax_rate": 0.26,               # blended federal + state
    "itc_rate": 0.30,               # investment tax credit on eligible (storage) capex
    "itc_transfer_price": 0.92,     # $ received per $1 of credit sold (transferability)
    "itc_basis_reduction": 0.50,    # depreciable basis reduced by 50% of the ITC
}

FINANCING = {
    "senior_rate": 0.065,
    "senior_dscr": 1.35,
    "senior_max_gearing": 0.60,     # % of total capex
    "senior_fee": 0.015,            # upfront fee, % of amount
    "mezz_rate": 0.105,
    "combined_dscr": 1.20,          # senior + mezz debt service
    "total_max_gearing": 0.75,
    "mezz_fee": 0.020,
    "tenor": 15,                    # years, both tranches, from term conversion
    "cost_of_equity": 0.110,        # target sponsor equity return
}

# capex_unit: "$/kWh" multiplies by duration; "$/kW" does not
ASSETS = [
    # name, tech, MW, hours, capex cost, capex unit, construction months, COD,
    # accreditation, degradation/yr, fixed O&M $/kW-yr, energy & ancillary margin $/kW-yr,
    # site lease $k/yr, ITC eligible, MACRS life, useful life yrs
    dict(name="Alpha", tech="BESS", mw=100, hours=4, cost=330, unit="$/kWh", build=11, cod=date(2027, 12, 1),
         accred=0.90, degr=0.020, fom=28, margin=45, lease=150, itc=1, macrs=5, life=20),
    dict(name="Bravo", tech="BESS", mw=75, hours=4, cost=335, unit="$/kWh", build=12, cod=date(2028, 1, 1),
         accred=0.90, degr=0.020, fom=28, margin=45, lease=120, itc=1, macrs=5, life=20),
    dict(name="Charlie", tech="BESS", mw=50, hours=2, cost=390, unit="$/kWh", build=9, cod=date(2028, 4, 1),
         accred=0.60, degr=0.025, fom=24, margin=35, lease=90, itc=1, macrs=5, life=20),
    dict(name="Delta", tech="BESS", mw=125, hours=4, cost=325, unit="$/kWh", build=14, cod=date(2028, 7, 1),
         accred=0.90, degr=0.020, fom=27, margin=45, lease=180, itc=1, macrs=5, life=20),
    dict(name="Echo", tech="Backup gen", mw=50, hours=0, cost=1350, unit="$/kW", build=15, cod=date(2028, 7, 1),
         accred=0.95, degr=0.000, fom=22, margin=10, lease=80, itc=0, macrs=15, life=20),
]

# Live-case scenario drivers: capex adj, capacity price adj, COD delay (months), rate adj (bps), O&M adj
SCENARIOS = {
    "Base":     dict(capex=0.00, price=0.00, delay=0, rate=0, om=0.00),
    "Downside": dict(capex=0.10, price=-0.10, delay=6, rate=100, om=0.10),
    "Upside":   dict(capex=-0.05, price=0.05, delay=0, rate=-50, om=-0.05),
}

MACRS = {
    5: [0.2000, 0.3200, 0.1920, 0.1152, 0.1152, 0.0576],
    15: [0.0500, 0.0950, 0.0855, 0.0770, 0.0693, 0.0623, 0.0590, 0.0590,
         0.0591, 0.0590, 0.0591, 0.0590, 0.0591, 0.0590, 0.0591, 0.0295],
}
