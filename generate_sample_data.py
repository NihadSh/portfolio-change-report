"""Generate two synthetic portfolio snapshots (yesterday and today) for the demo.

The "today" snapshot contains every policy from yesterday plus new sales and a
few cancellations, and - like real exports - some data-quality problems:
duplicated invoice rows, an outdated channel code and sellers attached to the
wrong department. All data is random and fictional.
"""
from pathlib import Path

import numpy as np
import pandas as pd

LOBS = ["Motor Casco", "Agent Casco", "Motor TPL", "Property", "Travel"]
CHANNELS = ["Agents", "Regional", "Key Accounts", "SME", "Partners"]
USAGE = ["Private", "Taxi", "Rental", "Commercial"]


def make_policies(rng, start_id: int, n: int, dates) -> pd.DataFrame:
    lob = rng.choice(LOBS, n, p=[0.35, 0.2, 0.25, 0.1, 0.1])
    casco = np.round(rng.lognormal(6.5, 0.6, n), 2) * np.isin(lob, ["Motor Casco", "Agent Casco"])
    tpl = np.round(rng.uniform(40, 120, n), 2) * np.isin(lob, ["Motor Casco", "Motor TPL"])
    pa = np.round(rng.uniform(10, 40, n), 2) * (lob == "Motor Casco")
    other = np.round(rng.lognormal(5.5, 0.7, n), 2) * np.isin(lob, ["Property", "Travel"])
    sellers = [f"Seller {i:02d}" for i in range(1, 21)]
    seller = rng.choice(sellers, n)
    return pd.DataFrame({
        "policy_number": [f"P-{i:07d}" for i in range(start_id, start_id + n)],
        "invoice_number": [f"INV-{i:07d}" for i in range(start_id, start_id + n)],
        "issue_date": rng.choice(dates, n),
        "lob": lob,
        "client_type": rng.choice(["Individual", "Corporate"], n, p=[0.75, 0.25]),
        "channel": rng.choice(CHANNELS, n),
        "seller": seller,
        "department": [f"Department {int(s[-2:]) % 5 + 1}" for s in seller],
        "usage": np.where(np.isin(lob, ["Motor Casco", "Agent Casco"]), rng.choice(USAGE, n, p=[0.7, 0.15, 0.1, 0.05]), ""),
        "status": "Active",
        "casco_premium": casco, "pa_premium": pa, "tpl_premium": tpl, "other_premium": other,
    }).assign(gwp=lambda d: d[["casco_premium", "pa_premium", "tpl_premium", "other_premium"]].sum(axis=1).round(2))


def main(out_dir: Path, seed: int = 11) -> None:
    rng = np.random.default_rng(seed)
    out_dir.mkdir(parents=True, exist_ok=True)
    old_days = pd.date_range("2026-11-01", "2026-11-10")
    today = pd.Timestamp("2026-11-11")

    old = make_policies(rng, 1, 3000, old_days)

    new = make_policies(rng, 3001, 180, [today])
    new.loc[new.sample(6, random_state=1).index, "channel"] = "KK"   # outdated channel code
    new = pd.concat([new, new.sample(4, random_state=2)])           # duplicated invoice rows

    cancelled = old.sample(25, random_state=3).copy()                # cancellations issued today
    cancelled["status"] = "Cancelled"
    cancelled["issue_date"] = today

    current = pd.concat([old.drop(cancelled.index), cancelled, new]).sample(frac=1, random_state=4)
    old.to_excel(out_dir / "portfolio_2026-11-10.xlsx", index=False)
    current.to_excel(out_dir / "portfolio_2026-11-11.xlsx", index=False)

    # Corrections maintained by the business (instead of hard-coding them in the script).
    pd.DataFrame([
        {"field": "channel", "match_field": "channel", "match_value": "KK", "new_value": "SME"},
        {"field": "department", "match_field": "seller", "match_value": "Seller 07", "new_value": "Department 3"},
        {"field": "department", "match_field": "seller", "match_value": "Seller 14", "new_value": "Department 1"},
        {"field": "usage", "match_field": "usage", "match_value": "Rental", "new_value": "Taxi"},
    ]).to_csv(out_dir / "corrections.csv", index=False)
    print(f"Sample data written to {out_dir.resolve()}")


if __name__ == "__main__":
    main(Path(__file__).parent / "data")
