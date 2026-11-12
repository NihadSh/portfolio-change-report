"""Daily portfolio change report.

Compares today's portfolio export with yesterday's and reports what changed:
  * NEW          - policies that were not in yesterday's file
  * CANCELLED    - cancellations issued on the latest date in today's file
  * corrections  - channel / department / usage fixes from corrections.csv
  * data quality - duplicated invoice rows are removed and counted

Output sheets: Summary, Changes (row level), By LOB, By channel (incl. lost
premium from cancellations), Casco by usage, Seller pivot.

Usage:
    python portfolio_change_report.py data/portfolio_2026-11-10.xlsx data/portfolio_2026-11-11.xlsx
"""
import argparse
from pathlib import Path

import pandas as pd

BASE = Path(__file__).parent
CASCO_LOBS = ["Motor Casco", "Agent Casco"]
PREMIUM_COLS = ["casco_premium", "pa_premium", "tpl_premium", "other_premium"]


def load(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = df.columns.str.strip()
    for c in PREMIUM_COLS + ["gwp"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["issue_date"] = pd.to_datetime(df["issue_date"], errors="coerce")
    return df


def apply_corrections(df: pd.DataFrame, rules: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    df = df.copy()
    changed = 0
    for r in rules.itertuples(index=False):
        mask = df[r.match_field].astype(str) == str(r.match_value)
        changed += int((df.loc[mask, r.field] != r.new_value).sum())
        df.loc[mask, r.field] = r.new_value
    return df, changed


def find_changes(old: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    added = new[~new["policy_number"].isin(old["policy_number"])].assign(change="NEW")
    last_date = new["issue_date"].max()
    cancelled = new[(new["status"] == "Cancelled") & (new["issue_date"] == last_date)
                    & ~new["policy_number"].isin(added["policy_number"])].assign(change="CANCELLED")
    return pd.concat([added, cancelled], ignore_index=True)


def by_lob(ch: pd.DataFrame) -> pd.DataFrame:
    g = ch.groupby(["lob", "change"]).agg(policies=("policy_number", "nunique"), gwp=("gwp", "sum"))
    out = g.unstack("change", fill_value=0)
    out.columns = [f"{chg.lower()}_{m}" for m, chg in out.columns]
    return out.reset_index()


def by_channel(ch: pd.DataFrame, old: pd.DataFrame) -> pd.DataFrame:
    casco = ch[ch["lob"].isin(CASCO_LOBS)]
    new = casco[casco["change"] == "NEW"].groupby(["channel", "client_type"]).agg(
        new_policies=("policy_number", "nunique"), new_gwp=("gwp", "sum"))
    canc = casco[casco["change"] == "CANCELLED"][["channel", "client_type", "policy_number"]]
    # premium lost = the original premium of the cancelled policy (from yesterday's file)
    lost = canc.merge(old[["policy_number", "gwp"]], on="policy_number", how="left")
    lost = lost.groupby(["channel", "client_type"]).agg(cancelled=("policy_number", "nunique"), lost_gwp=("gwp", "sum"))
    out = new.join(lost, how="outer").fillna(0).reset_index()
    out["net_gwp"] = out["new_gwp"] - out["lost_gwp"]
    return out


def casco_usage(ch: pd.DataFrame) -> pd.DataFrame:
    casco = ch[ch["lob"].isin(CASCO_LOBS) & (ch["change"] == "NEW")]
    return casco.pivot_table(index=["client_type", "usage"], values=["policy_number", "gwp"],
                             aggfunc={"policy_number": "nunique", "gwp": "sum"}, fill_value=0) \
        .rename(columns={"policy_number": "policies"}).reset_index()


def seller_pivot(ch: pd.DataFrame) -> pd.DataFrame:
    new = ch[ch["change"] == "NEW"]
    return new.groupby(["channel", "department", "seller"]).agg(
        policies=("policy_number", "nunique"), gwp=("gwp", "sum")).reset_index() \
        .sort_values("gwp", ascending=False)


def write(out: Path, sheets: dict[str, pd.DataFrame]) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as xw:
        head = xw.book.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1})
        money = xw.book.add_format({"num_format": "#,##0.00"})
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for c, col in enumerate(df.columns):
                ws.write(0, c, col, head)
                is_money = "gwp" in col or "premium" in col
                ws.set_column(c, c, max(12, len(col) + 3), money if is_money else None)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(df), len(df.columns) - 1)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("previous", type=Path)
    ap.add_argument("current", type=Path)
    ap.add_argument("--corrections", type=Path, default=BASE / "data" / "corrections.csv")
    ap.add_argument("--out", type=Path, default=None)
    a = ap.parse_args()

    old, new = load(a.previous), load(a.current)

    before = len(new)
    new = new.drop_duplicates(subset="invoice_number", keep="first")
    duplicates = before - len(new)

    new, corrected = apply_corrections(new, pd.read_csv(a.corrections))
    changes = find_changes(old, new)
    report_date = new["issue_date"].max()

    n_new = changes["change"].eq("NEW")
    summary = pd.DataFrame({
        "metric": ["Report date", "Policies yesterday", "Policies today", "New policies", "New GWP",
                   "Cancellations", "Lost GWP (cancellations)", "Duplicate invoice rows removed", "Values corrected"],
        "value": [report_date.strftime("%Y-%m-%d"), old["policy_number"].nunique(), new["policy_number"].nunique(),
                  int(n_new.sum()), round(changes.loc[n_new, "gwp"].sum(), 2), int((~n_new).sum()),
                  round(old.loc[old["policy_number"].isin(changes.loc[~n_new, "policy_number"]), "gwp"].sum(), 2),
                  duplicates, corrected],
    })

    out = a.out or BASE / "output" / f"portfolio_changes_{report_date:%Y-%m-%d}.xlsx"
    write(out, {
        "Summary": summary,
        "Changes": changes[["change"] + [c for c in changes.columns if c != "change"]],
        "By LOB": by_lob(changes),
        "By channel": by_channel(changes, old),
        "Casco by usage": casco_usage(changes),
        "Seller pivot": seller_pivot(changes),
    })
    print(summary.to_string(index=False))
    print(f"\nReport saved to {out.resolve()}")


if __name__ == "__main__":
    main()
