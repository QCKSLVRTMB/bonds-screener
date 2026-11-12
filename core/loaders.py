from __future__ import annotations
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta

import pandas as pd
import streamlit as st

from config import (BOND_BOARDS, BOARD_PRIORITY, BOARD_WORKERS, ISS_COLUMNS,
                    TTL_MARKET, TTL_HISTORY, HISTORY_LOOKBACK_DAYS)
from core.http import fetch_paged

KEYS = ["SECID", "BOARDID"]


def _load_board_raw(board: str) -> pd.DataFrame:
    """securities и marketdata приходят в одном ответе — merge вместо 2 загрузок."""
    res = fetch_paged(
        f"engines/stock/markets/bonds/boards/{board}/securities",
        ("securities", "marketdata"), columns=ISS_COLUMNS,
    )
    sec, mkt = res["securities"], res["marketdata"]
    if sec.empty:
        return sec
    if "BOARDID" not in sec:
        sec["BOARDID"] = board
    if not mkt.empty:
        if "BOARDID" not in mkt:
            mkt["BOARDID"] = board
        sec = sec.merge(mkt.drop_duplicates(KEYS), on=KEYS, how="left",
                        suffixes=("", "_mkt"))
    return sec


@st.cache_data(ttl=TTL_MARKET, show_spinner=False, persist="disk")
def load_universe(boards: tuple[str, ...] = BOND_BOARDS) -> pd.DataFrame:
    with ThreadPoolExecutor(max_workers=BOARD_WORKERS) as ex:
        frames = [f for f in ex.map(_load_board_raw, boards) if not f.empty]
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)
    df = df.loc[:, ~df.columns.duplicated()]
    df["_prio"] = df["BOARDID"].map(BOARD_PRIORITY).fillna(99)
    df = (df.sort_values("_prio")
            .drop_duplicates("SECID", keep="first")
            .drop(columns="_prio")
            .reset_index(drop=True))
    return df


# ---------- ИСТОРИЯ: один набор запросов на доску вместо запроса на бумагу ----------

def _board_closes_raw(board: str, d_from: date, d_to: date) -> pd.DataFrame:
    return fetch_paged(
        f"history/engines/stock/markets/bonds/boards/{board}/securities",
        ("history",),
        params={"from": d_from.isoformat(), "till": d_to.isoformat()},
        columns={"history": ISS_COLUMNS["history"]},
    )["history"]


@st.cache_data(ttl=TTL_HISTORY, show_spinner=False, persist="disk")
def load_closes(boards: tuple[str, ...], target: date,
                lookback: int = HISTORY_LOOKBACK_DAYS) -> pd.Series:
    """SECID -> последняя цена закрытия на дату target (или ближайшую до неё)."""
    d_from = target - timedelta(days=lookback)
    with ThreadPoolExecutor(max_workers=BOARD_WORKERS) as ex:
        frames = list(ex.map(lambda b: _board_closes_raw(b, d_from, target), boards))
    frames = [f for f in frames if not f.empty]
    if not frames:
        return pd.Series(dtype=float)
    df = pd.concat(frames, ignore_index=True)
    df["CLOSE"] = pd.to_numeric(df["CLOSE"], errors="coerce")
    df["TRADEDATE"] = pd.to_datetime(df["TRADEDATE"], errors="coerce")
    df = df.dropna(subset=["CLOSE", "TRADEDATE"])
    return df.sort_values("TRADEDATE").groupby("SECID")["CLOSE"].last()


# ---------- RGBI ----------

@st.cache_data(ttl=600, show_spinner=False, persist="disk")
def load_rgbi(d_from: date, d_to: date) -> pd.DataFrame:
    df = fetch_paged(
        "history/engines/stock/markets/index/securities/RGBI",
        ("history",),
        params={"from": d_from.isoformat(), "till": d_to.isoformat()},
        columns={"history": "TRADEDATE,OPEN,HIGH,LOW,CLOSE"},
    )["history"]
    if df.empty:
        return df
    df["TRADEDATE"] = pd.to_datetime(df["TRADEDATE"])
    for c in ("OPEN", "HIGH", "LOW", "CLOSE"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return (df.dropna(subset=["CLOSE"]).drop_duplicates("TRADEDATE")
              .sort_values("TRADEDATE").reset_index(drop=True))
