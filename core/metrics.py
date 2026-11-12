import numpy as np
import pandas as pd
from config import RUB_CURRENCIES


def _join_labels(parts: list[np.ndarray]) -> np.ndarray:
    out = np.full(len(parts[0]), "", dtype=object)
    for p in parts:
        out = np.where(out == "", p, np.where(p == "", out, out + " · " + p))
    return out


def add_market_risk(df, rate_change_pp):
    is_rub = df["Валюта номинала"].astype(str).str.strip().str.upper().isin(RUB_CURRENCIES)
    fallback = np.where(is_rub, 2.0, 1.0)
    delta = rate_change_pp if abs(rate_change_pp) > 1e-3 else fallback
    df["Рыночный риск, %"] = -df["Дюрация, лет"] * delta
    return df


def add_liquidity_score(df):
    k = df.get("Коэф. ликвидности", pd.Series(np.nan, index=df.index))
    n = df.get("Количество сделок", pd.Series(np.nan, index=df.index))
    kk, nn = k / 100.0, n.clip(upper=100) / 100.0
    df["Ликвидность (комбинир.)"] = np.where(
        k.notna() & n.notna(), kk * 0.6 + nn * 0.4,
        np.where(k.notna(), kk, nn))
    return df


def add_nearest_event(df, settlement_date):
    o = df.get("Дата оферты", pd.Series(pd.NaT, index=df.index))
    m = df.get("Дата погашения", pd.Series(pd.NaT, index=df.index))
    nearest = o.where(o.notna() & (m.isna() | (o <= m)), m)
    df["Ближайшая дата, лет"] = (nearest - pd.Timestamp(settlement_date)).dt.days / 365.0
    df["Событие"] = np.where(
        nearest.notna() & nearest.eq(o), "Оферта",
        np.where(nearest.notna(), "Погашение", ""))
    return df


def add_bond_type(df):
    cp = df.get("Тип купона", pd.Series("", index=df.index)).fillna("").astype(str).str.lower()
    fv = df.get("Тип номинала", pd.Series("", index=df.index)).fillna("").astype(str).str.lower()
    sub = df.get("Субординированная", pd.Series("", index=df.index)).fillna("").astype(str).str.lower()

    is_float = cp.str.contains("плавающ|перемен|формул", na=False)
    parts = [
        np.where(is_float, "Флоатер", np.where(cp.str.contains("фикс"), "Фикс", "")),
        np.where(fv.str.contains("амортиз"), "Амортизируемая", ""),
        np.where(fv.str.contains("индекс"), "Индексируемая", ""),
        np.where(df.get("Дата погашения", pd.Series(pd.NaT, index=df.index)).isna(), "Вечная", ""),
        np.where(sub.isin(["да", "yes"]), "Суборд", ""),
    ]
    labels = _join_labels(parts)
    df["Тип бумаги"] = np.where(labels == "", "—", labels)
    df["Флоатер"] = np.where(is_float, "Да", "—")
    return df


def add_strength(df, closes: pd.Series, rgbi_change_pct: float):
    """Векторный расчёт вместо цикла с HTTP-запросом на каждую бумагу."""
    price = df["Цена последней сделки, %"]
    if "Текущая цена, %" in df:
        price = price.fillna(df["Текущая цена, %"])
    hist = df["SECID"].map(closes)
    chg = (price / hist - 1.0) * 100.0
    df["Сила актива (Δ%)"] = (chg - rgbi_change_pct).where(hist > 0)
    return df
