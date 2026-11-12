import html as _html
import numpy as np
import pandas as pd
import streamlit as st
from config import DOHOD_BOND_URL, TABLE_PAGE_SIZE

_BG = {"green": "#c6efce", "yellow": "#ffeb9c", "red": "#ffc7ce"}
_FG = {"green": "#006100", "yellow": "#9c6500", "red": "#9c0006"}


def _format_column(s: pd.Series, kind: str) -> pd.Series:
    if kind == "date":
        return s.dt.strftime("%d.%m.%Y").fillna("")
    if kind == "int":
        return s.round().astype("Int64").astype(str).replace("<NA>", "")
    if kind == "float":
        return s.map(lambda v: "" if pd.isna(v) else f"{v:,.2f}".replace(",", " "))
    return s.fillna("").astype(str).map(_html.escape)


@st.fragment
def table_block(df, tags, columns, date_cols, int_cols, use_colors=True):
    total = len(df)
    pages = max(1, (total + TABLE_PAGE_SIZE - 1) // TABLE_PAGE_SIZE)
    c1, c2 = st.columns([1, 3])
    page = c1.number_input("Страница", 1, pages, 1, key="tbl_page")
    c2.caption(f"Строк: {total}. Показано {min(TABLE_PAGE_SIZE, total)} на странице.")

    sl = slice((page - 1) * TABLE_PAGE_SIZE, page * TABLE_PAGE_SIZE)
    view = df.iloc[sl][columns]
    tg = tags.reindex(index=view.index, columns=columns, fill_value="").fillna("")

    fmt = pd.DataFrame(index=view.index, columns=columns, dtype=object)
    for c in columns:
        kind = ("date" if c in date_cols else
                "int" if c in int_cols else
                "float" if pd.api.types.is_numeric_dtype(view[c]) else "str")
        fmt[c] = _format_column(view[c], kind)

    # ссылка строится по ISIN, а не по SECID (у Доход.ру карточка по ISIN)
    links = df.iloc[sl].get("ISIN", pd.Series("", index=view.index)).fillna("").astype(str)

    head = "".join(f"<th>{_html.escape(c)}</th>" for c in columns)
    rows = []
    for i, cells in enumerate(fmt.itertuples(index=False, name=None)):
        colors = tg.iloc[i].to_numpy()
        tds = []
        for j, (col, val) in enumerate(zip(columns, cells)):
            style = ""
            if use_colors and colors[j] in _BG:
                style = f' style="background:{_BG[colors[j]]};color:{_FG[colors[j]]}"'
            if col == "SECID" and links.iloc[i]:
                val = (f'<a href="{DOHOD_BOND_URL}{_html.escape(links.iloc[i])}" '
                       f'target="_blank" rel="noopener">{val}</a>')
            tds.append(f"<td{style}>{val}</td>")
        rows.append("<tr>" + "".join(tds) + "</tr>")

    st.markdown(
        '<div style="overflow:auto;max-height:620px;border:1px solid #ddd">'
        '<table class="screener"><thead><tr>' + head + "</tr></thead><tbody>"
        + "".join(rows) + "</tbody></table></div>",
        unsafe_allow_html=True,
    )
