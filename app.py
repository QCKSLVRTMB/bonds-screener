import streamlit as st
from datetime import date, timedelta

st.set_page_config(page_title="Скринер облигаций", layout="wide")

from config import BOND_BOARDS, DEFAULT_SETTLEMENT
from core.loaders import load_universe, load_rgbi, load_closes
from core.dohod import load_dohod
from core.metrics import enrich, add_strength
from core.filters import apply_filters
from core.coloring import compute_color_tags
from ui import sidebar, charts, filters_panel, table, summary

st.title("Скринер облигаций")

params = sidebar.render()                       # даты, ставка, НДФЛ, Excel
raw = load_universe(BOND_BOARDS)                # cache_data + persist disk
dohod, dohod_key = load_dohod(params.uploaded)  # ключ = md5 файла
base = merge_and_prepare(raw, dohod)            # rename, dtype, category — один раз

charts.render(extremum_date=params.extremum)    # @st.fragment: не перерисовывается
                                                # при изменении фильтров

opts = filters_panel.render(base)               # всё внутри st.form -> 1 rerun

data_key = f"{len(base)}:{dohod_key}"
enriched = enrich_cached(_df=base, data_key=data_key,
                         settlement=params.settlement,
                         target=params.forecast_target,
                         ndfl=params.ndfl, delta=params.rate_change_pp)

filtered = apply_filters(enriched, opts)
closes = load_closes(tuple(sorted(filtered["Режим торгов"].unique())), params.extremum)
filtered = add_strength(filtered.copy(), closes, charts.rgbi_change(params.extremum))

tags = compute_color_tags(filtered, params.current_cbr, opts["max_duration"])
table.render(filtered, tags)                    # @st.fragment: сортировка/колонки/страницы
summary.render(filtered)                        # без изменений
