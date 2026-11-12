from __future__ import annotations
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from config import BASE_URL, PAGE_SIZE, MAX_PAGES, HTTP_WORKERS, TIMEOUT

_session: requests.Session | None = None


def session() -> requests.Session:
    global _session
    if _session is None:
        s = requests.Session()
        retry = Retry(
            total=3, backoff_factor=0.4,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset(["GET"]),
        )
        s.mount("https://", HTTPAdapter(
            pool_connections=HTTP_WORKERS * 2,
            pool_maxsize=HTTP_WORKERS * 4,
            max_retries=retry,
        ))
        s.headers.update({"User-Agent": "bond-screener/2.0",
                          "Accept-Encoding": "gzip, deflate"})
        _session = s
    return _session


def _get(path: str, blocks: tuple[str, ...], params: dict,
         columns: dict | None) -> dict[str, pd.DataFrame]:
    p = {"iss.meta": "off", "iss.only": ",".join(blocks)}
    if columns:
        p.update({f"{b}.columns": c for b, c in columns.items() if b in blocks})
    p.update(params)
    try:
        r = session().get(f"{BASE_URL}/{path}.json", params=p, timeout=TIMEOUT)
        r.raise_for_status()
    except requests.HTTPError:
        if columns:                       # неизвестная колонка -> полный набор
            return _get(path, blocks, params, None)
        raise
    j = r.json()
    out = {}
    for b in blocks:
        blk = j.get(b) or {}
        out[b] = pd.DataFrame(blk.get("data") or [], columns=blk.get("columns") or [])
    return out


def fetch_paged(path: str, blocks: tuple[str, ...], params: dict | None = None,
                columns: dict | None = None, workers: int = HTTP_WORKERS,
                max_pages: int = MAX_PAGES) -> dict[str, pd.DataFrame]:
    """Качает страницы пачками по `workers`, стоп на первой пустой/короткой."""
    params = dict(params or {})
    acc: dict[str, list[pd.DataFrame]] = {b: [] for b in blocks}
    main = blocks[0]

    with ThreadPoolExecutor(max_workers=workers) as ex:
        page = 0
        while page < max_pages:
            idx = range(page, min(page + workers, max_pages))
            futures = [
                ex.submit(_get, path, blocks,
                          {**params, "start": i * PAGE_SIZE, "limit": PAGE_SIZE},
                          columns)
                for i in idx
            ]
            stop = False
            for f in futures:                       # порядок сохранён
                try:
                    res = f.result()
                except requests.RequestException:
                    stop = True
                    break
                if res[main].empty:
                    stop = True
                    break
                for b in blocks:
                    if not res[b].empty:
                        acc[b].append(res[b])
                if len(res[main]) < PAGE_SIZE:
                    stop = True
                    break
            if stop:
                break
            page += workers

    return {b: (pd.concat(v, ignore_index=True) if v else pd.DataFrame())
            for b, v in acc.items()}
