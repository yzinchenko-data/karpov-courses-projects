"""Сверка ДСП / подрядчик (AdRiver) / Метрика.

Как пользоваться: правишь только блок НАСТРОЙКИ ниже и запускаешь файл
(кнопка ▷ в VS Code или `python discrepancy_check.py`). Аргументы не нужны.

Для каждого источника пишешь: ссылку на файл, название колонки с числом и,
если хочешь сверку по дням, название колонки с датой (иначе сравниваются итоги).
Файлы: .xlsx / .xls / .csv, можно путь внутри .zip (…архив.zip/файл.xlsx).
Шапка ищется сама, регистр / ё / пробелы / похожие латинские-кириллические
буквы в названиях колонок не важны.
"""

import re
import sys
import zipfile
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd

# ============================== НАСТРОЙКИ ==============================

# Допустимое расхождение, %. Всё, что больше, помечается «ПРОВЕРИТЬ».
THRESHOLD_AD = 25   # ДСП <-> подрядчик
THRESHOLD_WEB = 25  # рекламная система <-> веб-аналитика (Метрика)

# Куда сохранить результат (None = рядом с этим скриптом)
OUT_DIR = None

# Дата с / по (None = весь период). Формат "2026-09-01".
DATE_FROM = None
DATE_TO = None

# Кампании. Пути пиши как есть, в кавычках; можно и с обычными \ — это Python.
# "date" можно не указывать -> тогда сравниваются только итоги за период.
CAMPAIGNS = [
    {
        "name": "Joilion",
        "dsp": {"file": r"C:\Users\YZinch01\Desktop\WhiteClix_Joilion.xlsx",
                "col": "Clicks"},
        "ver": {"file": r"C:\Users\YZinch01\Desktop\AdriverStat_Joilion.xlsx",
                "col": "Кликов всего с учетом IVT"},
        "metrika": {"file": r"C:\Users\YZinch01\Desktop\metrika_joilion.csv",
                    "col": "Визиты"},
    },
    {
        "name": "Fx7",
        "dsp": {"file": r"C:\Users\YZinch01\Downloads\354fc64a-51a5-4cff-b72e-2eb757e3e7c5\Haval F7x WhiteClix Лето-Зима 2026 _ aug-dec_statistics_27-09-2026.xlsx",
                "col": "Clicks", "date": "Day"},
        "ver": {"file": r"C:\Users\YZinch01\Downloads\354fc64a-51a5-4cff-b72e-2eb757e3e7c5\AdriverStat_ID_ 14770659_days_01-09-26_29-09-26 (1).xlsx",
                "col": "Кликов всего с учетом IVT", "date": "Дата"},
        "metrika": {"file": r"C:\Users\YZinch01\Downloads\Посещаемость-2026-09-01-2026-09-30 (5).xlsx",
                    "col": "Визиты", "date": "Дата визита"},
    },
]

# ========================== дальше можно не читать ==========================

_LOOKALIKE = str.maketrans("acekmoptxyhb", "асекмортхунв")  # латиница -> кириллица
_TOTAL_WORDS = {"итого", "итог", "всего", "total", "grand total", "сумма"}


def norm(s):
    s = "" if s is None or (isinstance(s, float) and pd.isna(s)) else str(s)
    s = s.replace("\u00a0", " ").replace("\u202f", " ").replace("ё", "е").replace("Ё", "е")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s.translate(_LOOKALIKE)


def parse_number(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("\u00a0", "").replace("\u202f", "").replace(" ", "")
    s = s.replace("%", "").replace(",", ".")
    if s in ("", "-", "—", "–"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def read_raw(path):
    """Читает файл целиком без шапки. Понимает пути внутри .zip."""
    path = str(path).strip().strip('"')
    m = re.search(r"\.zip[\\/]", path, flags=re.I)
    if m:
        zip_path, inner = path[: m.end() - 1], path[m.end():]
        with zipfile.ZipFile(zip_path) as z:
            names = {n.replace("\\", "/"): n for n in z.namelist()}
            key = inner.replace("\\", "/")
            if key not in names:
                raise FileNotFoundError(f"В архиве нет файла: {inner}")
            data, name = z.read(names[key]), inner
        src = BytesIO(data)
    else:
        if not Path(path).exists():
            raise FileNotFoundError(f"Файл не найден: {path}")
        src, name = path, path

    if name.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "cp1251", "utf-8"):
            for sep in (";", ",", "\t"):
                try:
                    if hasattr(src, "seek"):
                        src.seek(0)
                    df = pd.read_csv(src, header=None, sep=sep, encoding=enc,
                                     dtype=str, engine="python", on_bad_lines="skip")
                    if df.shape[1] > 1:
                        return df
                except Exception:
                    continue
        raise ValueError(f"Не смог прочитать CSV: {path}")
    return pd.read_excel(src, header=None, dtype=object)


def find_header(raw, wanted):
    """Строка шапки = первая строка, где встречаются все нужные названия."""
    wanted = [norm(w) for w in wanted if w]
    for i in range(min(len(raw), 60)):
        cells = [norm(x) for x in raw.iloc[i].tolist()]
        if all(any(w == c for c in cells) for w in wanted):
            return i
    for i in range(min(len(raw), 60)):  # мягко: «содержит»
        cells = [norm(x) for x in raw.iloc[i].tolist()]
        if all(any(w in c for c in cells if c) for w in wanted):
            return i
    return None


def pick_col(cols, wanted):
    w = norm(wanted)
    normed = [norm(c) for c in cols]
    if w in normed:
        return cols[normed.index(w)]
    hits = [c for c, n in zip(cols, normed) if w and w in n]
    if len(hits) >= 1:
        return hits[0]
    raise KeyError(f"колонка «{wanted}» не найдена. Есть: {[str(c) for c in cols if str(c) != 'nan']}")


def load_source(spec, label):
    raw = read_raw(spec["file"])
    h = find_header(raw, [spec["col"], spec.get("date")])
    if h is None:
        raise KeyError(f"{label}: не нашёл шапку с «{spec['col']}»"
                       + (f" и «{spec['date']}»" if spec.get("date") else "")
                       + f" в файле {Path(spec['file']).name}")
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c) if not pd.isna(c) else f"_{i}" for i, c in enumerate(raw.iloc[h])]
    vcol = pick_col(list(df.columns), spec["col"])

    # убираем строки «Итого»
    is_total = df.apply(lambda r: any(norm(x) in _TOTAL_WORDS for x in r.tolist()), axis=1)
    df = df[~is_total]

    df["_v"] = df[vcol].map(parse_number)
    df = df[df["_v"].notna()]

    if spec.get("date"):
        dcol = pick_col(list(df.columns), spec["date"])
        d = parse_dates(df[dcol])
        df = df.assign(_d=d.dt.normalize()).dropna(subset=["_d"])
        if DATE_FROM:
            df = df[df["_d"] >= pd.Timestamp(DATE_FROM)]
        if DATE_TO:
            df = df[df["_d"] <= pd.Timestamp(DATE_TO)]
        return df.groupby("_d")["_v"].sum()
    return float(df["_v"].sum())


def parse_dates(col):
    """ISO (2026-09-01) читаем как есть, остальное (01.09.2026) — день первым."""
    def one(x):
        if isinstance(x, (datetime, pd.Timestamp)):
            return pd.Timestamp(x)
        t = str(x).strip()
        if re.match(r"^\d{4}-\d{2}-\d{2}", t):
            return pd.to_datetime(t[:10], errors="coerce")
        return pd.to_datetime(t, dayfirst=True, errors="coerce")
    return col.map(one).astype("datetime64[ns]")


def diff_pct(base, other):
    """Расхождение в % относительно base: (other - base) / base."""
    if base is None or other is None or pd.isna(base) or pd.isna(other) or base == 0:
        return None
    return (other - base) / base * 100


def verdict(p, thr):
    if p is None:
        return "нет данных"
    return "ПРОВЕРИТЬ" if abs(p) > thr else "ок"


def run_campaign(c):
    res = {}
    for key, label in (("dsp", "ДСП"), ("ver", "Подрядчик"), ("metrika", "Метрика")):
        if key in c and c[key].get("file"):
            try:
                res[key] = load_source(c[key], label)
            except Exception as e:
                print(f"  ! {c['name']} / {label}: {e}")
                res[key] = None
        else:
            res[key] = None

    pairs = [("ДСП → Подрядчик", "dsp", "ver", THRESHOLD_AD),
             ("Подрядчик → Метрика", "ver", "metrika", THRESHOLD_WEB),
             ("ДСП → Метрика", "dsp", "metrika", THRESHOLD_WEB)]

    summary, daily = [], []
    series = {k: v for k, v in res.items() if isinstance(v, pd.Series)}
    for title, a, b, thr in pairs:
        va, vb = res.get(a), res.get(b)
        if va is None or vb is None:
            summary.append([c["name"], title, None, None, None, thr, "нет данных"])
            continue
        ta = float(va.sum()) if isinstance(va, pd.Series) else va
        tb = float(vb.sum()) if isinstance(vb, pd.Series) else vb
        p = diff_pct(ta, tb)
        summary.append([c["name"], title, ta, tb, p, thr, verdict(p, thr)])

        if isinstance(va, pd.Series) and isinstance(vb, pd.Series):
            d = pd.concat([va.rename("A"), vb.rename("B")], axis=1).fillna(0).sort_index()
            for day, r in d.iterrows():
                dp = diff_pct(r["A"], r["B"])
                daily.append([c["name"], title, day.date(), r["A"], r["B"], dp, verdict(dp, thr)])
    return summary, daily


def main():
    all_sum, all_day = [], []
    for c in CAMPAIGNS:
        print(f"\n== {c['name']} ==")
        s, d = run_campaign(c)
        all_sum += s
        all_day += d
        for row in s:
            name, title, a, b, p, thr, v = row
            if p is None:
                print(f"  {title:<22} {v}")
            else:
                print(f"  {title:<22} {a:>12,.0f} → {b:>12,.0f}   {p:+6.1f}%  (порог {thr}%)  {v}")

    sum_df = pd.DataFrame(all_sum, columns=["Кампания", "Сравнение", "Было (база)", "Стало",
                                            "Расхождение, %", "Порог, %", "Вердикт"])
    day_df = pd.DataFrame(all_day, columns=["Кампания", "Сравнение", "Дата", "Было (база)", "Стало",
                                            "Расхождение, %", "Вердикт"])
    out_dir = Path(OUT_DIR) if OUT_DIR else Path(__file__).resolve().parent
    out = out_dir / f"discrepancy_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        sum_df.round(1).to_excel(w, sheet_name="Итоги", index=False)
        if len(day_df):
            day_df.round(1).to_excel(w, sheet_name="По дням", index=False)
        for ws in w.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 20
            ws.freeze_panes = "A2"
    print(f"\nГотово: {out}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Ошибка: {e}")
        sys.exit(1)
