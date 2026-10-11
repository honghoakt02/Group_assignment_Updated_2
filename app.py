# -*- coding: utf-8 -*-
"""
========================================================================================
HOSE QUANT V4: STREAMLIT QUANTITATIVE PORTFOLIO BACKTESTING WEB APP
Hệ thống kiểm định quy trình đầu tư định lượng trên cổ phiếu HOSE
Chống rò rỉ thông tin tương lai (Look-ahead Bias Prevention) với Nested Walk-Forward
========================================================================================
"""

import os
import re
import math
import io
import warnings
from pathlib import Path
from itertools import product

import numpy as np
import pandas as pd
import streamlit as st
from scipy.optimize import minimize
from scipy.cluster.hierarchy import linkage
from scipy.spatial.distance import squareform
from sklearn.covariance import LedoitWolf
import plotly.graph_objects as go
import plotly.express as px

warnings.filterwarnings("ignore")

# ======================================================================================
# 0. STREAMLIT CONFIGURATION & CUSTOM STYLES
# ======================================================================================
st.set_page_config(
    page_title="HOSE Quant Backtest V4 | Định Lượng Cổ Phiếu",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main-title {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.2rem;
    }
    .badge-container {
        display: flex;
        gap: 0.6rem;
        flex-wrap: wrap;
        margin-bottom: 1.2rem;
    }
    .badge {
        padding: 0.3rem 0.75rem;
        border-radius: 9999px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    .badge-blue { background-color: #DBEAFE; color: #1E40AF; }
    .badge-green { background-color: #D1FAE5; color: #065F46; }
    .badge-amber { background-color: #FEF3C7; color: #92400E; }
    .badge-purple { background-color: #EDE9FE; color: #5B21B6; }
    .metric-card {
        background-color: #F9FAFB;
        border: 1px solid #E5E7EB;
        border-radius: 0.5rem;
        padding: 0.9rem;
        text-align: center;
    }
    .metric-val {
        font-size: 1.5rem;
        font-weight: 700;
        margin-top: 0.2rem;
    }
    .metric-label {
        font-size: 0.82rem;
        color: #6B7280;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
</style>
""", unsafe_allow_html=True)

# Hằng số mặc định hệ thống
BENCHMARK_ALIASES = {"VNINDEX", "VN-INDEX", "VNI", "^VNINDEX", "VNINDEXINDEX"}
TRADING_DAYS = 252

FA_ALIASES = {
    "ROE": ["roe", "return_on_equity", "returnonequity"],
    "ROA": ["roa", "return_on_assets", "returnonassets"],
    "RevenueGrowth": ["revenue_growth", "sales_growth", "revenuegrowth", "tangtruongdoanhthu"],
    "EPSGrowth": ["eps_growth", "epsgrowth", "tangtruongeps"],
    "GrossMargin": ["gross_margin", "grossmargin", "bienloinhuangop"],
    "OperatingMargin": ["operating_margin", "operatingmargin"],
    "NetMargin": ["net_margin", "netmargin", "bienloinhuanrong"],
    "CurrentRatio": ["current_ratio", "currentratio"],
    "DebtToEquity": ["debt_to_equity", "de", "d_e", "debttoequity"],
    "PE": ["pe", "p_e", "price_to_earnings", "pricetoearnings"],
    "PB": ["pb", "p_b", "price_to_book", "pricetobook"],
    "EVEBITDA": ["ev_ebitda", "evebitda"],
}

# ======================================================================================
# 1. CORE QUANT DATA UTILITIES & UNIVERSAL FILE LOADER
# ======================================================================================

def _canon(s):
    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

def find_col(df, candidates):
    lookup = {_canon(c): c for c in df.columns}
    for c in candidates:
        if _canon(c) in lookup:
            return lookup[_canon(c)]
    return None

def parse_dates_safely(s):
    """Ưu tiên mm/dd/yyyy với fallback tự động."""
    s = pd.Series(s)
    x = pd.to_datetime(s, format="%m/%d/%Y", errors="coerce")
    if x.notna().mean() < 0.80:
        x = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if x.notna().mean() < 0.80:
        x = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return x

def read_any_tabular_data(file_input):
    """
    Hàm đọc dữ liệu bảng vạn năng, tự động xử lý:
    - File Excel (.xlsx, .xls) dù viết hoa/thường (.XLSX, .Xlsx) hay không có đuôi.
    - Nhận diện Magic Bytes nhị phân (PK.. cho zip/xlsx và 0xd0cf11e0 cho ole/xls) để không bao giờ nhầm sang text/CSV.
    - File CSV với đa dạng bảng mã (UTF-8, UTF-8-BOM/utf-8-sig, Windows-1258, CP1252, Latin-1, UTF-16).
      Đặc biệt 'windows-1258' và 'cp1252' triệt tiêu hoàn toàn lỗi 'utf-8 codec can't decode byte 0xe3 in position 1'.
    - Tự động phát hiện dấu phân cách (dấu phẩy ',', chấm phẩy ';', hoặc tab '\t').
    """
    content = b""
    filename = ""

    # 1. Trích xuất nội dung bytes và tên file an toàn
    if isinstance(file_input, (str, Path)):
        filename = str(file_input).lower()
        if os.path.exists(file_input):
            with open(file_input, "rb") as f:
                content = f.read()
    elif hasattr(file_input, "getvalue"):  # Streamlit UploadedFile
        filename = getattr(file_input, "name", "").lower()
        content = file_input.getvalue()
    elif hasattr(file_input, "read"):
        if hasattr(file_input, "seek"):
            file_input.seek(0)
        content = file_input.read()
        if hasattr(file_input, "seek"):
            file_input.seek(0)
        filename = getattr(file_input, "name", "").lower()
    elif isinstance(file_input, (bytes, bytearray)):
        content = bytes(file_input)

    if not content:
        raise ValueError("File dữ liệu rỗng hoặc không thể đọc được nội dung.")

    # 2. Nhận diện Magic Bytes để phát hiện file Excel nhị phân
    is_zip_xlsx = content.startswith(b"PK\x03\x04")
    is_ole_xls = content.startswith(b"\xd0\xcf\x11\xe0")
    is_excel_ext = any(filename.endswith(ext) for ext in [".xlsx", ".xls", ".xlsm", ".xlsb"])

    if is_zip_xlsx or is_ole_xls or is_excel_ext:
        # Thử engine openpyxl cho xlsx
        try:
            return pd.read_excel(io.BytesIO(content), engine="openpyxl")
        except Exception:
            pass
        # Thử engine mặc định / xlrd
        try:
            return pd.read_excel(io.BytesIO(content))
        except Exception:
            pass

    # 3. Thử đọc dạng CSV/Text với các bảng mã (Encodings)
    encodings_to_try = [
        "utf-8-sig",
        "utf-8",
        "windows-1258",  # Bảng mã tiếng Việt chuẩn của Microsoft Windows (khắc phục triệt để byte 0xe3 - ký tự 'ã')
        "cp1252",        # Western European ANSI
        "latin1",        # ISO-8859-1 (nhận mọi byte 0x00 - 0xFF mà không bao giờ báo lỗi decode)
        "utf-16",
        "utf-16le",
        "utf-16be",
        "cp1250",
    ]

    separators_to_try = [None, ",", ";", "\t"]

    for enc in encodings_to_try:
        for sep in separators_to_try:
            try:
                if sep is None:
                    df = pd.read_csv(io.BytesIO(content), encoding=enc, sep=None, engine="python")
                else:
                    df = pd.read_csv(io.BytesIO(content), encoding=enc, sep=sep)
                if df is not None and len(df.columns) >= 2 and len(df) > 0:
                    return df
            except Exception:
                continue

    # 4. Fallback: Thử đọc Excel nếu file CSV thực chất là định dạng Excel
    try:
        return pd.read_excel(io.BytesIO(content))
    except Exception:
        pass

    # 5. Fallback tuyệt đối với latin1 (đảm bảo không bao giờ crash decode)
    return pd.read_csv(io.BytesIO(content), encoding="latin1", errors="replace")

def standardize_ohlcv(df_raw, use_adjusted=True):
    """Chuẩn hóa dữ liệu về chuẩn Ticker - Date - Open - High - Low - Close - Volume."""
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    ticker_col = find_col(df, ["ticker", "symbol", "code", "stock_code", "mack", "mã", "mã cp"])
    date_col = find_col(df, ["date", "datetime", "trading_date", "time", "ngay", "ngày"])
    volume_col = find_col(df, ["volume", "vol", "trading_volume", "khoiluong", "khối lượng khớp lệnh", "khối lượng"])

    raw_cols = {
        "Open": find_col(df, ["open", "open_price", "giá mở cửa", "mở cửa"]),
        "High": find_col(df, ["high", "high_price", "giá cao nhất", "cao nhất"]),
        "Low": find_col(df, ["low", "low_price", "giá thấp nhất", "thấp nhất"]),
        "Close": find_col(df, ["close", "close_price", "giá đóng cửa", "đóng cửa"]),
    }
    adj_cols = {
        "Open": find_col(df, ["adj_open", "adjusted_open"]),
        "High": find_col(df, ["adj_high", "adjusted_high"]),
        "Low": find_col(df, ["adj_low", "adjusted_low"]),
        "Close": find_col(df, ["adj_close", "adjusted_close", "adjclose", "giá điều chỉnh", "giá đóng cửa điều chỉnh"]),
    }
    if ticker_col is None or date_col is None:
        raise ValueError(f"Thiếu cột mã cổ phiếu hoặc ngày. Các cột hiện có: {list(df.columns)}")

    use_adj = bool(use_adjusted and all(v is not None for v in adj_cols.values()))
    pcols = adj_cols if use_adj else raw_cols
    if not all(v is not None for v in pcols.values()):
        pcols = raw_cols
        use_adj = False
        if not all(v is not None for v in pcols.values()):
            raise ValueError("Bộ dữ liệu thiếu ít nhất một trong các cột OHLC bắt buộc.")

    out = pd.DataFrame({
        "Ticker": df[ticker_col].astype(str).str.strip().str.upper(),
        "Date": parse_dates_safely(df[date_col]),
        "Open": pd.to_numeric(df[pcols["Open"]].astype(str).str.replace(",", "."), errors="coerce"),
        "High": pd.to_numeric(df[pcols["High"]].astype(str).str.replace(",", "."), errors="coerce"),
        "Low": pd.to_numeric(df[pcols["Low"]].astype(str).str.replace(",", "."), errors="coerce"),
        "Close": pd.to_numeric(df[pcols["Close"]].astype(str).str.replace(",", "."), errors="coerce"),
        "Volume": pd.to_numeric(df[volume_col].astype(str).str.replace(",", "."), errors="coerce") if volume_col else 0.0,
    })
    out = out.dropna(subset=["Ticker", "Date", "Open", "High", "Low", "Close"])
    out = out[(out[["Open", "High", "Low", "Close"]] > 0).all(axis=1)]
    out["Volume"] = out["Volume"].fillna(0).clip(lower=0)

    valid_ohlc = (
        (out["High"] >= out[["Open", "Close", "Low"]].max(axis=1)) &
        (out["Low"] <= out[["Open", "Close", "High"]].min(axis=1))
    )
    out = out.loc[valid_ohlc]
    out = out.drop_duplicates(["Ticker", "Date"], keep="last")
    out = out.sort_values(["Ticker", "Date"]).reset_index(drop=True)
    return out, "ADJUSTED" if use_adj else "RAW"

def data_quality_table(data):
    rows = []
    for t, g in data.groupby("Ticker"):
        rows.append({
            "Ticker": t,
            "Start": g["Date"].min().strftime("%Y-%m-%d"),
            "End": g["Date"].max().strftime("%Y-%m-%d"),
            "Sessions": g["Date"].nunique(),
            "MissingVolumePct": round(100 * g["Volume"].isna().mean(), 2),
            "MedianTradedValue": round((g["Close"] * g["Volume"]).median(), 0),
        })
    return pd.DataFrame(rows).sort_values("Sessions", ascending=False).reset_index(drop=True)

# ======================================================================================
# 2. BENCHMARK & PRICE UTILS
# ======================================================================================

def build_price_dict(data):
    return {
        t: g[["Date", "Open", "High", "Low", "Close", "Volume"]]
             .drop_duplicates("Date").sort_values("Date").set_index("Date")
        for t, g in data.groupby("Ticker")
    }

def build_synthetic_market_proxy(price_dict, exclude=None):
    exclude = set(exclude or [])
    closes = [px["Close"].rename(t) for t, px in price_dict.items() if t not in exclude]
    m = pd.concat(closes, axis=1).sort_index()
    ret = m.pct_change(fill_method=None)
    ewret = ret.mean(axis=1, skipna=True).fillna(0)
    close = 1000 * (1 + ewret).cumprod()
    return pd.DataFrame({"Open": close, "High": close, "Low": close, "Close": close, "Volume": 0.0}, index=close.index)

# ======================================================================================
# 3. TECHNICAL INDICATORS & TA SCORE (GROUP 1 FILTER)
# ======================================================================================

def ema(s, span):
    return s.ewm(span=span, adjust=False, min_periods=span).mean()

def rsi_wilder(close, window=14):
    d = close.diff()
    gain, loss = d.clip(lower=0), -d.clip(upper=0)
    ag = gain.ewm(alpha=1/window, adjust=False, min_periods=window).mean()
    al = loss.ewm(alpha=1/window, adjust=False, min_periods=window).mean()
    rs = ag / al.replace(0, np.nan)
    rsi = 100 - 100/(1+rs)
    rsi.loc[(al == 0) & (ag > 0)] = 100
    rsi.loc[(ag == 0) & (al > 0)] = 0
    return rsi

def atr_wilder(df, window=14):
    prev = df["Close"].shift(1)
    tr = pd.concat([
        df["High"] - df["Low"],
        (df["High"] - prev).abs(),
        (df["Low"] - prev).abs()
    ], axis=1).max(axis=1)
    return tr.ewm(alpha=1/window, adjust=False, min_periods=window).mean()

def adx_wilder(df, window=14):
    high, low = df["High"], df["Low"]
    up, down = high.diff(), -low.diff()
    plus_dm = pd.Series(np.where((up > down) & (up > 0), up, 0.0), index=df.index)
    minus_dm = pd.Series(np.where((down > up) & (down > 0), down, 0.0), index=df.index)
    atr = atr_wilder(df, window)
    plus_di = 100 * plus_dm.ewm(alpha=1/window, adjust=False, min_periods=window).mean() / atr.replace(0, np.nan)
    minus_di = 100 * minus_dm.ewm(alpha=1/window, adjust=False, min_periods=window).mean() / atr.replace(0, np.nan)
    dx = 100 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    adx = dx.ewm(alpha=1/window, adjust=False, min_periods=window).mean()
    return adx, plus_di, minus_di

def calculate_ta_features(df):
    out = df.copy().sort_index()
    out["EMA20"] = ema(out["Close"], 20)
    out["EMA50"] = ema(out["Close"], 50)
    out["EMA200"] = ema(out["Close"], 200)

    ema12, ema26 = ema(out["Close"], 12), ema(out["Close"], 26)
    out["MACD"] = ema12 - ema26
    out["MACD_Signal"] = ema(out["MACD"], 9)

    out["ADX"], out["Plus_DI"], out["Minus_DI"] = adx_wilder(out, 14)
    out["ATR"] = atr_wilder(out, 14)
    out["ATR_pct"] = 100 * out["ATR"] / out["Close"].replace(0, np.nan)
    out["ATR_pct_Q90"] = out["ATR_pct"].rolling(126, min_periods=40).quantile(0.90)

    out["Volume_MA20"] = out["Volume"].rolling(20, min_periods=20).mean()
    out["Volume_Ratio"] = out["Volume"] / out["Volume_MA20"].replace(0, np.nan)
    out["OBV"] = (np.sign(out["Close"].diff()).fillna(0) * out["Volume"]).cumsum()
    out["OBV_EMA20"] = ema(out["OBV"], 20)

    out["RSI"] = rsi_wilder(out["Close"], 14)
    out["ROC10"] = out["Close"].pct_change(10) * 100
    out["High_52W"] = out["High"].rolling(252, min_periods=150).max()
    out["Close_to_52W_High"] = out["Close"] / out["High_52W"].replace(0, np.nan)
    out["Avg_Traded_Value_60"] = (out["Close"] * out["Volume"]).rolling(60, min_periods=20).mean()

    out["Momentum63"] = out["Close"].pct_change(63)
    out["Momentum126"] = out["Close"].pct_change(126)
    out["Momentum252"] = out["Close"].pct_change(252)
    out["Vol63"] = out["Close"].pct_change().rolling(63, min_periods=40).std() * np.sqrt(TRADING_DAYS)
    rollmax = out["Close"].rolling(126, min_periods=60).max()
    out["DD126"] = out["Close"] / rollmax - 1

    # TA Score 0-10 cấu trúc chuẩn Nhóm 1
    out["Trend_Score"] = (
        (out["Close"] > out["EMA50"]).astype(int)*2 +
        (out["EMA50"] > out["EMA200"]).astype(int)*2 +
        (out["MACD"] > out["MACD_Signal"]).astype(int)
    ).clip(0, 5)
    momentum_core = ((out["RSI"] >= 45) & (out["RSI"] <= 75)) | (out["ROC10"] > 0)
    out["Momentum_Score"] = (
        momentum_core.astype(int)*2 +
        (out["Close_to_52W_High"] >= 0.85).astype(int)
    ).clip(0, 3)
    out["Volume_Score"] = ((out["Volume_Ratio"] >= 1.0) | (out["OBV"] > out["OBV_EMA20"])).astype(int)
    out["Volatility_Score"] = (
        out["ATR_pct"].notna() & out["ATR_pct_Q90"].notna() &
        (out["ATR_pct"] <= out["ATR_pct_Q90"])
    ).astype(int)
    out["TA_Score"] = (
        out["Trend_Score"] + out["Momentum_Score"] +
        out["Volume_Score"] + out["Volatility_Score"]
    ).clip(0, 10).astype(int)
    return out

def stateful_active_from_score(score, buy_threshold, exit_threshold):
    state = 0
    vals = []
    for v in score.fillna(-np.inf):
        if v >= buy_threshold:
            state = 1
        elif v <= exit_threshold:
            state = 0
        vals.append(state)
    return pd.Series(vals, index=score.index, dtype=int)

# ======================================================================================
# 4. TIME SPLITTING & CAUSAL INTEGRITY
# ======================================================================================

def resolve_auto_splits(data, split_mode="auto", train_prop=0.6, val_prop=0.2, test_prop=0.2,
                        min_total=420, min_train=252, min_val=63, min_test=100,
                        prefer_last_year=True, last_year_min=180):
    dates = pd.DatetimeIndex(sorted(pd.to_datetime(data["Date"].unique())))
    n = len(dates)
    if n < min_total:
        raise ValueError(f"Dữ liệu có {n} phiên. Cần tối thiểu {min_total} phiên để chia 3 tập.")

    if split_mode == "proportions":
        i_train = max(min_train, int(n * train_prop))
        i_val_end = max(i_train + min_val, int(n * (train_prop + val_prop)))
        i_val_end = min(i_val_end, n - min_test)
        if i_val_end <= i_train:
            raise ValueError("Không đủ dữ liệu phân bổ theo tỷ lệ.")
        core_end = dates[i_train - 1]
        val_start, val_end = dates[i_train], dates[i_val_end - 1]
        test_start, test_end = dates[i_val_end], dates[-1]
        return dates, core_end, val_start, val_end, test_start, test_end

    # Chế độ Auto: ưu tiên dùng năm cuối cùng làm OOS Final Test
    last_year = dates[-1].year
    last_year_dates = dates[dates.year == last_year]
    pre_last_year_count = int((dates < last_year_dates[0]).sum()) if len(last_year_dates) else 0

    if (prefer_last_year and len(last_year_dates) >= last_year_min and
        pre_last_year_count >= min_train + min_val):
        test_start, test_end = last_year_dates[0], last_year_dates[-1]
        pretest = dates[dates < test_start]
    else:
        test_n = max(min_test, min(252, int(round(test_prop * n))))
        test_start, test_end = dates[-test_n], dates[-1]
        pretest = dates[dates < test_start]

    val_n = max(min_val, min(252, int(round(0.25 * len(pretest)))))
    val_n = min(val_n, len(pretest) - min_train)
    if val_n < min_val:
        raise ValueError("Không đủ dữ liệu lịch sử trước Test để tách Validation.")

    val_start, val_end = pretest[-val_n], pretest[-1]
    core_end = pretest[-val_n - 1]
    return dates, core_end, val_start, val_end, test_start, test_end

# ======================================================================================
# 5. METRICS & SINGLE STOCK SIMULATION
# ======================================================================================

def annualized_sharpe(r):
    r = pd.Series(r).replace([np.inf, -np.inf], np.nan).dropna()
    return np.sqrt(TRADING_DAYS) * r.mean() / r.std(ddof=1) if len(r) > 1 and r.std(ddof=1) > 1e-12 else np.nan

def annualized_sortino(r):
    r = pd.Series(r).replace([np.inf, -np.inf], np.nan).dropna()
    neg = r[r < 0]
    return np.sqrt(TRADING_DAYS) * r.mean() / neg.std(ddof=1) if len(neg) > 1 and neg.std(ddof=1) > 1e-12 else np.nan

def metrics_from_equity(eq):
    eq = pd.Series(eq).dropna()
    if len(eq) < 2:
        return {}
    r = eq.pct_change().dropna()
    total = eq.iloc[-1] / eq.iloc[0] - 1
    years = max((eq.index[-1] - eq.index[0]).days / 365.25, 1/365.25)
    cagr = (eq.iloc[-1] / eq.iloc[0]) ** (1 / years) - 1
    dd = eq / eq.cummax() - 1
    mdd = dd.min()
    vol = r.std(ddof=1) * np.sqrt(TRADING_DAYS) if len(r) > 1 else np.nan
    return {
        "Return [%]": total * 100,
        "CAGR [%]": cagr * 100,
        "Sharpe": annualized_sharpe(r),
        "Sortino": annualized_sortino(r),
        "Volatility [%]": vol * 100 if pd.notna(vol) else np.nan,
        "MaxDD [%]": mdd * 100 if pd.notna(mdd) else np.nan,
        "Calmar": cagr / abs(mdd) if pd.notna(mdd) and mdd < 0 else np.nan,
    }

def simulate_single_stock(features, start, end, buy_threshold, exit_threshold,
                          initial_cash=1_000_000.0, commission=0.0015, slippage=0.0005):
    """
    Backtest 1 cổ phiếu độc lập: Tín hiệu tại Close(t) chỉ được giao dịch ở Open(t+1).
    Chống hoàn toàn việc dùng giá Close hiện tại để mua ngay trong phiên.
    """
    active_close = stateful_active_from_score(features["TA_Score"], buy_threshold, exit_threshold)
    active_open = active_close.shift(1).fillna(0).astype(int)
    x = features.loc[(features.index >= pd.Timestamp(start)) & (features.index <= pd.Timestamp(end))].copy()
    active_open = active_open.reindex(x.index).fillna(0).astype(int)
    if x.empty:
        return {}

    cash, shares, state, trades = float(initial_cash), 0.0, 0, 0
    eq = []
    for dt, row in x.iterrows():
        d = int(active_open.loc[dt])
        op, cl = float(row["Open"]), float(row["Close"])
        if d != state:
            if d == 1:
                px = op * (1 + slippage)
                qty = math.floor(cash / (px * (1 + commission)))
                if qty > 0:
                    cash -= qty * px * (1 + commission)
                    shares = qty
                    state = 1
                    trades += 1
            else:
                px = op * (1 - slippage)
                cash += shares * px * (1 - commission)
                shares = 0
                state = 0
                trades += 1
        eq.append((dt, cash + shares * cl))

    m = metrics_from_equity(pd.Series(dict(eq)).sort_index())
    m["Trades"] = trades
    return m

# ======================================================================================
# 6. CORE TRAIN GRID SEARCH CHO NGƯỠNG BUY / EXIT
# ======================================================================================

def make_equal_folds(dates, start, end, n_folds=3):
    d = pd.DatetimeIndex(dates)
    d = d[(d >= pd.Timestamp(start)) & (d <= pd.Timestamp(end))]
    if len(d) < n_folds * 20:
        n_folds = max(1, len(d) // 20)
    chunks = np.array_split(d, n_folds)
    folds = []
    for i, c in enumerate(chunks, 1):
        if len(c) >= 20:
            folds.append({"Fold": f"F{i}", "start": pd.Timestamp(c[0]), "end": pd.Timestamp(c[-1])})
    return folds

def build_universe(price_dict, asof, benchmark_ticker=None, min_hist_floor=180,
                   min_hist_cap=400, lookback=60, min_pct=0.20):
    all_dates = set()
    for px in price_dict.values():
        all_dates.update(px.loc[:asof].index)
    pre_sessions = len(all_dates)
    min_hist = min(min_hist_cap, max(min_hist_floor, int(0.55 * pre_sessions)))

    rows = []
    for t, px in price_dict.items():
        if t == benchmark_ticker:
            continue
        h = px.loc[:asof]
        if h.empty:
            continue
        avg_value = (h["Close"] * h["Volume"]).tail(lookback).mean()
        rows.append({
            "Ticker": t,
            "HistorySessions": len(h),
            "AvgTradedValue": avg_value,
            "EligibleHistory": len(h) >= min_hist,
        })
    u = pd.DataFrame(rows)
    if u.empty:
        return u
    ok = u["EligibleHistory"] & u["AvgTradedValue"].notna()
    u.loc[ok, "LiquidityPercentile"] = u.loc[ok, "AvgTradedValue"].rank(pct=True)
    u["PassLiquidity"] = u["LiquidityPercentile"].fillna(0) >= min_pct
    u["Candidate"] = u["EligibleHistory"] & u["PassLiquidity"]
    return u.sort_values("AvgTradedValue", ascending=False).reset_index(drop=True)

def choose_global_ta_thresholds(feature_store, candidates, folds, penalty=0.20):
    buy_choices = [5, 6, 7]
    exit_choices = [2, 3, 4]
    trials = []
    for b, e in product(buy_choices, exit_choices):
        if e >= b:
            continue
        stock_scores = []
        for t in candidates:
            ms = []
            for f in folds:
                m = simulate_single_stock(feature_store[t], f["start"], f["end"], b, e)
                if m:
                    sh = m.get("Sharpe", np.nan)
                    sh = -3.0 if not np.isfinite(sh) else float(sh)
                    ms.append((sh, float(m.get("Return [%]", 0)), float(m.get("MaxDD [%]", -100))))
            if ms:
                a = np.array(ms, float)
                stock_scores.append([np.nanmean(a[:, 0]), np.nanmean(a[:, 1]), np.nanmin(a[:, 2])])
        if not stock_scores:
            continue

        z = np.array(stock_scores, float)
        med_sh = float(np.nanmedian(z[:, 0]))
        sd_sh = float(np.nanstd(z[:, 0]))
        med_ret = float(np.nanmedian(z[:, 1]))
        med_dd = float(np.nanmedian(z[:, 2]))
        score = med_sh - penalty * sd_sh + 0.005 * med_ret + 0.003 * med_dd
        trials.append({
            "BuyThreshold": b,
            "ExitThreshold": e,
            "GlobalCVScore": score,
            "MedianSharpe": med_sh,
            "SharpeStd": sd_sh,
            "MedianReturnPct": med_ret,
            "MedianMaxDDPct": med_dd,
            "NStocks": len(stock_scores),
        })

    tr = pd.DataFrame(trials).sort_values("GlobalCVScore", ascending=False).reset_index(drop=True)
    if tr.empty:
        return 5, 2, tr
    best = tr.iloc[0]
    return int(best["BuyThreshold"]), int(best["ExitThreshold"]), tr

# ======================================================================================
# 7. MULTI-FACTOR RANKING & SELECTION (NESTED WALK-FORWARD)
# ======================================================================================

def trailing_stock_quality(features, asof, b, e, lookback=126):
    h = features.loc[:asof]
    if len(h) < 40:
        return {"TrailingSharpe": -3.0, "TrailingReturnPct": -100.0, "TrailingMaxDDPct": -100.0}
    sub_idx = h.index[-min(lookback, len(h)):]
    m = simulate_single_stock(features, sub_idx[0], sub_idx[-1], b, e)
    sh = m.get("Sharpe", np.nan) if m else np.nan
    return {
        "TrailingSharpe": -3.0 if not np.isfinite(sh) else float(sh),
        "TrailingReturnPct": float(m.get("Return [%]", -100.0)) if m else -100.0,
        "TrailingMaxDDPct": float(m.get("MaxDD [%]", -100.0)) if m else -100.0,
    }

def rank_and_select_asof(feature_store, price_dict, benchmark_features, benchmark_ticker,
                         asof, b, e, top_n=5, max_vol_pct=0.75, min_dd_pct=0.15,
                         max_rsi=80.0, max_pairwise_corr=0.80):
    """
    Chọn Top N cổ phiếu tại thời điểm asof (phiên t-1).
    Hoàn toàn không truy cập bất kỳ dữ liệu nào sau asof.
    """
    universe = build_universe(price_dict, asof, benchmark_ticker)
    if universe.empty:
        raise ValueError(f"Universe rỗng tại {asof}")
    candidates = universe.loc[universe["Candidate"], "Ticker"].tolist()
    bench = benchmark_features.loc[:asof]
    bench_m126 = bench["Close"].pct_change(126).iloc[-1] if len(bench) > 126 else np.nan
    uidx = universe.set_index("Ticker")

    rows = []
    for t in candidates:
        feat = feature_store[t].loc[:asof]
        if feat.empty:
            continue
        snap = feat.iloc[-1]
        tq = trailing_stock_quality(feature_store[t], asof, b, e)
        rows.append({
            "Ticker": t,
            **tq,
            "TA_Score": float(snap.get("TA_Score", np.nan)),
            "Momentum63": float(snap.get("Momentum63", np.nan)),
            "Momentum126": float(snap.get("Momentum126", np.nan)),
            "RelativeStrength126": (
                float(snap.get("Momentum126", np.nan) - bench_m126)
                if pd.notna(bench_m126) else float(snap.get("Momentum126", np.nan))
            ),
            "Vol63": float(snap.get("Vol63", np.nan)),
            "RSI": float(snap.get("RSI", np.nan)),
            "LiquidityPercentile": float(uidx.loc[t, "LiquidityPercentile"]),
        })

    r = pd.DataFrame(rows)
    if r.empty:
        raise ValueError(f"Không có cổ phiếu đủ điều kiện tại {asof}")

    # Gate 1: Lọc TA Score tối thiểu
    r["PassTA"] = r["TA_Score"].fillna(0) >= 4
    if int(r["PassTA"].sum()) >= top_n:
        r = r[r["PassTA"]].copy()

    # Gate 2: Risk Gates (Biến động, Drawdown, RSI)
    vol_cap = r["Vol63"].quantile(max_vol_pct)
    dd_floor = r["TrailingMaxDDPct"].quantile(min_dd_pct)
    r["PassRiskGate"] = (
        (r["Vol63"].isna() | (r["Vol63"] <= vol_cap)) &
        (r["TrailingMaxDDPct"].isna() | (r["TrailingMaxDDPct"] >= dd_floor)) &
        (r["RSI"].isna() | (r["RSI"] <= max_rsi))
    )
    if int(r["PassRiskGate"].sum()) >= top_n:
        r = r[r["PassRiskGate"]].copy()

    # Xếp hạng tổng hợp đa yếu tố bằng Percentile Rank
    r["P_Trailing"] = r["TrailingSharpe"].rank(pct=True)
    r["P_TA"] = r["TA_Score"].rank(pct=True)
    r["P_RS"] = r["RelativeStrength126"].rank(pct=True)
    r["P_LowVol"] = (-r["Vol63"]).rank(pct=True)
    r["P_Liq"] = r["LiquidityPercentile"].rank(pct=True)
    mommix = 0.6 * r["Momentum126"].fillna(r["Momentum63"]) + 0.4 * r["Momentum63"].fillna(0)
    r["P_Momentum"] = mommix.rank(pct=True)

    r["FinalScore"] = (
        0.25 * r["P_Trailing"] +
        0.20 * r["P_TA"] +
        0.25 * r["P_RS"] +
        0.10 * r["P_LowVol"] +
        0.05 * r["P_Liq"] +
        0.15 * r["P_Momentum"]
    ) * 100

    r = r.sort_values(["FinalScore", "TrailingSharpe"], ascending=False).reset_index(drop=True)

    # Đa dạng hóa: Correlation cap để tránh chọn các mã quá tương quan
    corr_tickers = r["Ticker"].tolist()
    close_list = [feature_store[t].loc[:asof, "Close"].rename(t) for t in corr_tickers]
    corr_px = pd.concat(close_list, axis=1, join="inner").tail(127) if close_list else pd.DataFrame()
    corr = corr_px.pct_change(fill_method=None).dropna().corr() if len(corr_px) > 20 else pd.DataFrame()

    selected = []
    for _, row in r.iterrows():
        t = row["Ticker"]
        if selected and not corr.empty:
            too_close = any(
                t in corr.index and x in corr.columns and
                pd.notna(corr.loc[t, x]) and abs(corr.loc[t, x]) > max_pairwise_corr
                for x in selected
            )
            if too_close:
                continue
        selected.append(t)
        if len(selected) >= top_n:
            break

    if len(selected) < top_n:
        for t in r["Ticker"]:
            if t not in selected:
                selected.append(t)
            if len(selected) >= top_n:
                break

    locked = r[r["Ticker"].isin(selected)].copy()
    locked["SelectedOrder"] = locked["Ticker"].map({t: i+1 for i, t in enumerate(selected)})
    locked["AsOfDate"] = pd.Timestamp(asof)
    locked = locked.sort_values("SelectedOrder")
    return r, locked, selected

def first_session_each_period(calendar, start, end, freq="Q"):
    idx = pd.DatetimeIndex(calendar)
    idx = idx[(idx >= pd.Timestamp(start)) & (idx <= pd.Timestamp(end))]
    if len(idx) == 0:
        return []
    s = pd.Series(idx, index=idx)
    key = idx.to_period(freq)
    firsts = s.groupby(key).first().tolist()
    if idx[0] not in firsts:
        firsts = [idx[0]] + firsts
    return sorted(set(pd.Timestamp(x) for x in firsts))

def previous_session(calendar, dt):
    idx = pd.DatetimeIndex(calendar)
    prev = idx[idx < pd.Timestamp(dt)]
    return pd.Timestamp(prev[-1]) if len(prev) else None

def build_selection_schedule(feature_store, price_dict, benchmark_ticker, benchmark_features,
                             calendar, start, end, b, e, top_n=5, selection_freq="Q"):
    effective_dates = first_session_each_period(calendar, start, end, selection_freq)
    rows = []
    details = {}
    for eff in effective_dates:
        asof = previous_session(calendar, eff)
        if asof is None:
            continue
        ranking, locked, selected = rank_and_select_asof(
            feature_store, price_dict, benchmark_features, benchmark_ticker,
            asof, b, e, top_n=top_n
        )
        details[pd.Timestamp(eff)] = {"ranking": ranking, "locked": locked}
        row = {
            "EffectiveDate": pd.Timestamp(eff),
            "AsOfDate": pd.Timestamp(asof),
        }
        for i in range(top_n):
            row[f"Top{i+1}"] = selected[i] if i < len(selected) else None
        rows.append(row)

    schedule = pd.DataFrame(rows).sort_values("EffectiveDate").reset_index(drop=True)
    return schedule, details

def selected_tickers_asof(selection_schedule, effective_date, top_n=5):
    x = selection_schedule[selection_schedule["EffectiveDate"] <= pd.Timestamp(effective_date)]
    if x.empty:
        return []
    row = x.iloc[-1]
    return [row[f"Top{i+1}"] for i in range(top_n) if pd.notna(row.get(f"Top{i+1}"))]

# ======================================================================================
# 8. PORTFOLIO OPTIMIZATION MODELS: EW / MPT / MINVAR / HRP / BLEND
# ======================================================================================

def close_matrix(price_dict, tickers, end=None, start=None):
    series = [price_dict[t]["Close"].rename(t) for t in tickers if t in price_dict]
    if not series:
        return pd.DataFrame()
    m = pd.concat(series, axis=1, join="inner").sort_index().dropna(how="any")
    if start is not None:
        m = m.loc[m.index >= pd.Timestamp(start)]
    if end is not None:
        m = m.loc[m.index <= pd.Timestamp(end)]
    return m

def equal_weights(tickers):
    tickers = list(tickers)
    return pd.Series(1/len(tickers), index=tickers, dtype=float) if tickers else pd.Series(dtype=float)

def robust_mu(prices, lookback=252, shrinkage=0.50):
    r = prices.pct_change(fill_method=None).dropna().tail(lookback)
    if r.empty:
        return pd.Series(0.0, index=prices.columns)
    span = max(20, min(lookback, len(r)))
    raw = r.ewm(span=span, adjust=False).mean().iloc[-1] * TRADING_DAYS
    grand = float(raw.mean())
    return shrinkage * raw + (1 - shrinkage) * grand

def lw_cov(prices, lookback=252):
    r = prices.pct_change(fill_method=None).dropna().tail(lookback)
    if len(r) < 30:
        return r.cov() * TRADING_DAYS
    est = LedoitWolf().fit(r.values)
    return pd.DataFrame(est.covariance_ * TRADING_DAYS, index=r.columns, columns=r.columns)

def max_sharpe_weights(mu, cov, max_weight=0.50, l2_gamma=0.03):
    tickers = list(mu.index)
    n = len(tickers)
    if n == 0:
        return pd.Series(dtype=float)
    max_weight = max(max_weight, 1/n)
    muv = mu.values.astype(float)
    cv = cov.reindex(index=tickers, columns=tickers).fillna(0).values.astype(float) + np.eye(n)*1e-8

    def obj(w):
        ret = float(w @ muv)
        vol = math.sqrt(max(float(w @ cv @ w), 1e-12))
        return -(ret / vol) + l2_gamma * float(np.sum(w * w))

    x0 = np.repeat(1/n, n)
    bounds = [(0, max_weight)] * n
    cons = ({"type": "eq", "fun": lambda w: np.sum(w) - 1},)
    res = minimize(obj, x0, method="SLSQP", bounds=bounds, constraints=cons, options={"maxiter": 1000, "ftol": 1e-10})
    if not res.success or np.any(~np.isfinite(res.x)):
        return equal_weights(tickers)
    w = np.clip(res.x, 0, None)
    return pd.Series(w / w.sum(), index=tickers) if w.sum() > 0 else equal_weights(tickers)

def min_variance_weights(cov, max_weight=0.50):
    tickers = list(cov.index)
    n = len(tickers)
    if n == 0:
        return pd.Series(dtype=float)
    max_weight = max(max_weight, 1/n)
    cv = cov.values.astype(float) + np.eye(n)*1e-8

    def obj(w):
        return float(w @ cv @ w)

    x0 = np.repeat(1/n, n)
    bounds = [(0, max_weight)] * n
    cons = ({"type": "eq", "fun": lambda w: np.sum(w) - 1},)
    res = minimize(obj, x0, method="SLSQP", bounds=bounds, constraints=cons, options={"maxiter": 1000, "ftol": 1e-12})
    if not res.success:
        return equal_weights(tickers)
    w = np.clip(res.x, 0, None)
    return pd.Series(w / w.sum(), index=tickers)

def _cluster_var(cov, items):
    sub = cov.loc[items, items].values
    ivp = 1 / np.diag(sub)
    ivp = ivp / ivp.sum()
    return float(ivp @ sub @ ivp)

def _quasi_diag(link):
    link = link.astype(int)
    sort_ix = pd.Series([link[-1, 0], link[-1, 1]])
    num_items = int(link[-1, 3])
    while sort_ix.max() >= num_items:
        sort_ix.index = range(0, sort_ix.shape[0] * 2, 2)
        df0 = sort_ix[sort_ix >= num_items]
        i = df0.index
        j = df0.values - num_items
        sort_ix.loc[i] = link[j, 0]
        df1 = pd.Series(link[j, 1], index=i + 1)
        sort_ix = pd.concat([sort_ix, df1]).sort_index()
        sort_ix.index = range(sort_ix.shape[0])
    return sort_ix.tolist()

def hrp_weights(prices, method="single"):
    r = prices.pct_change(fill_method=None).dropna()
    if len(r) < 30 or prices.shape[1] < 2:
        return equal_weights(prices.columns)
    cov = r.cov() * TRADING_DAYS
    corr = r.corr().clip(-1, 1)
    dist = np.sqrt((1 - corr) / 2)
    link = linkage(squareform(dist.values, checks=False), method=method)
    ordered = corr.index[_quasi_diag(link)].tolist()

    w = pd.Series(1.0, index=ordered)
    clusters = [ordered]
    while clusters:
        nxt = []
        for c in clusters:
            if len(c) <= 1:
                continue
            s = len(c) // 2
            c0, c1 = c[:s], c[s:]
            v0, v1 = _cluster_var(cov, c0), _cluster_var(cov, c1)
            a = 1 - v0 / (v0 + v1) if v0 + v1 > 0 else 0.5
            w[c0] *= a
            w[c1] *= (1 - a)
            nxt += [c0, c1]
        clusters = nxt
    return w.reindex(prices.columns).fillna(0) / w.sum()

def compute_weights(model_family, prices):
    tickers = list(prices.columns)
    ew = equal_weights(tickers)
    if len(tickers) == 0 or model_family == "EW":
        return ew

    cov = lw_cov(prices, 252)
    if model_family == "MINVAR_LW":
        return min_variance_weights(cov)
    if model_family == "HRP":
        return hrp_weights(prices.tail(252))

    mu = robust_mu(prices, 252)
    mpt = max_sharpe_weights(mu, cov)
    if model_family == "MPT_LW":
        return mpt
    if model_family == "BLEND50":
        w = 0.5 * mpt + 0.5 * ew
        return w / w.sum()
    return ew

# ======================================================================================
# 9. REGIME, MASKS & DYNAMIC REBALANCING SCHEDULE
# ======================================================================================

def build_market_regime(benchmark_features):
    x = benchmark_features.copy()
    ema_m = x["Close"].ewm(span=200, adjust=False, min_periods=200).mean()
    slope = ema_m - ema_m.shift(20)
    risk_close = (x["Close"] > ema_m) & (slope > 0)
    return pd.DataFrame({
        "EMA_Market": ema_m,
        "EMA_Slope": slope,
        "RiskOn_Close": risk_close.fillna(False),
        "RiskOn_Open": risk_close.shift(1).fillna(False),
    }, index=x.index)

def build_stock_open_masks(feature_store, tickers, b, e):
    masks = {}
    for t in tickers:
        active_close = stateful_active_from_score(feature_store[t]["TA_Score"], b, e)
        masks[t] = active_close.shift(1).fillna(0).astype(bool)
    return pd.DataFrame(masks).sort_index()

def build_dynamic_weight_schedule(price_dict, selection_schedule, calendar, start, end,
                                  model_family, top_n=5, allocation_freq="M"):
    rebalance_dates = first_session_each_period(calendar, start, end, allocation_freq)
    rows = []
    all_tickers = set()

    for eff in rebalance_dates:
        asof = previous_session(calendar, eff)
        if asof is None:
            continue
        top = selected_tickers_asof(selection_schedule, eff, top_n=top_n)
        if not top:
            continue
        all_tickers.update(top)
        prices = close_matrix(price_dict, top, end=asof)
        if prices.empty or len(prices) < 30:
            w = equal_weights(top)
        else:
            w = compute_weights(model_family, prices.tail(253))
        row = {"EffectiveDate": pd.Timestamp(eff), "AsOfDate": pd.Timestamp(asof)}
        row.update({t: float(w.get(t, 0.0)) for t in top})
        rows.append(row)

    if not rows:
        raise ValueError("Không tạo được weight schedule")

    meta = pd.DataFrame(rows).sort_values("EffectiveDate").reset_index(drop=True)
    tickers = sorted(all_tickers)
    weights = meta.set_index("EffectiveDate").reindex(columns=tickers).fillna(0.0)
    return weights, meta[["EffectiveDate", "AsOfDate"]]

def build_static_buyhold_schedule(price_dict, selection_schedule, calendar, start, end, top_n=5):
    start = pd.Timestamp(start)
    top = selected_tickers_asof(selection_schedule, start, top_n=top_n)
    if not top:
        raise ValueError("Không có Top N ở đầu kỳ cho Buy & Hold")
    return pd.DataFrame([equal_weights(top)], index=[start])

# ======================================================================================
# 10. EVENT-DRIVEN PORTFOLIO SIMULATION ENGINE
# ======================================================================================

def run_portfolio_engine_dynamic(price_dict, weight_schedule, calendar, start, end,
                                 market_regime_open=None, stock_active_open=None,
                                 use_market_regime=True, use_stock_mask=True,
                                 initial_cash=1_000_000_000.0, commission=0.0015,
                                 slippage=0.0005, fractional=False):
    tickers = list(weight_schedule.columns)
    idx = pd.DatetimeIndex(calendar)
    idx = idx[(idx >= pd.Timestamp(start)) & (idx <= pd.Timestamp(end))]
    if len(idx) == 0:
        raise ValueError("Không có ngày giao dịch trong kỳ.")

    opens = pd.DataFrame(index=idx, columns=tickers, dtype=float)
    closes = pd.DataFrame(index=idx, columns=tickers, dtype=float)
    for t in tickers:
        if t in price_dict:
            opens[t] = price_dict[t]["Open"].reindex(idx)
            closes[t] = price_dict[t]["Close"].reindex(idx)

    ws = weight_schedule.sort_index().reindex(columns=tickers)
    ws = ws.reindex(ws.index.union(idx)).sort_index().ffill().reindex(idx).fillna(0.0)

    regime = (
        pd.Series(True, index=idx)
        if market_regime_open is None
        else pd.Series(market_regime_open).reindex(idx).ffill().fillna(False).astype(bool)
    )
    masks = (
        pd.DataFrame(True, index=idx, columns=tickers)
        if stock_active_open is None
        else stock_active_open.reindex(index=idx, columns=tickers).fillna(False).astype(bool)
    )

    cash = float(initial_cash)
    shares = pd.Series(0.0, index=tickers)
    last_close = pd.Series(np.nan, index=tickers)
    last_target = pd.Series(np.nan, index=tickers)

    orders, eq_rows, exposure_rows = [], [], []
    turnover, rebals = 0.0, 0

    for dt in idx:
        op = opens.loc[dt].astype(float)
        cl = closes.loc[dt].astype(float)

        valid_close = cl.notna()
        last_close.loc[valid_close] = cl.loc[valid_close]
        mark_open = op.where(op.notna(), last_close)
        mark_close = cl.where(cl.notna(), last_close)

        base = ws.loc[dt].clip(lower=0)
        if base.sum() > 1 + 1e-8:
            base = base / base.sum()

        risk_on = bool(regime.loc[dt]) if use_market_regime else True
        active = masks.loc[dt].astype(bool) if use_stock_mask else pd.Series(True, index=tickers)

        target = base.copy()
        if use_stock_mask:
            target *= active.astype(float)
        if use_market_regime and not risk_on:
            target[:] = 0.0

        # Lưu ý cốt lõi: Không re-normalize lại sau mask -> phần thiếu trở thành CASH buffer
        target = target.clip(lower=0)
        if target.sum() > 1 + 1e-8:
            target /= target.sum()

        changed = last_target.isna().any() or (target - last_target.fillna(0)).abs().max() > 1e-10
        if changed:
            equity_open = cash + float((shares * mark_open.fillna(0)).sum())
            desired = shares.copy()

            for t in tickers:
                if pd.isna(op[t]) or op[t] <= 0:
                    continue
                if target[t] <= 0:
                    desired[t] = 0.0
                else:
                    px = op[t] * (1 + slippage)
                    raw_qty = equity_open * target[t] / (px * (1 + commission))
                    desired[t] = raw_qty if fractional else math.floor(raw_qty)

            # Bán trước
            for t in tickers:
                d = desired[t] - shares[t]
                if d < -1e-12 and pd.notna(op[t]) and op[t] > 0:
                    qty = -d
                    px = op[t] * (1 - slippage)
                    notional = qty * px
                    fee = notional * commission
                    cash += notional - fee
                    shares[t] -= qty
                    turnover += notional
                    orders.append({"Date": dt, "Ticker": t, "Side": "SELL", "Shares": qty, "ExecPrice": px, "Fee": fee, "Reason": "TARGET"})

            # Mua sau
            for t in tickers:
                d = desired[t] - shares[t]
                if d > 1e-12 and pd.notna(op[t]) and op[t] > 0:
                    px = op[t] * (1 + slippage)
                    unit = px * (1 + commission)
                    qty = min(d, cash / unit) if fractional else min(d, math.floor(cash / unit))
                    if qty > 0:
                        notional = qty * px
                        fee = notional * commission
                        cash -= notional + fee
                        shares[t] += qty
                        turnover += notional
                        orders.append({"Date": dt, "Ticker": t, "Side": "BUY", "Shares": qty, "ExecPrice": px, "Fee": fee, "Reason": "TARGET"})

            last_target = target.copy()
            rebals += 1

        equity = float(cash + (shares * mark_close.fillna(0)).sum())
        invested = float((shares * mark_close.fillna(0)).sum())
        exposure = invested / equity if equity > 0 else 0.0
        exposure_rows.append(exposure)
        eq_rows.append((dt, equity))

    eq = pd.Series(dict(eq_rows)).sort_index()
    st_res = metrics_from_equity(eq)
    st_res.update({
        "#Orders": len(orders),
        "#Rebalances": rebals,
        "Turnover [%]": 100 * turnover / initial_cash,
        "AvgExposure [%]": 100 * np.mean(exposure_rows) if exposure_rows else 100.0,
        "AvgCash [%]": 100 * (1 - np.mean(exposure_rows)) if exposure_rows else 0.0,
    })
    return st_res, eq, pd.DataFrame(orders)

def benchmark_nav(benchmark_px, start, end, initial_cash=1_000_000_000.0):
    x = benchmark_px.loc[(benchmark_px.index >= pd.Timestamp(start)) & (benchmark_px.index <= pd.Timestamp(end)), "Close"].dropna()
    if x.empty:
        return pd.Series(dtype=float)
    return initial_cash * x / x.iloc[0]

# ======================================================================================
# 11. MODEL VALIDATION & FINAL TEST EVALUATIONS
# ======================================================================================

ACTIVE_MODEL_SPECS = {
    "DYN_EW_REGIME":      {"family": "EW",        "regime": True, "mask": True},
    "DYN_MPT_LW_REGIME":  {"family": "MPT_LW",    "regime": True, "mask": True},
    "DYN_MINVAR_REGIME":  {"family": "MINVAR_LW", "regime": True, "mask": True},
    "DYN_HRP_REGIME":     {"family": "HRP",       "regime": True, "mask": True},
    "DYN_BLEND50_REGIME": {"family": "BLEND50",   "regime": True, "mask": True},
}

def model_validation_score(stats, benchmark_stats):
    sh = stats.get("Sharpe", np.nan)
    sh = -3 if not np.isfinite(sh) else sh
    so = stats.get("Sortino", np.nan)
    so = -3 if not np.isfinite(so) else so
    ret = stats.get("Return [%]", -100)
    dd = stats.get("MaxDD [%]", -100)
    turnover = stats.get("Turnover [%]", 0)
    bret = benchmark_stats.get("Return [%]", 0)
    return 0.45 * sh + 0.15 * so + 0.012 * (ret - bret) + 0.004 * dd - 0.0005 * turnover

def evaluate_dynamic_models(price_dict, feature_store, benchmark_ticker, benchmark_features,
                            selection_schedule, calendar, start, end, b, e, top_n=5,
                            use_market_regime=True, use_stock_mask=True):
    regime = build_market_regime(benchmark_features)
    union_tickers = sorted({
        selection_schedule.loc[i, f"Top{j+1}"]
        for i in selection_schedule.index for j in range(top_n)
        if pd.notna(selection_schedule.loc[i, f"Top{j+1}"])
    })
    masks = build_stock_open_masks(feature_store, union_tickers, b, e)
    bnav = benchmark_nav(price_dict[benchmark_ticker], start, end)
    bstats = metrics_from_equity(bnav)

    rows, curves, orders, schedules = [], {"BENCHMARK": bnav}, {}, {}
    for name, spec in ACTIVE_MODEL_SPECS.items():
        sched, _ = build_dynamic_weight_schedule(
            price_dict, selection_schedule, calendar, start, end, spec["family"], top_n=top_n
        )
        st_res, eq, od = run_portfolio_engine_dynamic(
            price_dict, sched, calendar, start, end,
            market_regime_open=regime["RiskOn_Open"],
            stock_active_open=masks,
            use_market_regime=(use_market_regime and spec["regime"]),
            use_stock_mask=(use_stock_mask and spec["mask"]),
        )
        st_res["Model"] = name
        st_res["ValidationScore"] = model_validation_score(st_res, bstats)
        rows.append(st_res)
        curves[name] = eq
        orders[name] = od
        schedules[name] = sched

    bstats.update({"Model": "BENCHMARK", "ValidationScore": np.nan})
    rows.append(bstats)
    res = pd.DataFrame(rows).set_index("Model")
    champion = res.drop(index="BENCHMARK")["ValidationScore"].idxmax()
    return res, champion, curves, orders, schedules

def evaluate_final_test(price_dict, feature_store, benchmark_ticker, benchmark_features,
                        selection_schedule, calendar, start, end, b, e, champion, top_n=5,
                        use_market_regime=True, use_stock_mask=True):
    regime = build_market_regime(benchmark_features)
    union_tickers = sorted({
        selection_schedule.loc[i, f"Top{j+1}"]
        for i in selection_schedule.index for j in range(top_n)
        if pd.notna(selection_schedule.loc[i, f"Top{j+1}"])
    })
    masks = build_stock_open_masks(feature_store, union_tickers, b, e)
    bnav = benchmark_nav(price_dict[benchmark_ticker], start, end)
    rows = []
    curves = {"VNINDEX_or_MarketBenchmark": bnav}
    order_map = {}

    bstats = metrics_from_equity(bnav)
    bstats["Model"] = "VNINDEX_or_MarketBenchmark"
    rows.append(bstats)

    # Static Buy & Hold EW
    bh_sched = build_static_buyhold_schedule(price_dict, selection_schedule, calendar, start, end, top_n=top_n)
    bh_stats, bh_eq, bh_od = run_portfolio_engine_dynamic(
        price_dict, bh_sched, calendar, start, end,
        market_regime_open=None, stock_active_open=None,
        use_market_regime=False, use_stock_mask=False,
    )
    bh_stats["Model"] = "STATIC_EW_BH"
    rows.append(bh_stats)
    curves["STATIC_EW_BH"] = bh_eq
    order_map["STATIC_EW_BH"] = bh_od

    # Dynamic EW không Regime
    ew_sched, _ = build_dynamic_weight_schedule(price_dict, selection_schedule, calendar, start, end, "EW", top_n=top_n)
    ew_stats, ew_eq, ew_od = run_portfolio_engine_dynamic(
        price_dict, ew_sched, calendar, start, end,
        market_regime_open=None, stock_active_open=None,
        use_market_regime=False, use_stock_mask=False,
    )
    ew_stats["Model"] = "DYN_EW_NO_REGIME"
    rows.append(ew_stats)
    curves["DYN_EW_NO_REGIME"] = ew_eq
    order_map["DYN_EW_NO_REGIME"] = ew_od

    # Champion và các mô hình chẩn đoán
    for name, spec in ACTIVE_MODEL_SPECS.items():
        sched, _ = build_dynamic_weight_schedule(price_dict, selection_schedule, calendar, start, end, spec["family"], top_n=top_n)
        st_res, eq, od = run_portfolio_engine_dynamic(
            price_dict, sched, calendar, start, end,
            market_regime_open=regime["RiskOn_Open"],
            stock_active_open=masks,
            use_market_regime=use_market_regime,
            use_stock_mask=use_stock_mask,
        )
        label = "LOCKED_CHAMPION" if name == champion else "DIAG_" + name
        st_res["Model"] = label
        st_res["OriginalModel"] = name
        rows.append(st_res)
        curves[label] = eq
        order_map[label] = od

    res = pd.DataFrame(rows).set_index("Model")
    if "STATIC_EW_BH" in res.index:
        res["ExcessCAGR_vs_StaticEW [pp]"] = res["CAGR [%]"] - res.loc["STATIC_EW_BH", "CAGR [%]"]
        res["ExcessSharpe_vs_StaticEW"] = res["Sharpe"] - res.loc["STATIC_EW_BH", "Sharpe"]
    bname = "VNINDEX_or_MarketBenchmark"
    res["ExcessCAGR_vs_Benchmark [pp]"] = res["CAGR [%]"] - res.loc[bname, "CAGR [%]"]
    res["DD_Improvement_vs_Benchmark [pp]"] = res["MaxDD [%]"] - res.loc[bname, "MaxDD [%]"]
    return res, curves, order_map

def generate_leakage_audit(core_end, val_end, test_end):
    return pd.DataFrame([
        ["1. TA BUY / EXIT Selection", str(core_end.date()), False, "Ngưỡng kỹ thuật chỉ tối ưu trên Core Train qua K-Fold CV."],
        ["2. Model Family Selection", str(val_end.date()), False, "Họ phân bổ (EW/MPT/MinVar/HRP/Blend) khóa tại Validation, không đổi."],
        ["3. Quarterly Top N Reselection", "t - 1", False, "Top N tại phiên t chỉ tính bằng chỉ báo và giá đến hết phiên t-1."],
        ["4. Monthly Weight Rebalancing", "t - 1", False, "Ma trận hiệp phương sai Ledoit-Wolf chỉ tính bằng dữ liệu đến t-1."],
        ["5. Execution Timing (Open-Close)", "t-1 Close -> t Open", False, "Tín hiệu sinh từ Close(t-1) khớp lệnh tại Open(t) với Slippage + Phí."],
        ["6. Market Regime / Risk Overlay", "t-1 Close -> t Open", False, "Đường EMA200 & Slope của chỉ số tính tại t-1, áp dụng Open(t). Phần loại thành Cash."],
        ["7. Final Test Evaluation", str(test_end.date()), True, "Final Test đóng vai trò OOS thuần túy; hoàn toàn không retune lại tham số."],
    ], columns=["Quy Trình / Giai Đoạn", "Ngày Dữ Liệu Tối Đa Sử Dụng", "Cho Phép Retune?", "Nguyên Tắc Chống Rò Rỉ"])


# ======================================================================================
# 12. STREAMLIT INTERFACE & RUNNER
# ======================================================================================

st.markdown('<div class="main-title">📈 HOSE QUANT V4: ĐỊNH LƯỢNG CỔ PHIẾU & QUẢN TRỊ RỦI RO</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Nền tảng kiểm định quy trình đầu tư định lượng trên rổ cổ phiếu HOSE | Chống Future Leakage với Nested Walk-Forward</div>', unsafe_allow_html=True)

st.markdown("""
<div class="badge-container">
    <span class="badge badge-blue">🛡️ Causal Integrity: Zero Future Leakage</span>
    <span class="badge badge-green">🔄 Nested Walk-Forward (Quarterly Top N)</span>
    <span class="badge badge-purple">⚖️ Allocation: EW / MPT / MinVar / HRP / Blend</span>
    <span class="badge badge-amber">📊 Risk Overlay: Market Regime + Cash Buffer</span>
</div>
""", unsafe_allow_html=True)

# ----------------- SIDEBAR CONTROLS -----------------
st.sidebar.header("📁 Dữ Liệu Đầu Vào")

uploaded_file = st.sidebar.file_uploader("Tải file dữ liệu (Excel .xlsx hoặc CSV)", type=["xlsx", "xls", "csv"])

# Tự động tìm kiếm file dữ liệu mặc định ở thư mục hiện tại hoặc thư mục lân cận
search_folders = [Path("."), Path("..")]
found_default = None
target_names = [
    "HOSE - 01.01.2021 - 31.12.2025 (97 ma).xlsx",
    "HOSE - 01.01.2021 - 31.12.2025 (97 mã) (1).xlsx",
    "HOSE - 01.01.2021 - 31.12.2025.xlsx",
]

for folder in search_folders:
    for tname in target_names:
        p = folder / tname
        if p.exists() and p.is_file():
            found_default = str(p.resolve())
            break
    if found_default:
        break

if not found_default:
    for folder in search_folders:
        for ext in ["*.xlsx", "*.xls", "*.csv"]:
            for match in folder.glob(ext):
                if "hose" in match.name.lower():
                    found_default = str(match.resolve())
                    break
            if found_default:
                break
        if found_default:
            break

if not found_default:
    for folder in search_folders:
        xlsx_list = list(folder.glob("*.xlsx"))
        if xlsx_list:
            found_default = str(xlsx_list[0].resolve())
            break

data_source = None
if uploaded_file is not None:
    data_source = uploaded_file
    st.sidebar.success(f"Đã chọn: {uploaded_file.name}")
elif found_default:
    data_source = found_default
    st.sidebar.info(f"Dữ liệu mặc định: {os.path.basename(found_default)}")
else:
    st.sidebar.warning("Vui lòng tải lên file dữ liệu HOSE dạng Excel hoặc CSV.")

st.sidebar.header("⚙️ Cấu Hình Chiến Lược")

col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    top_n_input = st.number_input("Top N Cổ Phiếu", min_value=3, max_value=15, value=5, step=1)
with col_sb2:
    split_mode_input = st.selectbox("Phân Chia Dữ Liệu", ["auto", "proportions"], index=0, help="Auto: Ưu tiên năm cuối làm Final Test")

use_adj_price = st.sidebar.checkbox("Ưu tiên giá điều chỉnh (Adjusted)", value=True)
use_regime = st.sidebar.checkbox("Market Regime Filter (EMA200)", value=True, help="Nếu VNINDEX < EMA200 hoặc dốc âm, tỷ trọng về 0 và chuyển sang Cash")
use_mask = st.sidebar.checkbox("Stock Active Mask (TA Score)", value=True, help="Nếu cổ phiếu bị tín hiệu EXIT, hạ tỷ trọng về 0 và đưa vào Cash")

st.sidebar.subheader("Chi Phí Giao Dịch")
col_c1, col_c2 = st.sidebar.columns(2)
with col_c1:
    commission_pct = st.number_input("Phí GD (%)", min_value=0.0, max_value=1.0, value=0.15, step=0.01) / 100.0
with col_c2:
    slippage_pct = st.number_input("Slippage (%)", min_value=0.0, max_value=1.0, value=0.05, step=0.01) / 100.0

st.sidebar.subheader("Tùy Chọn Tần Suất")
selection_freq = st.sidebar.selectbox("Tái chọn Top N", ["Q", "M", "2Q", "Y"], index=0, help="Q = Theo Quý, M = Theo Tháng")
allocation_freq = st.sidebar.selectbox("Tái cân bằng Tỷ trọng", ["M", "Q", "2W"], index=0, help="M = Hàng Tháng")

run_btn = st.sidebar.button("🚀 Chạy Kiểm Định Toàn Diện", type="primary", use_container_width=True)

# ----------------- CACHED PIPELINE EXECUTION -----------------
@st.cache_data(show_spinner=False)
def load_and_preprocess_data(file_bytes_or_path, use_adj):
    raw = read_any_tabular_data(file_bytes_or_path)
    clean_data, source_type = standardize_ohlcv(raw, use_adj)
    q_table = data_quality_table(clean_data)
    return clean_data, source_type, q_table

@st.cache_data(show_spinner=False)
def compute_all_indicators(data_df):
    p_dict = build_price_dict(data_df)
    bench_tk = next((t for t in p_dict if t in BENCHMARK_ALIASES), None)
    if bench_tk is None:
        bench_tk = "__MARKET_PROXY__"
        p_dict[bench_tk] = build_synthetic_market_proxy(p_dict)
    f_store = {t: calculate_ta_features(px) for t, px in p_dict.items()}
    return p_dict, f_store, bench_tk

# ----------------- MAIN APP LOGIC -----------------
if data_source is None:
    st.info("👋 Hãy tải lên file dữ liệu HOSE (ví dụ: `HOSE - 01.01.2021 - 31.12.2025 (97 ma).xlsx`) từ thanh bên trái để bắt đầu kiểm định.")
else:
    with st.spinner("Đang chuẩn hóa dữ liệu và tính toán chỉ báo kỹ thuật..."):
        try:
            clean_data, source_type, q_table = load_and_preprocess_data(data_source, use_adj_price)
            price_dict, feature_store, benchmark_ticker = compute_all_indicators(clean_data)
            dates, core_end, val_start, val_end, test_start, test_end = resolve_auto_splits(clean_data, split_mode=split_mode_input)
            
            split_df = pd.DataFrame([
                {"Giai Đoạn": "Core Train", "Bắt Đầu": dates[0].strftime("%Y-%m-%d"), "Kết Thúc": core_end.strftime("%Y-%m-%d"), "Số Phiên": int((dates <= core_end).sum()), "Nhiệm Vụ": "Tối ưu hóa ngưỡng BUY/EXIT chung qua K-Fold CV"},
                {"Giai Đoạn": "Validation", "Bắt Đầu": val_start.strftime("%Y-%m-%d"), "Kết Thúc": val_end.strftime("%Y-%m-%d"), "Số Phiên": int(((dates >= val_start) & (dates <= val_end)).sum()), "Nhiệm Vụ": "Chạy Nested Walk-Forward để chọn và Khóa Champion"},
                {"Giai Đoạn": "Final Test (OOS)", "Bắt Đầu": test_start.strftime("%Y-%m-%d"), "Kết Thúc": test_end.strftime("%Y-%m-%d"), "Số Phiên": int((dates >= test_start).sum()), "Nhiệm Vụ": "Đánh giá mô hình Champion OOS thuần túy, không retune"},
            ])
        except Exception as e:
            st.error(f"Lỗi khi đọc hoặc xử lý dữ liệu: {e}")
            st.stop()

    # Lưu trữ trạng thái chạy pipeline vào session state
    if "pipeline_run" not in st.session_state:
        st.session_state.pipeline_run = False

    if run_btn or not st.session_state.pipeline_run:
        with st.spinner("Đang chạy Grid Search Core Train, Validation Champion và Final Test OOS..."):
            # 1. Core Train Thresholds
            core_universe = build_universe(price_dict, core_end, benchmark_ticker)
            core_candidates = core_universe.loc[core_universe["Candidate"], "Ticker"].tolist()
            core_dates = dates[dates <= core_end]
            inner_n = max(63, min(126, int(0.30 * len(core_dates))))
            inner_start = core_dates[-inner_n]
            threshold_folds = make_equal_folds(core_dates, inner_start, core_end, 3)
            best_buy, best_exit, threshold_trials = choose_global_ta_thresholds(feature_store, core_candidates, threshold_folds)

            # 2. Validation Schedule & Selection
            val_selection, val_selection_details = build_selection_schedule(
                feature_store, price_dict, benchmark_ticker, feature_store[benchmark_ticker],
                dates, val_start, val_end, best_buy, best_exit, top_n=top_n_input, selection_freq=selection_freq
            )

            # 3. Validation Model Evaluation & Lock Champion
            val_results, champion, val_curves, val_orders, val_schedules = evaluate_dynamic_models(
                price_dict, feature_store, benchmark_ticker, feature_store[benchmark_ticker],
                val_selection, dates, val_start, val_end, best_buy, best_exit, top_n=top_n_input,
                use_market_regime=use_regime, use_stock_mask=use_mask
            )

            # 4. Final Test Schedule & Evaluation
            test_selection, test_selection_details = build_selection_schedule(
                feature_store, price_dict, benchmark_ticker, feature_store[benchmark_ticker],
                dates, test_start, test_end, best_buy, best_exit, top_n=top_n_input, selection_freq=selection_freq
            )

            test_results, test_curves, test_orders = evaluate_final_test(
                price_dict, feature_store, benchmark_ticker, feature_store[benchmark_ticker],
                test_selection, dates, test_start, test_end, best_buy, best_exit, champion, top_n=top_n_input,
                use_market_regime=use_regime, use_stock_mask=use_mask
            )

            leakage_df = generate_leakage_audit(core_end, val_end, test_end)

            st.session_state.pipeline_results = {
                "best_buy": best_buy,
                "best_exit": best_exit,
                "threshold_trials": threshold_trials,
                "val_selection": val_selection,
                "val_selection_details": val_selection_details,
                "val_results": val_results,
                "champion": champion,
                "val_curves": val_curves,
                "test_selection": test_selection,
                "test_selection_details": test_selection_details,
                "test_results": test_results,
                "test_curves": test_curves,
                "test_orders": test_orders,
                "leakage_df": leakage_df,
            }
            st.session_state.pipeline_run = True

    res = st.session_state.pipeline_results
    best_buy = res["best_buy"]
    best_exit = res["best_exit"]
    champion = res["champion"]
    test_results = res["test_results"]
    test_curves = res["test_curves"]
    test_orders = res["test_orders"]
    val_results = res["val_results"]
    leakage_df = res["leakage_df"]
    test_selection = res["test_selection"]

    # ----------------- TOP KPI METRICS DISPLAY -----------------
    champ_stats = test_results.loc["LOCKED_CHAMPION"] if "LOCKED_CHAMPION" in test_results.index else None
    bench_stats = test_results.loc["VNINDEX_or_MarketBenchmark"] if "VNINDEX_or_MarketBenchmark" in test_results.index else None
    bh_stats = test_results.loc["STATIC_EW_BH"] if "STATIC_EW_BH" in test_results.index else None

    if champ_stats is not None:
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.markdown(f"""<div class="metric-card"><div class="metric-label">Champion Model</div>
            <div class="metric-val" style="color:#1E3A8A; font-size:1.15rem;">{champion}</div></div>""", unsafe_allow_html=True)
        with c2:
            cagr_val = champ_stats.get('CAGR [%]', 0.0)
            cagr_color = "#10B981" if cagr_val > 0 else "#EF4444"
            st.markdown(f"""<div class="metric-card"><div class="metric-label">CAGR (Final Test)</div>
            <div class="metric-val" style="color:{cagr_color}">{cagr_val:.2f}%</div></div>""", unsafe_allow_html=True)
        with c3:
            sh_val = champ_stats.get('Sharpe', 0.0)
            st.markdown(f"""<div class="metric-card"><div class="metric-label">Sharpe Ratio</div>
            <div class="metric-val" style="color:#3B82F6">{sh_val:.2f}</div></div>""", unsafe_allow_html=True)
        with c4:
            mdd_val = champ_stats.get('MaxDD [%]', 0.0)
            st.markdown(f"""<div class="metric-card"><div class="metric-label">Max Drawdown</div>
            <div class="metric-val" style="color:#F59E0B">{mdd_val:.2f}%</div></div>""", unsafe_allow_html=True)
        with c5:
            exc_cagr = champ_stats.get('ExcessCAGR_vs_Benchmark [pp]', 0.0)
            exc_color = "#10B981" if exc_cagr > 0 else "#EF4444"
            st.markdown(f"""<div class="metric-card"><div class="metric-label">Excess vs VN-Index</div>
            <div class="metric-val" style="color:{exc_color}">{exc_cagr:+.2f} pp</div></div>""", unsafe_allow_html=True)

    st.write("")

    # ----------------- DETAILED TABS -----------------
    tabs = st.tabs([
        "📊 1. Dữ Liệu & Phân Kỳ",
        "⚙️ 2. Ngưỡng TA Core Train",
        "🔄 3. Top N Walk-Forward",
        "🏆 4. Khóa Champion (Validation)",
        "📈 5. Kiểm Định OOS Final Test",
        "🛡️ 6. Kiểm Toán Rò Rỉ Tương Lai",
        "📥 7. Xuất Báo Cáo"
    ])

    # ===== TAB 1: DỮ LIỆU & PHÂN KỲ =====
    with tabs[0]:
        st.subheader("1. Tổng Quan Dữ Liệu Thị Trường")
        col_t1, col_t2 = st.columns([1, 1.3])
        with col_t1:
            st.markdown(f"""
            - **Tổng số mã cổ phiếu:** `{clean_data['Ticker'].nunique()}` mã
            - **Tổng số phiên giao dịch:** `{clean_data['Date'].nunique()}` phiên
            - **Khoảng thời gian:** `{clean_data['Date'].min().strftime('%d/%m/%Y')}` ➔ `{clean_data['Date'].max().strftime('%d/%m/%Y')}`
            - **Loại giá sử dụng:** `{source_type}` (ưu tiên `Adjusted` nếu có đủ OHLC)
            - **Chỉ số tham chiếu (Benchmark):** `{benchmark_ticker}`
            """)
            st.markdown("#### Bảng Phân Chia Thời Gian (No-Leak Time Split)")
            st.dataframe(split_df, use_container_width=True)

        with col_t2:
            st.markdown("#### Bảng Kiểm Tra Chất Lượng Dữ Liệu (Data Quality)")
            st.dataframe(q_table.head(15), use_container_width=True)

    # ===== TAB 2: NGƯỠNG TA CORE TRAIN =====
    with tabs[1]:
        st.subheader("2. Tối Ưu Hóa Ngưỡng BUY / EXIT Chung Trên Core Train")
        st.info(f"💡 Cặp ngưỡng kỹ thuật đã khóa: **BUY = {best_buy}** | **EXIT = {best_exit}**. Toàn bộ quá trình chọn ngưỡng chỉ thực hiện trên tập Core Train bằng 3-Fold Cross Validation.")
        
        col_cv1, col_cv2 = st.columns([1.2, 1])
        with col_cv1:
            st.markdown("#### Bảng Thử Nghiệm Grid Search")
            st.dataframe(res["threshold_trials"], use_container_width=True)
        with col_cv2:
            fig_cv = px.bar(
                res["threshold_trials"],
                x=[f"B:{r.BuyThreshold} | E:{r.ExitThreshold}" for _, r in res["threshold_trials"].iterrows()],
                y="GlobalCVScore",
                color="GlobalCVScore",
                title="Điểm Ổn Định Cross-Validation Theo Cặp Ngưỡng",
                labels={"x": "Cặp Ngưỡng (Buy / Exit)", "GlobalCVScore": "Điểm CV Chung"},
                color_continuous_scale="Blues"
            )
            fig_cv.update_layout(xaxis_tickangle=-45, height=380)
            st.plotly_chart(fig_cv, use_container_width=True)

    # ===== TAB 3: TOP N WALK-FORWARD =====
    with tabs[2]:
        st.subheader("3. Lịch Trình Tái Lựa Chọn Cổ Phiếu Theo Quý (Nested Walk-Forward)")
        st.markdown("""
        > **Nguyên tắc nhân quả:** Tại mỗi `EffectiveDate`, danh mục Top N chỉ được xếp hạng dựa trên `AsOfDate` (phiên liền trước $t-1$).
        > Thông tin giá và biến động trong tương lai hoàn toàn không được biết trước.
        """)

        tab_sub1, tab_sub2 = st.tabs(["Lịch Trình Final Test", "Lịch Trình Validation"])
        with tab_sub1:
            st.dataframe(test_selection, use_container_width=True)
            sel_date = st.selectbox(
                "Xem chi tiết xếp hạng đa yếu tố tại kỳ:",
                options=test_selection["EffectiveDate"].tolist(),
                format_func=lambda d: pd.Timestamp(d).strftime("%Y-%m-%d"),
                key="sb_test_rank"
            )
            if sel_date in res["test_selection_details"]:
                st.markdown(f"**Danh sách Top 5 được chọn tại kỳ {pd.Timestamp(sel_date).strftime('%Y-%m-%d')}:**")
                st.dataframe(res["test_selection_details"][sel_date]["locked"][[
                    "Ticker", "SelectedOrder", "FinalScore", "TrailingSharpe", "RelativeStrength126", "Vol63", "TA_Score"
                ]], use_container_width=True)
        with tab_sub2:
            st.dataframe(res["val_selection"], use_container_width=True)

    # ===== TAB 4: KHÓA CHAMPION VALIDATION =====
    with tabs[3]:
        st.subheader("4. Đánh Giá Các Họ Mô Hình Trên Tập Validation")
        st.success(f"🏆 Họ mô hình phân bổ danh mục chiến thắng: **{champion}** (được khóa cho Final Test).")
        st.markdown("""
        **Công thức tính điểm Validation:**
        $$\\text{ValidationScore} = 0.45 \\times \\text{Sharpe} + 0.15 \\times \\text{Sortino} + 0.012 \\times (\\text{Return} - \\text{BenchReturn}) + 0.004 \\times \\text{MaxDD} - 0.0005 \\times \\text{Turnover}$$
        """)

        st.dataframe(val_results.sort_values("ValidationScore", ascending=False), use_container_width=True)

        # Biểu đồ NAV trên Validation
        fig_val = go.Figure()
        for m_name, curve in res["val_curves"].items():
            if len(curve) > 0:
                norm_nav = curve / curve.iloc[0]
                fig_val.add_trace(go.Scatter(x=norm_nav.index, y=norm_nav.values, mode='lines', name=m_name))
        fig_val.update_layout(
            title="Đường Tăng Trưởng Tài Sản Trên Tập Validation (Normalized NAV)",
            xaxis_title="Thời Gian", yaxis_title="NAV Chuẩn Hóa",
            hovermode="x unified", height=450
        )
        st.plotly_chart(fig_val, use_container_width=True)

    # ===== TAB 5: KIỂM ĐỊNH OOS FINAL TEST =====
    with tabs[4]:
        st.subheader("5. Kết Quả Kiểm Định Ngoài Mẫu (Out-of-Sample Final Test)")
        st.markdown("So sánh **Champion Đã Khóa** với **Benchmark (VN-Index)** và **Static Buy & Hold Top N**:")

        cols_display = [
            "Return [%]", "CAGR [%]", "Sharpe", "Sortino", "MaxDD [%]", "Calmar",
            "Turnover [%]", "AvgExposure [%]", "AvgCash [%]",
            "ExcessCAGR_vs_Benchmark [pp]", "DD_Improvement_vs_Benchmark [pp]"
        ]
        st.dataframe(test_results[[c for c in cols_display if c in test_results.columns]], use_container_width=True)

        # Biểu đồ tương tác Plotly cho Final Test
        fig_test = go.Figure()
        colors = {
            "LOCKED_CHAMPION": "#1E3A8A",
            "VNINDEX_or_MarketBenchmark": "#DC2626",
            "STATIC_EW_BH": "#059669",
            "DYN_EW_NO_REGIME": "#7C3AED",
        }
        for m_name in ["LOCKED_CHAMPION", "STATIC_EW_BH", "VNINDEX_or_MarketBenchmark", "DYN_EW_NO_REGIME"]:
            if m_name in test_curves and len(test_curves[m_name]) > 0:
                c = test_curves[m_name]
                norm_nav = c / c.iloc[0]
                fig_test.add_trace(go.Scatter(
                    x=norm_nav.index, y=norm_nav.values, mode='lines',
                    name=m_name, line=dict(color=colors.get(m_name, None), width=2.5 if m_name == "LOCKED_CHAMPION" else 1.8)
                ))

        fig_test.update_layout(
            title="So Sánh Tăng Trưởng Tài Sản Final Test (Normalized NAV)",
            xaxis_title="Ngày", yaxis_title="Tài Sản Chuẩn Hóa (Bắt Đầu = 1.0)",
            hovermode="x unified", height=480,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_test, use_container_width=True)

        # Biểu đồ Drawdown
        fig_dd = go.Figure()
        for m_name in ["LOCKED_CHAMPION", "VNINDEX_or_MarketBenchmark"]:
            if m_name in test_curves and len(test_curves[m_name]) > 0:
                c = test_curves[m_name]
                dd = (c / c.cummax() - 1) * 100
                fig_dd.add_trace(go.Scatter(
                    x=dd.index, y=dd.values, mode='lines',
                    name=m_name, fill='tozeroy' if m_name == "LOCKED_CHAMPION" else None,
                    line=dict(width=1.5)
                ))
        fig_dd.update_layout(
            title="Đồ Thị Sụt Giảm Giá Trị (Underwater Drawdown %)",
            xaxis_title="Ngày", yaxis_title="Drawdown (%)",
            height=320, hovermode="x unified"
        )
        st.plotly_chart(fig_dd, use_container_width=True)

        # Lịch sử lệnh giao dịch
        st.subheader("Nhật Ký Lệnh Giao Dịch (Order Log)")
        if "LOCKED_CHAMPION" in test_orders and not test_orders["LOCKED_CHAMPION"].empty:
            st.dataframe(test_orders["LOCKED_CHAMPION"], use_container_width=True)
        else:
            st.info("Không có lệnh giao dịch phát sinh.")

    # ===== TAB 6: KIỂM TOÁN RÒ RỈ TƯƠNG LAI =====
    with tabs[5]:
        st.subheader("6. Báo Cáo Kiểm Toán Chống Rò Rỉ Thông Tin Tương Lai (Leakage Audit)")
        st.markdown("""
        Bảng kiểm toán chứng minh tính toàn vẹn của mô hình qua 7 chốt kiểm soát nhân quả:
        """)
        st.dataframe(leakage_df, use_container_width=True)

        st.markdown("""
        ### Diễn giải các nguyên tắc cốt lõi:
        1. **Không tối ưu tham số trên Test:** Ngưỡng $(BUY, EXIT)$ được chọn hoàn toàn tại Core Train; Model Champion được chọn tại Validation. Tập Final Test chỉ được 'chạy một lần' để chấm điểm.
        2. **Cơ chế khớp lệnh Open(t):** Tín hiệu tính toán tại thời điểm đóng cửa phiên $t-1$ chỉ được phép khớp vào giá mở cửa phiên $t$. Điều này phản ánh chính xác thực tế nhà đầu tư đặt lệnh trước giờ mở cửa hôm sau.
        3. **Co Covariance & Mu Shrinkage:** Tối ưu hóa Markowitz (MPT) và MinVar sử dụng phương pháp co rút ma trận Ledoit-Wolf và co rút trung bình tỷ suất sinh lời để tránh sai số ước lượng (estimation error).
        4. **Cơ chế tiền mặt (Cash Buffer):** Khi thị trường rơi vào xu hướng giảm (Risk-off) hoặc cổ phiếu kích hoạt lệnh thoát (Exit), vốn không bị ép phân bổ lại sang cổ phiếu khác mà chuyển thẳng vào tiền mặt để bảo toàn vốn.
        """)

    # ===== TAB 7: XUẤT BÁO CÁO =====
    with tabs[6]:
        st.subheader("7. Xuất Kết Quả Kiểm Định (Excel & CSV)")
        st.markdown("Bạn có thể tải toàn bộ kết quả phân tích về máy tính dưới dạng file Excel tổng hợp đa sheet.")

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            q_table.to_excel(writer, sheet_name="DataQuality", index=False)
            split_df.to_excel(writer, sheet_name="TimeSplit", index=False)
            res["threshold_trials"].to_excel(writer, sheet_name="CoreTAThresholds", index=False)
            res["val_selection"].to_excel(writer, sheet_name="ValSelectionHistory", index=False)
            val_results.reset_index().to_excel(writer, sheet_name="ValidationModels", index=False)
            test_selection.to_excel(writer, sheet_name="TestSelectionHistory", index=False)
            test_results.reset_index().to_excel(writer, sheet_name="FinalTest", index=False)
            leakage_df.to_excel(writer, sheet_name="LeakageAudit", index=False)
        excel_buffer.seek(0)

        st.download_button(
            label="📥 Tải Báo Cáo Excel Tổng Hợp (HOSE_V4_RESULTS.xlsx)",
            data=excel_buffer,
            file_name="HOSE_V4_RESULTS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
