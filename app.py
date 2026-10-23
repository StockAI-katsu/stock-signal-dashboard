from datetime import datetime, timedelta
import os
import pandas as pd
import streamlit as st
import yfinance as yf
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ページの設定
st.set_page_config(
    page_title="トレード手法 検証・マスター管理ダッシュボード",
    layout="wide",
)

# --- Googleスプレッドシート接続設定 ---
SPREADSHEET_NAME = "トレード手法検証マスター"  # ご自身のスプレッドシート名
EXCEL_MASTER_FILE = "株式市場・テーマ監視ボード.xlsx"

@st.cache_resource
def get_gspread_client():
    """Googleスプレッドシートのクライアント認証を行う"""
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    
    # 1. ローカルの credentials.json を優先して探す
    json_path = "credentials.json"
    if os.path.exists(json_path):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"credentials.json の読み込みに失敗しました: {e}")
            return None

    # 2. Streamlit Cloud等の secrets を探す（ローカルにファイルがない場合のフォールバック）
    try:
        if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
    except Exception:
        pass

    return None

@st.cache_data
def load_excel_stock_master():
    """Excelの『銘柄データ』シートからコード、銘柄名、市場、セクターの辞書を作成する"""
    name_dict = {}
    market_dict = {}
    sector_dict = {}

    # ご自身のスプレッドシートのURLに書き換えてください
    sheet_url = "https://docs.google.com/spreadsheets/d/11qtf0L__69eW3b_jNfXMz6L5uKcZWRZnUtKeVFuirrE/export?format=csv&gid=0"

    try:
        # Excelの読み込みから、スプレッドシートのCSV読み込みに変更
        df_ex = pd.read_csv(sheet_url, dtype={"コード": str})
        
        for _, row in df_ex.iterrows():
            code_raw = str(row.get("コード", "")).replace(".T", "").strip()
            if not code_raw or code_raw.lower() == "nan" or not code_raw.isdigit():
                continue
            code = code_raw.zfill(4)

            name = str(row.get("銘柄名", "")).strip()
            market_raw = str(row.get("市場・商品区分", ""))
            sector_raw = str(
                row.get(
                    "33業種区分",
                    row.get(
                        "セクター", row.get("33業種", row.get("業種", "未分類"))
                    ),
                )
            )

            name_dict[code] = name

            if "プライム" in market_raw:
                market_dict[code] = "プライム"
            elif "スタンダード" in market_raw:
                market_dict[code] = "スタンダード"
            elif "グロース" in market_raw:
                market_dict[code] = "グロース"
            else:
                market_dict[code] = (
                    market_raw if market_raw != "nan" else "プライム"
                )

            sector_dict[code] = (
                sector_raw if sector_raw != "nan" and sector_raw != "" else "未分類"
            )
    except Exception as e:
        print(f"スプレッドシート読み込みエラー: {e}")
        
    return name_dict, market_dict, sector_dict


EXCEL_NAME_DICT, EXCEL_MARKET_DICT, EXCEL_SECTOR_DICT = (
    load_excel_stock_master()
)

# 20営業日分の列名を定義
DAY_COLUMNS = []
for i in range(1, 21):
    DAY_COLUMNS.extend([f"{i}日目_始値", f"{i}日目_終値", f"{i}日目_騰落率(%)"])

BASE_COLUMNS = [
    "発生日",
    "シグナル内容",
    "コード",
    "銘柄名",
    "市場",
    "セクター",
    "発生日終値",
    "発生日出来高",
    "手法",
]
ALL_COLUMNS = ["アラートID", "ティッカー"] + BASE_COLUMNS + DAY_COLUMNS


def load_master():
    """Googleスプレッドシートからマスターデータを読み込む"""
    df = pd.DataFrame(columns=ALL_COLUMNS)
    client = get_gspread_client()
    if client is None:
        st.error("Googleスプレッドシートの認証情報（credentials.json）が見つかりません。")
        return df

    try:
        sheet = client.open(SPREADSHEET_NAME).worksheet("master_signals")
        data = sheet.get_all_records()
        if data:
            df = pd.DataFrame(data)
            if "コード" in df.columns:
                df["コード"] = df["コード"].astype(str).str.replace(".0", "", regex=False).str.zfill(4)
            if "アラートID" in df.columns:
                df["アラートID"] = df["アラートID"].astype(str)
        else:
            df = pd.DataFrame(columns=ALL_COLUMNS)
    except Exception as e:
        try:
            spreadsheet = client.open(SPREADSHEET_NAME)
            sheet = spreadsheet.add_worksheet(title="master_signals", rows="1000", cols="100")
            sheet.append_row(ALL_COLUMNS)
            df = pd.DataFrame(columns=ALL_COLUMNS)
        except Exception as ex:
            st.error(f"スプレッドシートの読み込み/初期化に失敗しました: {ex}")

    for col in ALL_COLUMNS:
        if col not in df.columns:
            df[col] = None

    numeric_cols = ["発生日終値", "発生日出来高"] + DAY_COLUMNS
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def save_master(df):
    """Googleスプレッドシートへデータを保存・上書きする"""
    client = get_gspread_client()
    if client is None:
        st.error("Googleスプレッドシートの認証情報が見つかりません。")
        return

    try:
        sheet = client.open(SPREADSHEET_NAME).worksheet("master_signals")
        sheet.clear()
        df_to_save = df.fillna("")
        sheet.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
    except Exception as e:
        st.error(f"Googleスプレッドシートへの保存に失敗しました: {e}")


def parse_tradingview_csv(uploaded_files):
    """トレーディングビューCSVをパースし、シグナル抽出と重複を完全解消する"""
    all_new_rows = []
    if not isinstance(uploaded_files, list):
        uploaded_files = [uploaded_files]

    for uploaded_file in uploaded_files:
        try:
            df_raw = pd.read_csv(uploaded_file)
            for _, row in df_raw.iterrows():
                ticker = str(row.get("ティッカー", row.get("Ticker", "")))
                if not ticker or ticker == "nan":
                    ticker = "TSE:9984, 1日"

                time_str = str(
                    row.get("日時", row.get("Time", row.get("発生日", datetime.now())))
                )
                date_part = str(time_str).split("T")[0].split(" ")[0]

                name_val = str(row.get("名前", ""))
                desc_val = str(row.get("名称", row.get("Message", row.get("コメント", ""))))
                
                full_text = f"{name_val} {desc_val}".strip()
                full_text_lower = full_text.lower()

                # --- ことの手法の判定 ---
                if "kotonov" in full_text_lower or "こと" in full_text_lower or "5kai7" in full_text_lower:
                    method = "ことの手法"
                    if "kotonov5kai7" in full_text_lower or "5kai7" in full_text_lower:
                        sig_name = "kotonov5kai7"
                    else:
                        sig_name = desc_val if desc_val and desc_val != "nan" else "ことの手法シグナル"

                # --- 押し目プロ風手法の判定 ---
                else:
                    method = "押し目プロ風手法"
                    if "昇格gc" in full_text_lower or "上昇gc" in full_text_lower:
                        sig_name = "昇格GC（出来高あり）" if "出来高" in full_text_lower else "昇格GC"
                    elif "青gc" in full_text_lower:
                        sig_name = "青GC（出来高あり）" if "出来高" in full_text_lower else "青GC"
                    elif "売られすぎ" in full_text_lower:
                        sig_name = "売られすぎ（出来高あり）" if "出来高" in full_text_lower else "売られすぎ"
                    else:
                        clean_sig = desc_val.replace("{{ticker}}", "").replace("{{", "").replace("}}", "").strip()
                        sig_name = clean_sig if clean_sig and clean_sig != "nan" else (name_val if name_val != "nan" else "押し目プロ風シグナル")

                code_clean = ticker.split(":")[-1].split(",")[0].strip()
                numeric_code = code_clean.replace(".T", "").zfill(4)

                stock_name = EXCEL_NAME_DICT.get(numeric_code, f"銘柄_{numeric_code}")
                market_val = EXCEL_MARKET_DICT.get(numeric_code, "プライム")
                sector_val = EXCEL_SECTOR_DICT.get(numeric_code, "未分類")

                alert_id = abs(hash(f"{numeric_code}_{date_part}_{method}_{sig_name}")) % (10**10)

                all_new_rows.append({
                    "アラートID": alert_id,
                    "ティッカー": ticker,
                    "コード": numeric_code,
                    "手法": method,
                    "シグナル内容": sig_name,
                    "発生日": date_part,
                    "銘柄名": stock_name,
                    "市場": market_val,
                    "セクター": sector_val,
                    "発生日終値": 0.0,
                    "発生日出来高": 0,
                })
        except Exception as e:
            st.error(f"ファイル {uploaded_file.name} の読み込みに失敗しました: {e}")

    df_parsed = pd.DataFrame(all_new_rows)
    if not df_parsed.empty:
        df_parsed = df_parsed.drop_duplicates(subset=["コード", "発生日", "手法", "シグナル内容"], keep="first").reset_index(drop=True)

    return df_parsed


def fetch_stock_data_for_master(df, target_indices=None):
    """株価データ（終値・出来高・20日分）の取得と同期を行う"""
    if df.empty:
        return df, 0, 0

    numeric_cols = ["発生日終値", "発生日出来高"] + DAY_COLUMNS
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].astype(float)

    progress_bar = st.progress(0)
    indices_to_process = (
        target_indices if target_indices is not None else df.index
    )
    total = len(indices_to_process)

    success_count = 0
    fail_count = 0

    for count, idx in enumerate(indices_to_process):
        progress_bar.progress((count + 1) / max(total, 1))
        row = df.loc[idx]
        raw_code = str(row["コード"]).replace(".T", "").strip().zfill(4)
        yf_code = f"{raw_code}.T"
        start_date_str = str(row["発生日"]).split("T")[0].split(" ")[0]

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        except:
            fail_count += 1
            continue

        stock_name = EXCEL_NAME_DICT.get(raw_code, row["銘柄名"])
        market_val = EXCEL_MARKET_DICT.get(raw_code, row["市場"])
        sector_val = EXCEL_SECTOR_DICT.get(
            raw_code, row.get("セクター", "未分類")
        )

        df.at[idx, "コード"] = raw_code
        df.at[idx, "市場"] = market_val
        df.at[idx, "セクター"] = sector_val
        df.at[idx, "銘柄名"] = stock_name

        try:
            ticker_obj = yf.Ticker(yf_code)
            start_uni = start_date - timedelta(days=12)
            end_uni = datetime.now() + timedelta(days=2)

            hist = ticker_obj.history(start=start_uni, end=end_uni, auto_adjust=True)
            if hist is None or hist.empty:
                hist = ticker_obj.history(period="max", auto_adjust=True)

            if hist is None or hist.empty:
                fail_count += 1
                continue

            if isinstance(hist.columns, pd.MultiIndex):
                hist.columns = hist.columns.get_level_values(0)

            if hist.index.tz is not None:
                hist.index = hist.index.tz_localize(None)

            future_days = hist[hist.index >= start_date]
            if len(future_days) == 0:
                future_days = hist[hist.index >= (start_date - timedelta(days=6))]
                if len(future_days) == 0:
                    fail_count += 1
                    continue

            base_row = future_days.iloc[0]
            base_close = base_row["Close"]
            base_vol = base_row["Volume"]
            if isinstance(base_close, pd.Series):
                base_close = base_close.iloc[0]
            if isinstance(base_vol, pd.Series):
                base_vol = base_vol.iloc[0]

            if pd.isna(base_close) or float(base_close) == 0:
                fail_count += 1
                continue

            df.at[idx, "発生日終値"] = float(base_close)
            df.at[idx, "発生日出来高"] = float(base_vol) if pd.notna(base_vol) else 0.0

            for d in range(1, 21):
                if d < len(future_days):
                    day_row = future_days.iloc[d]
                    open_val = float(day_row["Open"])
                    close_val = float(day_row["Close"])
                    ret_val = (
                        (close_val - float(base_close)) / float(base_close)
                    ) * 100

                    df.at[idx, f"{d}日目_始値"] = open_val
                    df.at[idx, f"{d}日目_終値"] = close_val
                    df.at[idx, f"{d}日目_騰落率(%)"] = round(ret_val, 2)

            success_count += 1
        except Exception as e:
            print(f"Error fetching stock hist for {yf_code}: {e}")
            fail_count += 1
            continue

    progress_bar.empty()
    return df, success_count, fail_count


def calculate_optimal_strategies_for_df(dataframe, method_name):
    """最適なイン・アウトの組み合わせを自動算出する"""
    strategies = []
    if dataframe.empty or "1日目_始値" not in dataframe.columns:
        return pd.DataFrame()

    for in_day in range(1, 6):
        for out_day in range(in_day + 1, 21):
            returns = []
            for _, row in dataframe.iterrows():
                try:
                    buy_val = row[f"{in_day}日目_始値"]
                    sell_val = row[f"{out_day}日目_終値"]
                    if pd.notna(buy_val) and pd.notna(sell_val) and buy_val > 0:
                        ret = (sell_val - buy_val) / buy_val
                        returns.append(ret)
                except:
                    continue

            if len(returns) >= 3:
                win_count = sum(1 for r in returns if r > 0)
                win_rate = (win_count / len(returns)) * 100
                avg_ret = sum(returns) / len(returns) * 100

                strategies.append({
                    "対象手法": method_name,
                    "イン": f"{in_day}営業日後（始値）",
                    "アウト": f"{out_day}営業日後（終値）",
                    "勝率": round(win_rate, 1),
                    "平均リターン": round(avg_ret, 2),
                    "サンプル数": len(returns),
                })

    strat_df = pd.DataFrame(strategies)
    if not strat_df.empty:
        strat_df = strat_df.sort_values(by=["勝率", "平均リターン"], ascending=False)
        return strat_df.head(3)
    return pd.DataFrame()


def get_filtered_trade_candidates(
    dataframe, max_business_days=3, selected_market="すべて", selected_signal="すべて", vol_range=None
):
    """営業日を考慮して候補銘柄を抽出し、価格反映時の終値とその日付（または経過日数）を構築する"""
    if dataframe.empty:
        return pd.DataFrame()

    candidates = []
    today = pd.Timestamp(datetime.now().date())

    for _, row in dataframe.iterrows():
        発生日_str = row.get("発生日")
        if pd.isna(発生日_str):
            continue

        try:
            start_dt = pd.to_datetime(発生日_str)
            if start_dt > today:
                continue
            
            biz_days = pd.bdate_range(start=start_dt, end=today)
            business_days_diff = len(biz_days) - 1
            if business_days_diff < 0:
                business_days_diff = 0

            if business_days_diff <= max_business_days:
                market_val = str(row.get("市場", "プライム"))
                sig_val = str(row.get("シグナル内容", ""))
                vol_val = float(row.get("発生日出来高", 0))

                market_match = (selected_market == "すべて" or selected_market in market_val)
                signal_match = (selected_signal == "すべて" or selected_signal == sig_val)
                
                vol_match = True
                if vol_range is not None:
                    min_v, max_v = vol_range
                    if min_v == 0 and max_v == 0:
                        vol_match = True
                    else:
                        vol_match = (min_v <= vol_val <= max_v)

                if market_match and signal_match and vol_match:
                    current_close = row.get("発生日終値", 0.0)
                    price_date_label = f"発生日 ({str(start_dt.date())})"
                    
                    target_day_idx = 0
                    for d in range(20, 0, -1):
                        col_n = f"{d}日目_終値"
                        if col_n in row and pd.notna(row[col_n]) and float(row[col_n]) > 0:
                            current_close = float(row[col_n])
                            target_day_idx = d
                            break
                    
                    if target_day_idx > 0 and len(biz_days) > target_day_idx:
                        actual_date = biz_days[target_day_idx].strftime("%Y-%m-%d")
                        price_date_label = f"{target_day_idx}営業日後 ({actual_date})"
                    elif business_days_diff == 0:
                        price_date_label = f"当日 ({str(start_dt.date())})"

                    candidates.append({
                        "発生日": row["発生日"],
                        "手法": row.get("手法", "不明"),
                        "シグナル内容": sig_val,
                        "コード": row["コード"],
                        "銘柄名": row["銘柄名"],
                        "市場": market_val,
                        "セクター": row.get("セクター", "未分類"),
                        "発生日出来高": int(vol_val),
                        "発生日終値": row.get("発生日終値", 0.0),
                        "価格反映時終値": current_close,
                        "価格反映日": price_date_label,
                    })
        except:
            continue

    return pd.DataFrame(candidates)


# --- UI メイン ---
st.title("📊 トレード手法 検証・マスター管理ダッシュボード")
master_df = load_master()

# --- サイドバー: データ管理 ---
st.sidebar.header("⚙️ データ管理・操作 (Googleスプレッドシート連携)")

uploaded_files = st.sidebar.file_uploader(
    "トレーディングビューCSVをアップロード（複数選択可）",
    type=["csv"],
    accept_multiple_files=True,
)

if uploaded_files:
    if st.sidebar.button("🔄 マスター台帳に統合・更新する"):
        new_df = parse_tradingview_csv(uploaded_files)
        if not new_df.empty:
            combined = pd.concat([master_df, new_df]).drop_duplicates(
                subset=["コード", "発生日", "手法", "シグナル内容"], keep="first"
            )
            save_master(combined)
            master_df = combined
            st.sidebar.success(f"スプレッドシートのマスター台帳を更新しました！（総件数: {len(master_df)}件）")
            st.rerun()

if not master_df.empty:
    st.sidebar.markdown("### 📥 株価データ取得・更新")
    if st.sidebar.button("🚀 全件一括自動取得・更新"):
        with st.spinner("株価データを同期中..."):
            updated_df, succ, fail = fetch_stock_data_for_master(master_df)
            save_master(updated_df)
            master_df = updated_df
            st.sidebar.success(f"一括同期完了！ (成功: {succ}件)")
            st.rerun()

    if st.sidebar.button("🗑️ マスター台帳をリセット（初期化）"):
        client = get_gspread_client()
        if client:
            try:
                sheet = client.open(SPREADSHEET_NAME).worksheet("master_signals")
                sheet.clear()
                sheet.append_row(ALL_COLUMNS)
                st.sidebar.warning("スプレッドシートのマスター台帳をリセットしました。")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"リセットに失敗しました: {e}")

st.sidebar.markdown("---")
st.sidebar.header("🔍 絞り込み条件（経過日数）")
filter_max_days = st.sidebar.slider("発生からの経過営業日数（以内）", min_value=0, max_value=30, value=3)

st.sidebar.markdown("""
<small style="color: #666; line-height: 1.4;">
💡 <b>確認タイミングの目安:</b><br>
・<b>0日</b>: 引け後の確認（当日の引けで発生）<br>
・<b>1日</b>: 翌朝・寄付前の確認（前日の引けで発生）
</small>
""", unsafe_allow_html=True)

# --- 手法リストの取得 ---
methods = []
if not master_df.empty and "手法" in master_df.columns:
    methods = [str(m) for m in master_df["手法"].dropna().unique() if str(m).strip() != ""]

all_tab_labels = ["総合（全手法合算）"] + methods if methods else ["総合（全手法合算）"]
selected_tabs = st.tabs(all_tab_labels)


for idx, tab_label in enumerate(all_tab_labels):
    with selected_tabs[idx]:
        if idx == 0:
            st.subheader("🌐 総合（全手法合算）の分析結果 - 明日のトレード候補一覧")
            st.markdown("主要な2つの手法（押し目プロ風手法・ことの手法）ごとの明日のトレード候補をそれぞれ独立したフィルターで絞り込んで確認できます。")

            target_methods = ["押し目プロ風手法", "ことの手法"]
            
            for m_name in target_methods:
                st.markdown("---")
                st.markdown(f"### 📌 【 {m_name} 】 の明日のトレード候補")
                
                m_df = master_df[master_df["手法"] == m_name] if not master_df.empty else pd.DataFrame()
                
                if not m_df.empty:
                    c_col1, c_col2 = st.columns(2)
                    with c_col1:
                        c_markets = ["すべて"] + list(m_df["市場"].dropna().unique())
                        sel_c_market = st.selectbox(f"対象市場（{m_name}）", c_markets, key=f"comp_mkt_{m_name}")
                    with c_col2:
                        c_signals = ["すべて"] + list(m_df["シグナル内容"].dropna().unique())
                        sel_c_signal = st.selectbox(f"発生シグナル（{m_name}）", c_signals, key=f"comp_sig_{m_name}")

                    df_c_vol = m_df.copy()
                    if sel_c_market != "すべて":
                        df_c_vol = df_c_vol[df_c_vol["市場"] == sel_c_market]
                    if sel_c_signal != "すべて":
                        df_c_vol = df_c_vol[df_c_vol["シグナル内容"] == sel_c_signal]

                    valid_c_vols = pd.to_numeric(df_c_vol["発生日出来高"], errors="coerce").dropna()
                    c_min_vol = int(valid_c_vols.min()) if not valid_c_vols.empty and valid_c_vols.min() > 0 else 0
                    c_max_vol = int(valid_c_vols.max()) if not valid_c_vols.empty and valid_c_vols.max() > 0 else 10000000

                    st.markdown(f"📊 **出来高範囲: {c_min_vol:,} 株 〜 {c_max_vol:,} 株** （選択条件に合わせて自動変動）")

                    c_min_key = f"comp_vol_min_{m_name}"
                    c_max_key = f"comp_vol_max_{m_name}"

                    if c_min_key not in st.session_state or st.session_state.get(f"comp_prev_mkt_{m_name}") != sel_c_market or st.session_state.get(f"comp_prev_sig_{m_name}") != sel_c_signal:
                        st.session_state[c_min_key] = c_min_vol
                        st.session_state[c_max_key] = c_max_vol
                        st.session_state[f"comp_prev_mkt_{m_name}"] = sel_c_market
                        st.session_state[f"comp_prev_sig_{m_name}"] = sel_c_signal

                    c_col_v1, c_col_v2 = st.columns(2)
                    with c_col_v1:
                        c_vol_min = st.number_input("出来高 下限 (株)", min_value=0, max_value=max(c_max_vol, c_min_vol, 1), step=1000, key=c_min_key)
                    with c_col_v2:
                        c_vol_max = st.number_input("出来高 上限 (株)", min_value=0, max_value=max(c_max_vol * 2, 10000000, 1), step=1000, key=c_max_key)

                    m_cand = get_filtered_trade_candidates(
                        m_df, max_business_days=filter_max_days, selected_market=sel_c_market, selected_signal=sel_c_signal, vol_range=(c_vol_min, c_vol_max)
                    )

                    if not m_cand.empty:
                        st.success(f"条件に合致する候補銘柄: **{len(m_cand)}件**")
                        st.dataframe(m_cand, width="stretch")
                        
                        # CSVダウンロードボタン（総合タブ内）
                        csv_data_m = m_cand.to_csv(index=False).encode('utf-8-sig')
                        st.download_button(
                            label=f"📥 【{m_name}】候補銘柄をCSVダウンロード",
                            data=csv_data_m,
                            file_name=f"trade_candidates_{m_name}.csv",
                            mime="text/csv",
                            key=f"dl_comp_{m_name}"
                        )
                    else:
                        st.info(f"💡 現在、指定した条件に合致する「{m_name}」のエントリー候補銘柄はありません。")
                else:
                    st.info(f"💡 「{m_name}」のデータが登録されていません。")

        else:
            current_method = methods[idx - 1]
            target_df = master_df[master_df["手法"] == current_method]
            st.subheader(f"🏷️ 手法：{current_method} の分析結果")

            st.markdown("### 🔍 明日のトレード候補 絞り込みフィルター")
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                markets = ["すべて"] + list(target_df["市場"].dropna().unique())
                sel_market = st.selectbox("対象市場", markets, key=f"market_{current_method}")
            with col_f2:
                signals = ["すべて"] + list(target_df["シグナル内容"].dropna().unique())
                sel_signal = st.selectbox("発生シグナル", signals, key=f"signal_{current_method}")

            df_for_vol = target_df.copy()
            if sel_market != "すべて":
                df_for_vol = df_for_vol[df_for_vol["市場"] == sel_market]
            if sel_signal != "すべて":
                df_for_vol = df_for_vol[df_for_vol["シグナル内容"] == sel_signal]

            valid_vols = pd.to_numeric(df_for_vol["発生日出来高"], errors="coerce").dropna()
            min_possible_vol = int(valid_vols.min()) if not valid_vols.empty and valid_vols.min() > 0 else 0
            max_possible_vol = int(valid_vols.max()) if not valid_vols.empty and valid_vols.max() > 0 else 10000000

            st.markdown(f"📊 **出来高範囲: {min_possible_vol:,} 株 〜 {max_possible_vol:,} 株** （選択条件に合わせて自動変動）")
            
            min_key = f"vol_min_{current_method}"
            max_key = f"vol_max_{current_method}"
            
            if min_key not in st.session_state or st.session_state.get(f"prev_market_{current_method}") != sel_market or st.session_state.get(f"prev_signal_{current_method}") != sel_signal:
                st.session_state[min_key] = min_possible_vol
                st.session_state[max_key] = max_possible_vol
                st.session_state[f"prev_market_{current_method}"] = sel_market
                st.session_state[f"prev_signal_{current_method}"] = sel_signal

            col_v1, col_v2 = st.columns(2)
            with col_v1:
                vol_min = st.number_input(
                    "出来高 下限 (株)", 
                    min_value=0, 
                    max_value=max(max_possible_vol, min_possible_vol, 1), 
                    step=1000, 
                    key=min_key
                )
            with col_v2:
                vol_max = st.number_input(
                    "出来高 上限 (株)", 
                    min_value=0, 
                    max_value=max(max_possible_vol * 2, 10000000, 1), 
                    step=1000, 
                    key=max_key
                )

            candidates_df = get_filtered_trade_candidates(
                target_df, 
                max_business_days=filter_max_days,
                selected_market=sel_market, 
                selected_signal=sel_signal, 
                vol_range=(vol_min, vol_max)
            )

            st.markdown("### 🎯 明日のトレード候補（絞り込み抽出）")
            if not candidates_df.empty:
                st.success(f"条件に合致する候補銘柄が **{len(candidates_df)}件** 見つかりました！")
                st.dataframe(candidates_df, width="stretch")
                
                # CSVダウンロードボタン（トレード候補用）
                csv_cand = candidates_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 抽出されたトレード候補をCSVダウンロード",
                    data=csv_cand,
                    file_name=f"candidates_{current_method}.csv",
                    mime="text/csv",
                    key=f"dl_cand_{current_method}"
                )
            else:
                st.info("💡 現在、指定した条件に合致する翌日エントリー候補銘柄はありません。")

            st.markdown("---")

            st.markdown(f"### 🏆 手法「{current_method}」× セクター・市場別 最適戦略")

            st.markdown("#### 🏢 セクター別最適戦略")
            sectors = [
                s
                for s in target_df["セクター"].dropna().unique()
                if str(s).strip() != "" and str(s) != "nan"
            ]
            if sectors:
                for sector in sectors:
                    sector_df = target_df[target_df["セクター"] == sector]
                    if len(sector_df) >= 3:
                        st.markdown(
                            f"**・セクター: 【 {sector} 】** (検証数: {len(sector_df)}件)"
                        )
                        sec_top = calculate_optimal_strategies_for_df(
                            sector_df, method_name=f"{current_method} ({sector})"
                        )
                        if not sec_top.empty:
                            c1, c2, c3 = st.columns(3)
                            cols = [c1, c2, c3]
                            for s_idx, (_, strat) in enumerate(sec_top.iterrows()):
                                with cols[s_idx]:
                                    cols[s_idx].markdown(f"""
                                        **{s_idx+1}位**: {strat['イン']}〜{strat['アウト']} | 勝率:**{strat['勝率']}%** (平均:+{strat['平均リターン']}%)
                                        """)
            else:
                st.info("データが不足しているためセクター別最適戦略を表示できません。")

            st.markdown("")
            st.markdown("#### 🏛️ 市場区分別最適戦略")
            mkt_list = [
                mk
                for mk in ["プライム", "スタンダード", "グロース"]
                if mk in target_df["市場"].values
            ]
            if not mkt_list:
                mkt_list = [
                    mk
                    for mk in target_df["市場"].dropna().unique()
                    if str(mk).strip() != ""
                ]

            if mkt_list:
                for market in mkt_list:
                    market_df = target_df[target_df["市場"].str.contains(market, na=False)]
                    if len(market_df) >= 3:
                        st.markdown(
                            f"**・市場: 【 {market} 】** (検証数: {len(market_df)}件)"
                        )
                        mkt_top = calculate_optimal_strategies_for_df(
                            market_df, method_name=f"{current_method} ({market})"
                        )
                        if not mkt_top.empty:
                            c1, c2, c3 = st.columns(3)
                            cols = [c1, c2, c3]
                            for s_idx, (_, strat) in enumerate(mkt_top.iterrows()):
                                with cols[s_idx]:
                                    cols[s_idx].markdown(f"""
                                        **{s_idx+1}位**: {strat['イン']}〜{strat['アウト']} | 勝率:**{strat['勝率']}%** (平均:+{strat['平均リターン']}%)
                                        """)

            st.markdown("---")

            st.markdown("### 📋 シグナル一覧・トラッキング台帳（出来高・発生シグナル付き）")
            
            col4, col5 = st.columns(2)
            with col4:
                track_signal_filter = st.selectbox("発生シグナル（台帳用）", ["すべて"] + list(target_df["シグナル内容"].dropna().unique()), key=f"track_sig_{current_method}")
            with col5:
                track_market_filter = st.selectbox("対象市場（台帳用）", ["すべて"] + list(target_df["市場"].dropna().unique()), key=f"track_mkt_{current_method}")

            filtered_target_df = target_df.copy()
            if track_signal_filter != "すべて":
                filtered_target_df = filtered_target_df[filtered_target_df["シグナル内容"] == track_signal_filter]
            if track_market_filter != "すべて":
                filtered_target_df = filtered_target_df[filtered_target_df["市場"] == track_market_filter]

            if not filtered_target_df.empty:
                st.markdown(f"表示件数: **{len(filtered_target_df)}件**")
                display_cols = [
                    col
                    for col in filtered_target_df.columns
                    if col not in ["アラートID", "ティッカー"]
                ]
                st.dataframe(filtered_target_df[display_cols], width="stretch")
                
                # CSVダウンロードボタン（トラッキング台帳用）
                csv_track = filtered_target_df[display_cols].to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 トラッキング台帳（全データ）をCSVダウンロード",
                    data=csv_track,
                    file_name=f"tracking_master_{current_method}.csv",
                    mime="text/csv",
                    key=f"dl_track_{current_method}"
                )
            else:
                st.info("該当するデータはありません。")
