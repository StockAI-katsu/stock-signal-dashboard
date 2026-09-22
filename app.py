from datetime import datetime, timedelta
import hashlib
import os
import re
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection
import yfinance as yf

# ページの設定
st.set_page_config(
    page_title="トレード手法 検証・マスター管理ダッシュボード",
    layout="wide",
)

EXCEL_MASTER_FILE = "株式市場・テーマ監視ボード.xlsx"

@st.cache_data
def load_excel_stock_master():
    name_dict = {}
    market_dict = {}
    sector_dict = {}
    if os.path.exists(EXCEL_MASTER_FILE):
        try:
            df_ex = pd.read_excel(EXCEL_MASTER_FILE, sheet_name="銘柄データ", dtype={"コード": str})
            for _, row in df_ex.iterrows():
                raw_code = str(row["コード"])
                code = raw_code.replace(".T", "").replace("TSE:", "").strip().zfill(4)
                name = str(row["銘柄名"]).strip()
                market_raw = str(row.get("市場・商品区分", ""))
                sector_raw = str(row.get("33業種区分", row.get("17業種区分", "")))
                
                name_dict[code] = name
                if "プライム" in market_raw:
                    market_dict[code] = "プライム"
                elif "スタンダード" in market_raw:
                    market_dict[code] = "スタンダード"
                elif "グロース" in market_raw:
                    market_dict[code] = "グロース"
                else:
                    market_dict[code] = market_raw if market_raw != "nan" else "プライム"
                
                if sector_raw and sector_raw != "nan" and sector_raw != "-":
                    sector_dict[code] = sector_raw
                else:
                    sector_dict[code] = "未分類"
        except Exception as e:
            print(f"Excel読み込みエラー: {e}")
    return name_dict, market_dict, sector_dict

EXCEL_NAME_DICT, EXCEL_MARKET_DICT, EXCEL_SECTOR_DICT = load_excel_stock_master()

DAY_COLUMNS = []
for i in range(1, 21):
    DAY_COLUMNS.extend([f"{i}日目_始値", f"{i}日目_終値", f"{i}日目_騰落率(%)"])

BASE_COLUMNS = ["発生日", "シグナル内容", "コード", "銘柄名", "市場", "セクター", "発生日終値", "手法"]
ALL_COLUMNS = ["アラートID", "ティッカー"] + BASE_COLUMNS + DAY_COLUMNS

conn = st.connection("gsheets", type=GSheetsConnection)

def load_master():
    try:
        df = conn.read(worksheet="master_signals", ttl=0)
        if df is None or df.empty:
            return pd.DataFrame(columns=ALL_COLUMNS)
        if "コード" in df.columns:
            df["コード"] = df["コード"].astype(str).str.replace(".T", "").str.replace("TSE:", "").str.strip().str.zfill(4)
        for col in ALL_COLUMNS:
            if col not in df.columns:
                df[col] = None
        return df
    except Exception as e:
        return pd.DataFrame(columns=ALL_COLUMNS)

def save_master(df):
    try:
        conn.update(worksheet="master_signals", data=df)
    except Exception as e:
        st.error(f"スプレッドシートへの保存に失敗しました: {e}")

def parse_tradingview_csv(uploaded_files):
    """トレーディングビューの変則CSVから確実にデータを抽出する強化版パーサー"""
    all_new_rows = []
    if not isinstance(uploaded_files, list):
        uploaded_files = [uploaded_files]

    for uploaded_file in uploaded_files:
        try:
            df_raw = pd.read_csv(uploaded_file)
            st.write("--- デバッグ: 読み込んだCSVの列名 ---", df_raw.columns.tolist())
            st.write("--- デバッグ: CSVの最初の数行 ---", df_raw.head(5))

            for _, row in df_raw.iterrows():
                # 行全体の文字列を結合して検索しやすくする
                row_values = [str(val) for val in row.values if pd.notna(val)]
                row_str = " ".join(row_values)
                
                # 1. 銘柄コード（4桁の数字、または TSE:XXXX）を強力に探索
                numeric_code = ""
                ticker_val = ""
                
                for val in row.values:
                    v_str = str(val).strip()
                    # "TSE:3549" や "3549" などのパターンをチェック
                    if ":" in v_str:
                        parts = v_str.split(":")
                        candidate = parts[-1].strip()
                        if candidate.isdigit() and len(candidate) == 4:
                            numeric_code = candidate
                            ticker_val = v_str
                            break
                    elif v_str.isdigit() and len(v_str) == 4:
                        numeric_code = v_str
                        ticker_val = f"TSE:{v_str}"
                        break
                
                # セル単位で見つからない場合は行全体の文字列から4桁数字を正規表現で探す
                if not numeric_code:
                    # 日本の株価コードによくある4桁の数字（1000〜9999）を抽出
                    matches = re.findall(r'\b([1-9][0-9]{3})\b', row_str)
                    for m in matches:
                        # アラートIDやパラメータの数字（例: 75, 20, 40など）を除外するため、1000以上の4桁に限定
                        if int(m) >= 1000:
                            numeric_code = m
                            ticker_val = f"TSE:{m}"
                            break

                if not numeric_code:
                    continue

                # 2. 日時の抽出（ISO形式や "2026-09-18" 形式を探す）
                date_part = datetime.now().strftime("%Y-%m-%d")
                for val in row.values:
                    v_str = str(val)
                    if "202" in v_str or "203" in v_str:
                        # 日付部分のパターンマッチ (YYYY-MM-DD)
                        date_match = re.search(r'(20[2-3][0-9]-[0-1][0-9]-[0-3][0-9])', v_str)
                        if date_match:
                            date_part = date_match.group(1)
                            break
                        elif "T" in v_str:
                            date_part = v_str.split("T")[0].split(" ")[0]
                            break

                # 3. 手法の判定
                combined_text = row_str.lower()
                if "kotonov" in combined_text or "こと" in combined_text:
                    method = "ことの手法"
                elif "スイング押し目" in combined_text or "押し目" in combined_text or "v15" in combined_text:
                    method = "押し目プロ風手法"
                else:
                    method = "押し目プロ風手法" # デフォルト

                stock_name = EXCEL_NAME_DICT.get(numeric_code, f"銘柄_{numeric_code}")
                market_val = EXCEL_MARKET_DICT.get(numeric_code, "プライム")
                sector_val = EXCEL_SECTOR_DICT.get(numeric_code, "未分類")
                
                unique_key = f"{numeric_code}_{date_part}_{method}"
                alert_id = int(hashlib.md5(unique_key.encode()).hexdigest(), 16) % (10**10)

                # 重複防止しつつ追加
                if not any(r["アラートID"] == alert_id for r in all_new_rows):
                    all_new_rows.append({
                        "アラートID": alert_id,
                        "ティッカー": ticker_val,
                        "コード": numeric_code,
                        "手法": method,
                        "シグナル内容": row_str[:100],
                        "発生日": date_part,
                        "銘柄名": stock_name,
                        "市場": market_val,
                        "セクター": sector_val,
                        "発生日終値": 0.0,
                    })
        except Exception as e:
            st.error(f"ファイル読み込みエラー: {e}")

    return pd.DataFrame(all_new_rows)

def fetch_stock_data_for_master(df):
    if df.empty:
        return df

    progress_bar = st.progress(0)
    total = len(df)

    for idx, row in df.iterrows():
        progress_bar.progress((idx + 1) / total)
        raw_code = str(row["コード"]).replace(".T", "").replace("TSE:", "").strip().zfill(4)
        yf_code = f"{raw_code}.T"
        start_date_str = str(row["発生日"]).split("T")[0].split(" ")[0]

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        except:
            continue

        stock_name = EXCEL_NAME_DICT.get(raw_code, row.get("銘柄名", f"銘柄_{raw_code}"))
        market_val = EXCEL_MARKET_DICT.get(raw_code, row.get("市場", "プライム"))
        sector_val = EXCEL_SECTOR_DICT.get(raw_code, row.get("セクター", "未分類"))

        df.at[idx, "コード"] = raw_code
        df.at[idx, "市場"] = market_val
        df.at[idx, "セクター"] = sector_val
        df.at[idx, "銘柄名"] = stock_name

        end_date = datetime.now() + timedelta(days=5)

        try:
            hist = yf.download(yf_code, start=start_date - timedelta(days=5), end=end_date, progress=False)
            if hist.empty:
                continue

            if isinstance(hist.columns, pd.MultiIndex):
                hist.columns = hist.columns.get_level_values(0)

            hist.index = hist.index.tz_localize(None)
            future_days = hist[hist.index >= start_date]
            if len(future_days) == 0:
                continue

            base_close = future_days.iloc[0]["Close"]
            if isinstance(base_close, pd.Series):
                base_close = base_close.iloc[0]
            df.at[idx, "発生日終値"] = float(base_close)

            for d in range(1, 21):
                if d < len(future_days):
                    day_row = future_days.iloc[d]
                    open_val = float(day_row["Open"])
                    close_val = float(day_row["Close"])
                    ret_val = ((close_val - float(base_close)) / float(base_close)) * 100

                    df.at[idx, f"{d}日目_始値"] = open_val
                    df.at[idx, f"{d}日目_終値"] = close_val
                    df.at[idx, f"{d}日目_騰落率(%)"] = round(ret_val, 2)
        except Exception as e:
            continue

    progress_bar.empty()
    return df

def calculate_optimal_strategies(dataframe):
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

            if len(returns) >= 2:
                win_count = sum(1 for r in returns if r > 0)
                win_rate = (win_count / len(returns)) * 100
                avg_ret = sum(returns) / len(returns) * 100
                strategies.append({
                    "イン": f"{in_day}営業日後（始値）",
                    "イン日数": in_day,
                    "アウト": f"{out_day}営業日後（終値）",
                    "勝率": round(win_rate, 1),
                    "平均リターン": round(avg_ret, 2),
                    "サンプル数": len(returns)
                })

    strat_df = pd.DataFrame(strategies)
    if not strat_df.empty:
        return strat_df.sort_values(by=["勝率", "平均リターン"], ascending=False).head(3)
    return pd.DataFrame()

def show_custom_dataframe(df):
    if not df.empty:
        df_disp = df.copy()
        df_disp.index = range(1, len(df_disp) + 1)
        st.dataframe(df_disp, use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)


# --- UI メイン ---
st.title("📊 トレード手法 検証・マスター管理ダッシュボード")

master_df = load_master()

st.sidebar.header("⚙️ データ管理・操作")
uploaded_files = st.sidebar.file_uploader(
    "トレーディングビューCSVをアップロード（複数選択可）",
    type=["csv"],
    accept_multiple_files=True,
)

if uploaded_files:
    if st.sidebar.button("🔄 マスター台帳に統合・更新する"):
        new_df = parse_tradingview_csv(uploaded_files)
        st.write("--- デバッグ: 抽出しえた新規データ行数 ---", len(new_df))
        if not new_df.empty:
            st.write(new_df)
            combined = pd.concat([master_df, new_df]).drop_duplicates(subset=["アラートID"], keep="first")
            save_master(combined)
            master_df = combined
            st.sidebar.success(f"マスター台帳を更新しました！（総件数: {len(master_df)}件）")
            st.rerun()
        else:
            st.sidebar.warning("CSVから有効なシグナルデータを検出できませんでした。上のデバッグ情報をご確認ください。")

if not master_df.empty:
    if st.sidebar.button("🚀 株価データ・市場情報を一括自動取得・更新"):
        with st.spinner("株価データおよび市場・セクター情報を取得中..."):
            updated_df = fetch_stock_data_for_master(master_df)
            save_master(updated_df)
            master_df = updated_df
            st.sidebar.success("株価データと市場情報の更新が完了しました！")
            st.rerun()

    if st.sidebar.button("🗑️ マスター台帳をリセット（初期化）"):
        save_master(pd.DataFrame(columns=ALL_COLUMNS))
        st.sidebar.warning("マスター台帳をリセットしました。")
        st.rerun()

st.header("📋 シグナル一覧・トラッキング台帳")
if not master_df.empty:
    show_custom_dataframe(master_df)
else:
    st.info("データがありません。CSVをアップロードしてください。")