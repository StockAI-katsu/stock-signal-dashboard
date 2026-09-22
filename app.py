from datetime import datetime, timedelta
import hashlib
import os
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
    """Excelの『銘柄データ』シートからコードと銘柄名・市場・33業種区分の辞書を作成する"""
    name_dict = {}
    market_dict = {}
    sector_dict = {}
    if os.path.exists(EXCEL_MASTER_FILE):
        try:
            df_ex = pd.read_excel(EXCEL_MASTER_FILE, sheet_name="銘柄データ", dtype={"コード": str})
            for _, row in df_ex.iterrows():
                # .TやTSE:などを除去し、4桁の文字列に正規化
                raw_code = str(row["コード"])
                code = raw_code.replace(".T", "").replace("TSE:", "").strip().zfill(4)
                name = str(row["銘柄名"]).strip()
                market_raw = str(row.get("市場・商品区分", ""))
                # 33業種区分をセクターとして優先取得
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

# Googleスプレッドシート接続の初期化
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

                time_str = str(row.get("日時", row.get("Time", row.get("発生日", datetime.now()))))
                date_part = str(time_str).split("T")[0].split(" ")[0]

                name_val = str(row.get("名前", ""))
                desc_val = str(row.get("名称", row.get("Message", "")))
                combined_text = f"{name_val} {desc_val}".lower()

                if "kotonov" in combined_text or "こと" in combined_text:
                    method = "ことの手法"
                elif "スイング押し目" in combined_text or "押し目" in combined_text:
                    method = "押し目プロ風手法"
                elif name_val and name_val != "nan":
                    method = name_val.strip()
                else:
                    method = "押し目プロ風手法"

                # コードの抽出 (T がついていてもいなくても綺麗に4桁の数字を取り出す)
                code_clean = ticker.split(":")[-1].split(",")[0].strip()
                numeric_code = code_clean.replace(".T", "").replace("TSE:", "").strip().zfill(4)

                stock_name = EXCEL_NAME_DICT.get(numeric_code, f"銘柄_{numeric_code}")
                market_val = EXCEL_MARKET_DICT.get(numeric_code, "プライム")
                sector_val = EXCEL_SECTOR_DICT.get(numeric_code, "未分類")
                
                # サーバー再起動や環境に依存しない安定した一意IDを生成（重複防止対策）
                unique_key = f"{numeric_code}_{date_part}_{method}"
                alert_id = int(hashlib.md5(unique_key.encode()).hexdigest(), 16) % (10**10)

                all_new_rows.append({
                    "アラートID": alert_id,
                    "ティッカー": ticker,
                    "コード": numeric_code,
                    "手法": method,
                    "シグナル内容": f"{name_val} / {desc_val}",
                    "発生日": date_part,
                    "銘柄名": stock_name,
                    "市場": market_val,
                    "セクター": sector_val,
                    "発生日終値": 0.0,
                })
        except Exception as e:
            st.error(f"ファイル {uploaded_file.name} の読み込みに失敗しました: {e}")

    return pd.DataFrame(all_new_rows)

def fetch_stock_data_for_master(df):
    if df.empty:
        return df

    progress_bar = st.progress(0)
    total = len(df)

    for idx, row in df.iterrows():
        progress_bar.progress((idx + 1) / total)
        # コードの正規化と yfinance 用の .T 付きコード作成
        raw_code = str(row["コード"]).replace(".T", "").replace("TSE:", "").strip().zfill(4)
        yf_code = f"{raw_code}.T"
        start_date_str = str(row["発生日"]).split("T")[0].split(" ")[0]

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        except Exception as e:
            continue

        # 常にExcelマスターの最新情報を反映させる
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

            # 20営業日未満であっても、取得できた日数分だけデータを埋める
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
            print(f"Error fetching stock hist for {yf_code}: {e}")
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
        strat_df = strat_df.sort_values(by=["勝率", "平均リターン"], ascending=False)
        return strat_df.head(3)
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
st.markdown("手法別の最適戦略、セクター別分析、および翌日エントリー対象銘柄の自動判定を行います。")

master_df = load_master()

# サイドバー: アップロードと操作
st.sidebar.header("⚙️ データ管理・操作")
uploaded_files = st.sidebar.file_uploader(
    "トレーディングビューCSVをアップロード（複数選択可）",
    type=["csv"],
    accept_multiple_files=True,
)

if uploaded_files:
    if st.sidebar.button("🔄 マスター台帳に統合・更新する"):
        new_df = parse_tradingview_csv(uploaded_files)
        if not new_df.empty:
            combined = pd.concat([master_df, new_df]).drop_duplicates(subset=["アラートID"], keep="first")
            save_master(combined)
            master_df = combined
            st.sidebar.success(f"マスター台帳を更新しました！（総件数: {len(master_df)}件）")
            st.rerun()

if not master_df.empty:
    if st.sidebar.button("🚀 株価データ・市場情報を一括自動取得・更新"):
        with st.spinner("株価データおよび市場・セクター情報を取得中..."):
            updated_df = fetch_stock_data_for_master(master_df)
            save_master(updated_df)
            master_df = updated_df
            st.sidebar.success("株価データと市場情報の更新が完了しました！")
            st.rerun()

    if st.sidebar.button("🗑️ マスター台帳をリセット（初期化）"):
        empty_df = pd.DataFrame(columns=ALL_COLUMNS)
        save_master(empty_df)
        master_df = empty_df
        st.sidebar.warning("マスター台帳をリセットしました。")
        st.rerun()


# --- メイン機能 1：翌日エントリー対象銘柄の自動判定 ---
st.header("🔔 明日エントリー（イン）対象の注目銘柄")
st.markdown("各手法の**勝率第1位の戦略（何営業日後エントリーか）**に基づき、明日がまさにそのエントリー日（イン）にあたる銘柄を自動抽出します。")

if not master_df.empty and "手法" in master_df.columns:
    methods = [str(m) for m in master_df["手法"].dropna().unique() if str(m).strip() != ""]
    today = datetime.now().date()
    
    action_items = []
    for method in methods:
        m_df = master_df[master_df["手法"] == method]
        top_strat = calculate_optimal_strategies(m_df)
        if top_strat.empty:
            continue
        
        target_in_day = top_strat.iloc[0]["イン日数"]
        best_in_str = top_strat.iloc[0]["イン"]
        
        for _, row in m_df.iterrows():
            start_date_str = str(row["発生日"])
            try:
                start_dt = datetime.strptime(start_date_str, "%Y-%m-%d").date()
            except:
                continue
            
            business_days = pd.bdate_range(start=start_dt, end=today)
            elapsed_business_days = len(business_days) - 1
            
            if elapsed_business_days == (target_in_day - 1):
                action_items.append({
                    "手法": method,
                    "コード": row.get("コード", ""),
                    "銘柄名": row.get("銘柄名", ""),
                    "市場": row.get("市場", ""),
                    "セクター": row.get("セクター", ""),
                    "発生日": start_date_str,
                    "推奨エントリー": best_in_str,
                    "判定": "🔥 明日イン！"
                })

    if action_items:
        df_action = pd.DataFrame(action_items)
        show_custom_dataframe(df_action)
    else:
        st.info("💡 現在、明日がエントリー最適日（第1位戦略基準）に該当する銘柄はありません。")

st.markdown("---")


# --- メイン機能 2：手法別・セクター別の最適トレード戦略 ---
st.header("🏆 最適トレード戦略（手法・セクター別 自動算出 Top 3）")

if not master_df.empty and "手法" in master_df.columns:
    methods = [str(m) for m in master_df["手法"].dropna().unique() if str(m).strip() != ""]
    
    selected_method = st.selectbox("📌 検証する手法を選択してください", methods)
    method_df = master_df[master_df["手法"] == selected_method]
    
    st.subheader(f"📊 【{selected_method}】全体の最適戦略（全期間の全パターン検証結果）")
    top_strategies = calculate_optimal_strategies(method_df)
    
    if not top_strategies.empty:
        c1, c2, c3 = st.columns(3)
        cols = [c1, c2, c3]
        for idx, (_, strat) in enumerate(top_strategies.iterrows()):
            with cols[idx]:
                st.markdown(f"""
                **第 {idx+1} 位** 🎯
                * **エントリー**: {strat['イン']}
                * **エグジット**: {strat['アウト']}
                * **勝率**: **{strat['勝率']}%**
                * **平均リターン**: +{strat['平均リターン']}%
                * **サンプル数**: {strat['サンプル数']}件
                """)
    else:
        st.info("💡 該当手法のデータまたは株価実績が不足しています。")

    st.markdown("---")
    st.subheader(f"🏢 【{selected_method}】セクター別 最適戦略")
    sectors = [str(s) for s in method_df["セクター"].dropna().unique() if str(s).strip() != ""]
    if sectors:
        selected_sector = st.selectbox("📌 セクターを選択して詳細を確認", ["全セクター合算"] + sectors)
        if selected_sector != "全セクター合算":
            target_sector_df = method_df[method_df["セクター"] == selected_sector]
        else:
            target_sector_df = method_df
            
        sector_strategies = calculate_optimal_strategies(target_sector_df)
        if not sector_strategies.empty:
            show_custom_dataframe(sector_strategies)
        else:
            st.info(f"セクター「{selected_sector}」のデータが十分にありません。")

st.markdown("---")


# --- メイン画面：手法別・総合のタブ切り替え表示 ---
st.header("📋 シグナル一覧・トラッキング台帳")
if not master_df.empty and "手法" in master_df.columns:
    methods = [str(m) for m in master_df["手法"].dropna().unique() if str(m).strip() != ""]
    if methods:
        tab_labels = ["総合一覧"] + methods
        tabs = st.tabs(tab_labels)

        display_cols = [col for col in master_df.columns if col not in ["アラートID", "ティッカー"]]

        def render_table(df_target):
            show_custom_dataframe(df_target[display_cols])

        with tabs[0]:
            st.subheader("🌐 総合一覧（全手法）")
            render_table(master_df)

        for i, method in enumerate(methods):
            with tabs[i + 1]:
                st.subheader(f"🏷️ 手法：{method}")
                sub_df = master_df[master_df["手法"] == method]
                render_table(sub_df)