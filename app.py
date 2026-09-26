from datetime import datetime, timedelta
import os
import pandas as pd
import streamlit as st
import yfinance as yf

# ページの設定
st.set_page_config(
    page_title="トレード手法 検証・マスター管理ダッシュボード",
    layout="wide",
)

# ファイルパスの定義
MASTER_EXCEL_OUT = "トレード手法検証マスター.xlsx"
EXCEL_MASTER_FILE = "株式市場・テーマ監視ボード.xlsx"


@st.cache_data
def load_excel_stock_master():
  """Excelの『銘柄データ』シートからコード、銘柄名、市場、セクターの辞書を作成する"""
  name_dict = {}
  market_dict = {}
  sector_dict = {}

  if os.path.exists(EXCEL_MASTER_FILE):
    try:
      df_ex = pd.read_excel(
          EXCEL_MASTER_FILE, sheet_name="銘柄データ", dtype={"コード": str}
      )
      for _, row in df_ex.iterrows():
        code = str(row["コード"]).replace(".T", "").strip().zfill(4)
        name = str(row["銘柄名"]).strip()
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
      print(f"Excel読み込みエラー: {e}")
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
    "手法",
]
ALL_COLUMNS = ["アラートID", "ティッカー"] + BASE_COLUMNS + DAY_COLUMNS


def load_master():
  """マスターデータを読み込み、すべての数値列を強制的にfloat64に変換する"""
  df = pd.DataFrame(columns=ALL_COLUMNS)
  if os.path.exists(MASTER_EXCEL_OUT):
    try:
      df = pd.read_excel(
          MASTER_EXCEL_OUT, sheet_name="master_signals", dtype=str
      )
    except Exception as e:
      try:
        xls = pd.ExcelFile(MASTER_EXCEL_OUT)
        if xls.sheet_names:
          df = pd.read_excel(
              MASTER_EXCEL_OUT, sheet_name=xls.sheet_names[0], dtype=str
          )
      except:
        pass

  for col in ALL_COLUMNS:
    if col not in df.columns:
      df[col] = None

  # 発生日終値および20営業日分の列を強制的にfloat型へ変換
  numeric_cols = ["発生日終値"] + DAY_COLUMNS
  for col in numeric_cols:
    if col in df.columns:
      df[col] = pd.to_numeric(df[col], errors="coerce")

  return df


def save_master(df):
  """マスターデータをExcelファイルへ保存・永続化する"""
  try:
    with pd.ExcelWriter(MASTER_EXCEL_OUT, engine="openpyxl") as writer:
      df.to_excel(writer, sheet_name="master_signals", index=False)
  except Exception as e:
    st.error(f"マスターExcelファイルの保存に失敗しました: {e}")


def parse_tradingview_csv(uploaded_files):
  """トレーディングビューCSVをパースする"""
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

        code_clean = ticker.split(":")[-1].split(",")[0].strip()
        numeric_code = code_clean.replace(".T", "").zfill(4)

        stock_name = EXCEL_NAME_DICT.get(numeric_code, f"銘柄_{numeric_code}")
        market_val = EXCEL_MARKET_DICT.get(numeric_code, "プライム")
        sector_val = EXCEL_SECTOR_DICT.get(numeric_code, "未分類")

        alert_id = abs(hash(f"{ticker}_{date_part}_{method}")) % (10**10)

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


def fetch_stock_data_for_master(df, target_indices=None):
  """株価データの取得と、Excelマスターからの情報再同期を行う（型安全・完全修正版）"""
  if df.empty:
    return df, 0, 0

  # ★重要対策: 代入エラーを防ぐため、処理対象DataFrameの数値列を明示的にfloat型にキャストする
  numeric_cols = ["発生日終値"] + DAY_COLUMNS
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

      base_close = future_days.iloc[0]["Close"]
      if isinstance(base_close, pd.Series):
        base_close = base_close.iloc[0]

      if pd.isna(base_close) or float(base_close) == 0:
        fail_count += 1
        continue

      df.at[idx, "発生日終値"] = float(base_close)

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
    dataframe, max_days=10, selected_market="すべて"
):
  """フィルター条件に合致する候補銘柄を抽出する"""
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
      days_diff = (today - start_dt).days

      if days_diff <= max_days:
        market_val = str(row.get("市場", "プライム"))
        if selected_market == "すべて" or selected_market in market_val:
          candidates.append({
              "発生日": row["発生日"],
              "手法": row.get("手法", "不明"),
              "コード": row["コード"],
              "銘柄名": row["銘柄名"],
              "市場": market_val,
              "セクター": row.get("セクター", "未分類"),
              "発生日終値": row["発生日終値"],
          })
    except:
      continue

  return pd.DataFrame(candidates)


# --- UI メイン ---
st.title("📊 トレード手法 検証・マスター管理ダッシュボード")
st.markdown(
    f"Excelファイル (`{EXCEL_MASTER_FILE}`) の銘柄データをベースに、シグナルと20営業日分の値動きを自動追跡し、`{MASTER_EXCEL_OUT}`"
    f" の **`master_signals` シート** に永続保存します。"
)

master_df = load_master()

# --- サイドバー: データ管理 ＆ 絞り込みフィルター ---
st.sidebar.header("⚙️ データ管理・操作")

if os.path.exists(MASTER_EXCEL_OUT):
  with open(MASTER_EXCEL_OUT, "rb") as f:
    st.sidebar.download_button(
        label="📦 GitHub用マスターExcelをダウンロード",
        data=f,
        file_name=MASTER_EXCEL_OUT,
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
else:
  st.sidebar.info("💡 マスターExcelファイルはまだ生成されていません。")

st.sidebar.markdown("---")

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
          subset=["アラートID"], keep="first"
      )
      save_master(combined)
      master_df = combined
      st.sidebar.success(
          f"マスター台帳を更新し、Excelの master_signals"
          f" に保存しました！（総件数: {len(master_df)}件）"
      )
      st.rerun()

if not master_df.empty:
  st.sidebar.markdown("### 📥 株価データ取得・更新")
  if st.sidebar.button("🚀 全件一括自動取得・更新"):
    with st.spinner(
        "すべての銘柄の株価データおよびExcel上のセクター・市場情報を同期中..."
    ):
      updated_df, succ, fail = fetch_stock_data_for_master(master_df)
      save_master(updated_df)
      master_df = updated_df
      st.sidebar.success(
          f"一括同期完了！ (成功: {succ}件 / 失敗・データなし: {fail}件)"
      )
      st.rerun()

  # 未取得・0円の銘柄だけをピンポイントで再取得するボタン
  missing_df_indices = master_df[
      (master_df["発生日終値"].isna())
      | (master_df["発生日終値"] == 0)
      | (master_df["1日目_始値"].isna())
  ].index

  if len(missing_df_indices) > 0:
    if st.sidebar.button(
        f"⚠️ 未取得・エラー銘柄のみ再取得 ({len(missing_df_indices)}件)"
    ):
      with st.spinner("未取得の銘柄データをピンポイントで再取得中..."):
        updated_df, succ, fail = fetch_stock_data_for_master(
            master_df, target_indices=missing_df_indices
        )
        save_master(updated_df)
        master_df = updated_df
        st.sidebar.success(
            f"再取得完了！ (成功: {succ}件 / 失敗: {fail}件)"
        )
        st.rerun()
  else:
    st.sidebar.success("✅ 未取得のデータはありません（全件取得済み）")

  if st.sidebar.button("🗑️ マスター台帳をリセット（初期化）"):
    if os.path.exists(MASTER_EXCEL_OUT):
      os.remove(MASTER_EXCEL_OUT)
    st.sidebar.warning("マスター台帳のExcelファイルをリセットしました。")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("🔍 候補銘柄の絞り込み条件")
filter_max_days = st.sidebar.slider(
    "発生からの経過日数（以内）", min_value=1, max_value=30, value=10
)
filter_market = st.sidebar.selectbox(
    "対象市場フィルター", ["すべて", "プライム", "スタンダード", "グロース"]
)


# --- 手法（シグナル）リストの取得 ---
methods = []
if not master_df.empty and "手法" in master_df.columns:
  methods = [
      str(m) for m in master_df["手法"].dropna().unique() if str(m).strip() != ""
  ]

all_tab_labels = (
    ["総合（全手法合算）"] + methods if methods else ["総合（全手法合算）"]
)
selected_tabs = st.tabs(all_tab_labels)


for idx, tab_label in enumerate(all_tab_labels):
  with selected_tabs[idx]:
    if idx == 0:
      st.subheader("🌐 総合（全手法合算）の分析結果")

      st.markdown("### 🎯 明日のトレード候補（絞り込み結果一覧）")
      all_candidates = []
      if methods:
        for m in methods:
          m_df = master_df[master_df["手法"] == m]
          m_cand = get_filtered_trade_candidates(
              m_df, max_days=filter_max_days, selected_market=filter_market
          )
          if not m_cand.empty:
            all_candidates.append(m_cand)

      if all_candidates:
        combined_candidates_df = pd.concat(
            all_candidates, ignore_index=True
        )
        st.success(
            f"条件に合致する候補銘柄が **{len(combined_candidates_df)}件**"
            " 見つかりました！"
        )
        st.dataframe(combined_candidates_df, width="stretch")
      else:
        st.info(
            "💡 現在、指定した条件に合致するエントリー候補銘柄はありません。"
        )

      st.markdown("---")
      st.markdown("### 🏆 【手法別】セクター別 ＆ 市場別 最適戦略 (Top 3)")

      if methods and not master_df.empty:
        for m in methods:
          st.markdown(f"#### 🏷️ 手法：`{m}` の深掘り分析")
          m_df = master_df[master_df["手法"] == m]

          st.markdown("##### 🏢 セクター別最適戦略")
          sectors = [
              s
              for s in m_df["セクター"].dropna().unique()
              if str(s).strip() != "" and str(s) != "nan"
          ]
          if sectors:
            for sector in sectors:
              sector_df = m_df[m_df["セクター"] == sector]
              if len(sector_df) >= 3:
                st.markdown(
                    f"**・セクター: 【 {sector} 】** (検証数:"
                    f" {len(sector_df)}件)"
                )
                sec_top = calculate_optimal_strategies_for_df(
                    sector_df, method_name=f"{m} ({sector})"
                )
                if not sec_top.empty:
                  c1, c2, c3 = st.columns(3)
                  cols = [c1, c2, c3]
                  for s_idx, (_, strat) in enumerate(sec_top.iterrows()):
                    with cols[s_idx]:
                      cols[s_idx].markdown(f"""
                                    **{s_idx+1}位**: {strat['イン']}〜{strat['アウト']} | 勝率:**{strat['勝率']}%** (平均:+{strat['平均リターン']}%)
                                    """)

          st.markdown("")
          st.markdown("##### 🏛️ 市場区分別最適戦略")
          markets = [
              mk
              for mk in ["プライム", "スタンダード", "グロース"]
              if mk in m_df["市場"].values
          ]
          if not markets:
            markets = [
                mk
                for mk in m_df["市場"].dropna().unique()
                if str(mk).strip() != ""
            ]

          if markets:
            for market in markets:
              market_df = m_df[m_df["市場"].str.contains(market, na=False)]
              if len(market_df) >= 3:
                st.markdown(
                    f"**・市場: 【 {market} 】** (検証数: {len(market_df)}件)"
                )
                mkt_top = calculate_optimal_strategies_for_df(
                    market_df, method_name=f"{m} ({market})"
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
      else:
        st.info("💡 手法データが登録されていません。")

      st.markdown("### 📋 シグナル一覧・トラッキング台帳（全件）")
      if not master_df.empty:
        st.markdown(f"表示件数: **{len(master_df)}件**")
        display_cols = [
            col
            for col in master_df.columns
            if col not in ["アラートID", "ティッカー"]
        ]
        st.dataframe(master_df[display_cols], width="stretch")
      else:
        st.info("該当するデータはありません。")

    else:
      current_method = methods[idx - 1]
      target_df = master_df[master_df["手法"] == current_method]
      st.subheader(f"🏷️ 手法：{current_method} の分析結果")

      candidates_df = get_filtered_trade_candidates(
          target_df, max_days=filter_max_days, selected_market=filter_market
      )

      st.markdown("### 🎯 明日のトレード候補（絞り込み抽出）")
      if not candidates_df.empty:
        st.success(
            f"条件に合致する候補銘柄が **{len(candidates_df)}件** 見つかりました！"
        )
        st.dataframe(candidates_df, width="stretch")
      else:
        st.info(
            "💡 現在、指定した条件に合致する翌日エントリー候補銘柄はありません。"
        )

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

      st.markdown("")
      st.markdown("#### 🏛️ 市場区分別最適戦略")
      markets = [
          mk
          for mk in ["プライム", "スタンダード", "グロース"]
          if mk in target_df["市場"].values
      ]
      if not markets:
        markets = [
            mk
            for mk in target_df["市場"].dropna().unique()
            if str(mk).strip() != ""
        ]

      if markets:
        for market in markets:
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
      st.markdown("### 📋 シグナル一覧・トラッキング台帳")
      if not target_df.empty:
        st.markdown(f"表示件数: **{len(target_df)}件**")
        display_cols = [
            col
            for col in target_df.columns
            if col not in ["アラートID", "ティッカー"]
        ]
        st.dataframe(target_df[display_cols], width="stretch")
      else:
        st.info("該当するデータはありません。")