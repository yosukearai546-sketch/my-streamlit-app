import streamlit as st
import openpyxl
import pandas as pd

st.set_page_config(page_title="見積書計算チェックアプリ", layout="wide")

st.title("📄 見積書 Excel 自動計算チェック")
st.write("Excelファイル（.xlsx）をドラッグ＆ドロップすると、計算ミスや数式抜けをチェックします。")

# 1. ファイルアップローダー（ドラッグ＆ドロップ）
uploaded_file = st.file_uploader("見積書Excelファイルをアップロード", type=["xlsx"])

if uploaded_file is not None:
    # -------------------------------------------------------------
    # 2. openpyxlでExcelを読み込み（数式と値の両方を取得）
    # -------------------------------------------------------------
    # 値（計算結果）を取得用
    wb_val = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws_val = wb_val.active

    # 数式文字列を取得用（data_only=False）
    uploaded_file.seek(0)
    wb_formula = openpyxl.load_workbook(uploaded_file, data_only=False)
    ws_formula = wb_formula.active

    errors = []
    data_rows = []

    # -------------------------------------------------------------
    # 3. 明細行の計算チェック (例: 5行目〜10行目をチェック対象とする)
    #    ※実際のみ積書フォーマットに合わせて行番号・列番号を調整してください
    #    B列: 品名, C列: 単価, D列: 数量, E列: 金額
    # -------------------------------------------------------------
    st.subheader("🔍 チェック結果")

    for row in range(5, 11):
        item_name = ws_val[f"B{row}"].value
        unit_price = ws_val[f"C{row}"].value
        quantity = ws_val[f"D{row}"].value
        amount = ws_val[f"E{row}"].value

        # 行が空ならスキップ
        if not item_name and not unit_price and not amount:
            continue

        # 数値チェック（単価 × 数量 == 金額）
        if isinstance(unit_price, (int, float)) and isinstance(quantity, (int, float)):
            expected_amount = unit_price * quantity
            
            if amount != expected_amount:
                errors.append(
                    f"❌ **【明細計算ミス】 行 {row}（{item_name}）**: "
                    f"単価({unit_price:,}) × 数量({quantity}) ＝ **{expected_amount:,}円** であるべきですが、"
                    f"**{amount:,}円** と入力されています。"
                )

        data_rows.append({
            "行": row,
            "品名": item_name,
            "単価": unit_price,
            "数量": quantity,
            "記載金額": amount
        })

    # -------------------------------------------------------------
    # 4. 合計セルの数式チェック (例: E11セルが合計欄とする)
    # -------------------------------------------------------------
    sum_cell_formula = str(ws_formula["E11"].value)
    
    if sum_cell_formula and not sum_cell_formula.startswith("="):
        errors.append("⚠️ **【数式抜け警告】 E11セル（合計欄）**: 合計が関数（SUM等）ではなく**「手入力の値」**になっています。前後の行変更でズレるリスクがあります。")

    # -------------------------------------------------------------
    # 5. 画面への表示
    # -------------------------------------------------------------
    col1, col2 = st.columns([1, 1])

    with col1:
        st.write("### 📋 抽出データ")
        if data_rows:
            st.dataframe(pd.DataFrame(data_rows), use_container_width=True)

    with col2:
        st.write("### 🚨 判定レポート")
        if errors:
            for err in errors:
                st.error(err)
        else:
            st.success("✅ **問題ありません！** すべての明細計算および合計欄の不備は見つかりませんでした。")