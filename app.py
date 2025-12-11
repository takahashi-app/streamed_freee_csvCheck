"""
freee会計インポート支援アプリ

STREAMEDからのCSVをfreee会計へインポートするための前処理を行う
"""
import streamlit as st
import pandas as pd
import os
import sys
import subprocess
import io
from pathlib import Path
from datetime import datetime
from utils.csv_processor import CSVProcessor
from utils.name_matcher import NameMatcher
from utils.excel_writer import ExcelWriter


# ページ設定
st.set_page_config(
    page_title="STREAMED→freee会計 インポート用CSV修正アプリ",
    page_icon="📄",
    layout="wide"
)

# セッション状態の初期化
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'stage' not in st.session_state:
    st.session_state.stage = 1
if 'processed_df' not in st.session_state:
    st.session_state.processed_df = None
if 'master_data' not in st.session_state:
    st.session_state.master_data = None
if 'stage2_result_df' not in st.session_state:
    st.session_state.stage2_result_df = None
if 'stage2_original_df' not in st.session_state:
    st.session_state.stage2_original_df = None


def open_file(file_path):
    """
    ファイルを自動的に開く（クロスプラットフォーム対応）

    Args:
        file_path: 開くファイルのパス
    """
    try:
        if sys.platform == 'win32':
            os.startfile(file_path)
        elif sys.platform == 'darwin':
            subprocess.run(['open', file_path])
        else:
            subprocess.run(['xdg-open', file_path])
    except Exception as e:
        st.warning(f"⚠️ ファイルを自動で開けませんでした: {str(e)}")


def check_password():
    """パスワード認証画面"""
    st.title("STREAMED→freee会計  \nインポート用CSV修正アプリ")
    st.markdown("---")
    st.subheader("パスワードを入力してください")

    # フォームを使用してエンターキーでの送信に対応
    with st.form("password_form"):
        password = st.text_input("パスワード", type="password", key="password_input")
        submitted = st.form_submit_button("ログイン", type="primary")

    # フォームが送信された場合のみログイン処理を実行
    if submitted:
        # Streamlit Secretsからパスワードを取得
        try:
            correct_password = st.secrets["passwords"]["system_password"]
        except:
            # Streamlit Secretsが設定されていない場合はエラー
            st.error("❌ パスワードが設定されていません。.streamlit/secrets.tomlを設定してください。")
            return

        if password == correct_password:
            st.session_state.authenticated = True
            st.rerun()
        elif password:  # パスワードが入力されている場合のみエラー表示
            st.error("❌ パスワードが正しくありません")


def main():
    # パスワード認証チェック
    if not st.session_state.authenticated:
        check_password()
        return

    st.title("STREAMED→freee会計  \nインポート用CSV修正アプリ")
    st.markdown("---")

    # サイドバーにステージ選択
    with st.sidebar:
        st.header("処理ステージ")
        stage = st.radio(
            "ステージを選択",
            [1, 2],
            format_func=lambda x: f"ステージ {x}: {'初回処理' if x == 1 else 'freeeインポート用CSV生成'}",
            index=st.session_state.stage - 1
        )
        st.session_state.stage = stage

    # ステージ1: 初回処理
    if st.session_state.stage == 1:
        stage1_process()

    # ステージ2: 再アップロード処理
    elif st.session_state.stage == 2:
        stage2_process()


def stage1_process():
    """ステージ1: 初回処理"""

    st.header("ステージ1: 初回処理")
    st.markdown("""
    1. STREAMED CSVと freee仕訳帳CSVをアップロード
    2. 取引先名・部門名の表記ゆれをチェック
    3. 候補付きExcelファイルを出力
    """)

    # 新規処理ボタン（処理済みの場合のみ表示）
    if st.session_state.processed_df is not None:
        if st.button("🔄 新規処理を開始", type="secondary"):
            st.session_state.processed_df = None
            st.session_state.master_data = None
            st.rerun()

    # ファイルアップロード
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("STREAMED CSV")
        streamed_file = st.file_uploader(
            "STREAMEDからのCSVをアップロード",
            type=['csv'],
            key='streamed_upload'
        )

    with col2:
        st.subheader("freee仕訳帳CSV（新方式）")
        freee_files = st.file_uploader(
            "freee仕訳帳CSVをアップロード",
            type=['csv'],
            accept_multiple_files=True,
            key='freee_upload'
        )
        st.caption("※ 複数ファイルを選択可能です（過年度分を含む場合）")

    # 処理実行
    if streamed_file and freee_files:
        if st.button("🚀 処理を実行", type="primary"):
            with st.spinner("処理中..."):
                try:
                    # 処理を実行
                    result_df = process_stage1(streamed_file, freee_files)

                    # 成功メッセージ
                    st.success("✅ 処理が完了しました！")

                except Exception as e:
                    st.error(f"❌ エラーが発生しました: {str(e)}")
                    st.exception(e)
                    st.session_state.processed_df = None

    # 処理結果の表示（セッション状態に保存されている場合）
    if st.session_state.processed_df is not None:
        st.markdown("---")

        # 完了メッセージ
        st.success("✅ チェック完了しました。下のボタンからファイルを出力してチェックしてください。")

        # 統計情報
        show_statistics(st.session_state.processed_df)

        # Excel出力
        output_section(st.session_state.processed_df)


def process_stage1(streamed_file, freee_files):
    """
    ステージ1の処理を実行

    Args:
        streamed_file: STREAMEDのCSVファイル
        freee_files: freee仕訳帳のCSVファイルリスト

    Returns:
        pd.DataFrame: 処理後のデータフレーム
    """
    processor = CSVProcessor()
    matcher = NameMatcher()

    # STREAMED CSVを読み込み
    st.info("📖 STREAMED CSVを読み込んでいます...")
    streamed_df = pd.read_csv(streamed_file, encoding='cp932')

    # freee仕訳帳CSVを読み込み
    st.info(f"📖 freee仕訳帳CSV（{len(freee_files)}ファイル）を読み込んでいます...")
    freee_dfs = []
    for file in freee_files:
        df = pd.read_csv(file, encoding='cp932')
        freee_dfs.append(df)

    # マスタデータを抽出
    st.info("🔍 取引先・部門マスタを抽出しています...")
    master_data = processor.extract_master_data(freee_dfs)
    st.session_state.master_data = master_data

    st.success(f"✅ 取引先: {len(master_data['partners'])}件、部門: {len(master_data['departments'])}件")

    # STREAMED CSVを処理
    st.info("⚙️ STREAMED CSVを処理しています...")
    processed_df = processor.process_streamed_csv(streamed_df)

    # 表記ゆれチェック
    st.info("🔎 取引先名・部門名の表記ゆれをチェックしています...")
    result_df = processor.match_names(processed_df, master_data, matcher)

    # セッションに保存
    st.session_state.processed_df = result_df

    return result_df


def show_statistics(df):
    """統計情報を表示"""

    st.subheader("📈 統計情報")

    col1, col2 = st.columns(2)

    with col1:
        perfect_match_partner = df['_取引先完全一致'].sum() if '_取引先完全一致' in df.columns else 0
        total_partner = df['STREAMED元の取引先'].notna().sum() if 'STREAMED元の取引先' in df.columns else 0
        st.metric("取引先 完全一致", f"{perfect_match_partner} / {total_partner}件")

    with col2:
        perfect_match_dept = df['_部門完全一致'].sum() if '_部門完全一致' in df.columns else 0
        total_dept = df['STREAMED元の部門'].notna().sum() if 'STREAMED元の部門' in df.columns else 0
        st.metric("部門 完全一致", f"{perfect_match_dept} / {total_dept}件")


def output_section(df):
    """Excel出力セクション"""

    st.subheader("💾 Excel出力")

    # Excelファイルをメモリ上で生成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"freee_import_check_{timestamp}.xlsx"

    # ExcelWriterを使ってメモリ上でファイルを生成
    buffer = io.BytesIO()
    writer_obj = ExcelWriter()

    # 一時ファイルとして保存してからメモリに読み込む
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        writer_obj.write_to_excel(df, tmp.name)
        tmp.seek(0)
        with open(tmp.name, 'rb') as f:
            buffer.write(f.read())
        os.unlink(tmp.name)

    buffer.seek(0)

    # ダウンロードボタン
    st.download_button(
        label="📥 Excelファイルをダウンロード",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

    st.info("💡 ボタンをクリックすると、ブラウザのダウンロードフォルダに保存されます")


def stage2_process():
    """ステージ2: 再アップロード処理"""

    st.header("ステージ2: freeeインポート用CSV生成")
    st.markdown("""
    1. 目視確認後のExcelファイルをアップロード
    2. 候補1を自動適用
    3. freeeインポート用CSVを生成
    """)

    # 新規処理ボタン（処理済みの場合のみ表示）
    if st.session_state.stage2_result_df is not None:
        if st.button("🔄 新規処理を開始", type="secondary"):
            st.session_state.stage2_result_df = None
            st.rerun()

    # ファイルアップロード
    excel_file = st.file_uploader(
        "📊 目視確認後のExcelファイルをアップロード",
        type=['xlsx'],
        key='excel_upload'
    )

    if excel_file:
        if st.button("🚀 freeeインポート用CSV生成", type="primary"):
            with st.spinner("処理中..."):
                try:
                    # 処理を実行
                    result_df = process_stage2(excel_file)

                    # セッション状態に保存
                    st.session_state.stage2_result_df = result_df

                    # 成功メッセージ
                    st.success("✅ 処理が完了しました！")

                except Exception as e:
                    st.error(f"❌ エラーが発生しました: {str(e)}")
                    st.exception(e)
                    st.session_state.stage2_result_df = None

    # 処理結果の表示（セッション状態に保存されている場合）
    if st.session_state.stage2_result_df is not None:
        st.markdown("---")

        # 完了メッセージ
        st.success("✅ チェック完了しました。下のボタンからファイルを出力してチェックしてください。")

        # CSV・Excel出力
        output_stage2_section(st.session_state.stage2_result_df, st.session_state.get('stage2_original_df'))


def process_stage2(excel_file):
    """
    ステージ2の処理を実行

    Args:
        excel_file: 目視確認後のExcelファイル

    Returns:
        pd.DataFrame: freeeインポート用データフレーム
    """
    # Excelを読み込み
    st.info("📖 Excelファイルを読み込んでいます...")
    df = pd.read_excel(excel_file)

    # 元のデータをセッション状態に保存
    st.session_state.stage2_original_df = df.copy()

    # 候補1を適用
    st.info("⚙️ 候補1を適用しています...")

    # freee取引先名候補1 → 貸方取引先
    if 'freee取引先名候補1' in df.columns and '貸方取引先' in df.columns:
        mask = (df['freee取引先名候補1'].notna()) & (df['freee取引先名候補1'] != '')
        df.loc[mask, '貸方取引先'] = df.loc[mask, 'freee取引先名候補1']

    # 貸方取引先 → 借方取引先（空欄の場合のみ）
    if '借方取引先' in df.columns and '貸方取引先' in df.columns:
        mask = df['借方取引先'].isna() | (df['借方取引先'] == '')
        df.loc[mask, '借方取引先'] = df.loc[mask, '貸方取引先']

    # 複合仕訳の場合、同じ伝票番号内の取引先を全行にコピー
    if '伝票番号' in df.columns and '借方取引先' in df.columns and '貸方取引先' in df.columns:
        st.info("📋 複合仕訳の取引先を統一しています...")
        for voucher_num in df['伝票番号'].unique():
            # 同じ伝票番号の行を取得
            voucher_mask = df['伝票番号'] == voucher_num
            voucher_rows = df[voucher_mask]

            # 取引先名を取得（貸方取引先または借方取引先から）
            partner_name = None
            for _, row in voucher_rows.iterrows():
                if pd.notna(row.get('貸方取引先')) and row.get('貸方取引先') != '':
                    partner_name = row.get('貸方取引先')
                    break
                elif pd.notna(row.get('借方取引先')) and row.get('借方取引先') != '':
                    partner_name = row.get('借方取引先')
                    break

            # 取引先名を同じ伝票番号のすべての行にコピー
            if partner_name:
                df.loc[voucher_mask, '借方取引先'] = partner_name
                df.loc[voucher_mask, '貸方取引先'] = partner_name

    # freee部門候補1 → 借方部門・貸方部門
    if 'freee部門候補1' in df.columns:
        mask = (df['freee部門候補1'].notna()) & (df['freee部門候補1'] != '')
        if '借方部門' in df.columns:
            df.loc[mask, '借方部門'] = df.loc[mask, 'freee部門候補1']
        if '貸方部門' in df.columns:
            df.loc[mask, '貸方部門'] = df.loc[mask, 'freee部門候補1']

    # 複合仕訳の場合、同じ伝票番号内の部門を全行にコピー
    if '伝票番号' in df.columns and '借方部門' in df.columns and '貸方部門' in df.columns:
        st.info("📋 複合仕訳の部門を統一しています...")
        for voucher_num in df['伝票番号'].unique():
            # 同じ伝票番号の行を取得
            voucher_mask = df['伝票番号'] == voucher_num
            voucher_rows = df[voucher_mask]

            # 部門名を取得（借方部門または貸方部門から）
            dept_name = None
            for _, row in voucher_rows.iterrows():
                if pd.notna(row.get('借方部門')) and row.get('借方部門') != '':
                    dept_name = row.get('借方部門')
                    break
                elif pd.notna(row.get('貸方部門')) and row.get('貸方部門') != '':
                    dept_name = row.get('貸方部門')
                    break

            # 部門名を同じ伝票番号のすべての行にコピー
            if dept_name:
                df.loc[voucher_mask, '借方部門'] = dept_name
                df.loc[voucher_mask, '貸方部門'] = dept_name

    # 候補列とフラグ列、STREAMED元の列を削除
    cols_to_drop = [col for col in df.columns if '候補' in col or '_' in col or 'STREAMED元' in col]
    df = df.drop(columns=cols_to_drop, errors='ignore')

    return df


def output_stage2_section(processed_df, original_df):
    """ステージ2のCSV・Excel出力セクション"""

    # CSV出力
    st.subheader("💾 CSV出力（freeeインポート用）")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"freee_import_{timestamp}.csv"

    # CSVをメモリ上で生成（CP932でExcelとfreeeで正しく開ける）
    buffer = io.BytesIO()
    # float_format='%.0f'で小数点以下を出力しない（freeeインポート用）
    processed_df.to_csv(buffer, index=False, encoding='cp932', float_format='%.0f')
    csv_data = buffer.getvalue()

    st.download_button(
        label="📥 CSVファイルをダウンロード",
        data=csv_data,
        file_name=csv_filename,
        mime="text/csv",
        type="primary"
    )

    st.info("💡 ボタンをクリックすると、ブラウザのダウンロードフォルダに保存されます")

    # Excel出力（2シート構成）
    st.markdown("---")
    st.subheader("Excel出力（参考用）")
    st.caption("※ CSVファイルをfreeeにインポートしてください。このExcelファイルは内容確認用です。")

    excel_filename = f"freee_import_{timestamp}.xlsx"

    # Excelファイルをメモリ上で生成（色分け付き）
    import tempfile
    excel_buffer = io.BytesIO()

    # 一時ファイルとして保存してからメモリに読み込む
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        writer_obj = ExcelWriter()
        writer_obj.write_stage2_excel(
            original_df if original_df is not None else processed_df,
            processed_df,
            tmp.name
        )
        tmp.seek(0)
        with open(tmp.name, 'rb') as f:
            excel_buffer.write(f.read())
        os.unlink(tmp.name)

    excel_buffer.seek(0)

    st.download_button(
        label="📥 Excelファイルをダウンロード（参考用）",
        data=excel_buffer,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="secondary"
    )

    st.info("💡 ボタンをクリックすると、ブラウザのダウンロードフォルダに保存されます")




if __name__ == "__main__":
    main()
