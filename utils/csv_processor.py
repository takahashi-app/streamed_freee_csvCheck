"""
CSV処理ユーティリティ
STREAMEDとfreeeのCSVを読み込み、処理する
"""
import pandas as pd
import chardet
from datetime import datetime


class CSVProcessor:
    """CSV処理を行うクラス"""

    def __init__(self):
        pass

    def detect_encoding(self, file_path):
        """
        ファイルのエンコーディングを自動検出

        Args:
            file_path: ファイルパス

        Returns:
            str: 検出されたエンコーディング
        """
        with open(file_path, 'rb') as f:
            raw_data = f.read(100000)  # 最初の100KBを読む
            result = chardet.detect(raw_data)
            return result['encoding']

    def read_csv_auto(self, file_path):
        """
        エンコーディングを自動検出してCSVを読み込む

        Args:
            file_path: CSVファイルパス

        Returns:
            pd.DataFrame: 読み込んだデータフレーム

        Raises:
            Exception: 読み込みに失敗した場合
        """
        encodings = ['cp932', 'shift-jis', 'utf-8', 'utf-8-sig']

        # まず自動検出を試みる
        try:
            detected = self.detect_encoding(file_path)
            if detected:
                encodings.insert(0, detected)
        except:
            pass

        # 各エンコーディングで試行
        last_error = None
        for enc in encodings:
            try:
                df = pd.read_csv(file_path, encoding=enc)
                return df
            except Exception as e:
                last_error = e
                continue

        raise Exception(f"CSVの読み込みに失敗しました: {last_error}")

    def process_streamed_csv(self, df):
        """
        STREAMED CSVを処理

        処理内容:
        1. 列名の変更: 借方補助科目→借方取引先、貸方補助科目→貸方取引先
        2. 伝票番号を生成

        Args:
            df: STREAMED CSVのDataFrame

        Returns:
            pd.DataFrame: 処理後のDataFrame
        """
        df = df.copy()

        # 列名の変更
        rename_map = {
            '借方補助科目': '借方取引先',
            '貸方補助科目': '貸方取引先'
        }
        df.rename(columns=rename_map, inplace=True)

        # 伝票番号を生成
        df = self.generate_voucher_numbers(df)

        return df

    def generate_voucher_numbers(self, df):
        """
        伝票番号を生成（複合仕訳対応）

        形式: 月日時分 + 連番3桁
        例: 12081508001, 12081508002

        元の伝票番号でグループ化し、同じグループには同じ番号を振る

        Args:
            df: DataFrame

        Returns:
            pd.DataFrame: 伝票番号が更新されたDataFrame
        """
        df = df.copy()

        # 現在時刻を取得
        now = datetime.now()
        prefix = now.strftime('%m%d%H%M')  # 月日時分

        # 元の伝票番号の出現順序を取得（重複を排除）
        unique_vouchers = df['伝票番号'].unique()

        # 元の伝票番号 → 新しい伝票番号のマッピングを作成
        voucher_map = {}
        for idx, old_voucher in enumerate(unique_vouchers, 1):
            new_number = f"{prefix}{idx:03d}"
            voucher_map[old_voucher] = new_number

        # すべての行に新しい伝票番号を適用
        df['伝票番号'] = df['伝票番号'].map(voucher_map)

        return df

    def extract_master_data(self, freee_dfs):
        """
        freee仕訳帳CSVから取引先と部門のマスタデータを抽出

        Args:
            freee_dfs: freee仕訳帳CSVのDataFrameのリスト

        Returns:
            dict: {
                'partners': 取引先の重複なしリスト,
                'departments': 部門の重複なしリスト
            }
        """
        # すべてのデータフレームを結合
        if not freee_dfs:
            return {'partners': [], 'departments': []}

        if len(freee_dfs) == 1:
            combined_df = freee_dfs[0]
        else:
            combined_df = pd.concat(freee_dfs, ignore_index=True)

        # 取引先を抽出
        partners = set()
        if '借方取引先名' in combined_df.columns:
            partners.update(combined_df['借方取引先名'].dropna().unique())
        if '貸方取引先名' in combined_df.columns:
            partners.update(combined_df['貸方取引先名'].dropna().unique())

        # 部門を抽出
        departments = set()
        if '借方部門' in combined_df.columns:
            departments.update(combined_df['借方部門'].dropna().unique())
        if '貸方部門' in combined_df.columns:
            departments.update(combined_df['貸方部門'].dropna().unique())

        return {
            'partners': sorted(list(partners)),
            'departments': sorted(list(departments))
        }

    def match_names(self, df, master_data, matcher):
        """
        取引先名と部門名をマスタデータと照合し、候補を追加

        Args:
            df: 処理対象のDataFrame
            master_data: マスタデータ {'partners': [...], 'departments': [...]}
            matcher: NameMatcherインスタンス

        Returns:
            pd.DataFrame: 候補列が追加されたDataFrame
        """
        df = df.copy()

        # 取引先の照合（借方と貸方を統合）
        if master_data['partners']:
            df = self._match_partners_unified(df, master_data['partners'], matcher)

        # 部門の照合（借方と貸方を統合）
        if master_data['departments']:
            df = self._match_departments_unified(df, master_data['departments'], matcher)

        return df

    def _match_partners_unified(self, df, candidates, matcher):
        """
        取引先の照合（借方と貸方を統合して処理）

        Args:
            df: DataFrame
            candidates: 候補リスト
            matcher: NameMatcherインスタンス

        Returns:
            pd.DataFrame: 候補列が追加されたDataFrame
        """
        # 結果列を初期化
        df['STREAMED元の取引先'] = ''
        df['freee取引先名候補1'] = ''
        df['freee取引先名候補2'] = ''
        df['freee取引先名候補3'] = ''
        df['_取引先完全一致'] = False

        for idx, row in df.iterrows():
            # 借方と貸方から取引先を取得（どちらか空でない方を優先）
            debit_partner = row.get('借方取引先', '')
            credit_partner = row.get('貸方取引先', '')

            # 空でない値を優先して使用
            partner = credit_partner if pd.notna(credit_partner) and credit_partner != '' else debit_partner
            if pd.isna(partner) or partner == '':
                continue

            df.at[idx, 'STREAMED元の取引先'] = partner

            # 完全一致チェック（候補は表示しない）
            if partner in candidates:
                df.at[idx, '_取引先完全一致'] = True
                continue

            # 候補を検索
            results = matcher.find_candidates(partner, candidates, top_n=3, threshold=0.0)

            # 結果を格納
            for i, result in enumerate(results[:3], 1):
                df.at[idx, f'freee取引先名候補{i}'] = result['candidate']

        return df

    def _match_departments_unified(self, df, candidates, matcher):
        """
        部門の照合（借方と貸方を統合して処理）

        Args:
            df: DataFrame
            candidates: 候補リスト
            matcher: NameMatcherインスタンス

        Returns:
            pd.DataFrame: 候補列が追加されたDataFrame
        """
        # 結果列を初期化
        df['STREAMED元の部門'] = ''
        df['freee部門候補1'] = ''
        df['freee部門候補2'] = ''
        df['freee部門候補3'] = ''
        df['_部門完全一致'] = False

        for idx, row in df.iterrows():
            # 借方と貸方から部門を取得（どちらか空でない方を優先）
            debit_dept = row.get('借方部門', '')
            credit_dept = row.get('貸方部門', '')

            # 空でない値を優先して使用
            dept = credit_dept if pd.notna(credit_dept) and credit_dept != '' else debit_dept
            if pd.isna(dept) or dept == '':
                continue

            df.at[idx, 'STREAMED元の部門'] = dept

            # 完全一致チェック（候補は表示しない）
            if dept in candidates:
                df.at[idx, '_部門完全一致'] = True
                continue

            # 候補を検索
            results = matcher.find_candidates(dept, candidates, top_n=3, threshold=0.0)

            # 結果を格納
            for i, result in enumerate(results[:3], 1):
                df.at[idx, f'freee部門候補{i}'] = result['candidate']

        return df
