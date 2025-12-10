"""
表記ゆれチェックライブラリ
企業名・部門名の類似度を計算し、候補を提示する
"""
import re
import unicodedata
import jaconv
from Levenshtein import distance as levenshtein_distance


class NameMatcher:
    """名称のマッチングを行うクラス"""

    # 法人格のパターン
    LEGAL_PATTERNS = [
        r'株式会社',
        r'\(株\)',
        r'㈱',
        r'有限会社',
        r'\(有\)',
        r'㈲',
        r'合名会社',
        r'合資会社',
        r'合同会社',
        r'LLC',
        r'Co\.,?\s*Ltd\.?',
        r'Holdings?',
        r'HD',
        r'Corporation',
        r'Corp\.?',
        r'Inc\.?',
        r'Limited',
        r'Ltd\.?',
    ]

    # 除去する記号
    REMOVE_SYMBOLS = r'[×・／\-\s\.\,\(\)（）]'

    def __init__(self, ngram_weight=0.5, prefix_weight=0.3, edit_weight=0.2):
        """
        Args:
            ngram_weight: N-gram類似度の重み
            prefix_weight: 前方一致スコアの重み
            edit_weight: 編集距離スコアの重み
        """
        self.ngram_weight = ngram_weight
        self.prefix_weight = prefix_weight
        self.edit_weight = edit_weight

    def normalize(self, text):
        """
        テキストを正規化する

        処理内容:
        1. Unicode正規化 (NFKC)
        2. 全角→半角変換
        3. カタカナ→ひらがな統一
        4. 大文字→小文字変換
        5. 法人格の除去
        6. 記号の除去
        7. 先頭の×除去

        Args:
            text: 正規化対象のテキスト

        Returns:
            tuple: (正規化後のテキスト, 元のテキスト)
        """
        if not text or pd.isna(text):
            return "", text

        original = str(text)
        normalized = original

        # Unicode正規化（NFKC: 互換文字を統一）
        normalized = unicodedata.normalize('NFKC', normalized)

        # 全角英数字→半角
        normalized = jaconv.z2h(normalized, kana=False, digit=True, ascii=True)

        # カタカナ→ひらがな
        normalized = jaconv.kata2hira(normalized)

        # 大文字→小文字
        normalized = normalized.lower()

        # 法人格の除去
        for pattern in self.LEGAL_PATTERNS:
            normalized = re.sub(pattern, '', normalized, flags=re.IGNORECASE)

        # 記号の除去
        normalized = re.sub(self.REMOVE_SYMBOLS, '', normalized)

        # 先頭の×除去
        normalized = normalized.lstrip('×')

        # 前後の空白除去
        normalized = normalized.strip()

        return normalized, original

    def ngram_similarity(self, text1, text2, n=2):
        """
        N-gram (デフォルト2-gram) のJaccard類似度を計算

        Args:
            text1: テキスト1
            text2: テキスト2
            n: N-gramのN（デフォルト: 2）

        Returns:
            float: Jaccard類似度 (0.0 ~ 1.0)
        """
        if not text1 or not text2:
            return 0.0

        # N-gramセットを作成
        def create_ngrams(text, n):
            if len(text) < n:
                return {text}
            return {text[i:i+n] for i in range(len(text) - n + 1)}

        ngrams1 = create_ngrams(text1, n)
        ngrams2 = create_ngrams(text2, n)

        # Jaccard類似度 = 積集合 / 和集合
        intersection = len(ngrams1 & ngrams2)
        union = len(ngrams1 | ngrams2)

        if union == 0:
            return 0.0

        return intersection / union

    def prefix_match_score(self, text1, text2):
        """
        前方一致スコアを計算

        先頭から何文字一致するかを、短い方の文字数で割る

        Args:
            text1: テキスト1
            text2: テキスト2

        Returns:
            float: 前方一致スコア (0.0 ~ 1.0)
        """
        if not text1 or not text2:
            return 0.0

        min_len = min(len(text1), len(text2))
        if min_len == 0:
            return 0.0

        # 先頭から何文字一致するか
        match_count = 0
        for c1, c2 in zip(text1, text2):
            if c1 == c2:
                match_count += 1
            else:
                break

        return match_count / min_len

    def edit_distance_score(self, text1, text2):
        """
        レーベンシュタイン距離をスコアに変換

        score = 1 - distance / max_length

        Args:
            text1: テキスト1
            text2: テキスト2

        Returns:
            float: 編集距離スコア (0.0 ~ 1.0)
        """
        if not text1 or not text2:
            return 0.0

        max_len = max(len(text1), len(text2))
        if max_len == 0:
            return 1.0

        dist = levenshtein_distance(text1, text2)
        return 1.0 - (dist / max_len)

    def calculate_similarity(self, text1, text2):
        """
        2つのテキストの類似度を計算

        最終スコア = ngram_weight * N-gram類似度
                   + prefix_weight * 前方一致スコア
                   + edit_weight * 編集距離スコア

        Args:
            text1: テキスト1
            text2: テキスト2

        Returns:
            dict: {
                'score': 最終スコア,
                'ngram_score': N-gram類似度,
                'prefix_score': 前方一致スコア,
                'edit_score': 編集距離スコア,
                'normalized1': 正規化後のtext1,
                'normalized2': 正規化後のtext2
            }
        """
        # 正規化
        norm1, _ = self.normalize(text1)
        norm2, _ = self.normalize(text2)

        # 完全一致の場合
        if norm1 == norm2:
            return {
                'score': 1.0,
                'ngram_score': 1.0,
                'prefix_score': 1.0,
                'edit_score': 1.0,
                'normalized1': norm1,
                'normalized2': norm2
            }

        # 各スコアを計算
        ngram_score = self.ngram_similarity(norm1, norm2)
        prefix_score = self.prefix_match_score(norm1, norm2)
        edit_score = self.edit_distance_score(norm1, norm2)

        # 最終スコア
        final_score = (
            self.ngram_weight * ngram_score +
            self.prefix_weight * prefix_score +
            self.edit_weight * edit_score
        )

        return {
            'score': final_score,
            'ngram_score': ngram_score,
            'prefix_score': prefix_score,
            'edit_score': edit_score,
            'normalized1': norm1,
            'normalized2': norm2
        }

    def find_candidates(self, target, candidates, top_n=3, threshold=0.0):
        """
        対象テキストに対して、候補リストから類似度の高いものを抽出

        Args:
            target: 対象テキスト
            candidates: 候補リスト
            top_n: 返す候補数（デフォルト: 3）
            threshold: スコアの閾値（デフォルト: 0.0）

        Returns:
            list: [
                {
                    'candidate': 候補テキスト,
                    'score': スコア,
                    'details': 詳細情報
                },
                ...
            ]
        """
        results = []

        for candidate in candidates:
            similarity = self.calculate_similarity(target, candidate)

            if similarity['score'] >= threshold:
                results.append({
                    'candidate': candidate,
                    'score': similarity['score'],
                    'details': similarity
                })

        # スコアの降順でソート
        results.sort(key=lambda x: x['score'], reverse=True)

        return results[:top_n]


# pandasがインポートされていない場合の対応
try:
    import pandas as pd
except ImportError:
    class pd:
        @staticmethod
        def isna(x):
            return x is None or (isinstance(x, float) and x != x)
