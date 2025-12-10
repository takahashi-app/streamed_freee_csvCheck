# STREAMED→freee会計インポート用csv修正アプリ

STREAMEDからのCSVをfreee会計へインポートするための前処理を行うWebアプリケーションです。

## セキュリティ機能

- **パスワード認証**: アプリ起動時にパスワードの入力が必要です（エンターキーでログイン可能）
- **Streamlit Secrets**: パスワードはコードに直接記載せず、Streamlit Secretsで管理します

## 機能概要

### ステージ1: 初回処理
1. STREAMED CSVとfreee仕訳帳CSVをアップロード
2. 列名の変更（借方補助科目→借方取引先、貸方補助科目→貸方取引先）
3. 伝票番号を自動生成（複合仕訳対応、月日時分+連番3桁、例：12081508001）
4. 取引先名・部門名の表記ゆれをチェック
5. 候補付きExcelファイルを出力（色分け付き）
6. 出力したExcelファイルを自動で開く

### ステージ2: freeeインポート用CSV生成
1. 目視確認後のExcelファイルをアップロード
2. 候補1を自動適用
3. freeeインポート用CSVを生成 + 編集前・編集後の2シート構成Excelを生成
4. 出力したファイルを自動で開く

## 表記ゆれチェックロジック

### 正規化処理
1. Unicode正規化（NFKC）
2. 全角→半角変換
3. カタカナ→ひらがな統一
4. 大文字→小文字変換
5. 法人格の除去（株式会社、(株)、㈱、有限会社、LLC など）
6. 記号の除去（×、・、／、- など）

### スコアリング
最終スコア = 0.5 × N-gram類似度 + 0.3 × 前方一致スコア + 0.2 × 編集距離スコア

- **N-gram類似度**: 2-gramのJaccard類似度
- **前方一致スコア**: 先頭から何文字一致するか ÷ 短い方の文字数
- **編集距離スコア**: 1 - レーベンシュタイン距離 / 最大文字数

### Excel出力の色分け

#### 行の色分け
- **緑**: 完全一致
- **赤**: 不一致で候補なしor80点未満
- **色なし**: 不一致で80点以上の候補あり

#### 列の色分け
- **黄色**: 取引先候補列
- **青色**: 部門候補列

## インストール

### 必要要件
- Python 3.8以上

### セットアップ
```bash
# リポジトリをクローン（またはダウンロード）
cd freee_impt_streaamed

# 必要なパッケージをインストール
pip install -r requirements.txt
```

## 使い方

### アプリの起動
```bash
streamlit run app.py
```

ブラウザが自動的に開き、アプリが起動します。
（開かない場合は、http://localhost:8501 にアクセス）

### ステージ1: 初回処理

1. **STREAMED CSV**をアップロード
2. **freee仕訳帳CSV（新方式）**をアップロード（複数ファイル可）
3. **処理を実行**ボタンをクリック
4. 処理結果を確認
5. **デスクトップへ出力**または**指定フォルダへ出力**でExcelファイルを保存

### ステージ2: freeeインポート用CSV生成

1. サイドバーで「ステージ2」を選択
2. 目視確認後の**Excelファイル**をアップロード
3. **freeeインポート用CSV生成**ボタンをクリック
4. **デスクトップへ出力**または**指定フォルダへ出力**でCSVファイルを保存

## ファイル構成

```
freee_impt_streaamed/
├── app.py                      # Streamlitメインアプリ
├── requirements.txt            # 必要パッケージ
├── README.md                   # このファイル
├── utils/
│   ├── __init__.py
│   ├── csv_processor.py        # CSV読み込み・処理
│   ├── name_matcher.py         # 表記ゆれチェック
│   └── excel_writer.py         # Excel出力
└── sample/
    ├── streamed_20251208143118.csv
    └── 仕訳帳（新）CSV （2025年06月~2026年05月）.csv
```

## 注意事項

- CSVファイルのエンコーディングは自動検出されます（CP932/Shift-JIS/UTF-8対応）
- freee仕訳帳CSVは**新方式**のみ対応しています
- 伝票番号は処理実行時の日時を基準に生成されます
- 表記ゆれチェックの重み（0.5, 0.3, 0.2）は`utils/name_matcher.py`で変更可能です

## トラブルシューティング

### CSVが読み込めない
- ファイルのエンコーディングを確認してください（CP932/Shift-JIS推奨）
- ファイルが破損していないか確認してください

### 取引先/部門が見つからない
- freee仕訳帳CSVに該当のマスタデータが含まれているか確認してください
- 過年度分のCSVも含めてアップロードすることをお勧めします

### アプリが起動しない
- Python 3.8以上がインストールされているか確認してください
- `pip install -r requirements.txt`を実行して、必要なパッケージがインストールされているか確認してください

## Streamlit Cloudへの展開

### 1. GitHubへのプッシュ
```bash
git add .
git commit -m "Initial commit"
git push origin main
```

**注意**: `.gitignore`により、`.streamlit/secrets.toml`はGitHubにプッシュされません。

### 2. Streamlit Cloudでの設定

1. [Streamlit Cloud](https://streamlit.io/cloud)にログイン
2. リポジトリを選択して展開
3. **Settings** → **Secrets** に移動
4. 以下の内容を追加:

```toml
[passwords]
system_password = "your_password_here"
```

### 3. ローカル開発での設定

ローカルで開発する場合は、`.streamlit/secrets.toml`ファイルが既に作成されています。
このファイルは`.gitignore`に含まれており、GitHubにはプッシュされません。

## ライセンス

MIT License

## 開発者

Created with Claude Code
