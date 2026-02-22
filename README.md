# 納入台帳テキスト解析ツール (rea)

概要
- 貼り付けた納入台帳のテキストを解析し、明細行（No / 品名 / 規格 / 単位 / 数量 / 単価 / 金額 / 備考）と最終ページ集計（課税対象額 / 消費税 / 合計）を抽出・表示・編集・印刷プレビューするツール群。

主な機能
- 貼り付けテキストから明細と集計を抽出（casks/ledger_paste_parser.js）
- 印刷レイアウト生成（ヘッダー、ページ分割、ページ小計、フッター） （casks/print.js）
- 業者情報自動補完（gyousya.txt を参照）および反映（casks/gyousya.js）
- SharePoint連携用スケルトン（js/sp_*.js）

主要ファイル
- index.html：アプリのエントリ。UIとの連携とデータ保持（currentRows, originalSummary 等）。
- casks/ledger_paste_parser.js：貼り付けテキスト解析のコア。
- casks/print.js：印刷プレビュー生成ロジック。
- casks/gyousya.js：gyousya.txt から業者情報を読み込みヘッダーに反映。
- gyousya.txt：業者データ（テキスト形式）。
- js/sp_base.js, js/sp_api.js, js/sp_db.js：SharePoint 連携の雛形（未実装部分あり）。

利用手順（簡易）
1. index.html をブラウザで開く。  
2. 納入台帳の該当テキストを貼り付け。  
3. 解析ボタンで rows と summary を生成・確認。  
4. 必要に応じて業者情報を gyousya.txt に登録し反映。  
5. 印刷プレビューでレイアウトを確認し出力。

注意点
- 金額・数量は原本表記を優先して抽出する設計。エッジケースは ledger_paste_parser.js のルールを調整してください。
- SharePoint 連携を使う場合は js/sp_*.js を実装・設定してください。

ライセンス・連絡
- 本リポジトリにライセンスファイルがある場合はそちらを参照してください。問題や要望はリポジトリの issue に記載してください。
