# sales-email-automation
Google Sheets + GAS + OpenAI APIを使った営業メール自動生成システム

# 営業メール自動生成システム

Google Sheets + GAS + OpenAI APIを使った営業メール自動化ツール

## 概要
スプレッドシートの営業リストをもとに、GPT-4oがパーソナライズされた
営業メールを自動生成してD列に書き込みます。

## 使用技術
- Google Apps Script
- OpenAI API（gpt-4o）
- Google Sheets

## 機能
- 会社名・URLをもとに営業メールを自動生成
- 実行ログをE列に自動記録
- エラーが出ても次の行に進む安全設計
- スプレッドシートのメニューから1クリック実行

## 環境構築
1. スプレッドシートを新規作成
2. 拡張機能→Apps Scriptでコードを貼り付け
3. スクリプトプロパティにOPENAI_API_KEYを登録
4. スプレッドシートをリロードしてメニューから実行

## 動作イメージ
| A列 | B列 | C列 | D列 | E列 |
|---|---|---|---|---|
| 会社名 | URL | 担当者名 | 生成メール | ログ |
