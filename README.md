# Kintone-Excel連携システム

Kintoneからレコードを取得し、テンプレートExcelに反映して個別のExcelファイルを生成するシステム

## 機能概要

- Kintoneからレコードのダウンロード
- レコードの選択
- テンプレートExcelへのデータ反映
- 個別Excelファイルの生成

## 開発環境

- VBA（Visual Basic for Applications）
- Excel
- Kintone API

## プロジェクト構成

### モジュール構成

1. `KintoneAPI.bas` - Kintone APIとの通信処理
2. `ExcelOperation.bas` - Excel操作関連の処理
3. `FileOperation.bas` - ファイル操作関連の処理
4. `ErrorHandling.bas` - エラー処理とログ出力
5. `Main.bas` - メイン処理の制御

## 開発状況

現在の開発状況は[Issues](https://github.com/iyell/dandori-vba/issues)で確認できます。

## セットアップ方法

1. リポジトリのクローン
2. Excelファイルの設定
3. Kintone APIトークンの設定
4. テンプレートExcelの配置 