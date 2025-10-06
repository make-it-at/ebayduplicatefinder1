# eBay Duplicate Finder Tool - Claude Memory

## Project Overview
eBay出品管理ツールの重複商品検出機能を持つGoogle Apps Scriptベースのプロジェクト

## Current Status
- Version: v1.6.33
- Last update: バージョン管理統一とUS絞り込み最適化
- Files: version-config.js, Code.js, Sidebar.html

## Key Commands
- Testing: (TBD - need to check codebase)
- Linting: (TBD - need to check codebase)  
- Deployment: (TBD - need to check codebase)

## Architecture Notes
- Google Apps Script project
- Sidebar HTML interface
- Main logic in Code.js

## Development Notes
- Recent changes focus on permission error handling
- Uses HTML sidebar for user interaction
- v1.6.3: Added eBay Mag support - US-only filtering for duplicate detection (Code.js:1040-1060)
- v1.6.4: Performance optimization - Direct sheet deletion for large data processing (Code.js:1042-1076)
- v1.6.5: Two-step processing - Added standalone US filtering function (Code.js:1147-1230, Sidebar.html:1144-1721)
- v1.6.6-v1.6.7: UI debugging - Fixed button activation issues
- v1.6.8: Performance optimization - Using Google Sheets standard filter for bulk row deletion (Code.js:1141-1178)
- v1.6.9: API対応 - setVisibleValues() の廃止対応でsetHiddenValues()使用に変更
- v1.6.10: 統一ボタン状態管理システム実装 - 複数FileReader競合問題解決
- v1.6.11: 手作業相当処理実装 - 一括データ取得→連続範囲削除でタイムアウト解決 (Code.js:1141-1184)
- v1.6.12: ボタン有効化問題修正 - updateUIState関数でUS絞り込みボタン状態を適切に設定 (Sidebar.html:2125-2140)
- v1.6.13: インポートのみボタン機能追加 - CSVデータをインポートのみ実行する機能を追加 (Sidebar.html:1150-1807)
- v1.6.14: checkAppState関数追加 - インポート後のUI状態更新エラーを修正 (Code.js:1220-1287)
- v1.6.15: US絞り込みボタン単独実行機能追加 - インポート済みデータでも単独でUS絞り込みが可能に (Sidebar.html:1733-1831)
- v1.6.16: 統一ボタン状態管理システム導入 - 繰り返し発生していたボタン無効化問題を根本解決 (Sidebar.html:1216-1333)
- v1.6.17: タイムアウト回避手法実装 - 重複検出フェーズにマイクロチャンク処理(1500行/chunk)、安全マージン60秒拡張、ポーズ間隔1秒短縮でタイムアウト問題解決 (Code.js:3483-3690, Sidebar.html:2053-2155)
- v1.6.18: タイムアウト問題緊急修正 - 状態の3重永続化(メモリ+キャッシュ+シート)、チャンクサイズ800行に縮小、処理状態シート追加で状態消失問題を完全解決 (Code.js:3186-3350)
- v1.6.19: 状態管理デバッグ強化 - 状態保存・取得の詳細ログ追加、シート永続化の実行状況を完全可視化、状態消失の根本原因特定のための診断機能追加 (Code.js:3272-3370)