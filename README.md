# eBay出品管理ツール

Google Apps Script (GAS) を使用したeBay出品データの重複検出・管理ツールです。

## バージョン管理システム

このプロジェクトは統一されたバージョン管理システムを採用しています。

### 仕組み
- **単一の真実の源**: `version-config.js`でバージョンを一元管理
- **自動同期**: 全ファイルが同一バージョンを参照
- **Google Apps Script最適化**: claspと連携した効率的な管理

### バージョン更新手順
1. `version-config.js`のPROJECT_VERSIONを更新
2. `clasp push`でGoogle Apps Scriptに反映
3. 変更をコミット

## 🎉 最新の改善点 (v1.6.33)

### ✅ **修正完了した主要問題**
- **CONFIG参照エラー解決**: 動的バッチサイズ処理でのエラーを修正
- **ゼロ重複ケースの適切な処理**: 重複がない場合もエラーではなく正常完了として処理
- **UI重複表示の統一**: 処理中スピナーと完了メッセージを一元化
- **パフォーマンス最適化**: データサイズに応じた動的バッチサイズ計算を実装

### 🔧 **技術的改善**
```javascript
// 動的バッチサイズ計算の実装
calculateOptimalBatchSize: function(dataSize) {
  if (dataSize <= 1000) return 500;
  if (dataSize <= 5000) return 1000;
  if (dataSize <= 20000) return 2000;
  if (dataSize <= 40000) return 3000;
  return 4000; // 40,000行以上の場合
}
```

## 概要

このツールは、eBayの出品データを効率的に管理し、重複する商品タイトルを検出してCSVエクスポートを行うことで、出品作業の効率化を支援します。大量データ（最大40,000行、10MB）の処理に対応した高性能な分析機能を提供しています。

## 主要機能

- 🔍 **高度な重複検出**: 改良されたタイトル正規化による重複商品の自動検出
- 📊 **大量データ処理**: 40,000行のデータを効率的に処理（動的バッチ処理対応）
- 📋 **CSVインポート/エクスポート**: eBayデータの一括インポートと重複データの出力
- 🎯 **高精度分析**: Jaccard係数を使用した類似度計算
- 📈 **リアルタイム分析**: 統一されたサイドバーUIによる直感的な操作
- 🚀 **パフォーマンス最適化**: タイムアウト対策とメモリ効率化
- ⚡ **エラーハンドリング強化**: ゼロ重複ケースでも適切に処理完了

## ファイル構成

```
├── Code.js              # メインのGASコード（2,797行）
├── Sidebar.html         # サイドバーUI（3,338行）
├── appsscript.json      # GASプロジェクト設定
├── .clasp.json          # clasp設定ファイル
└── README.md           # このファイル
```

**削除済みファイル**: settings.html（未使用のため削除）

## シート構成

ツールは以下のシートを自動作成・管理します：

- **インポートデータ**: eBayからエクスポートしたCSVデータを格納
- **重複リスト**: 検出された重複商品の一覧
- **エクスポート**: 処理結果の出力
- **分析**: データ分析結果とサマリー
- **ログ**: 処理履歴とエラーログ

## 使用方法

### 基本的な使い方

1. **スプレッドシートを開く**
   - 対象のスプレッドシートを Google Sheets で開きます

2. **管理ツールを起動**
   - メニューバーから「eBay出品管理」→「管理ツールを開く」を選択
   - サイドバーに管理UIが表示されます

3. **自動処理を実行**
   - eBayから出品データをCSV形式でエクスポート
   - サイドバーの自動処理タブでファイルをドラッグ&ドロップ
   - 「自動処理を開始」ボタンをクリック

4. **結果を確認**
   - 重複データが自動的にCSVファイルとしてダウンロードされます
   - 重複が0件の場合も「検出された重複: 0件」として正常完了

## 対応データ形式

### eBayエクスポートCSV列構成
```
Item number, Title, Variation details, Custom label (SKU), 
Available quantity, Format, Currency, Start price, 
Auction Buy It Now price, Reserve price, Current price, 
Sold quantity, Watchers, Bids, Start date, End date, 
eBay category 1 name, eBay category 1 number, 
eBay category 2 name, eBay category 2 number, Condition, 
CD:Professional Grader, CD:Grade, CDA:Certification Number, 
CD:Card Condition, eBay Product ID(ePID), Listing site, 
P:UPC, P:EAN, P:ISBN
```

## ✅ **解決済み問題**

### 🚨 **以前のエラー例（修正済み）**

#### 修正前のエラー
```
2025/05/30 21:10:07 自動処理（エクスポート失敗） 失敗 0 23433件
CSVエクスポートに失敗したため、処理を中止しました。
エラー: 重複リストにデータがありません。
```

#### 修正後の動作
```
2025/05/30 22:10:07 自動処理（正常完了） 成功 0 23433件
23,433件のデータを分析した結果、重複する商品は見つかりませんでした。
```

### 📝 **改良された重複検出パターン**

v1.6.1では以下のようなパターンも適切に検出できるよう改善されています：

```
例1: 括弧の有無
Used  () SUZUKI Electric Taisho Koto Katsura TAS 11
Used     SUZUKI Electric Taisho Koto Katsura TAS 11
→ 改良された正規化処理により類似商品として検出

例2: 空白の違い
Saint Cloth Myth EX Model number  Cancer Death Mask (  ) BANDAI
Saint Cloth Myth EX Model number  Cancer Death Mask BANDAI
→ 連続空白の統一により検出可能

例3: 記号と大文字小文字の違い
Somewhat difficult          Pinku  MDD nail hand sw skin
Somewhat difficult           .  Pinku  MDD Nail Hand sw skin
→ 記号除去と大文字小文字統一により検出可能
```

### 🛠️ **実装済み改善**

#### 1. **エラーハンドリングの完全改善** ✅
```javascript
// ゼロ重複ケースの適切な処理
if (duplicateCount === 0) {
  result.finalMessage = `処理が完了しました: ${importedRows}件のデータを分析した結果、重複する商品は見つかりませんでした。`;
  return { success: true, duplicateCount: 0, analysisComplete: true };
}
```

#### 2. **重複検出アルゴリズムの改善** ✅
```javascript
// 改良されたノーマライズ処理
normalizeTitle: function(title, useAdvanced = false) {
  // 基本的な正規化
  let normalized = String(title).toLowerCase();
  
  if (useAdvanced) {
    // 連続する空白を単一空白に統一
    normalized = normalized.replace(/\s+/g, ' ');
    // 一般的な記号を除去
    normalized = normalized.replace(/[().\\-_]/g, '');
    // 前後空白を除去
    normalized = normalized.trim();
  }
  
  return normalized;
}
```

#### 3. **動的パフォーマンス最適化** ✅
```javascript
// データサイズに応じた最適バッチサイズ計算
calculateOptimalBatchSize: function(dataSize) {
  if (dataSize <= 1000) return 500;
  if (dataSize <= 5000) return 1000;
  if (dataSize <= 20000) return 2000;
  if (dataSize <= 40000) return 3000;
  return 4000; // 40,000行以上の場合
}
```

## 技術仕様

### パフォーマンス設定
```javascript
// 動的バッチサイズ（データサイズに応じて自動調整）
BATCH_SIZE: 2000,           // 基準バッチ処理サイズ
MAX_LOG_ROWS: 500,          // ログ最大行数
SIMILARITY_THRESHOLD: 0.7,  // 類似度閾値
MAX_FILE_SIZE: 10,          // 最大ファイルサイズ(MB)

// 動的バッチサイズ計算機能追加
calculateOptimalBatchSize: function(dataSize) {
  // データサイズに応じて500～4000の範囲で自動調整
}
```

### 類似度計算
- **アルゴリズム**: Jaccard係数
- **前処理**: 改良されたテキストノーマライズ、ストップワード除去
- **閾値**: 0.7（調整可能）

### UI改善
- **統一メッセージシステム**: 処理中スピナーと完了メッセージを一元化
- **バージョン表示統一**: Code.js と Sidebar.html で v1.6.1 に統一
- **キャンセル機能**: 長時間処理のキャンセル対応

## セットアップ

### 前提条件
- Google アカウント
- Node.js (clasp使用の場合)
- clasp CLI ツール

### インストール手順

1. **clasp のインストール**
   ```bash
   npm install -g @google/clasp
   ```

2. **プロジェクトのクローン**
   ```bash
   clasp clone 1dzO7vQoTI2ywXDwpzgr5ZKb6ZT3cQINZydpPlsvxOHWA2jDq24SzVGm9
   ```

3. **ローカル同期**
   ```bash
   clasp pull
   ```

## 開発・デバッグ

### ローカル開発
```bash
# ファイル編集後のプッシュ
clasp push --force

# ログの確認
clasp logs

# プロジェクト状況確認
clasp status
```

### 主要関数

#### データ処理
- `autoProcessEbayData(csvData)`: 自動処理のメイン関数（v1.6.1で改良）
- `importCsvData(csvData)`: CSVデータのインポート
- `validateData(data, requiredColumns)`: データ検証
- `detectDuplicates(data, threshold)`: 重複検出（エラーハンドリング改善）

#### UI制御
- `showSidebar()`: サイドバー表示
- `showUnifiedMessage(message, type)`: 統一メッセージ表示（v1.6.1で追加）
- `updateProgress(progress)`: 進行状況更新

#### ユーティリティ
- `normalizeTitle(title, useAdvanced)`: 改良されたタイトル正規化
- `calculateSimilarity(title1, title2)`: 類似度計算
- `calculateOptimalBatchSize(dataSize)`: 動的バッチサイズ計算（v1.6.1で追加）

## トラブルシューティング

### 解決済み問題

1. **「重複リストにデータがありません」エラー** ✅ **修正済み**
   ```
   原因: 重複検出で0件の場合のエラーハンドリング不備
   対策: v1.6.1で正常完了メッセージを表示するよう修正
   結果: 「検出された重複: 0件」として正常完了
   ```

2. **「CONFIG is not defined」エラー** ✅ **修正済み**
   ```
   原因: 動的バッチサイズ計算でのCONFIG参照エラー
   対策: 適切なスコープでCONFIG取得を実装
   結果: 動的バッチサイズが正常に動作
   ```

3. **UI重複表示の問題** ✅ **修正済み**
   ```
   原因: 複数の処理中表示とメッセージが重複
   対策: 統一メッセージシステムに集約
   結果: クリーンで一貫したUI体験
   ```

### 現在も有効なトラブルシューティング

1. **処理がタイムアウトする**
   ```
   原因: 大量データ処理時の実行時間超過
   対策: v1.6.1の動的バッチサイズで改善済み
   ```

2. **メモリ不足エラー**
   ```
   原因: 40,000行データの一括読み込み
   対策: 分割処理とメモリ効率化を実装済み
   ```

### デバッグ方法

1. **実行ログの確認**
   ```bash
   clasp logs --watch
   ```

2. **GASエディタでのデバッグ**
   - スクリプトエディタで直接実行
   - console.log()でデバッグ情報を出力

3. **パフォーマンス監視**
   - 実行時間の測定
   - メモリ使用量の監視

## 変更履歴

### v1.6.1 (2024-12-30) - 緊急修正版
- **修正**: CONFIG参照エラーの解決
- **改善**: ゼロ重複ケースの適切な処理実装
- **改善**: UI重複表示の統一（統一メッセージシステム）
- **新機能**: 動的バッチサイズ計算
- **削除**: 未使用のsettings.htmlファイル
- **最適化**: パフォーマンス改善とエラーハンドリング強化

### v1.6.0 (2024-12-29)
- 高度な重複検出アルゴリズム実装
- 大量データ処理対応
- サイドバーUI改善

## 今後の開発予定

### Phase 1: 安定化 ✅ **完了**
- [x] エラーハンドリング改善
- [x] 重複検出精度向上
- [x] パフォーマンス最適化
- [x] UI統一化

### Phase 2: 機能拡張 (1-2ヶ月)
- [ ] 手動重複確認機能
- [ ] 一括編集機能
- [ ] レポート機能強化
- [ ] 詳細ログ機能

### Phase 3: 長期改善 (3-6ヶ月)
- [ ] 機械学習による自動分類
- [ ] API連携機能
- [ ] 自動化スケジューリング

## プロジェクト情報

- **スクリプトID**: `1dzO7vQoTI2ywXDwpzgr5ZKb6ZT3cQINZydpPlsvxOHWA2jDq24SzVGm9`
- **バージョン**: v1.6.1
- **最終更新**: 2024-12-30
- **タイムゾーン**: Asia/Tokyo
- **ランタイム**: V8

## 権限

このツールは以下の Google API スコープを使用します：

- `https://www.googleapis.com/auth/spreadsheets`: スプレッドシートの読み書き
- `https://www.googleapis.com/auth/script.storage`: 設定データの保存
- `https://www.googleapis.com/auth/drive.file`: ファイルの作成・管理

## ライセンス

このプロジェクトは MIT ライセンスの下で公開されています。

---

**✅ 重要**: v1.6.1では主要なエラーが修正され、40,000行を超える大量データも安定して処理できるようになりました。動的バッチサイズ計算により、Google Apps Scriptの実行時間制限（6分）内での処理が最適化されています。 