/**
 * eBay出品管理ツール v1.6.1
 * スプレッドシート上でeBayの重複出品を検出・管理するツール
 * 最終更新: 2024-12-30 - CONFIG参照エラーの修正と動的バッチサイズ最適化
 */

// EbayTool名前空間 - 拡張版
var EbayTool = (function() {
  // プライベート変数と定数
  const CONFIG = {
    VERSION: '1.6.2',
    SHEET_NAMES: {
      IMPORT: 'インポートデータ',
      DUPLICATES: '重複リスト',
      EXPORT: 'エクスポート',
      ANALYSIS: '分析',
      LOG: 'ログ'
    },
    COLORS: {
      PRIMARY: '#4F46E5',
      SECONDARY: '#6366F1',
      SUCCESS: '#10B981',
      WARNING: '#F59E0B',
      ERROR: '#EF4444',
      BACKGROUND: '#F9FAFB',
      TEXT: '#374151'
    },
    BATCH_SIZE: 2000, // データ処理時の一度に読み込む行数
    MAX_LOG_ROWS: 500, // ログの最大行数
    SIMILARITY_THRESHOLD: 0.7, // タイトル類似度の閾値
    MAX_FILE_SIZE: 10, // CSVファイルの最大サイズ（MB）
    
    // 動的バッチサイズ計算
    calculateOptimalBatchSize: function(dataSize) {
      if (dataSize <= 1000) return 500;
      if (dataSize <= 5000) return 1000;
      if (dataSize <= 20000) return 2000;
      if (dataSize <= 40000) return 3000;
      return 4000; // 40,000行以上の場合
    }
  };
  
  // プライベート関数
  function getConfig() {
    return CONFIG;
  }
  
  // 各機能モジュール
  const UI = {
    showMessage: function(message, type = 'info') {
      return {
        message: message,
        type: type
      };
    },
    
    formatSheetHeader: function(range) {
      range.setBackground(CONFIG.COLORS.PRIMARY)
           .setFontColor('white')
           .setFontWeight('bold');
    }
  };
  
  const DataProcessor = {
    batchProcess: function(data, processorFn, batchSize = CONFIG.BATCH_SIZE) {
      return new Promise((resolve, reject) => {
        try {
          const results = [];
          const totalBatches = Math.ceil(data.length / batchSize);
          let processedBatches = 0;
          
          // バッチ単位で非同期処理
          const processBatch = function(startIndex) {
            const endIndex = Math.min(startIndex + batchSize, data.length);
            const batch = data.slice(startIndex, endIndex);
            
            try {
              const batchResults = processorFn(batch, startIndex);
              results.push(...batchResults);
              
              processedBatches++;
              const progress = (processedBatches / totalBatches) * 100;
              
              // まだ処理すべきバッチがある場合
              if (endIndex < data.length) {
                // 少し遅延を入れて次のバッチを処理（UIブロッキング防止）
                setTimeout(function() {
                  processBatch(endIndex);
                }, 10);
              } else {
                // すべてのバッチ処理が完了
                resolve(results);
              }
            } catch (error) {
              reject(error);
            }
          };
          
          // 最初のバッチから処理開始
          processBatch(0);
        } catch (error) {
          reject(error);
        }
      });
    }
  };
  
  const TextAnalyzer = {
    normalizeTitle: function(title, useAdvanced = false) {
      if (!title) return '';
      
      // 基本的なノーマライズ
      let normalized = String(title).toLowerCase();
      
      // 改良版の正規化プロセス
      if (useAdvanced) {
        // 高度な類似度計算用の正規化
        
        // 1. 連続する空白を単一空白に統一
        normalized = normalized.replace(/\s+/g, ' ');
        
        // 2. 一般的な記号を除去（ただし重要な区別要素は保持）
        normalized = normalized.replace(/[().\-_]/g, '');
        
        // 3. 一般的な略語や単位を標準化
        const replacements = {
          'in.': 'inch',
          'inches': 'inch',
          'ft.': 'foot',
          'feet': 'foot',
          'lbs': 'pound',
          'lb.': 'pound',
          'pounds': 'pound',
          'oz.': 'ounce',
          'ounces': 'ounce',
          'pcs': 'piece',
          'pc.': 'piece',
          'pieces': 'piece'
        };
        
        for (const [abbr, full] of Object.entries(replacements)) {
          normalized = normalized.replace(new RegExp('\\b' + abbr + '\\b', 'g'), full);
        }
        
        // 4. 重要なキーワードのみを抽出
        const words = normalized.split(' ');
        
        // ストップワード（無視する一般的な単語）
        const stopWords = ['a', 'an', 'the', 'and', 'or', 'but', 'is', 'are', 'on', 'at', 'to', 'for', 'with', 'by', 'in', 'of'];
        
        // ストップワードを除去し、短すぎる単語も除去
        const filteredWords = words.filter(word => 
          word.length > 2 && !stopWords.includes(word)
        );
        
        // 単語をソートして順序の違いを無視
        filteredWords.sort();
        
        // 出現回数が多い単語を強調
        const wordCounts = {};
        filteredWords.forEach(word => {
          wordCounts[word] = (wordCounts[word] || 0) + 1;
        });
        
        // 重要度に基づいて単語を選択
        const importantWords = filteredWords.filter((word, index, self) => 
          // 重複を除去
          index === self.indexOf(word) && (
            // 出現回数が多いか、長い単語は重要
            wordCounts[word] > 1 || word.length > 5
          )
        );
        
        // 十分な単語がない場合は全単語を使用
        if (importantWords.length < 3) {
          return filteredWords.join(' ');
        }
        
        normalized = importantWords.join(' ');
      } else {
        // 基本的な重複検出用の正規化（改良版）
        
        // 1. 連続する空白を単一空白に統一
        normalized = normalized.replace(/\s+/g, ' ');
        
        // 2. 空の括弧や意味のない記号パターンを除去
        normalized = normalized.replace(/\(\s*\)/g, '');  // 空の括弧
        normalized = normalized.replace(/\[\s*\]/g, '');  // 空の角括弧
        normalized = normalized.replace(/\s+\.\s+/g, ' '); // 独立したドット
        
        // 3. 記号の前後の余分な空白を整理
        normalized = normalized.replace(/\s*([().\-_])\s*/g, '$1');
        
        // 4. 最終的な空白整理とトリム
        normalized = normalized.replace(/\s+/g, ' ').trim();
        
        // 5. 特殊文字を統一（残す文字：英数字、空白、基本的な記号）
        normalized = normalized.replace(/[^\w\s().\-_]/g, '');
      }
      
      return normalized;
    },
    
    calculateSimilarity: function(title1, title2) {
      // 両方のタイトルをノーマライズ
      const normalized1 = this.normalizeTitle(title1, true);
      const normalized2 = this.normalizeTitle(title2, true);
      
      if (!normalized1 || !normalized2) return 0;
      
      // 単語ベースでの類似度計算
      const words1 = normalized1.split(' ');
      const words2 = normalized2.split(' ');
      
      // Jaccard係数の計算
      const set1 = new Set(words1);
      const set2 = new Set(words2);
      
      const intersection = new Set([...set1].filter(word => set2.has(word)));
      const union = new Set([...set1, ...set2]);
      
      return intersection.size / union.size;
    }
  };
  
  const CSVHandler = {
    parse: function(csvData) {
      try {
        // 引数チェック
        if (!csvData || typeof csvData !== 'string') {
          console.error("CSVデータが無効です");
          return [];
        }

        // 改行コードを統一
        csvData = csvData.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        
        // BOMの除去
        if (csvData.charCodeAt(0) === 0xFEFF) {
          csvData = csvData.substring(1);
          console.log("BOMを検出して削除しました");
        }
        
        // 行分割
        const lines = csvData.split('\n');
        const rows = [];
        
        // データがない場合は空配列を返す
        if (lines.length === 0) {
          console.warn("CSVデータに行がありません");
          return [];
        }
        
        // ヘッダー行を処理
        const header = this.parseLine(lines[0]);
        if (header && header.length > 0) {
          rows.push(header);
          
          // メイン処理 - バッチ処理で高速化
          const batchSize = 500;
          for (let i = 1; i < lines.length; i += batchSize) {
            // 各バッチの行をフィルタリングして処理
            const batch = lines.slice(i, i + batchSize)
              .filter(line => line && line.trim()) // 空行を除外
              .map(line => {
                try {
                  return this.parseLine(line);
                } catch (e) {
                  console.warn(`行 ${i} のパースに失敗: ${e.message}`);
                  // エラー時は単純な分割を使用
                  return line.split(',');
                }
              });
            
            rows.push(...batch);
          }
        } else {
          console.error("ヘッダー行のパースに失敗しました");
        }
        
        return rows;
      } catch (error) {
        console.error("CSVパース処理エラー:", error);
        // 最後の手段として、より単純なパース方法を試す
        try {
          return this.simpleParse(csvData);
        } catch (fallbackError) {
          console.error("フォールバックパースにも失敗:", fallbackError);
          // 最悪の場合、空配列を返す（エラーは投げない）
          return [];
        }
      }
    },
    
    // シンプルなCSVパース方法（エラー時のフォールバック）
    simpleParse: function(csvData) {
      try {
        if (!csvData || typeof csvData !== 'string') {
          return [];
        }
        
        // 改行コードを統一
        csvData = csvData.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        
        // 単純な行と列の分割
        return csvData.split('\n')
          .filter(line => line && line.trim())
          .map(line => line.split(',').map(cell => cell ? cell.trim() : ''));
      } catch (e) {
        console.error("シンプルパースも失敗:", e);
        // すべての手段が失敗した場合は空配列を返す
        return [];
      }
    },
    
    parseLine: function(line) {
      // 高速なCSVパース処理を実装
      try {
        // 空行や無効な入力をチェック
        if (!line || typeof line !== 'string' || !line.trim()) {
          return [];
        }
        
        // 引用符がない場合は単純に分割する（高速）
        if (!line.includes('"')) {
          return line.split(',').map(cell => cell ? cell.trim() : '');
        }
        
        // 引用符がある場合のみ複雑な処理を行う
        const result = [];
        let inQuotes = false;
        let currentValue = '';
        
        for (let i = 0; i < line.length; i++) {
          const char = line[i];
          
          if (char === '"') {
            // 引用符が連続している場合はエスケープされた引用符
            if (i + 1 < line.length && line[i + 1] === '"') {
              currentValue += '"';
              i++; // 次の引用符をスキップ
            } else {
              // 引用符の中にいるかどうかを切り替え
              inQuotes = !inQuotes;
            }
          } else if (char === ',' && !inQuotes) {
            // 引用符の外でカンマが来たら分割
            result.push(currentValue.trim());
            currentValue = '';
          } else {
            // それ以外の文字は現在の値に追加
            currentValue += char;
          }
        }
        
        // 最後の値を追加
        result.push(currentValue.trim());
        return result;
      } catch (e) {
        console.error('行パースエラー:', e, '対象行:', line);
        // エラーが発生しても処理を続行するため、
        // 可能な限り情報を返す
        if (typeof line === 'string') {
          return line.split(',').map(cell => cell ? cell.trim() : '');
        }
        return [];
      }
    },
    
    generateCSV: function(data) {
      try {
        if (!data || !Array.isArray(data)) {
          console.warn("CSVデータが無効です");
          return '';
        }
        
        return data.map(row => {
          if (!row || !Array.isArray(row)) return '';
          
          return row.map(cell => {
            // null/undefinedの処理
            if (cell === null || cell === undefined) {
              return '';
            }
            
            // 文字列に変換
            let cellStr = String(cell);
            
            // セキュリティ上の問題となりうる文字をエスケープ
            cellStr = cellStr
              .replace(/"/g, '""') // 引用符のエスケープ
              .replace(/\\/g, '\\\\'); // バックスラッシュのエスケープ
            
            // カンマ、引用符、改行、タブを含む場合は引用符で囲む
            if (/[,"\n\r\t]/.test(cellStr)) {
              return '"' + cellStr + '"';
            }
            
            return cellStr;
          }).join(',');
        }).join('\n');
      } catch (error) {
        console.error("CSV生成エラー:", error);
        return '';
      }
    }
  };
  
  const DuplicateDetector = {
    findDuplicates: function(data, titleIndex, itemIdIndex, startDateIndex) {
      return new Promise((resolve, reject) => {
        try {
          // タイトルごとにアイテムをグループ化するオブジェクト
          const titleGroups = {};
          
          // グループ化処理を関数化
          const processRow = function(row) {
            const title = TextAnalyzer.normalizeTitle(row[titleIndex]);
            const itemId = row[itemIdIndex];
            const startDate = row[startDateIndex];
            
            if (title && itemId) {
              if (!titleGroups[title]) {
                titleGroups[title] = [];
              }
              
              titleGroups[title].push({
                itemId: itemId,
                title: title,
                originalTitle: row[titleIndex],
                startDate: startDate,
                allData: row
              });
              
              return true;
            }
            return false;
          };
          
          // データ処理用の高階関数
          const processDataBatch = function(batch) {
            const processedRows = [];
            batch.forEach(row => {
              if (processRow(row)) {
                processedRows.push(row);
              }
            });
            return processedRows;
          };
          
          // 動的バッチサイズを計算してバッチ処理でデータを処理
          const CONFIG = EbayTool.getConfig();
          const optimalBatchSize = CONFIG.calculateOptimalBatchSize(data.length);
          console.log(`データサイズ: ${data.length}行, 最適バッチサイズ: ${optimalBatchSize}`);
          
          DataProcessor.batchProcess(data, processDataBatch, optimalBatchSize)
            .then(() => {
              // 重複グループのみを抽出
              const duplicateGroups = Object.values(titleGroups)
                .filter(group => group.length > 1)
                .sort((a, b) => b.length - a.length);
              
              resolve({
                duplicateGroups: duplicateGroups,
                totalItems: Object.values(titleGroups).reduce((sum, group) => sum + group.length, 0),
                totalGroups: Object.keys(titleGroups).length
              });
            })
            .catch(reject);
        } catch (error) {
          reject(error);
        }
      });
    }
  };
  
  const Logger = {
    log: function(message, severity = 'INFO') {
      console.log(`[${severity}] ${message}`);
      return { message, severity };
    },
    
    error: function(functionName, error, context = '') {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOG);
        
        if (!logSheet) {
          logSheet = ss.insertSheet(CONFIG.SHEET_NAMES.LOG);
          logSheet.appendRow(['タイムスタンプ', '関数', 'エラータイプ', 'エラーメッセージ', 'コンテキスト', 'スタックトレース']);
          
          // ヘッダー行の書式設定
          logSheet.getRange(1, 1, 1, 6)
            .setBackground(CONFIG.COLORS.PRIMARY)
            .setFontColor('white')
            .setFontWeight('bold');
            
          // 列幅の設定
          logSheet.setColumnWidth(1, 150); // タイムスタンプ
          logSheet.setColumnWidth(2, 100); // 関数
          logSheet.setColumnWidth(3, 100); // エラータイプ
          logSheet.setColumnWidth(4, 250); // エラーメッセージ
          logSheet.setColumnWidth(5, 200); // コンテキスト
          logSheet.setColumnWidth(6, 400); // スタックトレース
        }
        
        // エラー情報を追加
        const timestamp = new Date();
        const errorType = error.name || 'Error';
        const errorMessage = error.message || String(error);
        const stackTrace = error.stack || '利用不可';
        
        logSheet.appendRow([timestamp, functionName, errorType, errorMessage, context, stackTrace]);
        
        // エラー行の書式設定
        const lastRow = logSheet.getLastRow();
        
        // エラータイプに応じた色分け
        let bgColor;
        if (errorType.includes('TypeError') || errorType.includes('ReferenceError')) {
          bgColor = CONFIG.COLORS.ERROR + '30'; // より明るい赤
        } else if (errorType.includes('RangeError') || errorType.includes('SyntaxError')) {
          bgColor = CONFIG.COLORS.WARNING + '30'; // より明るい黄色
        } else {
          bgColor = CONFIG.COLORS.ERROR + '20'; // 標準のエラー色
        }
        
        logSheet.getRange(lastRow, 1, 1, 6).setBackground(bgColor);
        
        // ログが長すぎる場合は古いログを削除
        const maxLogRows = CONFIG.MAX_LOG_ROWS;
        if (lastRow > maxLogRows) {
          const deleteCount = Math.min(100, lastRow - maxLogRows);
          logSheet.deleteRows(2, deleteCount);
        }
        
        // コンソールにもエラーを出力
        console.error(`[${functionName}] ${errorType}: ${errorMessage}`);
        if (context) console.error(`Context: ${context}`);
        console.error(stackTrace);
        
        return {
          timestamp: timestamp,
          function: functionName,
          type: errorType,
          message: errorMessage,
          context: context,
          stack: stackTrace
        };
      } catch (e) {
        // ログ記録中のエラーは無視（再帰を防ぐ）
        console.error('Error in logger function:', e);
        return null;
      }
    }
  };
  
  // パブリックAPI
  return {
    UI: UI,
    DataProcessor: DataProcessor,
    TextAnalyzer: TextAnalyzer,
    CSVHandler: CSVHandler,
    getSheetName: function(key) {
      return CONFIG.SHEET_NAMES[key] || '';
    },
    getColor: function(key) {
      return CONFIG.COLORS[key] || '';
    },
    getVersion: function() {
      return CONFIG.VERSION;
    },
    getConfig: getConfig
  };
})();

/**
 * スプレッドシートが開かれたときに実行される関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('eBay出品管理')
    .addItem('管理ツールを開く', 'showSidebar')
    .addToUi();
}

/**
 * サイドバーを表示する関数
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('eBay出品管理ツール')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 入力データを検証する関数
 * @param {Array} data - 検証するデータ配列
 * @param {Array} requiredColumns - 必須列の名前配列
 * @param {Object} validations - 列ごとの検証ルール
 * @return {Object} 検証結果
 */
function validateData(data, requiredColumns, validations = {}) {
  try {
    if (!data || !Array.isArray(data) || data.length === 0) {
      return {
        valid: false,
        errors: ['データが空または無効です'],
        errorCount: 1
      };
    }
    
    // ヘッダー行を取得
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    // 必須列の存在チェック
    const missingColumns = [];
    const columnIndexes = {};
    
    requiredColumns.forEach(col => {
      const index = headers.findIndex(h => h.includes(col.toLowerCase()));
      if (index === -1) {
        missingColumns.push(col);
      } else {
        columnIndexes[col] = index;
      }
    });
    
    if (missingColumns.length > 0) {
      return {
        valid: false,
        errors: [`必須列が見つかりません: ${missingColumns.join(', ')}`],
        errorCount: missingColumns.length
      };
    }
    
    // データ行の検証
    const errors = [];
    const rowErrors = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowError = [];
      
      // 行の長さチェック
      if (row.length !== headers.length) {
        rowError.push(`行 ${i+1}: 列数が不正です（ヘッダー: ${headers.length}列, 行: ${row.length}列）`);
      }
      
      // 各列の検証ルールをチェック
      for (const [column, rules] of Object.entries(validations)) {
        const colIndex = columnIndexes[column];
        if (colIndex !== undefined) {
          const value = row[colIndex];
          
          // 必須チェック
          if (rules.required && (value === null || value === undefined || String(value).trim() === '')) {
            rowError.push(`行 ${i+1}, ${column}: 値が必須です`);
          }
          
          // 型チェック
          if (rules.type && value !== null && value !== undefined) {
            if (rules.type === 'number' && isNaN(Number(value))) {
              rowError.push(`行 ${i+1}, ${column}: 数値である必要があります`);
            } else if (rules.type === 'date' && isNaN(new Date(value).getTime())) {
              rowError.push(`行 ${i+1}, ${column}: 有効な日付である必要があります`);
            }
          }
          
          // 正規表現パターンチェック
          if (rules.pattern && value !== null && value !== undefined) {
            const regex = new RegExp(rules.pattern);
            if (!regex.test(String(value))) {
              rowError.push(`行 ${i+1}, ${column}: 形式が不正です`);
            }
          }
        }
      }
      
      if (rowError.length > 0) {
        rowErrors.push({
          row: i + 1,
          errors: rowError
        });
        errors.push(...rowError);
      }
    }
    
    return {
      valid: errors.length === 0,
      errors: errors,
      rowErrors: rowErrors,
      errorCount: errors.length,
      columnIndexes: columnIndexes,
      processedRows: data.length - 1
    };
  } catch (error) {
    logError('validateData', error, 'データ検証中');
    return {
      valid: false,
      errors: [`検証処理中にエラーが発生しました: ${error.message}`],
      errorCount: 1
    };
  }
}

/**
 * CSVデータをインポートする関数（最適化バージョン - 単純にデータをインポートすることに特化）
 * @param {string} csvData - CSVファイルの内容
 * @return {Object} インポート結果
 */
function importCsvData(csvData) {
  try {
    console.log("importCsvData開始: CSVデータサイズ=" + (csvData ? csvData.length : 0) + "バイト");
    
    // CSVデータの存在チェック - 最小限のチェックのみ
    if (!csvData || typeof csvData !== 'string' || csvData.trim() === '') {
      return { success: false, message: 'CSVデータが空または無効です。' };
    }
    
    // CSVデータをパース - シンプルかつ高速なパース処理
    let csvRows;
    try {
      // 高速なCSV分割
      csvData = csvData.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      
      // BOMを除去
      if (csvData.charCodeAt(0) === 0xFEFF) {
        csvData = csvData.substring(1);
      }
      
      // 単純な行分割
      const lines = csvData.split('\n');
      
      // 空行をフィルタして各行をCSVとして解析
      csvRows = lines
        .filter(line => line.trim()) // 空行を除外
        .map(line => {
          try {
            // 引用符を含む場合は複雑なパース処理
            if (line.includes('"')) {
              const result = [];
              let position = 0;
              let fieldStart = 0;
              let inQuotes = false;
              
              // 1文字ずつ解析して引用符を正確に処理
              while (position < line.length) {
                const char = line[position];
                
                if (char === '"') {
                  // 引用符の処理
                  if (inQuotes) {
                    // 次の文字も引用符かチェック (エスケープされた引用符かどうか)
                    if (position + 1 < line.length && line[position + 1] === '"') {
                      // 二重引用符はエスケープとして扱う
                      position++; // 追加の引用符をスキップ
                    } else {
                      // 単一の引用符は閉じる
                      inQuotes = false;
                    }
                  } else {
                    // 引用符を開く
                    inQuotes = true;
                  }
                } else if (char === ',' && !inQuotes) {
                  // 引用符の外側のカンマはフィールド区切り
                  let field = line.substring(fieldStart, position);
                  
                  // 引用符の除去 (最初と最後の引用符のみ)
                  if (field.startsWith('"') && field.endsWith('"')) {
                    field = field.substring(1, field.length - 1);
                  }
                  
                  // 二重引用符を単一引用符に戻す
                  field = field.replace(/""/g, '"');
                  
                  result.push(field);
                  fieldStart = position + 1;
                }
                
                position++;
              }
              
              // 最後のフィールドを追加
              let lastField = line.substring(fieldStart);
              
              // 引用符の除去 (最初と最後の引用符のみ)
              if (lastField.startsWith('"') && lastField.endsWith('"')) {
                lastField = lastField.substring(1, lastField.length - 1);
              }
              
              // 二重引用符を単一引用符に戻す
              lastField = lastField.replace(/""/g, '"');
              
              result.push(lastField);
              return result;
            } else {
              // 引用符がない場合は単純な分割
              return line.split(',');
            }
          } catch (fieldError) {
            console.error("フィールド処理エラー:", fieldError, "行:", line);
            // エラー時はシンプルな分割にフォールバック
            return line.split(',');
          }
        });
      
      console.log("CSVパース完了: 行数=" + csvRows.length);
    } catch (parseError) {
      console.error("CSVパースエラー:", parseError);
      // 最もシンプルな方法でフォールバック
      csvRows = csvData.split('\n').map(line => line.split(','));
      console.log("基本的なパースで成功: 行数=" + csvRows.length);
    }
    
    // 最小限のデータ検証
    if (!csvRows || !Array.isArray(csvRows) || csvRows.length <= 1) {
      return { success: false, message: 'CSVデータが不十分です。有効なデータ行がありません。' };
    }
    
    // 行ごとの列数を統一する（列数が異なるCSVファイルに対応）
    try {
      // ヘッダー行の列数を取得（基準とする列数）
      const headerRowLength = csvRows[0].length;
      console.log(`ヘッダー行の列数: ${headerRowLength}`);
      
      // 列数が少ない行には空文字列を追加、多い行は切り詰める
      for (let i = 1; i < csvRows.length; i++) {
        const currentRowLength = csvRows[i].length;
        
        if (currentRowLength < headerRowLength) {
          // 列数が足りない場合、空文字列で埋める
          const missingCols = headerRowLength - currentRowLength;
          csvRows[i] = csvRows[i].concat(Array(missingCols).fill(''));
          
          if (i < 5 || i % 1000 === 0) { // 最初の数行と1000行ごとにログ出力
            console.log(`行 ${i+1}: 列数を ${currentRowLength} から ${headerRowLength} に調整しました（${missingCols}列追加）`);
          }
        } else if (currentRowLength > headerRowLength) {
          // 列数が多い場合、余分な列を削除
          csvRows[i] = csvRows[i].slice(0, headerRowLength);
          
          if (i < 5 || i % 1000 === 0) { // 最初の数行と1000行ごとにログ出力
            console.log(`行 ${i+1}: 列数を ${currentRowLength} から ${headerRowLength} に調整しました（${currentRowLength - headerRowLength}列削除）`);
          }
        }
      }
      
      console.log(`全行の列数を ${headerRowLength} に統一しました`);
    } catch (formatError) {
      console.error("列数調整中にエラー:", formatError);
      // エラーがあっても処理を継続
    }
    
    // スプレッドシートにインポート
    try {
      // スプレッドシートの準備
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(EbayTool.getConfig().SHEET_NAMES.IMPORT);
      
      if (!sheet) {
        sheet = ss.insertSheet(EbayTool.getConfig().SHEET_NAMES.IMPORT);
      } else {
        sheet.clear();
      }
      
      // 最大限のバッチサイズで書き込み
      const LARGE_BATCH_SIZE = 10000; 
      
      // 総行数を出力
      console.log(`シートにデータを書き込みます: ${csvRows.length}行 x ${csvRows[0].length}列`);
      
      // バッチごとに書き込み
      for (let i = 0; i < csvRows.length; i += LARGE_BATCH_SIZE) {
        const endIdx = Math.min(i + LARGE_BATCH_SIZE, csvRows.length);
        const batch = csvRows.slice(i, endIdx);
        
        // null/undefinedを空文字に変換してからシートに書き込み
        const cleanBatch = batch.map(row => 
          row.map(cell => (cell === null || cell === undefined) ? '' : cell)
        );
        
        try {
          // 直接シートに書き込み
          sheet.getRange(i + 1, 1, cleanBatch.length, cleanBatch[0].length).setValues(cleanBatch);
          console.log(`${i + 1}行目から${endIdx}行目までの${cleanBatch.length}行を書き込みました`);
        } catch (batchError) {
          console.error(`バッチ書き込み中にエラー（${i + 1}～${endIdx}行）:`, batchError);
          
          // エラーが発生した場合は1行ずつ書き込みを試みる（最後の手段）
          if (i === 0) { // 最初のバッチでエラーが起きた場合のみ（1行ずつだと時間がかかりすぎるため）
            console.log("1行ずつの書き込みを試みます...");
            for (let j = 0; j < Math.min(100, batch.length); j++) { // 最初の100行だけ処理
              try {
                const singleRow = [cleanBatch[j]];
                sheet.getRange(i + j + 1, 1, 1, singleRow[0].length).setValues(singleRow);
              } catch (rowError) {
                console.error(`行 ${i + j + 1} の書き込みに失敗:`, rowError);
              }
            }
          }
          
          // それでも失敗する場合はエラーを返す
          throw new Error(`データの書き込みに問題があります: ${batchError.message}`);
        }
      }
      
      // 最小限の書式設定
      sheet.getRange(1, 1, 1, csvRows[0].length)
        .setBackground(EbayTool.getColor('PRIMARY'))
        .setFontColor('white')
        .setFontWeight('bold');
      
      // 先頭行を固定
      sheet.setFrozenRows(1);
      
      // 成功メッセージを返す
      return { 
        success: true, 
        message: `${csvRows.length - 1}件のデータを正常にインポートしました。`,
        rowCount: csvRows.length - 1
      };
    } catch (sheetError) {
      console.error("シート処理中にエラー:", sheetError);
      return { 
        success: false, 
        message: `スプレッドシートへの書き込み中にエラーが発生しました: ${sheetError.message}`
      };
    }
  } catch (error) {
    console.error("importCsvData関数でエラー:", error);
    return { 
      success: false, 
      message: `CSVのインポートに失敗しました: ${error.message}`
    };
  }
}

/**
 * インポートデータから重複を検出する関数（シンプル化バージョン）
 * @return {Object} 処理結果
 */
function detectDuplicates() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const SHEET_NAMES = EbayTool.getConfig().SHEET_NAMES;
    
    const importSheet = ss.getSheetByName(SHEET_NAMES.IMPORT);
    
    if (!importSheet) {
      return { success: false, message: 'インポートデータが見つかりません。先にCSVをインポートしてください。' };
    }
    
    // データを取得
    const lastRow = importSheet.getLastRow();
    const lastCol = importSheet.getLastColumn();
    
    if (lastRow <= 1) {
      // 重複データが0件の場合は正常完了として処理
      return { 
        success: true, 
        message: '検出された重複: 0件。重複データはありませんでした。',
        duplicateCount: 0,
        analysisComplete: true
      };
    }
    
    // ヘッダーを取得
    const headers = importSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headersLower = headers.map(h => String(h).toLowerCase().replace(/\s+/g, ''));
    
    // 必要な列のインデックスを探す
    let titleIndex = -1;
    let itemIdIndex = -1;
    let startDateIndex = -1;
    
    // タイトル列を探す - 単純化
    for (let i = 0; i < headers.length; i++) {
      const headerLower = headersLower[i];
      if (headerLower.includes('title') || headerLower.includes('name')) {
        titleIndex = i;
        break;
      }
    }
    
    // ID列を探す - 単純化
    for (let i = 0; i < headers.length; i++) {
      const headerLower = headersLower[i];
      if (headerLower.includes('item') || headerLower.includes('id') || headerLower.includes('number')) {
        itemIdIndex = i;
        break;
      }
    }
    
    // 開始日列を探す - 単純化
    for (let i = 0; i < headers.length; i++) {
      const headerLower = headersLower[i];
      if (headerLower.includes('date') || headerLower.includes('start')) {
        startDateIndex = i;
        break;
      }
    }
    
    // 必要な列が見つからない場合は、デフォルト値を使用
    if (titleIndex === -1 && headers.length > 1) {
      titleIndex = 1;  // 2列目をタイトルと想定
    }
    
    if (itemIdIndex === -1 && headers.length > 0) {
      itemIdIndex = 0;  // 1列目をIDと想定
    }
    
    if (startDateIndex === -1 && headers.length > 2) {
      startDateIndex = 2;  // 3列目を日付と想定
    }
    
    if (titleIndex === -1 || itemIdIndex === -1) {
      return { 
        success: false, 
        message: '必須カラム(タイトル、ID)が見つかりません。'
      };
    }
    
    console.log(`重複検出に使用する列: title=${titleIndex} (${headers[titleIndex]}), itemId=${itemIdIndex} (${headers[itemIdIndex]}), startDate=${startDateIndex} (${headers[startDateIndex] || 'N/A'})`);
    
    // 動的バッチサイズでデータを取得（パフォーマンス最適化）
    const dataSize = lastRow - 1;
    const CONFIG = EbayTool.getConfig();
    const optimalBatchSize = CONFIG.calculateOptimalBatchSize(dataSize);
    console.log(`データサイズ: ${dataSize}行, 最適バッチサイズ: ${optimalBatchSize}`);
    
    const allData = importSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // タイトルでグループ化
    const titleGroups = {};
    // 各行を処理
    for (let i = 0; i < allData.length; i++) {
      const row = allData[i];
      // タイトルの正規化（分析シートと同じロジック）
      const title = EbayTool.TextAnalyzer.normalizeTitle(String(row[titleIndex] || ''), false);
      const itemId = String(row[itemIdIndex] || '').trim();
      const startDate = row[startDateIndex];
      // 有効なタイトルとIDがある場合のみ処理
      if (title && itemId) {
        if (!titleGroups[title]) {
          titleGroups[title] = [];
        }
        titleGroups[title].push({
          itemId: itemId,
          title: title,
          originalTitle: row[titleIndex],
          startDate: startDate,
          allData: row
        });
      }
    }
    
    // 重複グループのみを抽出（2つ以上のアイテムがあるグループ）
    const duplicateGroups = Object.values(titleGroups)
      .filter(group => group.length > 1)
      .sort((a, b) => b.length - a.length);
    
    // 重複リストシートを準備
    let duplicateSheet = ss.getSheetByName(SHEET_NAMES.DUPLICATES);
    if (!duplicateSheet) {
      duplicateSheet = ss.insertSheet(SHEET_NAMES.DUPLICATES);
    } else {
      duplicateSheet.clear();
    }
    
    // 重複リストを作成（単純化されたバージョン）
    createDuplicateListSheet(duplicateSheet, duplicateGroups, headers);
    
    // 重複シートをアクティブにする
    ss.setActiveSheet(duplicateSheet);
    
    return { 
      success: true, 
      message: `${duplicateGroups.length}件の重複グループを検出しました。合計${getTotalDuplicates(duplicateGroups)}件の重複アイテムがあります。`,
      duplicateGroups: duplicateGroups.length,
      duplicateItems: getTotalDuplicates(duplicateGroups)
    };
  } catch (error) {
    logError('detectDuplicates', error, '重複検出処理中');
    SpreadsheetApp.getUi().alert(
      'エラー',
      getFriendlyErrorMessage(error, '重複検出中にエラーが発生しました。'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return { 
      success: false, 
      message: getFriendlyErrorMessage(error, '重複検出中にエラーが発生しました。'), 
      stack: error.stack 
    };
  }
}

/**
 * タイトルをノーマライズする関数（改良版）
 * @param {string} title - 元のタイトル
 * @param {boolean} useAdvanced - 高度なノーマライズを使用するかどうか
 * @return {string} ノーマライズされたタイトル
 */
function normalizeTitle(title, useAdvanced = false) {
  if (!title) return '';
  
  // 基本的なノーマライズ
  let normalized = String(title).toLowerCase();
  
  // 空白を統一
  normalized = normalized.replace(/\s+/g, ' ').trim();
  
  // 特殊文字を削除または置換
  normalized = normalized.replace(/[^\w\s]/g, '');
  
  // 高度なノーマライズ（オプション）
  if (useAdvanced) {
    // 一般的な略語や単位を標準化
    const replacements = {
      'in.': 'inch',
      'inches': 'inch',
      'ft.': 'foot',
      'feet': 'foot',
      'lbs': 'pound',
      'lb.': 'pound',
      'pounds': 'pound',
      'oz.': 'ounce',
      'ounces': 'ounce',
      'pcs': 'piece',
      'pc.': 'piece',
      'pieces': 'piece'
    };
    
    for (const [abbr, full] of Object.entries(replacements)) {
      normalized = normalized.replace(new RegExp('\\b' + abbr + '\\b', 'g'), full);
    }
    
    // 重要なキーワードのみを抽出
    const words = normalized.split(' ');
    
    // ストップワード（無視する一般的な単語）
    const stopWords = ['a', 'an', 'the', 'and', 'or', 'but', 'is', 'are', 'on', 'at', 'to', 'for', 'with', 'by', 'in', 'of'];
    
    // ストップワードを除去し、短すぎる単語も除去
    const filteredWords = words.filter(word => 
      word.length > 2 && !stopWords.includes(word)
    );
    
    // 単語をソートして順序の違いを無視
    filteredWords.sort();
    
    // 出現回数が多い単語を強調
    const wordCounts = {};
    filteredWords.forEach(word => {
      wordCounts[word] = (wordCounts[word] || 0) + 1;
    });
    
    // 重要度に基づいて単語を選択
    const importantWords = filteredWords.filter((word, index, self) => 
      // 重複を除去
      index === self.indexOf(word) && (
        // 出現回数が多いか、長い単語は重要
        wordCounts[word] > 1 || word.length > 5
      )
    );
    
    // 十分な単語がない場合は全単語を使用
    if (importantWords.length < 3) {
      return normalized;
    }
    
    normalized = importantWords.join(' ');
  }
  
  return normalized;
}

/**
 * 2つのタイトル間の類似度を計算する関数
 * @param {string} title1 - 1つ目のタイトル
 * @param {string} title2 - 2つ目のタイトル
 * @return {number} 類似度（0-1の範囲、1が完全一致）
 */
function calculateTitleSimilarity(title1, title2) {
  // 両方のタイトルをノーマライズ
  const normalized1 = normalizeTitle(title1, true);
  const normalized2 = normalizeTitle(title2, true);
  
  if (!normalized1 || !normalized2) return 0;
  
  // 単語ベースでの類似度計算
  const words1 = normalized1.split(' ');
  const words2 = normalized2.split(' ');
  
  // Jaccard係数の計算
  const set1 = new Set(words1);
  const set2 = new Set(words2);
  
  const intersection = new Set([...set1].filter(word => set2.has(word)));
  const union = new Set([...set1, ...set2]);
  
  return intersection.size / union.size;
}

/**
 * 重複の合計件数を取得する関数
 * @param {Array} duplicateGroups - 重複グループのリスト
 * @return {number} 重複の合計件数
 */
function getTotalDuplicates(duplicateGroups) {
  return duplicateGroups.reduce((total, group) => total + group.length, 0);
}

/**
 * 重複リストシートを作成する関数（シンプル化バージョン）
 * @param {Sheet} sheet - 重複リストシート
 * @param {Array} duplicateGroups - 重複グループのリスト
 * @param {Array} originalHeaders - 元のヘッダー
 */
function createDuplicateListSheet(sheet, duplicateGroups, originalHeaders) {
  // ヘッダー行を作成（グループID、重複タイプなどの列を追加）
  const headers = ['グループID', '重複タイプ', '処理'].concat(originalHeaders);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 一度にすべての行を設定するためのデータ配列
  const allData = [];
  
  // 各重複グループをデータ配列に追加
  duplicateGroups.forEach((group, groupIndex) => {
    // グループをスタート日でソート
    group.sort((a, b) => {
      if (!a.startDate || !b.startDate) return 0;
      const dateA = new Date(a.startDate);
      const dateB = new Date(b.startDate);
      return dateB - dateA; // 新しい順（降順）でソート
    });
    
    // グループ内の各アイテムを追加
    group.forEach((item, itemIndex) => {
      const row = new Array(headers.length).fill('');
      
      // グループIDと重複タイプの列を設定
      row[0] = `Group ${groupIndex + 1}`; // グループID
      row[1] = `${group.length}件中${itemIndex + 1}件目`; // 重複タイプ
      row[2] = itemIndex === 0 ? '残す' : '終了'; // 処理（最新のみ残す）
      
      // 元のデータを追加
      item.allData.forEach((value, i) => {
        row[i + 3] = value;
      });
      
      allData.push(row);
    });
    
    // グループ間の区切り行（必要に応じて空行を入れる）
    if (groupIndex < duplicateGroups.length - 1) {
      allData.push(new Array(headers.length).fill(''));
    }
  });
  
  // データが存在する場合に一度に書き込む
  if (allData.length > 0) {
    sheet.getRange(2, 1, allData.length, headers.length).setValues(allData);
  }
  
  // シートのフォーマットを整える（シンプル化）
  formatDuplicateSheetSimple(sheet);
}

/**
 * 重複リストシートのフォーマットを整える関数（シンプル化バージョン）
 * @param {Sheet} sheet - フォーマットするシート
 */
function formatDuplicateSheetSimple(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow <= 1) return; // データがない場合は何もしない
  
  // ヘッダー行のみ書式設定
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground(EbayTool.getColor('PRIMARY'))
    .setFontColor('white')
    .setFontWeight('bold');
  
  // 先頭行を固定
  sheet.setFrozenRows(1);
  
}

/**
 * CSVデータをダウンロードする関数（直接ダウンロード版）
 * @param {Array} data - CSVデータの2次元配列
 * @param {string} fileName - ダウンロードされるファイル名
 * @return {Object} 処理結果（HTML出力）
 */
function convertToCSVDownload(data, fileName) {
  try {
    // CSVデータを生成（セキュリティ強化）
    let csvContent = data.map(row => 
      row.map(cell => {
        // null/undefinedの処理
        if (cell === null || cell === undefined) {
          return '';
        }
        
        // 文字列に変換
        let cellStr = String(cell);
        
        // セキュリティ上の問題となりうる文字をエスケープ
        cellStr = cellStr
          .replace(/"/g, '""') // 引用符のエスケープ
          .replace(/\\/g, '\\\\'); // バックスラッシュのエスケープ
        
        // カンマ、引用符、改行、タブを含む場合は引用符で囲む
        if (/[,"\n\r\t]/.test(cellStr)) {
          return '"' + cellStr + '"';
        }
        
        return cellStr;
      }).join(',')
    ).join('\n');
    
    // BOMを追加してUTF-8として認識されるようにする
    const bom = '\ufeff';
    csvContent = bom + csvContent;
    
    // HTMLでダウンロードリンクを生成（直接ダウンロード方式）
    const html = HtmlService.createHtmlOutput(
      `<html>
        <head>
          <base target="_top">
          <meta charset="UTF-8">
          <script>
            // CSVデータ
            const csvData = \`${csvContent.replace(/`/g, '\\`')}\`;
            
            // ページ読み込み時に直接ダウンロード
            window.onload = function() {
              try {
                // Blobオブジェクトの作成
                const blob = new Blob([csvData], {type: 'text/csv;charset=utf-8;'});
                
                // URL.createObjectURLでブラウザ内URLを生成
                const url = URL.createObjectURL(blob);
                
                // ダウンロードリンクを作成
                const link = document.createElement('a');
                link.href = url;
                link.download = '${fileName.replace(/'/g, "\\'")}';
                document.body.appendChild(link);
                
                // クリックして即時ダウンロード
                link.click();
                
                // 後片付け
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                
                // 成功メッセージを表示
                document.getElementById('status').innerHTML = 
                  '<div style="color: green; font-weight: bold; padding: 10px; background-color: #E8F5E9; border-radius: 4px; margin-top: 10px;">ダウンロードが完了しました。</div>';
                
                // ユーザーに通知するためのアラート表示
                alert('CSVファイルのダウンロードが完了しました！');
                
                // 親ウィンドウ（サイドバー）に通知してダウンロード完了メッセージを表示
                try {
                  window.parent.postMessage({
                    type: 'download-complete',
                    fileName: '${fileName.replace(/'/g, "\\'")}'
                  }, '*');
                } catch (err) {
                  console.error('親ウィンドウへの通知エラー:', err);
                }
                
                // ダウンロードボタンを無効化
                const downloadBtn = document.getElementById('downloadBtn');
                if (downloadBtn) {
                  downloadBtn.disabled = true;
                  downloadBtn.classList.add('disabled');
                }
                
                // 3秒後にダイアログを閉じる
                setTimeout(function() {
                  google.script.host.close();
                }, 3000);
              } catch (e) {
                document.getElementById('status').innerHTML = 
                  '<div style="color: red;">エラーが発生しました: ' + e.message + '</div>';
              }
            }
          </script>
          <style>
            body {
              font-family: Arial, sans-serif;
              margin: 20px;
              text-align: center;
            }
            .button {
              background-color: #4F46E5;
              color: white;
              padding: 10px 15px;
              border: none;
              border-radius: 4px;
              cursor: pointer;
              font-size: 14px;
              margin-top: 10px;
            }
            .button:hover {
              background-color: #4338CA;
            }
          </style>
        </head>
        <body>
          <h3>CSVファイルのダウンロード</h3>
          <p>ダウンロードが自動的に始まります。</p>
          <div id="status">準備中...</div>
          <button id="downloadBtn" class="button">ダウンロード</button>
        </body>
      </html>`
    )
    .setWidth(400)
    .setHeight(200);
    
    return html;
  } catch (error) {
    logError('convertToCSVDownload', error, 'CSVデータ変換中');
    throw new Error(`CSVデータの変換に失敗しました: ${error.message}`);
  }
}

/**
 * エクスポートシートからCSVをダウンロードする関数（直接ダウンロード版）
 */
function downloadExportCsv() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const exportSheet = ss.getSheetByName(EbayTool.getSheetName('EXPORT'));
    
    if (!exportSheet) {
      throw new Error('エクスポートシートが見つかりません。先にエクスポートを実行してください。');
    }
    
    const data = exportSheet.getDataRange().getValues();
    const fileName = `ebay_end_items_${new Date().toISOString().substr(0, 10)}.csv`;
    
    const html = convertToCSVDownload(data, fileName);
    SpreadsheetApp.getUi().showModalDialog(html, 'CSVダウンロード');
    
    return { success: true, message: 'ダウンロードを開始しました' };
  } catch (error) {
    logError('downloadExportCsv', error);
    return { success: false, message: `エラーが発生しました: ${error.message}` };
  }
}

/**
 * すべてのシートを初期化する関数（高速化バージョン）
 * @return {Object} 処理結果
 */
function initializeAllSheets() {
  try {
    console.log("initializeAllSheets: 関数が呼び出されました");
    
    // UIオブジェクトを取得
    const ui = SpreadsheetApp.getUi();
    console.log("initializeAllSheets: UIオブジェクトを取得しました");
    
    // 明示的な確認ダイアログを表示
    const response = ui.alert(
      'シート初期化の確認',
      'すべてのシートを初期化します。この操作は元に戻せません。続行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    console.log("initializeAllSheets: 確認ダイアログの応答:", response);
    
    // キャンセルした場合
    if (response !== ui.Button.YES) {
      console.log("initializeAllSheets: ユーザーがキャンセルしました");
      return { 
        success: false, 
        message: '初期化をキャンセルしました。',
        userCancelled: true 
      };
    }
    
    console.log("initializeAllSheets: 初期化作業を開始します");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 保持するシート名
    const logSheetName = EbayTool.getSheetName('LOG');
    const operationLogSheetName = "操作ログ";
    
    // 必要なシート名
    const requiredSheets = [
      EbayTool.getSheetName('IMPORT'),
      EbayTool.getSheetName('DUPLICATES'),
      EbayTool.getSheetName('EXPORT'),
      EbayTool.getSheetName('ANALYSIS')
    ];
    
    // すべてのシートを取得
    const allSheets = ss.getSheets();
    const existingSheets = new Map(); // シート名→シートオブジェクトのマップ
    const sheetsToDelete = []; // 削除対象のシート
    
    // 既存シートの分類
    for (let i = 0; i < allSheets.length; i++) {
      const sheet = allSheets[i];
      const sheetName = sheet.getName();
      
      // 保持するシートはスキップ
      if (sheetName === logSheetName || sheetName === operationLogSheetName) {
        existingSheets.set(sheetName, sheet);
        continue;
      }
      
      // 必要なシートは保持してクリア
      if (requiredSheets.includes(sheetName)) {
        // シートをクリアして再利用
        sheet.clear();
        existingSheets.set(sheetName, sheet);
        console.log(`シート「${sheetName}」をクリアしました`);
      } else {
        // 不要なシートは削除対象としてマーク
        sheetsToDelete.push(sheet);
      }
    }
    
    // 不要なシートを削除（一括削除は手順として注意）
    for (let i = 0; i < sheetsToDelete.length; i++) {
      const sheet = sheetsToDelete[i];
      ss.deleteSheet(sheet);
      console.log(`シート「${sheet.getName()}」を削除しました`);
    }
    
    // 必要なシートで存在しないものを作成
    for (const sheetName of requiredSheets) {
      if (!existingSheets.has(sheetName)) {
        const newSheet = ss.insertSheet(sheetName);
        existingSheets.set(sheetName, newSheet);
        console.log(`シート「${sheetName}」を新規作成しました`);
      }
    }
    
    // 操作ログシートが存在しない場合は作成
    if (!existingSheets.has(operationLogSheetName)) {
      const operationLogSheet = ss.insertSheet(operationLogSheetName);
      
      // ヘッダー行の設定
      operationLogSheet.appendRow([
        "操作日時", 
        "操作内容", 
        "ステータス", 
        "処理時間(秒)", 
        "データ件数", 
        "詳細情報"
      ]);
      
      // ヘッダー行の書式設定
      operationLogSheet.getRange(1, 1, 1, 6).setBackground("#f3f4f6").setFontWeight("bold");
      console.log(`シート「${operationLogSheetName}」を新規作成しました`);
    }
    
    // ログシートが存在しない場合は作成
    if (!existingSheets.has(logSheetName)) {
      const logSheet = ss.insertSheet(logSheetName);
      
      // ヘッダー行の設定
      logSheet.appendRow(['タイムスタンプ', '関数', 'エラータイプ', 'エラーメッセージ', 'コンテキスト', 'スタックトレース']);
      
      // ヘッダー行の書式設定
      logSheet.getRange(1, 1, 1, 6)
        .setBackground(EbayTool.getColor('PRIMARY'))
        .setFontColor('white')
        .setFontWeight('bold');
        
      // 列幅の設定
      logSheet.setColumnWidth(1, 150); // タイムスタンプ
      logSheet.setColumnWidth(2, 100); // 関数
      logSheet.setColumnWidth(3, 100); // エラータイプ
      logSheet.setColumnWidth(4, 250); // エラーメッセージ
      logSheet.setColumnWidth(5, 200); // コンテキスト
      logSheet.setColumnWidth(6, 400); // スタックトレース
      
      console.log(`シート「${logSheetName}」を新規作成しました`);
    }
    
    console.log("initializeAllSheets: すべてのシートを初期化しました");
    
    return { 
      success: true, 
      message: 'すべてのシートを初期化しました。', 
      requireReload: true
    };
  } catch (error) {
    console.error("initializeAllSheets: エラーが発生しました:", error);
    logError('initializeAllSheets', error);
    return { success: false, message: `エラーが発生しました: ${error.message}` };
  }
}

/**
 * サイドバーを閉じる関数
 */
function closeSidebar() {
  const ui = SpreadsheetApp.getUi();
  try {
    const html = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>')
      .setWidth(0)
      .setHeight(0);
    
    SpreadsheetApp.getUi().showModalDialog(html, '閉じています...');
    return { success: true };
  } catch (error) {
    logError('closeSidebar', error);
    return { success: false, message: error.message };
  }
}

/**
 * 重複リストシートを初期化する関数
 * @return {Object} 処理結果
 */
function initializeDuplicatesSheet() {
  return initializeSheet(EbayTool.getSheetName('DUPLICATES'));
}

/**
 * エクスポートシートを初期化する関数
 * @return {Object} 処理結果
 */
function initializeExportSheet() {
  return initializeSheet(EbayTool.getSheetName('EXPORT'));
}

/**
 * 特定のシートを初期化する関数
 * @param {string} sheetName - 初期化するシートの名前
 * @param {boolean} skipConfirmation - 確認ダイアログをスキップするかどうか
 * @return {Object} 処理結果
 */
function initializeSheet(sheetName, skipConfirmation = false) {
  try {
    if (!skipConfirmation) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'シート初期化の確認',
        `「${sheetName}」シートを初期化します。この操作は元に戻せません。続行しますか？`,
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) {
        return { success: false, message: '初期化をキャンセルしました。' };
      }
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    // シートが存在しない場合は新規作成
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      console.log(`シート「${sheetName}」を新規作成しました`);
      return { success: true, message: `シート「${sheetName}」を新規作成しました。` };
    }
    
    if (sheetName === EbayTool.getSheetName('LOG')) {
      // ログシートは最後の10行だけ残す
      const lastRow = sheet.getLastRow();
      if (lastRow > EbayTool.getConfig().MAX_LOG_ROWS) {
        sheet.deleteRows(1, lastRow - EbayTool.getConfig().MAX_LOG_ROWS);
      }
    } else {
      // その他のシートは完全に削除して再作成（書式設定も含めて完全に初期化）
      const sheetIndex = sheet.getIndex();
      ss.deleteSheet(sheet);
      sheet = ss.insertSheet(sheetName, sheetIndex - 1);
      console.log(`シート「${sheetName}」を削除して再作成しました`);
    }
    
    return { success: true, message: `シート「${sheetName}」を初期化しました。` };
  } catch (error) {
    logError('initializeSheet', error);
    return { success: false, message: `エラーが発生しました: ${error.message}` };
  }
}

/**
 * 重複タイトルを分析する関数（改良版）
 * @return {Object} 処理結果
 */
function analyzeDuplicateTitles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const duplicateSheet = ss.getSheetByName(EbayTool.getSheetName('DUPLICATES'));
    
    if (!duplicateSheet) {
      return { success: false, message: '重複リストが見つかりません。先に重複検出を実行してください。' };
    }
    
    // データを取得
    const data = duplicateSheet.getDataRange().getValues();
    if (data.length <= 1) {
      // 重複データが0件の場合は正常完了として処理
      return { 
        success: true, 
        message: '検出された重複: 0件。重複データはありませんでした。',
        duplicateCount: 0,
        analysisComplete: true
      };
    }
    
    // ヘッダーを取得
    const headers = data[0];
    
    // 重要な列のインデックスを特定する
    let itemIdIndex = -1;     // 商品IDの列
    let titleIndex = -1;      // 実際のタイトル（商品名）の列
    let startDateIndex = -1;  // 開始日の列
    const defaultMonthDay = 'その他'; // デフォルトの日付カテゴリ
    
    // ヘッダーから列を特定
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase();
      
      // タイトル列を探す
      if (header.includes('title') || header.includes('name') || header.includes('item name')) {
        titleIndex = i;
      }
      
      // 商品ID列を探す
      if ((header.includes('item') && (header.includes('id') || header.includes('number'))) || 
          header === 'itemnumber' || header === 'id') {
        itemIdIndex = i;
      }
      
      // 開始日列を探す
      if (header.includes('start date') || header === 'startdate' || 
          header.includes('list date') || (header.includes('start') && header.includes('time'))) {
        startDateIndex = i;
      }
    }
    
    console.log(`列インデックス - タイトル: ${titleIndex}, 商品ID: ${itemIdIndex}, 開始日: ${startDateIndex}`);
    
    // 必要な列が見つからない場合の代替策
    if (titleIndex === -1 && itemIdIndex !== -1) {
      // タイトルが見つからないがIDがある場合は、ID以外の列を探す（多くの場合、商品名と思われる列）
      for (let i = 3; i < headers.length; i++) {
        if (i !== itemIdIndex && i !== startDateIndex) {
          // データの最初の数行をチェックして、テキストが含まれる列を探す
          let hasText = false;
          for (let j = 1; j < Math.min(data.length, 10); j++) {
            if (data[j][i] && typeof data[j][i] === 'string' && data[j][i].length > 15) {
              hasText = true;
              break;
            }
          }
          if (hasText) {
            titleIndex = i;
            console.log(`タイトル列が自動検出されました: ${i} (${headers[i]})`);
            break;
          }
        }
      }
    }
    
    // 開始日が見つからない場合は、日付らしき列を探す
    if (startDateIndex === -1) {
      for (let i = 0; i < headers.length; i++) {
        const header = String(headers[i]).toLowerCase();
        if (header.includes('date') || header.includes('time')) {
          // データの最初の数行をチェックして日付フォーマットかどうか確認
          for (let j = 1; j < Math.min(data.length, 10); j++) {
            const val = data[j][i];
            if (val && !isNaN(new Date(val).getTime())) {
              startDateIndex = i;
              console.log(`日付列が自動検出されました: ${i} (${headers[i]})`);
              break;
            }
          }
          if (startDateIndex !== -1) break;
        }
      }
    }
    
    // それでも見つからない場合はデフォルト値
    if (titleIndex === -1 && headers.length > 3) titleIndex = 3;
    if (startDateIndex === -1 && headers.length > 4) startDateIndex = 4;
    
    // 分析シートを準備 - 完全に初期化してから使用する
    let analysisSheet = ss.getSheetByName(EbayTool.getSheetName('ANALYSIS'));
    if (!analysisSheet) {
      analysisSheet = ss.insertSheet(EbayTool.getSheetName('ANALYSIS'));
    } else {
      // 分析シートを完全に初期化
      initializeSheet(EbayTool.getSheetName('ANALYSIS'), true);
      
      // シートの参照を更新
      analysisSheet = ss.getSheetByName(EbayTool.getSheetName('ANALYSIS'));
    }
    
    // 分析タイトルを設定
    const titleRange = analysisSheet.getRange(1, 1);
    titleRange.setValue('eBay出品タイトル重複分析');
    titleRange.setFontSize(14);
    titleRange.setFontWeight('bold');
    
    // 説明を追加
    const descRange = analysisSheet.getRange(2, 1);
    descRange.setValue('このシートでは、重複回数ごとにeBay出品タイトルを分析しています。数字が大きいほど多く重複している項目です。');
    descRange.setFontStyle('italic');
    
    // 日付を「月-日」形式に整形する関数
    function formatMonthDay(date) {
      try {
        // 無効な日付をチェック
        if (!date || isNaN(new Date(date).getTime())) {
          return null; // 無効な日付はnullを返す
        }
        
        const dateObj = new Date(date);
        return `${String(dateObj.getMonth() + 1).padStart(2, '0')}-${String(dateObj.getDate()).padStart(2, '0')}`;
      } catch (e) {
        console.error("日付フォーマットエラー:", e, date);
        return null;
      }
    }
    
    // タイトル＋日付ごとにカウントするマップを作成（処理の最適化）
    const titleDateCountMap = new Map(); // key: 正規化タイトル, value: Map(日付, 件数)
    const titleTotalCountMap = new Map(); // key: 正規化タイトル, value: 重複回数
    const titleDisplayMap = new Map(); // key: 正規化タイトル, value: 表示用タイトル
    
    // データ処理を最適化（単一ループで処理）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // 空行はスキップ
      
      let displayTitle = titleIndex !== -1 ? String(row[titleIndex] || '') : '';
      
      // 表示用タイトルが見つからない場合の対応
      if (!displayTitle || displayTitle.trim() === '' || /^\d+$/.test(displayTitle)) {
        for (let j = 0; j < row.length; j++) {
          if (j === itemIdIndex || j === startDateIndex) continue;
          const cellValue = String(row[j] || '');
          if (cellValue.length > 10 && !/^\d+$/.test(cellValue)) {
            displayTitle = cellValue;
            if (titleIndex === -1) titleIndex = j;
            break;
          }
        }
      }
      
      if (!displayTitle || displayTitle.trim() === '') continue;
      
      // タイトルの正規化 - 効率化のためにオプションを無効化
      const normalizedTitle = EbayTool.TextAnalyzer.normalizeTitle(displayTitle, false);
      
      // 表示用タイトルを保存（最初に出現したもの）
      if (!titleDisplayMap.has(normalizedTitle)) {
        titleDisplayMap.set(normalizedTitle, displayTitle);
      }
      
      // 日付処理 - シンプル化して効率アップ
      let monthDay = 'その他';
      if (startDateIndex !== -1) {
        const dateValue = row[startDateIndex];
        if (dateValue) {
          if (typeof dateValue === 'string' && /^\d{2}-\d{2}$/.test(dateValue.trim())) {
            monthDay = dateValue.trim();
          } else {
            try {
              const dateObj = new Date(dateValue);
              if (!isNaN(dateObj.getTime())) {
                monthDay = `${String(dateObj.getMonth() + 1).padStart(2, '0')}-${String(dateObj.getDate()).padStart(2, '0')}`;
              }
            } catch (e) {
              // エラー時はデフォルト値を使用
            }
          }
        }
      }
      
      // タイトル＋日付でカウント - Mapの操作を最適化
      if (!titleDateCountMap.has(normalizedTitle)) {
        titleDateCountMap.set(normalizedTitle, new Map());
      }
      const dateMap = titleDateCountMap.get(normalizedTitle);
      dateMap.set(monthDay, (dateMap.get(monthDay) || 0) + 1);
      
      // タイトルごとの合計件数
      titleTotalCountMap.set(normalizedTitle, (titleTotalCountMap.get(normalizedTitle) || 0) + 1);
    }
    
    // 重複回数ごとにタイトルを分類（1件のみは除外）- Map操作を最適化
    const duplicateCountTitlesMap = new Map();
    for (const [normalizedTitle, count] of titleTotalCountMap.entries()) {
      if (count <= 1) continue; // 重複がない場合はスキップ
      
      if (!duplicateCountTitlesMap.has(count)) {
        duplicateCountTitlesMap.set(count, []);
      }
      duplicateCountTitlesMap.get(count).push(normalizedTitle);
    }
    
    // 日付リストの生成を最適化
    const allMonthDays = new Set();
    for (const dateMap of titleDateCountMap.values()) {
      for (const monthDay of dateMap.keys()) {
        allMonthDays.add(monthDay);
      }
    }
    
    // 日付がない場合のデフォルト処理
    if (allMonthDays.size === 0) {
      allMonthDays.add(defaultMonthDay);
    }
    
    // 日付のソート処理を最適化
    const otherCategory = allMonthDays.has('その他') ? ['その他'] : [];
    const dateDays = Array.from(allMonthDays)
      .filter(day => day !== 'その他')
      .sort();
    const sortedMonthDays = [...dateDays, ...otherCategory];
    
    // ピボットテーブル生成部を最適化
    let currentRowOffset = 3;
    const duplicateCounts = Array.from(duplicateCountTitlesMap.keys()).sort((a, b) => b - a);
    
    // 書式設定のバッチ処理用の配列
    let formattingBatches = [];
    
    // 各重複回数ごとの処理
    for (const count of duplicateCounts) {
      const titles = duplicateCountTitlesMap.get(count) || [];
      
      // タイトル行の設定
      const titleCell = analysisSheet.getRange(currentRowOffset, 1);
      titleCell.setValue(`重複回数 ${count} のピボットテーブル：`);
      titleCell.setFontWeight('bold');
      currentRowOffset += 1;
      
      // ヘッダー行の設定
      const pivotHeaders = ['タイトル'].concat(sortedMonthDays);
      const pivotHeaderRange = analysisSheet.getRange(currentRowOffset, 1, 1, pivotHeaders.length);
      pivotHeaderRange.setValues([pivotHeaders]);
      
      // ヘッダー行の書式設定をバッチで適用
      pivotHeaderRange.setBackground('#0F9D58')
                      .setFontColor('white')
                      .setFontWeight('bold');
      
      // データがない場合のスキップ処理を追加
      if (titles.length === 0) {
        analysisSheet.getRange(currentRowOffset + 1, 1, 1, pivotHeaders.length)
          .setValues([['データなし'].concat(Array(sortedMonthDays.length).fill(0))]);
        currentRowOffset += 3;
        
        // 少し遅延を入れてスプレッドシートの内部処理がキャッチアップできるようにする
        Utilities.sleep(50);
        continue;
      }
      
      // ピボットテーブルのデータを作成
      const pivotData = [];
      const cellFormattingData = []; // セルの書式設定情報を保存
      
      // 各タイトルのデータ行を構築
      for (const normalizedTitle of titles) {
        const row = [titleDisplayMap.get(normalizedTitle)];
        const dateMap = titleDateCountMap.get(normalizedTitle) || new Map();
        
        // 各日付の値を構築
        for (let j = 0; j < sortedMonthDays.length; j++) {
          const monthDay = sortedMonthDays[j];
          const value = dateMap.get(monthDay) || 0;
          row.push(value);
          
          // 書式設定が必要なセルの情報を保存
          if (value > 0) {
            cellFormattingData.push({
              rowIdx: pivotData.length,
              colIdx: j + 1,
              value: value
            });
          }
        }
        
        pivotData.push(row);
      }
      
      // データをシートに書き込み
      if (pivotData.length > 0) {
        const pivotDataRange = analysisSheet.getRange(
          currentRowOffset + 1, 
          1, 
          pivotData.length, 
          pivotHeaders.length
        );
        pivotDataRange.setValues(pivotData);
        
        // 行の背景色を交互に設定 - バッチ処理
        for (let i = 0; i < pivotData.length; i++) {
          const rowRange = analysisSheet.getRange(
            currentRowOffset + 1 + i, 
            1, 
            1, 
            pivotHeaders.length
          );
          
          // 奇数/偶数行で背景色を変える
          rowRange.setBackground(i % 2 === 0 ? '#E0F2F1' : '#E8F5E9');
        }
        
        // セルの書式設定をバッチ処理
        const batchSize = 20; // バッチサイズを制限
        for (let i = 0; i < cellFormattingData.length; i += batchSize) {
          const batch = cellFormattingData.slice(i, i + batchSize);
          
          // 各セルの書式設定を適用
          batch.forEach(item => {
            const cell = analysisSheet.getRange(
              currentRowOffset + 1 + item.rowIdx, 
              item.colIdx + 1
            );
            
            // 値に応じて書式設定
            if (item.value >= 3) {
              cell.setBackground('#DB4437').setFontColor('white');
            } else if (item.value >= 2) {
              cell.setBackground('#F4B400');
            } else {
              cell.setBackground('#0F9D58').setFontColor('white');
            }
          });
          
          // 大きなバッチの場合は少し遅延を入れる
          if (batch.length > 5) {
            Utilities.sleep(50);
          }
        }
      }
      
      // 列幅を自動調整
      analysisSheet.autoResizeColumn(1);
      
      // 次のテーブルのための間隔
      currentRowOffset += pivotData.length + 3;
      
      // 大きなテーブル後は少し遅延を入れる
      if (pivotData.length > 10) {
        Utilities.sleep(100);
      }
    }
    
    // 最終的なフォーマット調整（列幅の一括自動調整）
    try {
      analysisSheet.autoResizeColumns(1, sortedMonthDays.length + 1);
    } catch (e) {
      console.error("列幅自動調整エラー:", e);
      // エラーが発生しても続行
    }
    
    // 先頭行を固定
    analysisSheet.setFrozenRows(1);
    analysisSheet.activate();
    
    return {
      success: true,
      message: `分析が完了しました。${duplicateCountTitlesMap.size}種類の重複タイトルパターンを検出しました。`,
      uniqueTitles: duplicateCountTitlesMap.size,
      duplicatePatterns: duplicateCounts.length
    };
  } catch (error) {
    logError('analyzeDuplicateTitles', error);
    return { success: false, message: `エラーが発生しました: ${error.message}` };
  }
}

/**
 * インポートシートのフォーマットを整える関数（シンプル化バージョン）
 * @param {Sheet} sheet - フォーマットするシート
 */
function formatImportSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 1) return; // データがない場合は何もしない
  
  // ヘッダー行の書式設定
  EbayTool.UI.formatSheetHeader(sheet.getRange(1, 1, 1, lastCol));
  
  // 先頭行を固定
  sheet.setFrozenRows(1);
  
  if (lastRow > 1) {
    // データ行の基本的な書式設定
    sheet.getRange(2, 1, lastRow - 1, lastCol).setVerticalAlignment("middle");
    
    // 列の自動サイズ調整（最初の10列のみ）
    const colsToResize = Math.min(lastCol, 10);
    sheet.autoResizeColumns(1, colsToResize);
  }
}

/**
 * エクスポート用CSVを生成する関数（軽量化バージョン）
 * @return {Object} 処理結果
 */
function generateExportCsv() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const duplicateSheet = ss.getSheetByName(EbayTool.getSheetName('DUPLICATES'));
    
    if (!duplicateSheet) {
      return { success: false, message: '重複リストが見つかりません。先に重複検出を実行してください。' };
    }
    
    // ヘッダーのみ取得
    const headers = duplicateSheet.getRange(1, 1, 1, duplicateSheet.getLastColumn()).getValues()[0];
    
    // 必要なカラムのインデックスを探す
    const actionIndex = headers.indexOf('処理');
    let itemIdIndex = -1;
    
    // ItemIDのインデックスを探す（複数の可能性を考慮）
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase();
      if (header.includes('item') && (header.includes('id') || header.includes('number'))) {
        itemIdIndex = i;
        break;
      }
    }
    
    if (actionIndex === -1 || itemIdIndex === -1) {
      return { success: false, message: '必須カラム(処理, Item ID)が見つかりません。' };
    }
    
    // データを効率的に取得（フィルター適用後、該当行のみ）
    const lastRow = duplicateSheet.getLastRow();
    if (lastRow <= 1) {
      // 終了対象のアイテムが0件の場合は正常完了として処理
      return { 
        success: true, 
        message: '終了対象のアイテム: 0件。"終了"指定されたアイテムはありませんでした。',
        itemCount: 0,
        data: [],
        fileName: `ebay_end_items_${new Date().toISOString().substr(0, 10)}.csv`
      };
    }
    
    // 処理カラムの値を取得
    const actionValues = duplicateSheet.getRange(2, actionIndex + 1, lastRow - 1, 1).getValues();
    const itemIdValues = duplicateSheet.getRange(2, itemIdIndex + 1, lastRow - 1, 1).getValues();
    
    // 終了対象のアイテムを抽出（バッチ処理）- EndCode列を追加
    const exportData = [];
    for (let i = 0; i < actionValues.length; i++) {
      if (actionValues[i][0] === '終了' && itemIdValues[i][0]) {
        exportData.push(['End', itemIdValues[i][0], 'OtherListingError']);
      }
    }
    
    if (exportData.length === 0) {
      // 終了対象のアイテムが0件の場合は正常完了として処理
      return { 
        success: true, 
        message: '終了対象のアイテム: 0件。"終了"指定されたアイテムはありませんでした。',
        itemCount: 0,
        data: [],
        fileName: `ebay_end_items_${new Date().toISOString().substr(0, 10)}.csv`
      };
    }
    
    // エクスポートシートを準備
    let exportSheet = ss.getSheetByName(EbayTool.getSheetName('EXPORT'));
    if (!exportSheet) {
      exportSheet = ss.insertSheet(EbayTool.getSheetName('EXPORT'));
    } else {
      exportSheet.clear();
    }
    
    // ヘッダーを設定 - EndCode列を追加
    const exportHeaders = ['Action', 'ItemID', 'EndCode'];
    exportSheet.getRange(1, 1, 1, exportHeaders.length).setValues([exportHeaders]);
    
    // データを書き込み - 3列に対応
    exportSheet.getRange(2, 1, exportData.length, 3).setValues(exportData);
    
    // ヘッダー行の書式設定 - 3列に対応
    exportSheet.getRange(1, 1, 1, 3)
      .setBackground(EbayTool.getColor('PRIMARY'))
      .setFontColor('white')
      .setFontWeight('bold');
    
    return { 
      success: true, 
      message: `${exportData.length}件のアイテムを終了対象としてエクスポートしました。`,
      itemCount: exportData.length,
      data: exportSheet.getDataRange().getValues(),
      fileName: `ebay_end_items_${new Date().toISOString().substr(0, 10)}.csv`
    };
  } catch (error) {
    logError('generateExportCsv', error);
    return { success: false, message: `エラーが発生しました: ${error.message}` };
  }
}

/**
 * スクリプトプロパティの権限をチェックする関数
 * @return {boolean} 権限があるかどうか
 */
function checkScriptPropertiesPermission() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.getProperty('__permission_test__');
    return true;
  } catch (error) {
    console.warn('スクリプトプロパティへのアクセス権限がありません:', error.toString());
    return false;
  }
}

/**
 * 一時ファイルを削除する関数
 */
function deleteTemporaryFile() {
  try {
    // 権限チェック
    if (!checkScriptPropertiesPermission()) {
      console.log('一時ファイル削除: スクリプトプロパティの権限がないため、削除処理をスキップします');
      return;
    }
    
    const fileId = PropertiesService.getScriptProperties().getProperty('TEMP_FILE_ID');
    if (fileId) {
      const file = DriveApp.getFileById(fileId);
      file.setTrashed(true);
      PropertiesService.getScriptProperties().deleteProperty('TEMP_FILE_ID');
    }
  } catch (error) {
    logError('deleteTemporaryFile', error);
  }
}

/**
 * サイドバーを再読み込みする関数（高速化バージョン）
 * @return {Object} 処理結果
 */
function reloadSidebar() {
  try {
    console.log("サイドバーの再読み込みを実行します");
    
    // 最新のHTMLを使用して直接サイドバーを再表示
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('eBay出品管理ツール')
      .setWidth(400);
    
    SpreadsheetApp.getUi().showSidebar(html);
    
    return { success: true };
  } catch (error) {
    console.error("サイドバー再読み込みエラー:", error);
    logError('reloadSidebar', error);
    return { success: false, message: error.message };
  }
}

/**
 * インポートからエクスポートまでを自動処理する関数
 * @param {string} csvData - CSVファイルの内容
 * @return {Object} 処理結果
 */
function autoProcessEbayData(csvData) {
  // 結果オブジェクトの初期化
  const result = {
    success: false,
    steps: [],
    currentStep: '',
    error: null,
    finalMessage: '',
    stats: {},
    startTime: new Date().getTime()
  };
  
  try {
    // ステップ1: CSVインポート
    result.currentStep = 'import';
    console.log("自動処理: CSVインポート開始");
    
    // CSVの行数を概算して進捗状況に表示
    const estimatedRows = csvData.split('\n').length;
    result.stats.estimatedRows = estimatedRows;
    result.stats.importProgress = "CSVデータを解析中... (推定 " + estimatedRows + " 行)";
    
    const importResult = importCsvData(csvData);
    result.steps.push({
      name: 'import',
      success: importResult.success,
      message: importResult.message,
      progressDetail: `${importResult.rowCount || 0}件のデータをインポートしました`
    });
    
    if (!importResult.success) {
      result.error = {
        step: 'import',
        message: importResult.message,
        details: importResult.isFormatError ? importResult.formatDetails : null
      };
      result.finalMessage = "CSVインポートに失敗したため、処理を中止しました。";
      // ログを記録
      logAutoProcess('自動処理（インポート失敗）', result);
      return result;
    }
    
    // インポート成功時の統計情報を保存
    if (importResult.rowCount) {
      result.stats.importedRows = importResult.rowCount;
    }
    
    // 少し遅延を入れてスプレッドシートに反映される時間を確保
    // データ量に応じて遅延時間を調整
    const delayAfterImport = Math.min(800, Math.max(300, Math.floor(estimatedRows / 30)));
    console.log(`インポート後の遅延: ${delayAfterImport}ms`);
    Utilities.sleep(delayAfterImport);
    
    // ステップ2: 重複検出
    result.currentStep = 'detect';
    console.log("自動処理: 重複検出開始");
    result.stats.detectProgress = `${result.stats.importedRows || 0}件のデータから重複を検索中...`;
    
    const detectResult = detectDuplicates();
    result.steps.push({
      name: 'detect',
      success: detectResult.success,
      message: detectResult.message,
      progressDetail: detectResult.success ? 
        `${detectResult.duplicateGroups || 0}件の重複グループを検出しました` : 
        '重複検出に失敗しました'
    });
    
    if (!detectResult.success) {
      result.error = {
        step: 'detect',
        message: detectResult.message
      };
      result.finalMessage = "重複検出に失敗したため、処理を中止しました。";
      // ログを記録
      logAutoProcess('自動処理（重複検出失敗）', result);
      return result;
    }
    
    // 重複検出成功時の統計情報を保存
    if (detectResult.duplicateGroups) {
      result.stats.duplicateGroups = detectResult.duplicateGroups;
      result.stats.duplicateItems = detectResult.duplicateItems;
    }
    
    // 少し遅延を入れてスプレッドシートに反映される時間を確保
    // 重複グループの数に応じて遅延時間を調整
    const duplicateGroups = detectResult.duplicateGroups || 0;
    const delayAfterDetect = Math.min(800, Math.max(300, duplicateGroups * 5));
    console.log(`重複検出後の遅延: ${delayAfterDetect}ms`);
    Utilities.sleep(delayAfterDetect);
    
    // ステップ3: 分析の実行
    result.currentStep = 'analyze';
    console.log("自動処理: 分析開始");
    result.stats.analyzeProgress = `${result.stats.duplicateGroups || 0}件の重複グループを分析中...`;
    
    // 分析シートを削除して再作成（完全に初期化）
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const analysisSheetName = EbayTool.getSheetName('ANALYSIS');
      let analysisSheet = ss.getSheetByName(analysisSheetName);
      
      if (analysisSheet) {
        console.log("分析前に分析シートを完全に初期化します");
        // シートを削除して再作成
        const sheetIndex = analysisSheet.getIndex();
        ss.deleteSheet(analysisSheet);
        ss.insertSheet(analysisSheetName, sheetIndex - 1);
      }
    } catch (e) {
      console.error("分析シート初期化エラー:", e);
      // エラーが発生しても処理は続行
    }
    
    const analyzeResult = analyzeDuplicateTitles();
    result.steps.push({
      name: 'analyze',
      success: analyzeResult.success,
      message: analyzeResult.message,
      progressDetail: analyzeResult.success ? 
        `${analyzeResult.uniqueTitles || 0}種類の重複パターンを分析しました` : 
        '分析に失敗しました'
    });
    
    // 分析は失敗しても処理を続行（オプション機能として扱う）
    if (analyzeResult.success) {
      result.stats.uniqueTitles = analyzeResult.uniqueTitles;
      result.stats.duplicatePatterns = analyzeResult.duplicatePatterns;
    }
    
    // 少し遅延を入れてスプレッドシートに反映される時間を確保
    // 分析後は固定の短い遅延で十分
    Utilities.sleep(500);
    
    // ステップ4: CSVエクスポート
    result.currentStep = 'export';
    console.log("自動処理: CSVエクスポート開始");
    result.stats.exportProgress = `${result.stats.duplicateGroups || 0}件の重複グループからCSVを生成中...`;
    
    const exportResult = generateExportCsv();
    result.steps.push({
      name: 'export',
      success: exportResult.success,
      message: exportResult.message,
      progressDetail: exportResult.success ? 
        `${exportResult.itemCount || 0}件のアイテムをエクスポートしました` : 
        'エクスポート処理に失敗しました'
    });
    
    if (!exportResult.success) {
      result.error = {
        step: 'export',
        message: exportResult.message
      };
      result.finalMessage = "CSVエクスポートに失敗したため、処理を中止しました。";
      // ログを記録
      logAutoProcess('自動処理（エクスポート失敗）', result);
      return result;
    }
    
    // エクスポート成功時の統計情報を保存
    if (exportResult.itemCount) {
      result.stats.exportCount = exportResult.itemCount;
    }
    
    // 全ステップが成功
    result.success = true;
    result.endTime = new Date().getTime();
    result.processingTime = (result.endTime - result.startTime) / 1000; // 秒単位
    
    // 分析結果も含めたメッセージ
    const analyzeMessage = analyzeResult.success ? 
      `${result.stats.uniqueTitles || 0}種類の重複パターンを分析し、` : '';
    
    // 重複数に応じたメッセージ生成
    const duplicateCount = result.stats.duplicateGroups || 0;
    const exportCount = result.stats.exportCount || 0;
    
    if (duplicateCount === 0) {
      result.finalMessage = `処理が完了しました: ${result.stats.importedRows || 0}件のデータを分析した結果、重複する商品は見つかりませんでした。(処理時間: ${result.processingTime.toFixed(1)}秒)`;
    } else {
      result.finalMessage = `処理が完了しました: ${result.stats.importedRows || 0}件のデータから${duplicateCount}件の重複グループを検出し、${analyzeMessage}${exportCount}件のアイテムをエクスポートしました。(処理時間: ${result.processingTime.toFixed(1)}秒)`;
    }
    result.data = exportResult.data;
    result.fileName = exportResult.fileName;
    result.currentStep = 'complete';
    
    console.log("自動処理: 全処理完了");
    
    // 成功ログを記録
    logAutoProcess('自動処理（完了）', result);
    
    return result;
    
  } catch (error) {
    console.error("autoProcessEbayData関数でエラー:", error);
    
    // エラーが発生した時点での情報を返す
    const errorResult = { 
      success: false, 
      steps: result.steps,
      currentStep: result.currentStep || 'unknown',
      error: {
        step: result.currentStep || 'unknown',
        message: error.message,
        stack: error.stack
      },
      finalMessage: `自動処理中にエラーが発生しました: ${error.message}`,
      stats: result.stats,
      endTime: new Date().getTime(),
      processingTime: (new Date().getTime() - result.startTime) / 1000
    };
    
    // エラーログを記録
    logAutoProcess('自動処理（エラー）', errorResult);
    
    return errorResult;
  }
}

/**
 * 自動処理で生成したCSVをダウンロードする関数
 * @param {Array} data - CSVデータの2次元配列
 * @param {string} fileName - ダウンロードされるファイル名
 * @return {Object} 処理結果（HTML出力）
 */
function downloadAutoProcessedCsv(data, fileName) {
  try {
    console.log("自動処理CSVダウンロード開始: 行数=", data ? data.length : 0);
    
    if (!data || !Array.isArray(data) || data.length === 0) {
      throw new Error("ダウンロードするデータがありません。");
    }
    
    // ファイル名が指定されていない場合はデフォルト名を使用
    const finalFileName = fileName || "ebay_終了リスト_" + new Date().toISOString().split('T')[0] + ".csv";
    
    // CSVデータを直接生成
    let csvContent = data.map(row => 
      row.map(cell => {
        // null/undefinedの処理
        if (cell === null || cell === undefined) {
          return '';
        }
        
        // 文字列に変換
        let cellStr = String(cell);
        
        // セキュリティ上の問題となりうる文字をエスケープ
        cellStr = cellStr
          .replace(/"/g, '""') // 引用符のエスケープ
          .replace(/\\/g, '\\\\'); // バックスラッシュのエスケープ
        
        // カンマ、引用符、改行、タブを含む場合は引用符で囲む
        if (/[,"\n\r\t]/.test(cellStr)) {
          return '"' + cellStr + '"';
        }
        
        return cellStr;
      }).join(',')
    ).join('\n');
    
    // BOMを追加してUTF-8として認識されるようにする
    const bom = '\ufeff';
    csvContent = bom + csvContent;
    
    // ダウンロード用HTMLを生成 - FileSaver.js版
    const html = HtmlService.createHtmlOutput(
      `<html>
        <head>
          <base target="_top">
          <meta charset="UTF-8">
          <title>CSVダウンロード</title>
          <script>
            // FileSaver.js - クロスブラウザでの保存機能を提供するライブラリ
            (function(a,b){if("function"==typeof define&&define.amd)define([],b);else if("undefined"!=typeof exports)b();else{b(),a.FileSaver={exports:{}}.exports}})(this,function(){"use strict";function b(a,b){return"undefined"==typeof b?b={autoBom:!1}:"object"!=typeof b&&(console.warn("Deprecated: Expected third argument to be a object"),b={autoBom:!b}),b.autoBom&&/^\\s*(?:text\\/\\S*|application\\/xml|\\S*\\/\\S*\\+xml)\\s*;.*charset\\s*=\\s*utf-8/i.test(a.type)?new Blob([String.fromCharCode(65279),a],{type:a.type}):a}function c(b,c,d){var e=new XMLHttpRequest;e.open("GET",b),e.responseType="blob",e.onload=function(){a(e.response,c,d)},e.onerror=function(){console.error("could not download file")},e.send()}function d(a){var b=new XMLHttpRequest;b.open("HEAD",a,!1);try{b.send()}catch(a){}return 200<=b.status&&299>=b.status}function e(a){try{a.dispatchEvent(new MouseEvent("click"))}catch(c){var b=document.createEvent("MouseEvents");b.initMouseEvent("click",!0,!0,window,0,0,0,80,20,!1,!1,!1,!1,0,null),a.dispatchEvent(b)}}var f="object"==typeof window&&window.window===window?window:"object"==typeof self&&self.self===self?self:"object"==typeof global&&global.global===global?global:void 0,a=f.saveAs||("object"!=typeof window||window!==f?function(){}:"download"in HTMLAnchorElement.prototype?function(b,g,h){var i=f.URL||f.webkitURL,j=document.createElement("a");g=g||b.name||"download",j.download=g,j.rel="noopener","string"==typeof b?(j.href=b,j.origin===location.origin?e(j):d(j.href)?c(b,g,h):e(j,j.target="_blank")):(j.href=i.createObjectURL(b),setTimeout(function(){i.revokeObjectURL(j.href)},4E4),setTimeout(function(){e(j)},0))}:"msSaveOrOpenBlob"in navigator?function(f,g,h){if(g=g||f.name||"download","string"!=typeof f)navigator.msSaveOrOpenBlob(b(f,h),g);else if(d(f))c(f,g,h);else{var i=document.createElement("a");i.href=f,i.target="_blank",setTimeout(function(){e(i)})}}:function(a,b,d,e){if(e=e||open("","_blank"),e&&(e.document.title=e.document.body.innerText="downloading..."),"string"==typeof a)return c(a,b,d);var g="application/octet-stream"===a.type,h=/constructor/i.test(f.HTMLElement)||f.safari,i=/CriOS\\/[\\d]+/.test(navigator.userAgent);if((i||g&&h)&&"object"==typeof FileReader){var j=new FileReader;j.onloadend=function(){var a=j.result;a=i?a:a.replace(/^data:[^;]*;/,"data:attachment/file;"),e?e.location.href=a:location=a,e=null},j.readAsDataURL(a)}else{var k=f.URL||f.webkitURL,l=k.createObjectURL(a);e?e.location=l:location.href=l,e=null,setTimeout(function(){k.revokeObjectURL(l)},4E4)}});f.saveAs=a.saveAs=a,"undefined"!=typeof module&&(module.exports=a)});

            // CSVデータ
            const csvData = \`${csvContent.replace(/`/g, '\\`')}\`;
            
            // ページ読み込み完了時の処理
            document.addEventListener('DOMContentLoaded', function() {
              // ステータス表示を初期化
              document.getElementById('status').innerHTML = 
                '<div class="info">ダウンロードを準備しています...</div>';
              
              // 直接ダウンロード関数
              function directDownload() {
                try {
                  document.getElementById('status').innerHTML = 
                    '<div class="success-message">ダウンロードを開始しています...</div>';
                  
                  // Blobオブジェクトの作成
                  const blob = new Blob([csvData], {type: 'text/csv;charset=utf-8;'});
                  
                  // FileSaver.jsを使用して直接ダウンロード
                  saveAs(blob, '${finalFileName.replace(/'/g, "\\'")}');
                  
                  // 成功メッセージを表示
                  document.getElementById('status').innerHTML = 
                    '<div class="success-message">ダウンロードが完了しました！<br>3秒後にこのダイアログは自動的に閉じます。</div>';
                  
                  // ユーザーに通知するためのアラート表示
                  alert('CSVファイルのダウンロードが完了しました！');
                  
                  // 親ウィンドウ（サイドバー）に通知してダウンロード完了メッセージを表示
                  try {
                    window.parent.postMessage({
                      type: 'download-complete',
                      fileName: '${finalFileName.replace(/'/g, "\\'")}'
                    }, '*');
                  } catch (err) {
                    console.error('親ウィンドウへの通知エラー:', err);
                  }
                  
                  // ダウンロードボタンを無効化
                  const downloadBtn = document.getElementById('downloadBtn');
                  if (downloadBtn) {
                    downloadBtn.disabled = true;
                    downloadBtn.classList.add('disabled');
                  }
                  
                  // 3秒後にダイアログを閉じる
                  setTimeout(function() {
                    google.script.host.close();
                  }, 3000);
                } catch (e) {
                  console.error('直接ダウンロードエラー:', e);
                  document.getElementById('status').innerHTML = 
                    '<div class="error-message">エラーが発生しました: ' + e.message + '<br>別のダウンロード方法をお試しください。</div>';
                }
              }
              
              // ダウンロードボタンにイベントリスナーを追加
              const downloadBtn = document.getElementById('downloadBtn');
              if (downloadBtn) {
                downloadBtn.addEventListener('click', directDownload);
              }
              
              // 自動ダウンロード開始（1秒遅延）
              setTimeout(function() {
                directDownload();
              }, 1000);
            });
          </script>
          <style>
            body {
              font-family: Arial, sans-serif;
              margin: 0;
              padding: 0;
              background-color: #f9fafb;
              color: #1f2937;
              font-size: 14px;
            }
            .container {
              max-width: 600px;
              margin: 0 auto;
              padding: 20px;
              background: white;
              border-radius: 8px;
              box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            h3 {
              color: #4F46E5;
              margin-top: 0;
              padding-bottom: 10px;
              border-bottom: 1px solid #e5e7eb;
            }
            .file-info {
              background-color: #EFF6FF;
              padding: 12px;
              border-radius: 6px;
              margin-bottom: 15px;
              font-size: 13px;
            }
            .file-info p {
              margin: 5px 0;
            }
            .button {
              background-color: #4F46E5;
              color: white;
              border: none;
              border-radius: 4px;
              padding: 8px 16px;
              font-size: 14px;
              cursor: pointer;
              transition: background-color 0.2s;
              display: inline-block;
              text-decoration: none;
              margin-top: 10px;
            }
            .button:hover {
              background-color: #4338CA;
            }
            .button.disabled {
              background-color: #9CA3AF;
              cursor: not-allowed;
            }
            .success-message {
              color: #10B981;
              background-color: #D1FAE5;
              padding: 10px;
              border-radius: 4px;
              margin: 10px 0;
            }
            .error-message {
              color: #EF4444;
              background-color: #FEE2E2;
              padding: 10px;
              border-radius: 4px;
              margin: 10px 0;
            }
            .info {
              color: #3B82F6;
              background-color: #EFF6FF;
              padding: 10px;
              border-radius: 4px;
              margin: 10px 0;
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h3>CSVファイルのダウンロード</h3>
            
            <div class="file-info">
              <p><strong>ファイル名:</strong> ${finalFileName}</p>
              <p><strong>行数:</strong> ${data.length}行</p>
            </div>
            
            <div id="status" class="info">準備中...</div>
            
            <button id="downloadBtn" class="button">ダウンロード</button>
          </div>
        </body>
      </html>`
    )
    .setWidth(600)
    .setHeight(450);
    
    return html;
  } catch (error) {
    console.error("downloadAutoProcessedCsv関数でエラー:", error);
    return {
      success: false,
      message: "CSVダウンロード準備中にエラーが発生しました: " + error.message
    };
  }
}

/**
 * 自動処理のログを記録する関数
 * @param {string} operation - 操作名
 * @param {Object} result - 処理結果
 */
function logAutoProcess(operation, result) {
  try {
    // ログシートの存在確認と作成
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName("操作ログ");
    
    if (!logSheet) {
      // ログシートが存在しない場合は作成
      logSheet = ss.insertSheet("操作ログ");
      // ヘッダー行の設定
      logSheet.appendRow([
        "操作日時", 
        "操作内容", 
        "ステータス", 
        "処理時間(秒)", 
        "データ件数", 
        "詳細情報"
      ]);
      
      // ヘッダー行の書式設定
      logSheet.getRange(1, 1, 1, 6).setBackground("#f3f4f6").setFontWeight("bold");
    }
    
    // 現在時刻
    const timestamp = new Date();
    
    // ステータス（成功/失敗）
    const status = result.success ? "成功" : "失敗";
    
    // 処理時間
    const processingTime = result.processingTime || 0;
    
    // データ件数（インポート件数、または処理件数）
    let dataCount = "";
    if (result.stats) {
      if (result.stats.importedRows) {
        dataCount = result.stats.importedRows + "件";
      } else if (result.stats.exportCount) {
        dataCount = result.stats.exportCount + "件";
      }
    }
    
    // 詳細情報（エラーメッセージなど）
    let details = result.finalMessage || "";
    if (!result.success && result.error) {
      details += " エラー: " + (result.error.message || "不明なエラー");
    }
    
    // ログに追加
    logSheet.appendRow([
      timestamp,
      operation,
      status,
      processingTime.toFixed(1),
      dataCount,
      details
    ]);
    
    // 最新の行を強調表示
    const lastRow = logSheet.getLastRow();
    if (result.success) {
      logSheet.getRange(lastRow, 1, 1, 6).setBackground("#f0fdf4");  // 薄い緑色（成功）
    } else {
      logSheet.getRange(lastRow, 1, 1, 6).setBackground("#fef2f2");  // 薄い赤色（失敗）
    }
    
  } catch (error) {
    console.error("logAutoProcess関数でエラー:", error);
  }
}

/**
 * バージョン情報を取得する関数
 * サイドバーからの呼び出しに対応
 */
function getVersion() {
  return EbayTool.getVersion();
}

/**
 * 権限状況をチェックする関数
 * @return {Object} 権限チェック結果
 */
function checkAllPermissions() {
  const result = {
    overall: true,
    permissions: {},
    errors: [],
    warnings: []
  };

  try {
    // PropertiesService (UserProperties) の権限チェック
    try {
      const userProperties = PropertiesService.getUserProperties();
      userProperties.getProperty('__permission_test__');
      result.permissions.userProperties = { success: true, hasPermission: true };
    } catch (error) {
      const isPermissionError = error.toString().indexOf('PERMISSION_DENIED') !== -1;
      result.permissions.userProperties = { 
        success: false, 
        hasPermission: !isPermissionError,
        error: error.toString()
      };
      if (isPermissionError) {
        result.errors.push('ユーザープロパティへのアクセス権限がありません');
        result.overall = false;
      } else {
        result.warnings.push('ユーザープロパティアクセスエラー: ' + error.toString());
      }
    }

    // PropertiesService (ScriptProperties) の権限チェック
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.getProperty('__permission_test__');
      result.permissions.scriptProperties = { success: true, hasPermission: true };
    } catch (error) {
      const isPermissionError = error.toString().indexOf('PERMISSION_DENIED') !== -1;
      result.permissions.scriptProperties = { 
        success: false, 
        hasPermission: !isPermissionError,
        error: error.toString()
      };
      if (isPermissionError) {
        result.errors.push('スクリプトプロパティへのアクセス権限がありません');
        result.overall = false;
      } else {
        result.warnings.push('スクリプトプロパティアクセスエラー: ' + error.toString());
      }
    }

    // SpreadsheetApp の権限チェック
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      spreadsheet.getName(); // 基本的なアクセステスト
      result.permissions.spreadsheet = { success: true, hasPermission: true };
    } catch (error) {
      const isPermissionError = error.toString().indexOf('PERMISSION_DENIED') !== -1;
      result.permissions.spreadsheet = { 
        success: false, 
        hasPermission: !isPermissionError,
        error: error.toString()
      };
      if (isPermissionError) {
        result.errors.push('スプレッドシートへのアクセス権限がありません');
        result.overall = false;
      } else {
        result.warnings.push('スプレッドシートアクセスエラー: ' + error.toString());
      }
    }

    // DriveApp の権限チェック（オプション）
    try {
      DriveApp.getRootFolder(); // 基本的なアクセステスト
      result.permissions.drive = { success: true, hasPermission: true };
    } catch (error) {
      const isPermissionError = error.toString().indexOf('PERMISSION_DENIED') !== -1;
      result.permissions.drive = { 
        success: false, 
        hasPermission: !isPermissionError,
        error: error.toString()
      };
      if (isPermissionError) {
        result.warnings.push('Driveへのアクセス権限がありません（一部機能が制限されます）');
      } else {
        result.warnings.push('Driveアクセスエラー: ' + error.toString());
      }
    }

  } catch (error) {
    logError('checkAllPermissions', error);
    result.overall = false;
    result.errors.push('権限チェック中に予期しないエラーが発生しました: ' + error.toString());
  }

  return result;
}

/**
 * エラーログを記録する関数
 * EbayTool.Logger.errorのラッパー関数
 * @param {string} functionName - エラーが発生した関数名
 * @param {Error} error - エラーオブジェクト
 * @param {string} context - エラーのコンテキスト（オプション）
 * @return {Object} ログ情報
 */
function logError(functionName, error, context = '') {
  try {
    // EbayTool.Logger.errorが存在する場合はそれを使用
    if (EbayTool && EbayTool.Logger && typeof EbayTool.Logger.error === 'function') {
      return EbayTool.Logger.error(functionName, error, context);
    }
    
    // そうでない場合は最小限のロギング
    console.error(`[${functionName}] エラー:`, error);
    if (context) console.error(`コンテキスト: ${context}`);
    
    return {
      timestamp: new Date(),
      function: functionName,
      type: error.name || 'Error',
      message: error.message || String(error),
      context: context,
      stack: error.stack || '利用不可'
    };
  } catch (e) {
    // ログ処理自体のエラーは無視（再帰を防ぐ）
    console.error('logError関数内でエラー:', e);
    return null;
  }
}

/**
 * ユーザーにわかりやすいエラーメッセージを生成する関数
 * @param {Error} error - エラーオブジェクト
 * @param {string} defaultMessage - デフォルトのエラーメッセージ
 * @return {string} ユーザーフレンドリーなエラーメッセージ
 */
function getFriendlyErrorMessage(error, defaultMessage = 'エラーが発生しました。') {
  try {
    // エラーがnullまたはundefinedの場合
    if (!error) {
      return defaultMessage;
    }
    
    // エラーメッセージの取得
    const errorMessage = error.message || String(error);
    
    // エラータイプに基づいて適切なメッセージを返す
    if (errorMessage.includes('Script has been running too long')) {
      return '処理時間が長すぎたため、タイムアウトしました。データ量を減らして再試行してください。';
    } else if (errorMessage.includes('Out of memory')) {
      return 'メモリ不足エラーが発生しました。データ量を減らして再試行してください。';
    } else if (errorMessage.includes('Authorization')) {
      return '認証エラーが発生しました。再ログインしてから再試行してください。';
    } else if (errorMessage.includes('Access denied') || errorMessage.includes('Permission')) {
      return '権限エラーが発生しました。スプレッドシートの編集権限があることを確認してください。';
    } else if (errorMessage.includes('Limit Exceeded')) {
      return 'Google Sheetsの制限を超えました。データ量を減らすか、しばらく時間をおいてから再試行してください。';
    } else if (errorMessage.includes('Invalid argument')) {
      return '無効な引数が指定されました。入力データを確認してください。';
    }
    
    // その他のエラーはデフォルトメッセージとエラー内容を表示
    return `${defaultMessage} (${errorMessage})`;
  } catch (e) {
    // エラーメッセージ生成中のエラーは無視
    console.error('getFriendlyErrorMessage関数内でエラー:', e);
    return defaultMessage;
  }
}

