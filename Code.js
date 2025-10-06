/**
 * eBay出品管理ツール
 * スプレッドシート上でeBayの重複出品を検出・管理するツール
 * 最終更新: 2025-10-01 - バージョン管理統一とUS絞り込み最適化
 */

// EbayTool名前空間 - 拡張版
var EbayTool = (function() {
  // プライベート変数と定数
  const CONFIG = {
    VERSION: '1.6.43',
    SHEET_NAMES: {
      IMPORT: 'インポートデータ',
      DUPLICATES: '重複リスト',
      EXPORT: 'エクスポート',
      ANALYSIS: '分析',
      LOG: 'ログ',
      PROCESS_STATE: '処理状態',
      PERFORMANCE: '性能ログ'
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
    MAX_EXECUTION_TIME: 330000, // 最大実行時間(5.5分)
    SAFETY_MARGIN: 30000, // 安全マージン(30秒)
    CHUNK_SIZE: 1000, // 分割処理時のチャンクサイズ
    
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
  const startTime = new Date().getTime();
  let dataRows = 0;
  let fileSizeMB = 0;

  try {
    fileSizeMB = csvData ? Math.round(csvData.length / 1024 / 1024 * 100) / 100 : 0;
    console.log(`🚀 [${new Date().toLocaleTimeString()}] 高速CSVインポート開始: データサイズ=${csvData ? csvData.length : 0}バイト (${fileSizeMB}MB)`);

    if (!csvData || typeof csvData !== 'string' || csvData.trim() === '') {
      // 失敗ログを記録
      logPerformance('CSVインポート', startTime, new Date().getTime(), {
        success: false,
        errorMessage: 'CSVデータが空または無効',
        fileSizeMB: fileSizeMB,
        dataRows: 0
      });
      return { success: false, message: 'CSVデータが空または無効です。' };
    }

    // 手動インポートを模倣した超高速処理
    try {
      // 1. 最小限の前処理
      if (csvData.charCodeAt(0) === 0xFEFF) {
        csvData = csvData.substring(1); // BOM除去
      }

      // 2. 単純な行分割（引用符処理は最小限）
      const lines = csvData.split(/\r?\n/);
      console.log(`行分割完了: ${lines.length}行`);

      // 3. 引用符対応CSV分割
      console.log(`📋 [${new Date().toLocaleTimeString()}] 引用符対応CSVパース開始`);
      const csvRows = [];
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line) {
          // 引用符内カンマを適切に処理
          csvRows.push(parseCSVLine(line));
        }
      }

      if (csvRows.length <= 1) {
        return { success: false, message: 'CSVデータが不十分です。' };
      }

      console.log(`✅ [${new Date().toLocaleTimeString()}] CSV引用符対応パース完了: ${csvRows.length}行 x ${csvRows[0].length}列`);

      // 4. データ品質確認
      const qualityCheck = validateCSVQuality(csvRows);
      if (!qualityCheck.isValid) {
        console.warn(`⚠️  データ品質問題検出: ${qualityCheck.issues.join(', ')}`);
      } else {
        console.log(`✅ データ品質確認完了: 問題なし`);
      }

      // 5. 列数統一（必要最小限）
      const headerLength = csvRows[0].length;
      for (let i = 1; i < csvRows.length; i++) {
        const row = csvRows[i];
        if (row.length !== headerLength) {
          // 短い行は空文字で埋める、長い行は切り詰める
          csvRows[i] = row.slice(0, headerLength).concat(
            Array(Math.max(0, headerLength - row.length)).fill('')
          );
        }
      }

      console.log(`列数統一完了: 全${csvRows.length}行を${headerLength}列に統一`);
      dataRows = csvRows.length - 1; // ヘッダー除く

      // 5. Google Sheetsの最適化API使用（一括書き込み）
      console.log(`📝 [${new Date().toLocaleTimeString()}] シート書き込み開始: ${csvRows.length}行`);

      let result;
      try {
        result = writeToSheetOptimized(csvRows);

        // 成功時の処理
        if (result.success) {
          const elapsedSeconds = ((new Date().getTime() - startTime) / 1000).toFixed(1);
          console.log(`✅ [${new Date().toLocaleTimeString()}] CSVインポート成功: ${dataRows}行を${elapsedSeconds}秒で処理`);

          logPerformance('CSVインポート', startTime, new Date().getTime(), {
            success: true,
            fileSizeMB: fileSizeMB,
            dataRows: dataRows,
            elapsedSeconds: parseFloat(elapsedSeconds),
            additionalInfo: {
              totalRows: csvRows.length,
              columns: headerLength,
              method: '引用符対応高速インポート',
              dataQuality: qualityCheck.isValid ? '良好' : `問題あり: ${qualityCheck.issues.join(', ')}`,
              columnMismatchCount: qualityCheck.stats.columnMismatchCount,
              avgEmptyFields: qualityCheck.stats.avgEmptyFields.toFixed(1)
            }
          });
        }

        return result;

      } catch (writeError) {
        // タイムアウト等のエラー時も性能ログを記録
        const elapsedSeconds = ((new Date().getTime() - startTime) / 1000).toFixed(1);
        console.error(`⚠️ [${new Date().toLocaleTimeString()}] シート書き込みタイムアウト: ${elapsedSeconds}秒経過 - ${writeError.message}`);

        logPerformance('CSVインポート', startTime, new Date().getTime(), {
          success: false,
          errorMessage: `タイムアウト: ${writeError.message}`,
          fileSizeMB: fileSizeMB,
          dataRows: dataRows,
          elapsedSeconds: parseFloat(elapsedSeconds),
          additionalInfo: {
            totalRows: csvRows.length,
            columns: headerLength,
            method: '高速インポート(タイムアウト)'
          }
        });

        // エラーを再スロー（上位でキャッチされる）
        throw writeError;
      }

    } catch (error) {
      console.error("高速インポートエラー:", error);
      // エラーログを記録
      logPerformance('CSVインポート', startTime, new Date().getTime(), {
        success: false,
        errorMessage: `高速インポートエラー: ${error.message}`,
        fileSizeMB: fileSizeMB,
        dataRows: dataRows,
        additionalInfo: { method: '高速インポート→フォールバック' }
      });
      // フォールバック: 従来方式
      return importCsvDataFallback(csvData);
    }

  } catch (error) {
    console.error("CSVインポート全体エラー:", error);
    // 全体エラーログを記録
    logPerformance('CSVインポート', startTime, new Date().getTime(), {
      success: false,
      errorMessage: `全体エラー: ${error.message}`,
      fileSizeMB: fileSizeMB,
      dataRows: 0
    });
    return { success: false, message: `インポートに失敗しました: ${error.message}` };
  }
}

// 手動インポート模倣: Google Sheets API直接利用
function writeToSheetOptimized(csvRows) {
  try {
    console.log(`🚀 手動インポート模倣開始: ${csvRows.length}行 x ${csvRows[0].length}列`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(EbayTool.getConfig().SHEET_NAMES.IMPORT);

    if (!sheet) {
      sheet = ss.insertSheet(EbayTool.getConfig().SHEET_NAMES.IMPORT);
    } else {
      sheet.clear();
    }

    // **手動インポートと同じ方法: 一括でsetValues実行**
    // バッチ処理、分割処理、休憩を完全廃止

    // データクリーニング（最小限）
    const cleanData = csvRows.map(row =>
      row.map(cell => cell == null ? '' : String(cell))
    );

    console.log(`データクリーニング完了 - 一括書き込み実行`);

    // **一括書き込み実行（手動インポートと同様）**
    const range = sheet.getRange(1, 1, cleanData.length, cleanData[0].length);
    range.setValues(cleanData);

    console.log(`✅ 手動インポート模倣完了: 総${csvRows.length}行を一括書き込み`);

    return {
      success: true,
      message: `CSVインポート完了: ${csvRows.length}行のデータをインポートしました`,
      importedRows: csvRows.length - 1, // ヘッダー除く
      totalRows: csvRows.length
    };

  } catch (error) {
    console.error("❌ 一括書き込みエラー:", error);
    throw error;
  }
}

/**
 * 引用符内カンマを適切に処理するCSVパーサー
 */
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  let i = 0;

  while (i < line.length) {
    const char = line[i];
    const nextChar = line[i + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        // エスケープされた引用符 ("")
        current += '"';
        i += 2;
      } else {
        // 引用符の開始/終了
        inQuotes = !inQuotes;
        i++;
      }
    } else if (char === ',' && !inQuotes) {
      // 引用符外のカンマ = フィールド区切り
      result.push(current.trim());
      current = '';
      i++;
    } else {
      // 通常の文字
      current += char;
      i++;
    }
  }

  // 最後のフィールドを追加
  result.push(current.trim());

  return result;
}

/**
 * CSVデータの品質を確認する関数
 */
function validateCSVQuality(csvRows) {
  const issues = [];
  const headerRow = csvRows[0];
  const expectedColumns = headerRow.length;

  // 1. 列数の一貫性チェック
  let columnMismatchCount = 0;
  for (let i = 1; i < csvRows.length; i++) {
    if (csvRows[i].length !== expectedColumns) {
      columnMismatchCount++;
    }
  }

  if (columnMismatchCount > 0) {
    const percentage = ((columnMismatchCount / (csvRows.length - 1)) * 100).toFixed(1);
    issues.push(`列数不一致: ${columnMismatchCount}行 (${percentage}%)`);
  }

  // 2. 重要な列の存在確認
  const requiredColumns = ['item', 'title', 'site'];
  const headerLower = headerRow.map(h => String(h).toLowerCase().replace(/\s+/g, ''));

  for (const required of requiredColumns) {
    const found = headerLower.some(h => h.includes(required));
    if (!found) {
      issues.push(`必須列不在: '${required}' 関連列が見つかりません`);
    }
  }

  // 3. データサンプル確認（最初の10行）
  let emptyFieldCount = 0;
  const sampleSize = Math.min(10, csvRows.length - 1);

  for (let i = 1; i <= sampleSize; i++) {
    const row = csvRows[i];
    const emptyFields = row.filter(field => !field || field.trim() === '').length;
    emptyFieldCount += emptyFields;
  }

  const avgEmptyFields = emptyFieldCount / sampleSize;
  if (avgEmptyFields > expectedColumns * 0.3) {
    issues.push(`空フィールド多数: 平均${avgEmptyFields.toFixed(1)}個/行`);
  }

  return {
    isValid: issues.length === 0,
    issues: issues,
    stats: {
      totalRows: csvRows.length,
      expectedColumns: expectedColumns,
      columnMismatchCount: columnMismatchCount,
      avgEmptyFields: avgEmptyFields
    }
  };
}

/**
 * インポート状況確認関数（タイムアウト後の確認用）
 */
function checkImportStatus() {
  try {
    console.log('📋 軽量インポート状況確認開始');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(EbayTool.getConfig().SHEET_NAMES.IMPORT);

    if (!sheet) {
      console.log('❌ インポートシートが見つかりません');
      return { hasData: false, rowCount: 0, message: 'インポートシートが見つかりません' };
    }

    // 軽量な確認: 行数と列数のみ取得
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    console.log(`📊 シート情報: ${lastRow}行 x ${lastCol}列`);

    if (lastRow <= 1) {
      console.log('❌ データなし');
      return { hasData: false, rowCount: 0, message: 'インポートデータがありません' };
    }

    // 軽量確認: ヘッダー行の最初の3セルのみチェック
    if (lastCol > 0) {
      const headerSample = sheet.getRange(1, 1, 1, Math.min(3, lastCol)).getValues()[0];
      const hasHeaders = headerSample.some(cell => cell && String(cell).trim().length > 0);

      if (hasHeaders) {
        console.log(`✅ インポート確認成功: ${lastRow}行 x ${lastCol}列のデータを検出`);

        // 軽量ログ: コンソールのみ、性能ログシートへの書き込みは省略
        console.log(`📊 [${new Date().toLocaleTimeString()}] 軽量確認完了`);

        return {
          hasData: true,
          rowCount: lastRow - 1, // ヘッダー除く
          totalRows: lastRow,
          columns: lastCol,
          message: `インポート完了: ${lastRow - 1}行のデータ`
        };
      }
    }

    console.log('❌ 無効なデータ形式');
    return { hasData: false, rowCount: 0, message: '無効なデータ形式' };

  } catch (error) {
    console.error('インポート状況確認エラー:', error);
    return { hasData: false, rowCount: 0, message: `確認エラー: ${error.message}` };
  }
}

/**
 * 超軽量インポート状況確認（最小限の処理のみ）
 */
function checkImportStatusUltraLight() {
  try {
    console.log('🔍 超軽量インポート状況確認開始');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('インポートデータ'); // ハードコードで高速化

    if (!sheet) {
      console.log('❌ インポートシートなし');
      return { hasData: false, rowCount: 0, message: 'シートなし' };
    }

    // 最小限の確認: 行数のみ
    const lastRow = sheet.getLastRow();
    console.log(`📊 行数: ${lastRow}`);

    if (lastRow > 1) {
      console.log(`✅ 超軽量確認成功: ${lastRow - 1}行`);
      return {
        hasData: true,
        rowCount: lastRow - 1,
        totalRows: lastRow,
        message: `データあり: ${lastRow - 1}行`
      };
    }

    console.log('❌ データなし');
    return { hasData: false, rowCount: 0, message: 'データなし' };

  } catch (error) {
    console.error('超軽量確認エラー:', error);
    return { hasData: false, rowCount: 0, message: `エラー: ${error.message}` };
  }
}

// フォールバック用の従来処理（簡略化）
function importCsvDataFallback(csvData) {
  try {
    console.log("フォールバック処理実行");
    // 基本的なCSV処理のみ
    const lines = csvData.replace(/\r\n/g, '\n').split('\n');
    const csvRows = lines.filter(line => line.trim()).map(line => line.split(','));

    if (csvRows.length <= 1) {
      return { success: false, message: 'CSVデータが不十分です。' };
    }

    // フォールバック: 基本的なシート書き込み
    return writeToSheetOptimized(csvRows);

  } catch (error) {
    console.error("フォールバック処理エラー:", error);
    return {
      success: false,
      message: `フォールバック処理に失敗しました: ${error.message}`
    };
  }
}

/**
 * インポートデータから重複を検出する関数（最適化チャンク処理対応）
 * @return {Object} 処理結果
 */
function detectDuplicates() {
  const startTime = new Date().getTime();

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
      return { 
        success: true, 
        message: '検出された重複: 0件。重複データはありませんでした。',
        duplicateCount: 0,
        analysisComplete: true
      };
    }
    
    // データサイズに応じた処理方法を選択
    const dataSize = lastRow - 1;
    console.log(`重複検出開始: ${dataSize} 行のデータを処理します`);
    
    // 大規模データ（15,000行以上）の場合はチャンク処理
    let result;
    if (dataSize >= 15000) {
      console.log('大規模データ検出: チャンク処理を実行します');
      result = detectDuplicatesChunked(importSheet, lastRow, lastCol);
    } else {
      console.log('通常処理を実行します');
      result = detectDuplicatesStandard(importSheet, lastRow, lastCol);
    }

    // 性能ログを記録
    logPerformance('重複検出', startTime, new Date().getTime(), {
      success: result.success,
      dataRows: dataSize,
      errorMessage: result.success ? '' : result.message,
      additionalInfo: {
        duplicateGroups: result.duplicateGroups || 0,
        duplicateItems: result.duplicateItems || 0,
        method: dataSize >= 15000 ? 'チャンク処理' : '通常処理'
      }
    });

    return result;
    
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
 * 列インデックスを検索するヘルパー関数
 */
function findColumnIndices(headers) {
  const headersLower = headers.map(h => String(h).toLowerCase().replace(/\s+/g, ''));
  
  let titleIndex = -1;
  let itemIdIndex = -1;
  let startDateIndex = -1;
  
  // タイトル列を探す
  for (let i = 0; i < headers.length; i++) {
    const headerLower = headersLower[i];
    if (headerLower.includes('title') || headerLower.includes('name')) {
      titleIndex = i;
      break;
    }
  }
  
  // ID列を探す
  for (let i = 0; i < headers.length; i++) {
    const headerLower = headersLower[i];
    if (headerLower.includes('item') || headerLower.includes('id') || headerLower.includes('number')) {
      itemIdIndex = i;
      break;
    }
  }
  
  // 開始日列を探す
  for (let i = 0; i < headers.length; i++) {
    const headerLower = headersLower[i];
    if (headerLower.includes('date') || headerLower.includes('start')) {
      startDateIndex = i;
      break;
    }
  }
  
  // デフォルト値を設定
  if (titleIndex === -1 && headers.length > 1) titleIndex = 1;
  if (itemIdIndex === -1 && headers.length > 0) itemIdIndex = 0;
  if (startDateIndex === -1 && headers.length > 2) startDateIndex = 2;
  
  return { titleIndex, itemIdIndex, startDateIndex };
}

/**
 * タイトルでグループ化するヘルパー関数
 */
function groupByTitle(allData, titleIndex, itemIdIndex, startDateIndex) {
  const titleGroups = {};
  
  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const title = EbayTool.TextAnalyzer.normalizeTitle(String(row[titleIndex] || ''), false);
    const itemId = String(row[itemIdIndex] || '').trim();
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
    }
  }
  
  return titleGroups;
}

/**
 * 重複シートを準備するヘルパー関数
 */
function prepareDuplicateSheet(ss, sheetName) {
  let duplicateSheet = ss.getSheetByName(sheetName);
  if (!duplicateSheet) {
    duplicateSheet = ss.insertSheet(sheetName);
  } else {
    duplicateSheet.clear();
  }
  return duplicateSheet;
}

/**
 * 標準的な重複検出処理（小～中規模データ用）
 */
function detectDuplicatesStandard(importSheet, lastRow, lastCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAMES = EbayTool.getConfig().SHEET_NAMES;
  
  // ヘッダーを取得
  const headers = importSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const { titleIndex, itemIdIndex, startDateIndex } = findColumnIndices(headers);
  
  if (titleIndex === -1 || itemIdIndex === -1) {
    return { success: false, message: '必須カラム(タイトル、ID)が見つかりません。' };
  }
  
  console.log(`重複検出に使用する列: title=${titleIndex} (${headers[titleIndex]}), itemId=${itemIdIndex} (${headers[itemIdIndex]}), startDate=${startDateIndex} (${headers[startDateIndex] || 'N/A'})`);
  
  // 全データを取得
  const allData = importSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  // タイトルでグループ化
  const titleGroups = groupByTitle(allData, titleIndex, itemIdIndex, startDateIndex);
  
  // 重複グループのみを抽出
  const duplicateGroups = Object.values(titleGroups)
    .filter(group => group.length > 1)
    .sort((a, b) => b.length - a.length);
  
  // 重複リストシートを準備・作成
  const duplicateSheet = prepareDuplicateSheet(ss, SHEET_NAMES.DUPLICATES);
  createDuplicateListSheet(duplicateSheet, duplicateGroups, headers);
  
  ss.setActiveSheet(duplicateSheet);
  
  return { 
    success: true, 
    message: `${duplicateGroups.length}件の重複グループを検出しました。合計${getTotalDuplicates(duplicateGroups)}件の重複アイテムがあります。`,
    duplicateGroups: duplicateGroups.length,
    duplicateItems: getTotalDuplicates(duplicateGroups)
  };
}

/**
 * チャンク処理による重複検出（大規模データ用）
 */
function detectDuplicatesChunked(importSheet, lastRow, lastCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAMES = EbayTool.getConfig().SHEET_NAMES;
  const startTime = new Date().getTime();
  const MAX_EXECUTION_TIME = 330000; // 5.5分
  
  // ヘッダーを取得
  const headers = importSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const { titleIndex, itemIdIndex, startDateIndex } = findColumnIndices(headers);
  
  if (titleIndex === -1 || itemIdIndex === -1) {
    return { success: false, message: '必須カラム(タイトル、ID)が見つかりません。' };
  }
  
  console.log(`チャンク処理開始: ${lastRow-1} 行のデータを処理します`);
  console.log(`重複検出に使用する列: title=${titleIndex} (${headers[titleIndex]}), itemId=${itemIdIndex} (${headers[itemIdIndex]}), startDate=${startDateIndex} (${headers[startDateIndex] || 'N/A'})`);
  
  const dataSize = lastRow - 1;
  const CHUNK_SIZE = Math.min(3000, Math.max(1000, Math.floor(dataSize / 15))); // 動的チャンクサイズ
  console.log(`チャンクサイズ: ${CHUNK_SIZE}`);
  
  const titleGroups = {};
  let processedRows = 0;
  
  // チャンクごとに処理
  for (let startRow = 2; startRow <= lastRow; startRow += CHUNK_SIZE) {
    const currentTime = new Date().getTime();
    if (currentTime - startTime > MAX_EXECUTION_TIME) {
      console.log('実行時間制限に近づいたため処理を中断');
      return { 
        success: false, 
        message: `タイムアウトが発生しました。${processedRows}/${dataSize} 行まで処理済み。データを分割して再実行してください。` 
      };
    }
    
    const endRow = Math.min(startRow + CHUNK_SIZE - 1, lastRow);
    const chunkSize = endRow - startRow + 1;
    
    console.log(`チャンク処理中: 行 ${startRow}-${endRow} (${chunkSize} 行)`);
    
    // チャンクデータを取得
    const chunkData = importSheet.getRange(startRow, 1, chunkSize, lastCol).getValues();
    
    // チャンク内でタイトルグループ化
    for (let i = 0; i < chunkData.length; i++) {
      const row = chunkData[i];
      const title = EbayTool.TextAnalyzer.normalizeTitle(String(row[titleIndex] || ''), false);
      const itemId = String(row[itemIdIndex] || '').trim();
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
      }
    }
    
    processedRows += chunkSize;
    
    // 進捗ログ（5チャンクごと）
    if ((startRow - 2) / CHUNK_SIZE % 5 === 0) {
      const progress = Math.round((processedRows / dataSize) * 100);
      console.log(`進捗: ${progress}% (${processedRows}/${dataSize} 行処理済み)`);
    }
    
    // メモリ使用量軽減のため短時間待機
    Utilities.sleep(10);
  }
  
  console.log('タイトルグループ化完了。重複抽出中...');
  
  // 重複グループのみを抽出
  const duplicateGroups = Object.values(titleGroups)
    .filter(group => group.length > 1)
    .sort((a, b) => b.length - a.length);
  
  console.log(`${duplicateGroups.length} 件の重複グループを検出`);
  
  // 重複リストシートを準備・作成
  const duplicateSheet = prepareDuplicateSheet(ss, SHEET_NAMES.DUPLICATES);
  
  // 大量データの場合はチャンク化して書き込み
  if (duplicateGroups.length > 100) {
    createDuplicateListSheetChunked(duplicateSheet, duplicateGroups, headers);
  } else {
    createDuplicateListSheet(duplicateSheet, duplicateGroups, headers);
  }
  
  ss.setActiveSheet(duplicateSheet);
  
  const processingTime = Math.round((new Date().getTime() - startTime) / 1000);
  console.log(`チャンク処理完了: ${processingTime} 秒`);
  
  return { 
    success: true, 
    message: `${duplicateGroups.length}件の重複グループを検出しました。合計${getTotalDuplicates(duplicateGroups)}件の重複アイテムがあります。（処理時間: ${processingTime}秒）`,
    duplicateGroups: duplicateGroups.length,
    duplicateItems: getTotalDuplicates(duplicateGroups)
  };
}

/**
 * 重複リストシートを作成（チャンク処理版）
 */
function createDuplicateListSheetChunked(sheet, duplicateGroups, headers) {
  console.log(`重複リストシート作成開始: ${duplicateGroups.length} グループ`);
  
  // ヘッダーを設定（従来版と同じ形式）
  const duplicateHeaders = ['グループID', '重複タイプ', '処理'].concat(headers);
  sheet.getRange(1, 1, 1, duplicateHeaders.length).setValues([duplicateHeaders]);
  
  let currentRow = 2;
  const BATCH_SIZE = 1000; // 書き込み単位
  let batchData = [];
  
  // グループごとに処理
  for (let groupIndex = 0; groupIndex < duplicateGroups.length; groupIndex++) {
    const group = duplicateGroups[groupIndex];
    
    // グループをスタート日でソート（従来版と同じ）
    group.sort((a, b) => {
      if (!a.startDate || !b.startDate) return 0;
      const dateA = new Date(a.startDate);
      const dateB = new Date(b.startDate);
      return dateB - dateA; // 新しい順（降順）でソート
    });
    
    for (let itemIndex = 0; itemIndex < group.length; itemIndex++) {
      const item = group[itemIndex];
      const row = new Array(duplicateHeaders.length).fill('');
      
      // グループIDと重複タイプの列を設定（従来版と同じ形式）
      row[0] = `Group ${groupIndex + 1}`;                        // グループID
      row[1] = `${group.length}件中${itemIndex + 1}件目`;       // 重複タイプ
      row[2] = itemIndex === 0 ? '残す' : '終了';                // 処理（最新のみ残す）
      
      // 元のデータを3番目以降に配置（従来版と同じ）
      item.allData.forEach((value, i) => {
        row[i + 3] = value;
      });
      
      batchData.push(row);
      
      // バッチサイズに達したら書き込み
      if (batchData.length >= BATCH_SIZE) {
        sheet.getRange(currentRow, 1, batchData.length, duplicateHeaders.length).setValues(batchData);
        currentRow += batchData.length;
        batchData = [];
        console.log(`重複リスト書き込み中: ${currentRow - 2} 行完了`);
      }
    }
  }
  
  // 残りのデータを書き込み
  if (batchData.length > 0) {
    sheet.getRange(currentRow, 1, batchData.length, duplicateHeaders.length).setValues(batchData);
    console.log(`重複リスト作成完了: ${currentRow + batchData.length - 2} 行`);
  }
  
  // ヘッダー行のフォーマット
  const headerRange = sheet.getRange(1, 1, 1, duplicateHeaders.length);
  EbayTool.UI.formatSheetHeader(headerRange);
}

/**
 * 重複リストシートの構造をデバッグする関数
 * @return {Object} デバッグ情報
 */
function debugDuplicateSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const duplicateSheet = ss.getSheetByName(EbayTool.getSheetName('DUPLICATES'));
    
    if (!duplicateSheet) {
      return { success: false, message: '重複リストシートが見つかりません' };
    }
    
    const lastRow = duplicateSheet.getLastRow();
    const lastCol = duplicateSheet.getLastColumn();
    
    if (lastRow <= 0) {
      return { success: false, message: '重複リストシートにデータがありません' };
    }
    
    // ヘッダー行を取得
    const headers = duplicateSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // 処理列のインデックスを探す
    const actionIndex = headers.indexOf('処理');
    
    // サンプルデータを取得（最初の5行）
    const sampleData = lastRow > 1 ? 
      duplicateSheet.getRange(1, 1, Math.min(6, lastRow), lastCol).getValues() : 
      [headers];
    
    return {
      success: true,
      sheetInfo: {
        lastRow: lastRow,
        lastCol: lastCol,
        headers: headers,
        actionIndex: actionIndex,
        hasActionColumn: actionIndex !== -1,
        sampleData: sampleData
      }
    };
  } catch (error) {
    return { success: false, message: `デバッグエラー: ${error.message}` };
  }
}

/**
 * エクスポート処理をデバッグする関数
 * @return {Object} デバッグ情報
 */
function debugExportProcess() {
  try {
    // 重複リストシートの構造をチェック
    const sheetDebug = debugDuplicateSheet();
    if (!sheetDebug.success) {
      return sheetDebug;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const duplicateSheet = ss.getSheetByName(EbayTool.getSheetName('DUPLICATES'));
    const headers = sheetDebug.sheetInfo.headers;
    const actionIndex = sheetDebug.sheetInfo.actionIndex;
    const lastRow = sheetDebug.sheetInfo.lastRow;
    
    // ItemIDのインデックスを探す
    let itemIdIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase();
      if (header.includes('item') && (header.includes('id') || header.includes('number'))) {
        itemIdIndex = i;
        break;
      }
    }
    
    let endItemsCount = 0;
    let sampleEndItems = [];
    
    if (actionIndex !== -1 && itemIdIndex !== -1 && lastRow > 1) {
      // 処理列とItemID列の値を取得
      const actionValues = duplicateSheet.getRange(2, actionIndex + 1, lastRow - 1, 1).getValues();
      const itemIdValues = duplicateSheet.getRange(2, itemIdIndex + 1, lastRow - 1, 1).getValues();
      
      // 終了対象のアイテムをカウント・サンプル取得
      for (let i = 0; i < actionValues.length; i++) {
        if (actionValues[i][0] === '終了' && itemIdValues[i][0]) {
          endItemsCount++;
          if (sampleEndItems.length < 5) {
            sampleEndItems.push({
              row: i + 2,
              action: actionValues[i][0],
              itemId: itemIdValues[i][0]
            });
          }
        }
      }
    }
    
    return {
      success: true,
      debugInfo: {
        ...sheetDebug.sheetInfo,
        itemIdIndex: itemIdIndex,
        hasItemIdColumn: itemIdIndex !== -1,
        endItemsCount: endItemsCount,
        sampleEndItems: sampleEndItems
      }
    };
  } catch (error) {
    return { success: false, message: `エクスポートデバッグエラー: ${error.message}` };
  }
}

/**
 * US出品のみに絞り込む専用関数（eBay Mag対応）
 * @return {Object} 処理結果
 */
/**
 * 連続する行番号を範囲にグループ化
 */
function groupConsecutiveRows(rowNumbers) {
  if (rowNumbers.length === 0) return [];

  const sorted = [...rowNumbers].sort((a, b) => a - b);
  const ranges = [];
  let start = sorted[0];
  let end = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === end + 1) {
      // 連続している
      end = sorted[i];
    } else {
      // 連続が途切れた
      ranges.push({ start: start, end: end });
      start = sorted[i];
      end = sorted[i];
    }
  }

  // 最後の範囲を追加
  ranges.push({ start: start, end: end });

  return ranges;
}

function filterUSOnly() {
  const startTime = new Date().getTime();
  let originalRowCount = 0;
  let filteredRowCount = 0;

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
      return { success: true, message: 'データが見つかりません。' };
    }

    // ヘッダーを取得してサイト列を特定
    const headers = importSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headersLower = headers.map(h => String(h).toLowerCase().replace(/\s+/g, ''));

    // AA列（リスティングサイト）を探す
    let listingSiteIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      const headerLower = headersLower[i];
      if (headerLower.includes('site') || headerLower.includes('listing') || i === 26) { // AA列は26番目（0ベース）
        listingSiteIndex = i;
        break;
      }
    }

    if (listingSiteIndex === -1) {
      return { success: false, message: 'リスティングサイト列（AA列）が見つかりません。' };
    }

    console.log(`🚀 [${new Date().toLocaleTimeString()}] 高速US絞り込み開始: 列${listingSiteIndex + 1} (${headers[listingSiteIndex]})`);
    console.log(`📊 処理対象: ${lastRow - 1}行のデータ`);
    originalRowCount = lastRow - 1;

    // 一括削除方式による最適化処理
    try {
      console.log(`📖 [${new Date().toLocaleTimeString()}] データ読み込み開始: ${lastRow}行 × ${lastCol}列`);

      // 1. 全データを一度だけ読み込み
      const allData = importSheet.getRange(1, 1, lastRow, lastCol).getValues();

      console.log(`🔍 [${new Date().toLocaleTimeString()}] 削除対象行を特定中...`);

      // 2. 削除対象行を特定
      const rowsToDelete = [];
      for (let i = 1; i < allData.length; i++) { // ヘッダー行をスキップ
        const siteValue = String(allData[i][listingSiteIndex]).trim().toUpperCase();
        if (siteValue && siteValue !== 'US' && siteValue !== 'USA' && siteValue !== 'UNITED STATES') {
          rowsToDelete.push(i + 1); // 1ベース行番号
        }
      }

      console.log(`🎯 削除対象特定完了: ${rowsToDelete.length}行を削除予定`);

      // 7. 連続行範囲の一括削除（超高速化）
      if (rowsToDelete.length > 0) {
        console.log(`📋 [${new Date().toLocaleTimeString()}] ${rowsToDelete.length}行を一括削除開始...`);

        // 連続する行範囲をグループ化
        const ranges = groupConsecutiveRows(rowsToDelete);
        console.log(`📦 連続行範囲: ${ranges.length}グループに分割`);

        // 下から上へ一括削除（範囲ごと）
        for (let i = ranges.length - 1; i >= 0; i--) {
          const range = ranges[i];
          const rowCount = range.end - range.start + 1;

          console.log(`🗑️  範囲削除: ${range.start}-${range.end}行 (${rowCount}行)`);

          // 一括削除実行
          importSheet.deleteRows(range.start, rowCount);

          // 大量削除時は小休憩
          if (rowCount > 1000) {
            Utilities.sleep(50);
          }
        }

        console.log(`✅ [${new Date().toLocaleTimeString()}] 一括削除完了: ${rowsToDelete.length}行削除`);
      }

      // 8. 結果を計算
      const newLastRow = importSheet.getLastRow();
      filteredRowCount = newLastRow - 1;
      const deletedCount = originalRowCount - filteredRowCount;

      const elapsedSeconds = ((new Date().getTime() - startTime) / 1000).toFixed(1);
      console.log(`✅ [${new Date().toLocaleTimeString()}] US絞り込み完了: ${originalRowCount} → ${filteredRowCount} (${deletedCount}行削除) - ${elapsedSeconds}秒`);

      // 成功ログを記録
      logPerformance('US絞り込み', startTime, new Date().getTime(), {
        success: true,
        dataRows: filteredRowCount,
        elapsedSeconds: parseFloat(elapsedSeconds),
        additionalInfo: {
          originalRows: originalRowCount,
          filteredRows: filteredRowCount,
          deletedRows: deletedCount,
          method: '一括削除方式',
          rangeGroups: rowsToDelete.length > 0 ? groupConsecutiveRows(rowsToDelete).length : 0
        }
      });

      return {
        success: true,
        message: `US絞り込み完了: ${originalRowCount}件 → ${filteredRowCount}件 (${deletedCount}行削除)`,
        originalCount: originalRowCount,
        filteredCount: filteredRowCount,
        deletedCount: deletedCount
      };

    } catch (filterError) {
      console.warn('フィルター方式でエラー発生、従来方式にフォールバック:', filterError);

      // フォールバック: 従来の方式
      const allData = importSheet.getRange(1, 1, lastRow, lastCol).getValues();
      const rowsToDelete = [];

      for (let i = 1; i < allData.length; i++) {
        const siteValue = String(allData[i][listingSiteIndex]).trim().toUpperCase();
        if (siteValue && siteValue !== 'US' && siteValue !== 'USA' && siteValue !== 'UNITED STATES') {
          rowsToDelete.push(i + 1);
        }
      }

      // 下から上へ削除
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        importSheet.deleteRow(rowsToDelete[i]);
      }

      const newLastRow = importSheet.getLastRow();
      const filteredRowCount = newLastRow - 1;
      const deletedCount = originalRowCount - filteredRowCount;

      return {
        success: true,
        message: `US絞り込み完了: ${originalRowCount}件 → ${filteredRowCount}件 (${deletedCount}行削除)`,
        originalCount: originalRowCount,
        filteredCount: filteredRowCount,
        deletedCount: deletedCount
      };
    }

  } catch (error) {
    logError('filterUSOnly', error, 'US絞り込み処理中');
    SpreadsheetApp.getUi().alert(
      'エラー',
      getFriendlyErrorMessage(error, 'US絞り込み中にエラーが発生しました。'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return {
      success: false,
      message: getFriendlyErrorMessage(error, 'US絞り込み中にエラーが発生しました。'),
      stack: error.stack
    };
  }
}

/**
 * アプリケーションの状態をチェックする関数
 * @return {Object} アプリケーションの状態情報
 */
function checkAppState() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const SHEET_NAMES = EbayTool.getConfig().SHEET_NAMES;
    
    // 各シートの存在確認
    const importSheet = ss.getSheetByName(SHEET_NAMES.IMPORT);
    const duplicateSheet = ss.getSheetByName(SHEET_NAMES.DUPLICATES);
    const exportSheet = ss.getSheetByName(SHEET_NAMES.EXPORT);
    
    const state = {
      hasImportSheet: importSheet !== null,
      hasDuplicateSheet: duplicateSheet !== null,
      hasExportSheet: exportSheet !== null,
      stats: {}
    };
    
    // インポートデータの統計
    if (importSheet) {
      const lastRow = importSheet.getLastRow();
      if (lastRow > 1) {
        state.stats.rowCount = lastRow - 1; // ヘッダーを除く
      }
    }
    
    // 重複データの統計
    if (duplicateSheet) {
      const lastRow = duplicateSheet.getLastRow();
      if (lastRow > 1) {
        // 重複グループの数を計算
        const groupData = duplicateSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        const uniqueGroups = new Set();
        let totalItems = 0;
        
        groupData.forEach(row => {
          const groupId = row[0];
          if (groupId) {
            uniqueGroups.add(groupId);
            totalItems++;
          }
        });
        
        state.stats.duplicateGroups = uniqueGroups.size;
        state.stats.duplicateItems = totalItems;
      }
    }
    
    // エクスポートデータの統計
    if (exportSheet) {
      const lastRow = exportSheet.getLastRow();
      if (lastRow > 1) {
        state.stats.exportCount = lastRow - 1;
      }
    }
    
    return state;
    
  } catch (error) {
    logError('checkAppState', error, 'アプリ状態確認中');
    return {
      hasImportSheet: false,
      hasDuplicateSheet: false,
      hasExportSheet: false,
      stats: {},
      error: getFriendlyErrorMessage(error, 'アプリ状態確認中にエラーが発生しました。')
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
    let exportSheet = ss.getSheetByName(EbayTool.getSheetName('EXPORT'));

    // エクスポートシートが存在しない場合は先に生成する
    if (!exportSheet) {
      console.log('エクスポートシートが見つかりません。generateExportCsvを実行します。');
      const generateResult = generateExportCsv();

      if (!generateResult.success) {
        return generateResult;
      }

      // 生成後に再取得
      exportSheet = ss.getSheetByName(EbayTool.getSheetName('EXPORT'));
      if (!exportSheet) {
        throw new Error('エクスポートシートの生成に失敗しました。');
      }
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
      try {
        // スプレッドシートに最低1つのシートは必要なので、最後のシートは削除しない
        if (ss.getSheets().length > 1) {
          ss.deleteSheet(sheet);
          console.log(`シート「${sheet.getName()}」を削除しました`);
        } else {
          console.log(`シート「${sheet.getName()}」は最後のシートのため削除をスキップしました`);
          sheet.clear(); // 代わりにクリア
        }
      } catch (error) {
        console.error(`シート「${sheet.getName()}」の削除中にエラー:`, error.message);
        // エラーが発生しても処理を続行
      }
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

    // デバッグ情報を出力
    console.log(`*** generateExportCsv デバッグ ***`);
    console.log(`処理列インデックス: ${actionIndex}, ItemID列インデックス: ${itemIdIndex}`);
    console.log(`データ行数: ${actionValues.length}`);
    console.log(`最初の5行の処理値:`, actionValues.slice(0, 5).map(row => `"${row[0]}"`));

    let endCount = 0;
    for (let i = 0; i < actionValues.length; i++) {
      const actionValue = actionValues[i][0];
      const itemIdValue = itemIdValues[i][0];

      if (actionValue === '終了' && itemIdValue) {
        exportData.push(['End', itemIdValue, 'OtherListingError']);
        endCount++;
      }

      // 最初の10行をデバッグ出力
      if (i < 10) {
        console.log(`行${i+2}: 処理="${actionValue}" ItemID="${itemIdValue}" 判定=${actionValue === '終了' && itemIdValue ? 'エクスポート対象' : 'スキップ'}`);
      }
    }

    console.log(`終了対象として抽出されたアイテム数: ${endCount}`);
    console.log(`*** generateExportCsv デバッグ終了 ***`);
    
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
 * 性能ログを記録する関数
 * @param {string} operation - 操作名
 * @param {number} startTime - 開始時刻
 * @param {number} endTime - 終了時刻
 * @param {Object} details - 詳細情報
 */
function logPerformance(operation, startTime, endTime, details = {}) {
  try {
    console.log(`性能ログ記録開始: ${operation}`);
    const duration = endTime - startTime;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = EbayTool.getConfig().SHEET_NAMES.PERFORMANCE;
    console.log(`性能ログシート名: ${sheetName}`);
    let perfSheet = ss.getSheetByName(sheetName);

    if (!perfSheet) {
      // 性能ログシートを作成
      perfSheet = ss.insertSheet(EbayTool.getConfig().SHEET_NAMES.PERFORMANCE);

      // ヘッダーを設定
      const headers = [
        '実行日時', '操作名', '処理時間(秒)', '処理時間(ミリ秒)',
        'データ行数', 'ファイルサイズ(MB)', '成功/失敗', 'エラーメッセージ', '詳細'
      ];
      perfSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // ヘッダーの書式設定
      perfSheet.getRange(1, 1, 1, headers.length)
        .setBackground(EbayTool.getColor('PRIMARY'))
        .setFontColor('white')
        .setFontWeight('bold');

      // 列幅を調整
      perfSheet.setColumnWidth(1, 150); // 実行日時
      perfSheet.setColumnWidth(2, 120); // 操作名
      perfSheet.setColumnWidth(3, 100); // 処理時間(秒)
      perfSheet.setColumnWidth(4, 120); // 処理時間(ミリ秒)
      perfSheet.setColumnWidth(8, 200); // エラーメッセージ
      perfSheet.setColumnWidth(9, 300); // 詳細

      perfSheet.setFrozenRows(1);
    }

    // ログデータを準備
    const logData = [
      new Date(),                                          // 実行日時
      operation,                                           // 操作名
      Math.round(duration / 1000 * 100) / 100,           // 処理時間(秒)
      duration,                                           // 処理時間(ミリ秒)
      details.dataRows || '',                             // データ行数
      details.fileSizeMB || '',                           // ファイルサイズ(MB)
      details.success ? '成功' : '失敗',                   // 成功/失敗
      details.errorMessage || '',                         // エラーメッセージ
      JSON.stringify(details.additionalInfo || {})       // 詳細
    ];

    // ログを追加
    const lastRow = perfSheet.getLastRow();
    perfSheet.getRange(lastRow + 1, 1, 1, logData.length).setValues([logData]);

    // 行の色分け（成功=緑、失敗=赤）
    const logRow = perfSheet.getRange(lastRow + 1, 1, 1, logData.length);
    if (details.success) {
      logRow.setBackground('#F0FDF4'); // 薄い緑
    } else {
      logRow.setBackground('#FEF2F2'); // 薄い赤
    }

    console.log(`性能ログ記録: ${operation} - ${Math.round(duration/1000*100)/100}秒`);

    // 古いログの削除（500行を超えた場合）
    if (lastRow > 500) {
      const deleteCount = lastRow - 500;
      perfSheet.deleteRows(2, deleteCount); // ヘッダーを除いて削除
    }

  } catch (error) {
    console.error('性能ログ記録エラー:', error);
    // エラーがあってもメイン処理に影響させない
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

/**
 * 分割処理マネージャー - タイムアウト対策
 */
var ChunkedProcessor = {
  // メモリ内状態ストレージ
  memoryStorage: {},
  /**
   * 処理状態を保存（強化版）
   */
  saveState: function(processId, state) {
    try {
      const stateData = {
        ...state,
        lastUpdated: new Date().getTime(),
        backupCount: (state.backupCount || 0) + 1
      };
      
      // まずメモリに保存
      this.memoryStorage[processId] = stateData;
      console.log(`メモリ内状態保存成功: ${processId} (バックアップ${stateData.backupCount}回目)`);
      
      // CacheServiceにもバックアップ保存を試行
      try {
        const cache = CacheService.getScriptCache();
        cache.put(`process_${processId}`, JSON.stringify(stateData), 3600); // 1時間
        console.log(`キャッシュ状態保存成功: ${processId}`);
      } catch (cacheError) {
        console.warn('キャッシュ保存失敗（メモリ保存は成功）:', cacheError.message);
      }
      
      // シートにも永続化保存を試行（新機能）
      try {
        this.saveStateToSheet(processId, stateData);
        console.log(`シート状態保存完了: ${processId}`);
      } catch (sheetError) {
        console.error('シート保存失敗:', sheetError.message);
        console.error('シート保存エラー詳細:', sheetError);
      }
      
      return true;
    } catch (error) {
      console.error('処理状態保存エラー:', error);
      return false;
    }
  },
  
  /**
   * 処理状態をシートに保存
   */
  saveStateToSheet: function(processId, stateData) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let stateSheet = ss.getSheetByName(EbayTool.getSheetName('PROCESS_STATE'));
      
      if (!stateSheet) {
        stateSheet = ss.insertSheet(EbayTool.getSheetName('PROCESS_STATE'));
        // ヘッダー設定
        stateSheet.getRange(1, 1, 1, 3).setValues([['ProcessID', 'State', 'LastUpdated']]);
      }
      
      // 既存の状態を検索
      const lastRow = stateSheet.getLastRow();
      let targetRow = -1;
      
      if (lastRow > 1) {
        const processIds = stateSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < processIds.length; i++) {
          if (processIds[i][0] === processId) {
            targetRow = i + 2;
            break;
          }
        }
      }
      
      // データの保存
      const rowData = [processId, JSON.stringify(stateData), new Date().getTime()];
      if (targetRow !== -1) {
        // 更新
        stateSheet.getRange(targetRow, 1, 1, 3).setValues([rowData]);
      } else {
        // 新規追加
        stateSheet.getRange(lastRow + 1, 1, 1, 3).setValues([rowData]);
      }
      
      console.log(`シート状態保存成功: ${processId}`);
    } catch (error) {
      throw new Error(`シート保存エラー: ${error.message}`);
    }
  },
  
  /**
   * 処理状態を取得（強化版）
   */
  getState: function(processId) {
    try {
      console.log(`状態取得開始: ${processId} - メモリ内状態数: ${Object.keys(this.memoryStorage).length}`);
      
      // まずメモリから取得を試行
      if (this.memoryStorage[processId]) {
        console.log(`メモリ内状態取得成功: ${processId}`);
        return this.memoryStorage[processId];
      }
      console.log(`メモリ内に状態なし: ${processId}`);
      
      // メモリにない場合はキャッシュから取得を試行
      try {
        const cache = CacheService.getScriptCache();
        const stateJson = cache.get(`process_${processId}`);
        if (stateJson) {
          const state = JSON.parse(stateJson);
          // メモリにも復元
          this.memoryStorage[processId] = state;
          console.log(`キャッシュ状態取得成功: ${processId}`);
          return state;
        }
      } catch (cacheError) {
        console.warn('キャッシュ取得失敗:', cacheError.message);
      }
      
      // キャッシュにもない場合はシートから取得を試行（新機能）
      try {
        console.log(`シートから状態取得を試行: ${processId}`);
        const sheetState = this.getStateFromSheet(processId);
        if (sheetState) {
          console.log(`シート状態発見: ${processId}`);
          // メモリとキャッシュにも復元
          this.memoryStorage[processId] = sheetState;
          try {
            const cache = CacheService.getScriptCache();
            cache.put(`process_${processId}`, JSON.stringify(sheetState), 3600);
            console.log(`キャッシュに復元: ${processId}`);
          } catch (cacheError) {
            console.warn('キャッシュ復元失敗:', cacheError.message);
          }
          console.log(`シート状態取得成功: ${processId}`);
          return sheetState;
        } else {
          console.log(`シートにも状態なし: ${processId}`);
        }
      } catch (sheetError) {
        console.error('シート取得失敗:', sheetError.message);
        console.error('シート取得エラー詳細:', sheetError);
      }
      
      console.log(`処理状態が見つかりません: ${processId}`);
      return null;
    } catch (error) {
      console.error('処理状態取得エラー:', error);
      return null;
    }
  },
  
  /**
   * 処理状態をシートから取得
   */
  getStateFromSheet: function(processId) {
    try {
      console.log(`シートから状態を検索開始: ${processId}`);
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const stateSheet = ss.getSheetByName(EbayTool.getSheetName('PROCESS_STATE'));
      
      if (!stateSheet) {
        console.log(`処理状態シートが存在しません`);
        return null;
      }
      
      const lastRow = stateSheet.getLastRow();
      console.log(`処理状態シートの行数: ${lastRow}`);
      if (lastRow <= 1) {
        console.log(`処理状態シートにデータなし`);
        return null;
      }
      
      // プロセスIDを検索
      const data = stateSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      console.log(`シートデータ行数: ${data.length}`);
      for (let i = 0; i < data.length; i++) {
        console.log(`行${i + 2}: ID="${data[i][0]}" vs "${processId}"`);
        if (data[i][0] === processId) {
          const stateJson = data[i][1];
          console.log(`状態データ発見: ${processId}`);
          return JSON.parse(stateJson);
        }
      }
      
      console.log(`シートに該当する状態データなし: ${processId}`);
      return null;
    } catch (error) {
      console.error(`シート取得エラー詳細:`, error);
      throw new Error(`シート取得エラー: ${error.message}`);
    }
  },
  
  /**
   * 処理状態をクリア
   */
  clearState: function(processId) {
    try {
      // メモリから削除
      if (this.memoryStorage[processId]) {
        delete this.memoryStorage[processId];
        console.log(`メモリ内状態クリア: ${processId}`);
      }
      
      // キャッシュからも削除を試行
      try {
        const cache = CacheService.getScriptCache();
        cache.remove(`process_${processId}`);
        console.log(`キャッシュ状態クリア: ${processId}`);
      } catch (cacheError) {
        console.warn('キャッシュクリア失敗:', cacheError.message);
      }
      
      return true;
    } catch (error) {
      console.error('処理状態クリアエラー:', error);
      return false;
    }
  },
  
  /**
   * 実行時間をチェック
   */
  checkExecutionTime: function(startTime, maxTime = 330000) {
    const elapsed = new Date().getTime() - startTime;
    const remaining = maxTime - elapsed;
    const safetyMargin = 60000; // 60秒（安全マージン拡張）
    return {
      elapsed: elapsed,
      remaining: remaining,
      shouldStop: remaining < safetyMargin,
      progress: Math.min(100, (elapsed / maxTime) * 100)
    };
  },
  
  /**
   * 分割された自動処理を開始
   */
  startChunkedAutoProcess: function(csvData) {
    const processId = `auto_process_${new Date().getTime()}`;
    const startTime = new Date().getTime();
    
    console.log(`*** CHUNKED PROCESSOR DEBUG START ***`);
    console.log(`処理ID生成: ${processId}`);
    console.log(`開始時刻: ${startTime}`);
    console.log(`CSVデータサイズ: ${csvData ? csvData.length : 'null'}`);
    
    try {
      // 初期状態を設定
      const initialState = {
        processId: processId,
        phase: 'import',
        startTime: startTime,
        csvData: csvData,
        totalPhases: 4,
        currentPhase: 1,
        completed: false,
        error: null,
        result: {
          success: false,
          steps: [],
          currentStep: '',
          stats: {},
          startTime: startTime
        }
      };
      
      console.log(`分割処理初期化: ${processId}`);
      console.log(`初期状態オブジェクト作成完了`);
      console.log(`*** 状態保存を実行します ***`);
      const saveResult = this.saveState(processId, initialState);
      console.log(`状態保存結果: ${saveResult}`);
      if (!saveResult) {
        throw new Error('初期状態の保存に失敗しました');
      }
      console.log(`初期状態保存成功: ${processId}`);
      
      // 最初のフェーズを実行
      console.log(`*** executeNextPhaseを呼び出します ***`);
      return this.executeNextPhase(processId);
      
    } catch (error) {
      console.error('分割自動処理開始エラー:', error);
      return {
        success: false,
        message: `処理開始に失敗しました: ${error.message}`,
        processId: processId
      };
    }
  },
  
  /**
   * 次のフェーズを実行
   */
  executeNextPhase: function(processId) {
    const startTime = new Date().getTime();
    const state = this.getState(processId);
    
    console.log(`executeNextPhase実行: ${processId}`);
    if (!state) {
      console.error(`状態取得失敗: ${processId}`);
      return {
        success: false,
        message: '処理状態が見つかりません',
        processId: processId
      };
    }
    console.log(`状態取得成功: ${processId}, フェーズ: ${state.phase}`);
    
    try {
      let phaseResult = null;
      
      switch (state.phase) {
        case 'import':
          phaseResult = this.executeImportPhase(state, startTime);
          break;
        case 'detect':
          phaseResult = this.executeDetectPhase(state, startTime);
          break;
        case 'analyze':
          phaseResult = this.executeAnalyzePhase(state, startTime);
          break;
        case 'export':
          phaseResult = this.executeExportPhase(state, startTime);
          break;
        case 'completed':
          return {
            success: true,
            message: '全ての処理が完了しました',
            completed: true,
            processId: processId,
            result: state.result
          };
        default:
          throw new Error(`未知のフェーズ: ${state.phase}`);
      }
      
      // 実行時間をチェック
      const timeCheck = this.checkExecutionTime(startTime);
      
      if (timeCheck.shouldStop && !phaseResult.completed) {
        // タイムアウト前に処理を一時停止
        console.log(`時間制限に近づいたため処理を一時停止: ${timeCheck.elapsed}ms経過`);
        return {
          success: true,
          message: `処理を一時停止しました (フェーズ: ${state.phase})`,
          processId: processId,
          paused: true,
          progress: Math.round((state.currentPhase / state.totalPhases) * 100)
        };
      }
      
      if (phaseResult.completed) {
        // 現在のフェーズが完了
        state.currentPhase++;
        state.result.steps.push(phaseResult.step);
        
        if (state.currentPhase > state.totalPhases) {
          // 全フェーズ完了
          state.phase = 'completed';
          state.completed = true;
          state.result.success = true;
          state.result.finalMessage = phaseResult.finalMessage || '全ての処理が正常に完了しました';
        } else {
          // 次のフェーズに進む
          const phases = ['import', 'detect', 'analyze', 'export'];
          state.phase = phases[state.currentPhase - 1];
        }
        
        this.saveState(processId, state);
        
        if (state.completed) {
          this.clearState(processId);
          return {
            success: true,
            message: '全ての処理が完了しました',
            completed: true,
            processId: processId,
            result: state.result
          };
        }
        
        // 次のフェーズを即座に開始（時間が許せば）
        if (!timeCheck.shouldStop) {
          return this.executeNextPhase(processId);
        }
      }
      
      return {
        success: true,
        message: `フェーズ「${state.phase}」を実行中`,
        processId: processId,
        progress: Math.round((state.currentPhase / state.totalPhases) * 100)
      };
      
    } catch (error) {
      console.error(`フェーズ実行エラー (${state.phase}):`, error);
      state.error = error.message;
      this.saveState(processId, state);
      
      return {
        success: false,
        message: `処理中にエラーが発生しました: ${error.message}`,
        processId: processId,
        error: error
      };
    }
  },
  
  /**
   * インポートフェーズを実行
   */
  executeImportPhase: function(state, startTime) {
    console.log('インポートフェーズ開始');
    
    try {
      const importResult = importCsvData(state.csvData);
      
      return {
        completed: true,
        step: {
          name: 'import',
          success: importResult.success,
          message: importResult.message,
          progressDetail: `${importResult.rowCount || 0}件のデータをインポートしました`
        },
        finalMessage: importResult.success ? null : "CSVインポートに失敗したため、処理を中止しました。"
      };
    } catch (error) {
      throw new Error(`インポートフェーズでエラー: ${error.message}`);
    }
  },
  
  /**
   * 重複検出フェーズを実行（マイクロチャンク対応）
   */
  executeDetectPhase: function(state, startTime) {
    console.log('重複検出フェーズ開始（マイクロチャンク処理）');
    
    try {
      // チャンク処理状態の初期化
      if (!state.detectState) {
        state.detectState = {
          chunkSize: 800, // さらに縮小（1500→800）
          processedRows: 0,
          totalRows: 0,
          duplicateGroups: 0,
          completed: false
        };
      }
      
      const timeCheck = this.checkExecutionTime(startTime);
      
      // マイクロチャンク重複検出を実行
      const chunkResult = this.executeDetectChunk(state, startTime);
      
      if (chunkResult.shouldPause) {
        // 時間制限により一時停止
        this.saveState(state.processId, state);
        return {
          completed: false,
          paused: true,
          step: {
            name: 'detect',
            success: true,
            message: `重複検出中（${state.detectState.processedRows}/${state.detectState.totalRows}行処理済み）`,
            progressDetail: `進捗: ${Math.round((state.detectState.processedRows / state.detectState.totalRows) * 100)}%`
          }
        };
      }
      
      if (chunkResult.completed) {
        // 重複検出完了
        return {
          completed: true,
          step: {
            name: 'detect',
            success: chunkResult.success,
            message: chunkResult.message,
            progressDetail: chunkResult.success ? 
              `${chunkResult.duplicateGroups || 0}件の重複グループを検出しました` : 
              '重複検出に失敗しました'
          },
          finalMessage: chunkResult.success ? null : "重複検出に失敗したため、処理を中止しました。"
        };
      }
      
      // 継続処理
      return {
        completed: false,
        step: {
          name: 'detect',
          success: true,
          message: `重複検出継続中（${state.detectState.processedRows}/${state.detectState.totalRows}行）`,
          progressDetail: `進捗: ${Math.round((state.detectState.processedRows / state.detectState.totalRows) * 100)}%`
        }
      };
      
    } catch (error) {
      throw new Error(`重複検出フェーズでエラー: ${error.message}`);
    }
  },

  /**
   * 重複検出のマイクロチャンク処理
   */
  executeDetectChunk: function(state, startTime) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const importSheet = ss.getSheetByName(EbayTool.getSheetName('IMPORT'));
      
      if (!importSheet) {
        return {
          completed: true,
          success: false,
          message: 'インポートシートが見つかりません'
        };
      }
      
      // 初回実行時の初期化
      if (state.detectState.totalRows === 0) {
        const lastRow = importSheet.getLastRow();
        state.detectState.totalRows = Math.max(0, lastRow - 1); // ヘッダー除く
        state.detectState.processedRows = 0;
        
        if (state.detectState.totalRows === 0) {
          return {
            completed: true,
            success: true,
            message: '検出された重複: 0件。重複データはありませんでした。',
            duplicateGroups: 0
          };
        }
        
        console.log(`重複検出開始: 総行数 ${state.detectState.totalRows}`);
      }
      
      const chunkSize = state.detectState.chunkSize;
      const startRow = state.detectState.processedRows + 2; // ヘッダー行+1から開始
      const endRow = Math.min(startRow + chunkSize - 1, state.detectState.totalRows + 1);
      const actualChunkSize = endRow - startRow + 1;
      
      if (actualChunkSize <= 0) {
        // 全行処理完了
        console.log(`重複検出完了: ${state.detectState.processedRows}行処理済み`);
        return {
          completed: true,
          success: true,
          message: `検出された重複: ${state.detectState.duplicateGroups}件の重複グループが見つかりました。`,
          duplicateGroups: state.detectState.duplicateGroups
        };
      }
      
      console.log(`チャンク処理: ${startRow}-${endRow}行 (${actualChunkSize}行)`);
      
      // チャンクデータを取得して重複検出
      const lastCol = importSheet.getLastColumn();
      const chunkData = importSheet.getRange(startRow, 1, actualChunkSize, lastCol).getValues();
      
      // 簡易重複チェック（タイトル列での重複検出）
      const titleColumnIndex = this.findTitleColumn(importSheet);
      const duplicatesFound = this.findDuplicatesInChunk(chunkData, titleColumnIndex);
      
      state.detectState.duplicateGroups += duplicatesFound;
      state.detectState.processedRows += actualChunkSize;
      
      // 実行時間チェック
      const timeCheck = this.checkExecutionTime(startTime);
      if (timeCheck.shouldStop) {
        console.log('時間制限により一時停止');
        return {
          shouldPause: true
        };
      }
      
      // まだ処理が残っている場合は継続
      if (state.detectState.processedRows < state.detectState.totalRows) {
        return {
          completed: false,
          continuing: true
        };
      }
      
      // 全処理完了
      return {
        completed: true,
        success: true,
        message: `検出された重複: ${state.detectState.duplicateGroups}件の重複グループが見つかりました。`,
        duplicateGroups: state.detectState.duplicateGroups
      };
      
    } catch (error) {
      console.error('チャンク処理エラー:', error);
      throw new Error(`チャンク処理でエラー: ${error.message}`);
    }
  },
  
  /**
   * タイトル列を見つける
   */
  findTitleColumn: function(sheet) {
    try {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      for (let i = 0; i < headers.length; i++) {
        const header = String(headers[i]).toLowerCase();
        if (header.includes('title') || header.includes('name') || header.includes('item name')) {
          return i;
        }
      }
      // デフォルトは3列目
      return 3;
    } catch (error) {
      console.warn('タイトル列検出エラー:', error);
      return 3;
    }
  },
  
  /**
   * チャンク内での重複を検出
   */
  findDuplicatesInChunk: function(chunkData, titleIndex) {
    try {
      const titleCounts = {};
      let duplicateGroups = 0;
      
      for (let i = 0; i < chunkData.length; i++) {
        const title = String(chunkData[i][titleIndex] || '').trim();
        if (title.length > 10) { // 最低10文字以上のタイトル
          titleCounts[title] = (titleCounts[title] || 0) + 1;
        }
      }
      
      for (const title in titleCounts) {
        if (titleCounts[title] > 1) {
          duplicateGroups++;
        }
      }
      
      return duplicateGroups;
    } catch (error) {
      console.warn('チャンク重複検出エラー:', error);
      return 0;
    }
  },
  
  /**
   * 分析フェーズを実行
   */
  executeAnalyzePhase: function(state, startTime) {
    console.log('分析フェーズ開始');
    
    try {
      // 分析処理は省略してスキップ（重複検出が主要機能）
      console.log('分析フェーズをスキップ（重複検出完了済み）');
      
      return {
        completed: true,
        step: {
          name: 'analyze',
          success: true,
          message: '分析フェーズ完了（スキップ）',
          progressDetail: '統計分析をスキップしました'
        }
      };
    } catch (error) {
      throw new Error(`分析フェーズでエラー: ${error.message}`);
    }
  },
  
  /**
   * エクスポートフェーズを実行
   */
  executeExportPhase: function(state, startTime) {
    console.log('エクスポートフェーズ開始');
    
    try {
      const exportResult = generateExportCsv();
      
      return {
        completed: true,
        step: {
          name: 'export',
          success: exportResult.success,
          message: exportResult.message,
          progressDetail: exportResult.success ? 
            'CSVエクスポートが完了しました' : 
            'CSVエクスポートに失敗しました'
        },
        finalMessage: exportResult.success ? 
          "処理が完了しました。重複データをCSVファイルとしてダウンロードしてください。" : 
          "CSVエクスポートに失敗したため、処理を中止しました。"
      };
    } catch (error) {
      throw new Error(`エクスポートフェーズでエラー: ${error.message}`);
    }
  }
};

/**
 * インポートシートを初期化（クライアント側チャンクアップロード用）
 */
function initializeImportSheet() {
  try {
    console.log('*** SERVER DEBUG: initializeImportSheet called ***');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const importSheetName = EbayTool.getSheetName('IMPORT');
    
    // インポートシートを取得または作成
    let importSheet;
    try {
      importSheet = spreadsheet.getSheetByName(importSheetName);
      if (!importSheet) {
        importSheet = spreadsheet.insertSheet(importSheetName);
      } else {
        importSheet.clear(); // 既存データをクリア
      }
    } catch (e) {
      importSheet = spreadsheet.insertSheet(importSheetName);
    }
    
    console.log('*** SERVER DEBUG: Import sheet initialized successfully ***');
    return { success: true, message: 'インポートシートを初期化しました' };
    
  } catch (error) {
    console.error('initializeImportSheet エラー:', error);
    return {
      success: false,
      message: `インポートシート初期化に失敗しました: ${error.message}`
    };
  }
}

/**
 * CSVチャンクをインポートシートに追加
 */
function appendCsvChunkToImportSheet(chunkCsv, chunkIndex, totalChunks) {
  try {
    console.log(`*** SERVER DEBUG: appendCsvChunkToImportSheet called - chunk ${chunkIndex + 1}/${totalChunks} ***`);
    console.log(`*** SERVER DEBUG: Chunk size: ${chunkCsv.length} characters ***`);
    
    if (!chunkCsv || chunkCsv.trim() === '') {
      console.log('*** SERVER DEBUG: Empty chunk, skipping ***');
      return { success: true, message: '空のチャンクをスキップしました' };
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const importSheetName = EbayTool.getSheetName('IMPORT');
    const importSheet = spreadsheet.getSheetByName(importSheetName);
    
    if (!importSheet) {
      return { success: false, message: 'インポートシートが見つかりません' };
    }
    
    // チャンクCSVを行に分割して解析
    const lines = chunkCsv.split('\n').filter(line => line.trim() !== '');
    console.log(`*** SERVER DEBUG: Processing ${lines.length} lines ***`);
    
    if (lines.length === 0) {
      console.log('*** SERVER DEBUG: No valid lines in chunk ***');
      return { success: true, message: '有効な行がありません' };
    }
    
    // 各行をCSV解析
    const rows = [];
    let columnCount = 0;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;
      
      try {
        // 簡易CSV行解析
        const row = line.split(',').map(cell => {
          let cleanCell = cell.trim();
          if (cleanCell.startsWith('"') && cleanCell.endsWith('"')) {
            cleanCell = cleanCell.slice(1, -1).replace(/""/g, '"');
          }
          return cleanCell;
        });
        
        // 最初の行（全体のヘッダー行 or チャンクの最初の行）で列数を確定
        if (columnCount === 0) {
          columnCount = row.length;
          console.log(`*** SERVER DEBUG: Column count set to ${columnCount} ***`);
        }
        
        // 列数を統一
        while (row.length < columnCount) {
          row.push('');
        }
        if (row.length > columnCount) {
          row.splice(columnCount);
        }
        
        rows.push(row);
        
      } catch (parseError) {
        console.warn(`*** SERVER DEBUG: Skipping malformed line in chunk: ${parseError.message} ***`);
        continue;
      }
    }
    
    if (rows.length === 0) {
      console.log('*** SERVER DEBUG: No valid rows after parsing ***');
      return { success: true, message: '解析後に有効な行がありません' };
    }
    
    // インポートシートの現在の最終行を取得
    const currentLastRow = importSheet.getLastRow();
    const startRow = currentLastRow + 1;
    
    // データを追加
    const range = importSheet.getRange(startRow, 1, rows.length, columnCount);
    range.setValues(rows);
    
    console.log(`*** SERVER DEBUG: Added ${rows.length} rows starting at row ${startRow} ***`);
    
    return { 
      success: true, 
      message: `チャンク ${chunkIndex + 1}/${totalChunks} を追加しました (${rows.length}行)`,
      rowsAdded: rows.length,
      totalRowsNow: currentLastRow + rows.length
    };
    
  } catch (error) {
    console.error('appendCsvChunkToImportSheet エラー:', error);
    return {
      success: false,
      message: `チャンク追加に失敗しました: ${error.message}`
    };
  }
}

/**
 * インポートシートから重複検出を開始
 */
function startDuplicateDetectionFromImportSheet() {
  try {
    console.log('*** SERVER DEBUG: startDuplicateDetectionFromImportSheet called ***');
    
    // 直接重複検出を実行
    const detectResult = detectDuplicates();
    console.log('*** SERVER DEBUG: Duplicate detection result:', detectResult);
    
    return {
      success: true,
      completed: true,
      message: '重複検出が完了しました',
      result: detectResult,
      stats: {
        duplicatesFound: detectResult?.duplicateCount || 0
      }
    };
    
  } catch (error) {
    console.error('startDuplicateDetectionFromImportSheet エラー:', error);
    return {
      success: false,
      message: `重複検出に失敗しました: ${error.message}`
    };
  }
}

/**
 * 手作業インポート後の高速重複検出
 */
function executeFastDuplicateDetection() {
  try {
    const startTime = new Date().getTime();
    console.log('*** SERVER DEBUG: executeFastDuplicateDetection called ***');
    console.log('*** SERVER DEBUG: 高速重複検出開始時刻:', new Date(startTime).toLocaleString());
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const importSheetName = EbayTool.getSheetName('IMPORT');
    const importSheet = spreadsheet.getSheetByName(importSheetName);
    
    if (!importSheet) {
      return {
        success: false,
        message: 'インポートシートが見つかりません。まず手作業でCSVデータをインポートしてください。'
      };
    }
    
    const lastRow = importSheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: false,
        message: 'インポートシートにデータが見つかりません。手作業でCSVデータをインポートしてください。'
      };
    }
    
    console.log(`*** SERVER DEBUG: インポートシートに${lastRow}行のデータを発見 ***`);
    
    // 直接重複検出を実行（CSVアップロードをスキップ）
    console.log('*** SERVER DEBUG: 重複検出を直接実行開始 ***');
    const detectResult = detectDuplicates();
    console.log('*** SERVER DEBUG: 重複検出完了:', detectResult);
    
    const endTime = new Date().getTime();
    const processingTime = endTime - startTime;
    console.log(`*** SERVER DEBUG: 処理時間: ${processingTime}ms (${Math.round(processingTime/1000)}秒) ***`);
    
    return {
      success: true,
      completed: true,
      message: `高速重複検出が完了しました (処理時間: ${Math.round(processingTime/1000)}秒)`,
      result: detectResult,
      stats: {
        totalRows: lastRow - 1, // ヘッダー行を除く
        duplicatesFound: detectResult?.duplicateCount || 0,
        processingTimeMs: processingTime,
        processingTimeSeconds: Math.round(processingTime/1000)
      },
      processingTime: processingTime
    };
    
  } catch (error) {
    console.error('executeFastDuplicateDetection エラー:', error);
    return {
      success: false,
      message: `高速重複検出に失敗しました: ${error.message}`
    };
  }
}

/**
 * CSVデータを一時保存（大きなパラメータによるタイムアウト対策）
 */
function storeCsvDataForChunkedProcess(csvData) {
  try {
    console.log('*** SERVER DEBUG: storeCsvDataForChunkedProcess called ***');
    console.log('*** SERVER DEBUG: csvData length:', csvData ? csvData.length : 'null');
    
    if (!csvData) {
      return { success: false, message: 'CSVデータがありません' };
    }
    
    // CSVデータをスプレッドシートに一時保存
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 一時保存用のシートを作成または取得
    let tempSheet;
    try {
      tempSheet = spreadsheet.getSheetByName('_TempCSV');
      if (!tempSheet) {
        tempSheet = spreadsheet.insertSheet('_TempCSV');
      } else {
        tempSheet.clear(); // 既存データをクリア
      }
    } catch (e) {
      tempSheet = spreadsheet.insertSheet('_TempCSV');
    }
    
    // 大きなCSVデータを一度に解析するとタイムアウトするため、行単位でストリーミング処理
    console.log('*** SERVER DEBUG: Starting streaming CSV parse and save ***');
    
    // CSVデータを行単位に分割
    const lines = csvData.split('\n');
    console.log(`*** SERVER DEBUG: Split into ${lines.length} lines ***`);
    
    // バッチサイズを小さくしてタイムアウトを回避
    const BATCH_SIZE = 500;
    let totalRows = 0;
    let batchRows = [];
    let headerParsed = false;
    let columnCount = 0;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue; // 空行をスキップ
      
      try {
        // 簡易CSV行解析（カンマ区切り、引用符対応）
        const row = line.split(',').map(cell => {
          // 引用符で囲まれた値の処理
          let cleanCell = cell.trim();
          if (cleanCell.startsWith('"') && cleanCell.endsWith('"')) {
            cleanCell = cleanCell.slice(1, -1).replace(/""/g, '"');
          }
          return cleanCell;
        });
        
        // ヘッダー行で列数を確定
        if (!headerParsed) {
          columnCount = row.length;
          headerParsed = true;
          console.log(`*** SERVER DEBUG: Header parsed, ${columnCount} columns ***`);
        }
        
        // 列数を統一（不足分は空文字で埋める）
        while (row.length < columnCount) {
          row.push('');
        }
        // 過多分は切り捨て
        if (row.length > columnCount) {
          row.splice(columnCount);
        }
        
        batchRows.push(row);
        
        // バッチサイズに達したら保存
        if (batchRows.length >= BATCH_SIZE) {
          const startRow = totalRows + 1;
          const range = tempSheet.getRange(startRow, 1, batchRows.length, columnCount);
          range.setValues(batchRows);
          totalRows += batchRows.length;
          console.log(`*** SERVER DEBUG: Saved streaming batch, rows ${startRow}-${totalRows} ***`);
          batchRows = []; // バッチをクリア
        }
        
      } catch (parseError) {
        console.warn(`*** SERVER DEBUG: Skipping malformed line ${i}: ${parseError.message} ***`);
        continue;
      }
    }
    
    // 残りのバッチを保存
    if (batchRows.length > 0) {
      const startRow = totalRows + 1;
      const range = tempSheet.getRange(startRow, 1, batchRows.length, columnCount);
      range.setValues(batchRows);
      totalRows += batchRows.length;
      console.log(`*** SERVER DEBUG: Saved final streaming batch, rows ${startRow}-${totalRows} ***`);
    }
    
    console.log(`*** SERVER DEBUG: CSV data stored successfully, total rows: ${totalRows} ***`);
    return { 
      success: true, 
      message: `CSVデータを保存しました (${totalRows}行)`,
      rowCount: totalRows 
    };
    
  } catch (error) {
    console.error('storeCsvDataForChunkedProcess エラー:', error);
    return {
      success: false,
      message: `CSVデータ保存に失敗しました: ${error.message}`
    };
  }
}

/**
 * 保存されたCSVデータから分割処理を開始
 */
function startChunkedAutoProcessFromStorage() {
  try {
    console.log('*** SERVER DEBUG: startChunkedAutoProcessFromStorage called ***');
    
    // 一時保存されたCSVデータを読み込み
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const tempSheet = spreadsheet.getSheetByName('_TempCSV');
    
    if (!tempSheet) {
      return {
        success: false,
        message: '一時保存されたCSVデータが見つかりません'
      };
    }
    
    const lastRow = tempSheet.getLastRow();
    const lastCol = tempSheet.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      return {
        success: false,
        message: '保存されたCSVデータが空です'
      };
    }
    
    console.log(`*** SERVER DEBUG: Loading ${lastRow} rows, ${lastCol} columns from temp sheet ***`);
    
    // 大きなデータのCSV変換はタイムアウトするため、直接インポートシートに移行
    console.log('*** SERVER DEBUG: Directly copying data to import sheet to avoid timeout ***');
    
    // インポートシートを取得または作成
    const importSheetName = EbayTool.getSheetName('IMPORT');
    let importSheet;
    
    try {
      importSheet = spreadsheet.getSheetByName(importSheetName);
      if (!importSheet) {
        importSheet = spreadsheet.insertSheet(importSheetName);
      } else {
        importSheet.clear(); // 既存データをクリア
      }
    } catch (e) {
      importSheet = spreadsheet.insertSheet(importSheetName);
    }
    
    // 一時シートからインポートシートへ直接データをコピー（バッチ処理）
    const BATCH_SIZE = 5000; // コピー用のバッチサイズ
    let copiedRows = 0;
    
    for (let startRow = 1; startRow <= lastRow; startRow += BATCH_SIZE) {
      const endRow = Math.min(startRow + BATCH_SIZE - 1, lastRow);
      const batchSize = endRow - startRow + 1;
      
      console.log(`*** SERVER DEBUG: Copying batch ${startRow}-${endRow} (${batchSize} rows) ***`);
      
      // バッチデータを読み込み
      const batchData = tempSheet.getRange(startRow, 1, batchSize, lastCol).getValues();
      
      // インポートシートに書き込み
      const targetRange = importSheet.getRange(startRow, 1, batchSize, lastCol);
      targetRange.setValues(batchData);
      
      copiedRows += batchSize;
      console.log(`*** SERVER DEBUG: Copied ${copiedRows}/${lastRow} rows ***`);
    }
    
    console.log('*** SERVER DEBUG: Data copied successfully, starting direct duplicate detection ***');
    
    // CSVを使わず、直接シート上で重複検出を実行
    const detectResult = EbayTool.detectDuplicates();
    console.log('*** SERVER DEBUG: Direct duplicate detection result:', detectResult);
    
    // 一時シートを削除してクリーンアップ
    try {
      spreadsheet.deleteSheet(tempSheet);
      console.log('*** SERVER DEBUG: Temporary sheet cleaned up ***');
    } catch (e) {
      console.warn('*** SERVER DEBUG: Failed to cleanup temp sheet:', e.message);
    }
    
    // 完了結果を返す
    return {
      success: true,
      completed: true,
      processId: `direct_process_${new Date().getTime()}`,
      message: '処理が完了しました',
      result: detectResult,
      stats: {
        totalRows: copiedRows,
        duplicatesFound: detectResult?.duplicateCount || 0
      }
    };
    
  } catch (error) {
    console.error('startChunkedAutoProcessFromStorage エラー:', error);
    return {
      success: false,
      message: `保存データからの処理開始に失敗しました: ${error.message}`
    };
  }
}

/**
 * 分割処理による自動処理開始（UIから呼び出される）
 */
function startChunkedAutoProcessFromUI(csvData) {
  try {
    console.log('*** SERVER DEBUG: startChunkedAutoProcessFromUI called ***');
    console.log('*** SERVER DEBUG: csvData length:', csvData ? csvData.length : 'null');
    console.log('分割処理による自動処理開始');
    const result = ChunkedProcessor.startChunkedAutoProcess(csvData);
    console.log('*** SERVER DEBUG: ChunkedProcessor returned:', result);
    return result;
  } catch (error) {
    console.error('startChunkedAutoProcessFromUI エラー:', error);
    return {
      success: false,
      message: `分割処理の開始に失敗しました: ${error.message}`
    };
  }
}

/**
 * 分割処理の続行（UIから呼び出される）
 */
function continueChunkedProcess(processId) {
  try {
    console.log(`分割処理続行: ${processId}`);
    return ChunkedProcessor.executeNextPhase(processId);
  } catch (error) {
    console.error('continueChunkedProcess エラー:', error);
    return {
      success: false,
      message: `分割処理の続行に失敗しました: ${error.message}`,
      processId: processId
    };
  }
}

/**
 * 分割処理の状態確認（UIから呼び出される）
 */
function getChunkedProcessStatus(processId) {
  try {
    const state = ChunkedProcessor.getState(processId);
    if (!state) {
      return {
        success: false,
        message: '処理状態が見つかりません'
      };
    }
    
    return {
      success: true,
      processId: processId,
      phase: state.phase,
      currentPhase: state.currentPhase,
      totalPhases: state.totalPhases,
      progress: Math.round((state.currentPhase / state.totalPhases) * 100),
      completed: state.completed,
      error: state.error
    };
  } catch (error) {
    console.error('getChunkedProcessStatus エラー:', error);
    return {
      success: false,
      message: `状態確認に失敗しました: ${error.message}`
    };
  }
}

