/**
 * 統一バージョン管理設定
 * Google Apps Script専用バージョン管理システム
 */

// 単一の真実の源（Single Source of Truth）
const PROJECT_VERSION = '1.6.48';

// バージョン情報を取得する関数
function getProjectVersion() {
  return PROJECT_VERSION;
}

// バージョン情報をオブジェクトで返す関数
function getVersionInfo() {
  return {
    version: PROJECT_VERSION,
    fullName: `eBay出品管理ツール v${PROJECT_VERSION}`,
    displayName: `v${PROJECT_VERSION}`
  };
}