# 履修登録システム リファクタリング 作業ログ

**プロジェクト:** 通信制高校 履修登録システム（GAS + Spreadsheet）  
**リポジトリ:** https://github.com/jimitas/2027-rishu-touroku.git  
**元システム:** 京都芸術大学附属高等学校 履修登録システム  
**開始日:** 2026-04-08

---

## 進捗サマリー

| Phase | 内容 | 状態 | 完了日 |
|-------|------|------|--------|
| 準備 | リバースエンジニアリング・分析 | 完了 | 2026-04-08 |
| 1-1 | バックエンド ファイル分割 | 未着手 | - |
| 1-2 | 重複パターンの共通化 | 未着手 | - |
| 1-3 | approveStudent バグ修正 | 未着手 | - |
| 1-4 | 学校固有情報の外部化 | 未着手 | - |
| 2-1 | 科目IDの導入 | 未着手 | - |
| 2-2 | 登録データの縦持ち化 | 未着手 | - |
| 2-3 | 変換レイヤー実装 | 未着手 | - |
| 2-4 | データマイグレーション | 未着手 | - |
| 3-1 | doGet変更（テンプレート化） | 未着手 | - |
| 3-2 | フロントエンド ファイル分割 | 未着手 | - |
| 3-3 | モーダルコンポーネント化 | 未着手 | - |
| 3-4 | 状態管理の改善 | 未着手 | - |
| 3-5 | セキュリティ改善 | 未着手 | - |

---

## 作業履歴

### 2026-04-08: 準備フェーズ完了

**実施内容:**
- claspでGASコードをプル（コード.js, index.html, appsscript.json）
- Google Spreadsheet 7シートの構造を分析
- バックエンド（5,046行・134関数）の全体分析
- フロントエンド（19,224行・14管理オブジェクト・28モーダル）の全体分析
- 技術的負債の特定・分類
- リファクタリング計画の策定・承認

**生成ドキュメント:**
- `docs/spreadsheet-structure.md` — スプレッドシート構造
- `docs/reverse-engineering-report.md` — 詳細分析報告書

**発見した重大な問題:**
- `approveStudent()` 関数が未定義（承認機能のバグ）
- 全コードが2ファイルに集中（計24,270行）
- 科目名（日本語）が全シートの結合キー

---

## 再開ガイド

### 次に何をすべきか
1. Phase 1-1 から開始: `00_Config.gs` と `01_Utils.gs` を作成
2. 続いて `02_SheetAccess.gs`（キャッシュ基盤）を作成
3. `コード.js` の関数を各サービスファイルに分割

### ファイル構成（計画）
```
src/                         # Phase 1で作成
├── 00_Config.gs
├── 01_Utils.gs
├── 02_SheetAccess.gs
├── 03_Auth.gs
├── 04_CourseService.gs
├── 05_RegistrationService.gs
├── 06_StudentService.gs
├── 07_StaffService.gs
├── 08_SettingsService.gs
├── 09_AdminService.gs
├── 10_DataTransformer.gs    # Phase 2で作成
├── 11_WebApp.gs
└── appsscript.json

frontend/                    # Phase 3で作成
├── index.html
├── css_*.html
├── html_*.html
├── js_*.html
```

### 重要な参照先
- 計画書: `.claude/plans/glowing-stargazing-swan.md`
- 分析報告書: `docs/reverse-engineering-report.md`
- シート構造: `docs/spreadsheet-structure.md`
- 元コード: `コード.js`（5,046行）、`index.html`（19,224行）
