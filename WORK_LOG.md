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
| 1-1 | バックエンド ファイル分割 | 完了 | 2026-04-08 |
| 1-2 | 重複パターンの共通化 | 完了（1-1と同時） | 2026-04-08 |
| 1-3 | approveStudent バグ修正 | 完了（1-1と同時） | 2026-04-08 |
| 1-4 | 学校固有情報の外部化 | 完了 | 2026-04-08 |
| 2-1 | 科目IDの導入 | 完了 | 2026-04-08 |
| 2-2 | 登録データの縦持ち化 | 完了 | 2026-04-08 |
| 2-3 | 変換レイヤー実装 | 完了 | 2026-04-08 |
| 2-4 | データマイグレーション | 完了（スクリプト作成済） | 2026-04-08 |
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

### 2026-04-08: Phase 1-1/1-2/1-3 完了

**実施内容:**
- `コード.js`（5,046行）を11ファイル（計4,745行）に分割
- 共通ユーティリティ作成: `SheetUtils`, `Logger_`, `Validator` (01_Utils.gs)
- キャッシュ基盤を統一: 4つの`*Cached()`関数 → `SheetAccess.getCachedData()` 1関数 (02_SheetAccess.gs)
- 認証モジュール分離: `Auth` 名前空間 (03_Auth.gs)
- サービス層分割: CourseService, RegistrationService, StudentService, StaffService, SettingsService, AdminService
- `CONFIG`定数による一元管理: シート名、ステータス、ロール、基本カラム名
- `approveStudent()` バグ修正: 未定義だった関数を05_RegistrationService.gsに新規実装
- 全公開APIに後方互換ラッパーを追加（既存フロントエンドとの互換性維持）

**ファイル構成:**
```
src/
├── 00_Config.gs              (104行) 定数・設定
├── 01_Utils.gs               (293行) SheetUtils, Logger_, Validator
├── 02_SheetAccess.gs         (233行) キャッシュ基盤
├── 03_Auth.gs                (292行) 認証・認可
├── 04_CourseService.gs       (431行) 科目マスタ
├── 05_RegistrationService.gs (1007行) 履修登録
├── 06_StudentService.gs      (666行) 生徒データ
├── 07_StaffService.gs        (349行) 教職員管理
├── 08_SettingsService.gs     (125行) 設定管理
├── 09_AdminService.gs        (890行) 管理者機能
└── 11_WebApp.gs              (355行) WebApp・API
```

---

## 再開ガイド

### 2026-04-08: Phase 1-4 / Phase 2 完了

**Phase 1-4 実施内容:**
- `src/11_WebApp.gs` の学校名ハードコードを設定シート参照に変更
- `index.html` のフォールバック学校名4箇所を汎用化

**Phase 2 実施内容:**
- `src/00_Config.gs` にREG_HEADER/REG_DETAILシート名を追加
- `src/90_Migration.gs` 作成:
  - `migrateAddCourseIds()` — 科目IDの自動採番（C001-C999）
  - `migrateCreateVerticalSheets()` — 登録ヘッダー/登録明細シート作成
  - `migrateHorizontalToVertical()` — 横持ち→縦持ちデータ変換
  - `getCourseIdToNameMap()` / `getCourseNameToIdMap()` — ID変換ヘルパー
- `src/10_DataTransformer.gs` 作成:
  - `toHorizontal()` — 縦持ち→横持ち（マトリクス表示用）
  - `toVertical()` — 横持ち→縦持ち（保存用）
  - `buildMatrix()` — 全生徒マトリクス生成（教員画面用）
  - `saveHeader()` / `saveDetail()` — 縦持ちシートへの書き込み

---

### 次に何をすべきか
1. Phase 3-1: doGetの変更（createTemplateFromFile化）
2. Phase 3-2: フロントエンドのファイル分割（index.html → 15ファイル）
3. Phase 3-3以降: モーダルコンポーネント化、状態管理改善、セキュリティ改善

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
