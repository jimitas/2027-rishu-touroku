# 履修登録システム リファクタリング完了報告書

**作成日:** 2026-04-08  
**対象システム:** 京都芸術大学附属高等学校 通信制課程 履修登録システム  
**目的:** 自校（通信制高校、200-600人規模）への転用を目指したリファクタリング

---

## 1. 実施概要

### 1.1 リファクタリング前の課題

| 課題 | 詳細 |
|------|------|
| 巨大な単一ファイル | バックエンド 5,046行/フロントエンド 19,224行の2ファイル |
| コード重複 | キャッシュ4重複、ヘッダー検索20+重複、行検索10+重複 |
| 未定義関数バグ | `approveStudent()` が呼び出されるが未定義 |
| セキュリティ | innerHTML 60+箇所でXSSリスク |
| データモデル | 横持ちテーブル（78列）、科目名が結合キー |
| ハードコード | 学校名・ドメイン等が埋め込み |

### 1.2 実施フェーズと成果

| フェーズ | 内容 | 主な成果 |
|----------|------|----------|
| Phase 1-1 | バックエンド分割 | 5,046行 → 11ファイル（計4,745行） |
| Phase 1-2 | 重複の共通化 | 4重複→1関数、20+重複→ユーティリティ |
| Phase 1-3 | バグ修正 | approveStudent() 新規実装 |
| Phase 1-4 | 学校固有情報外部化 | CONFIG.DEFAULTS + 設定シート参照 |
| Phase 2-1 | 科目ID導入 | C001-C999形式、非破壊的追加 |
| Phase 2-2/2-3 | 縦持ち化 | 登録ヘッダー/明細シート + DataTransformer |
| Phase 2-4 | マイグレーション | 横持ち→縦持ち自動変換スクリプト |
| Phase 3-1 | テンプレート化 | createTemplateFromFile + include |
| Phase 3-2 | フロントエンド分割 | 19,224行 → 4ファイル |
| Phase 3-3 | モーダルコンポーネント化 | Modal ユーティリティ + CSSサイズバリアント |
| Phase 3-4 | 状態管理改善 | AppState 5カテゴリ整理 |
| Phase 3-5 | セキュリティ改善 | escapeHtml 103箇所適用 |

---

## 2. アーキテクチャ（リファクタリング後）

### 2.1 バックエンドファイル構成

```
src/
├── 00_Config.gs              (104行) 定数・設定値の一元管理
├── 01_Utils.gs               (293行) 共通ユーティリティ
├── 02_SheetAccess.gs         (233行) キャッシュ付きデータアクセス基盤
├── 03_Auth.gs                (292行) 認証・認可モジュール
├── 04_CourseService.gs       (431行) 科目マスタCRUD
├── 05_RegistrationService.gs (1007行) 履修登録フロー
├── 06_StudentService.gs      (666行) 生徒データ管理
├── 07_StaffService.gs        (349行) 教職員管理
├── 08_SettingsService.gs     (125行) 設定管理
├── 09_AdminService.gs        (890行) 管理者機能
├── 10_DataTransformer.gs     (---行) 縦持ち⇔横持ち変換レイヤー
├── 11_WebApp.gs              (355行) doGet・API・include
├── 90_Migration.gs           (---行) データマイグレーション
└── appsscript.json
```

### 2.2 モジュール依存関係

```
00_Config（定数）
    ↓
01_Utils（ユーティリティ）
    ↓
02_SheetAccess（データアクセス）
    ↓
03_Auth（認証・認可）
    ↓
┌───────────────┬───────────────┬───────────────┐
04_CourseService  05_Registration  06_StudentService
                  Service
└───────────────┴───────────────┴───────────────┘
    ↓
┌───────────────┬───────────────┬───────────────┐
07_StaffService  08_SettingsService 09_AdminService
└───────────────┴───────────────┴───────────────┘
    ↓
10_DataTransformer（変換レイヤー）
    ↓
11_WebApp（公開API）
```

### 2.3 フロントエンドファイル構成

```
├── index.html        (22行)     シェル（include読込）
├── css_main.html     (1,745行)  全CSS + モーダルサイズバリアント
├── html_main.html    (1,342行)  HTML本体（26モーダル含む）
└── js_app.html       (16,115行) 全JavaScript
```

### 2.4 名前空間パターン

GASではES modulesが使えないため、オブジェクトリテラルによる名前空間パターンを採用：

```javascript
// 各サービスは独立した名前空間オブジェクト
const CourseService = {
  getCourseData() { ... },
  addCourse(data) { ... },
  // ...
};

// 後方互換ラッパーでフロントエンドとの互換性を維持
function getCourseData() { return CourseService.getCourseData(); }
```

---

## 3. 主要な設計判断

### 3.1 後方互換性の維持

既存フロントエンドが `google.script.run.getCourseData()` のようにグローバル関数を呼び出すため、各サービスの末尾に後方互換ラッパーを配置。フロントエンド側の変更なしでバックエンドのリファクタリングが可能。

### 3.2 データモデル再設計（縦持ち化）

**問題:** 登録データが横持ち（78列）で、科目追加のたびにカラム追加が必要。

**解決:**
- **登録ヘッダーシート:** 生徒基本情報（学籍番号, 学年, 組, 番号, 名前, ステータス等）
- **登録明細シート:** 科目選択データ（学籍番号, 科目ID, 科目名, マーク, 学年, 更新日時）
- **DataTransformer:** フロントエンドのマトリクス表示に必要な横持ち形式へのリアルタイム変換

### 3.3 科目ID導入

科目名（日本語）が全シートの結合キーだった問題を、`C001`-`C999` 形式のIDで解決。既存の科目名ベースのコードとの互換性のため、双方向の変換マップを提供。

### 3.4 セキュリティ改善

innerHTML の全202箇所を監査し、ユーザー入力や外部データが流入する箇所に `Utils.escapeHtml()` を適用。

**修正対象カテゴリ:**
- 科目名・セクション名（スプレッドシートから取得）
- 生徒データ（Excel取込時の外部ファイル入力）
- サーバーエラーメッセージ
- PDF用紙の動的コンテンツ
- QR URL属性値

---

## 4. 技術的負債の解消状況

### 4.1 解消済み

| 負債 | 重大度 | 対処 |
|------|--------|------|
| 全コードが2ファイル | Critical | 11+4ファイルに分割 |
| approveStudent未定義 | Critical | 05_RegistrationService.gsに実装 |
| キャッシュ4重複 | High | SheetAccess.getCachedData()に統一 |
| ヘッダー検索20+重複 | High | SheetUtils.buildHeaderMap()に統一 |
| 行検索10+重複 | High | SheetUtils.findRow()に統一 |
| innerHTML XSSリスク | High | escapeHtml 103箇所適用 |
| 横持ちテーブル | Medium | 縦持ち化 + DataTransformer |
| 科目名が結合キー | Medium | 科目ID (C001-C999) 導入 |
| 学校固有ハードコード | Medium | CONFIG.DEFAULTS + 設定シート |

### 4.2 部分的に解消

| 負債 | 重大度 | 現状 |
|------|--------|------|
| モーダル28個未コンポーネント化 | Medium | Modal ユーティリティ追加、既存モーダルのHTML構造は維持 |
| インラインスタイル780+箇所 | Medium | CSSクラス追加済み、段階的移行が必要 |
| グローバル変数汚染 | Medium | AppState整理済み、完全なカプセル化は今後 |

### 4.3 未解消（低優先度）

| 負債 | 重大度 | 備考 |
|------|--------|------|
| イベントハンドラ分散（onclick混在） | Low | 動作に影響なし |
| 認証コードのブルートフォース対策 | Low | GAS WebAppの制約上、レート制限は困難 |
| エラーハンドリング不統一 | Low | 段階的に統一予定 |

---

## 5. 未実装機能（コア機能外）

以下の機能はスタブ関数（`'未実装（後回し）'` を返す）として `09_AdminService.gs` に定義済み：

| 機能 | 関数 | 元システムの役割 | 実装優先度 |
|------|------|------------------|------------|
| 教科書管理 | `getTextbookData()` / `saveTextbookData()` | 科目と教科書商品コードの対応管理 | 低（自校で必要なら実装） |
| レポート管理 | `getReportData()` / `saveReportData()` | レポートシステムの科目コード対応 | 低（自校で必要なら実装） |
| SM設定 | `getSMSettings()` / `saveSMSettings()` | スクーリンググループ分け | 中（スクーリングがある場合） |

**補足:** `StudentService.getStudentInfoForTextbook()` のみ実装済み（教科書発注用の生徒情報取得）。

---

## 6. 自校転用時の作業

### 6.1 必須作業

1. **`.clasp.json` の変更** — GASプロジェクトIDを自校のものに変更
2. **設定シートの更新** — 学校名、ドメイン、単位数設定等
3. **教職員データの投入** — メールアドレスと権限設定
4. **科目データの投入** — 自校のカリキュラムに合わせた科目マスタ
5. **`clasp push`** — コードをGASプロジェクトにデプロイ
6. **WebApp公開設定** — 「ウェブアプリとしてデプロイ」で公開

### 6.2 推奨作業

1. **データマイグレーション実行** — `migrateAddCourseIds()` → `migrateCreateVerticalSheets()`
2. **`00_Config.gs` のカスタマイズ** — DEFAULTS内の学校名・ドメインを変更
3. **卒業要件の設定** — 設定シートで最低単位数・教科別要件を調整

### 6.3 オプション

- SM設定・教科書・レポート機能の実装（スタブからの開発）
- インラインスタイルのCSS変数移行
- フロントエンドのさらなるファイル分割（js_app.html → 機能別分割）

---

## 7. ファイル一覧

| パス | 説明 |
|------|------|
| `src/00_Config.gs` - `src/11_WebApp.gs` | バックエンドモジュール群 |
| `src/10_DataTransformer.gs` | 縦持ち⇔横持ち変換 |
| `src/90_Migration.gs` | データマイグレーションスクリプト |
| `index.html` | フロントエンドシェル |
| `css_main.html` | CSS（モーダルバリアント含む） |
| `html_main.html` | HTML本体 |
| `js_app.html` | JavaScript全体 |
| `docs/spreadsheet-structure.md` | スプレッドシート構造 |
| `docs/reverse-engineering-report.md` | リバースエンジニアリング報告書 |
| `docs/refactoring-report.md` | 本書（リファクタリング完了報告書） |
| `docs/usage-guide.md` | 使い方ガイド |
| `WORK_LOG.md` | 作業ログ |
| `.claude/plans/glowing-stargazing-swan.md` | リファクタリング計画書 |
