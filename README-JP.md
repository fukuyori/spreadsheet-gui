# Spreadsheet GUI

Common LispとLTKで構築されたLisp式表計算アプリケーション

![Version](https://img.shields.io/badge/version-0.5.1-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[English README](README.md)

## 概要

数式をLispのS式で記述するユニークな表計算ソフトです。セルには数値、リスト、シンボル、文字列など、任意のLisp値を格納できます。ラムダ式、自動再計算、Undo/Redo、ファイル保存/読み込み、CSVインポート/エクスポートに対応しています。

## スクリーンショット

```
┌─────────────────────────────────────────────────────────────────────┐
│ [File] [Edit]                                                       │
├─────────────────────────────────────────────────────────────────────┤
│ =(+ A1 5)                                                           │
├────┬────────────┬────────────────────────┬──────────┬───────────────┤
│    │     A      │           B            │    C     │   ...         │
├────┼────────────┼────────────────────────┼──────────┼───────────────┤
│  1 │         10 │ (1 2 3 4 5)            │          │               │
│  2 │         15 │ ← 自動更新！           │          │               │
│  3 │         30 │ こんにちは             │          │               │
│ .. │            │                        │          │               │
└────┴────────────┴────────────────────────┴──────────┴───────────────┘
         ↑ 数値は右寄せ         ↑ テキストは左寄せ
```

## 機能

### スマート数式参照 (v0.5)
- 行・列の挿入・削除時に数式参照を自動更新
- 絶対参照 (A1, B2)、相対参照 (rel)、相対範囲 (rel-range) に対応
- 削除されたセルへの参照は #REF! エラー
- データ損失の可能性がある場合は確認ダイアログを表示

### ユーザーインターフェース
- ヘッダー上での右クリックメニュー
- 列幅・行高さの調整（ヘッダー境界をドラッグ）
- 列幅・行高さの数値指定ダイアログ
- 行・列の挿入・削除
- 数値は右寄せ、テキストは左寄せ
- ダイレクト入力モード（タイプするだけで編集開始）

### 数式システム
- S式による数式（フルLisp構文）
- カスタム変換のためのラムダ式
- 位置参照: `this-row`, `this-col`, `rel`, `rel-range`
- 80以上のホワイトリスト純粋関数
- 依存関係追跡による自動再計算

### データ管理
- 100操作までのUndo/Redo
- ネイティブ `.ssp` 形式でのファイル保存/読み込み
- Excel/Googleスプレッドシート互換のCSVインポート/エクスポート
- ドラッグまたはShift+矢印による範囲選択
- システムクリップボード対応（TSV形式）

## 必要環境

- **SBCL** (Steel Bank Common Lisp)
- **Quicklisp**
- **LTK** (Lisp Toolkit)
- **Tcl/Tk** 8.5以上

## インストール

```bash
# Ubuntu/Debian
sudo apt install sbcl tk

# macOS
brew install sbcl tcl-tk

# Quicklispのインストール
curl -O https://beta.quicklisp.org/quicklisp.lisp
sbcl --load quicklisp.lisp
```

```lisp
(quicklisp-quickstart:install)
(ql:add-to-init-file)
(quit)
```

## 使い方

```lisp
(load "spreadsheet-gui.lisp")
(ss-gui:start)
```

### 起動パラメータ

```lisp
(ss-gui:start)                           ; デフォルト: 26行, 14列
(ss-gui:start :rows 30 :cols 10)         ; カスタムサイズ
(ss-gui:start :input-lines 5)            ; 入力エリアを大きく
```

## キーボードショートカット

### ナビゲーション・選択

| キー | 動作 |
|------|------|
| 矢印キー | カーソル移動 |
| Shift+矢印 | 選択範囲の拡張 |
| クリック/ドラッグ | セル選択 |
| 任意の文字 | タイプして編集開始（セルをクリア） |
| F2 | 編集モード（内容を保持） |
| Escape | 編集をキャンセル |
| Delete/BackSpace | セルをクリア |

### 入力確定

| キー | 動作 |
|------|------|
| Enter | 確定 → 下へ移動 |
| Ctrl+Enter | 確定 → その場に留まる |
| Shift+Enter | 入力内で改行 |
| Alt+Enter | 確定 → 右へ移動 |
| Shift+Alt+Enter | 確定 → 左へ移動 |

### 編集操作

| キー | 動作 |
|------|------|
| Ctrl+Z | 元に戻す |
| Ctrl+Y | やり直し |
| Ctrl+X | 切り取り |
| Ctrl+C | コピー |
| Ctrl+V | 貼り付け |

### ファイル操作

| キー | 動作 |
|------|------|
| Ctrl+N | 新規シート |
| Ctrl+O | ファイルを開く |
| Ctrl+S | ファイルを保存 |

## メニュー

### Fileメニュー

| 項目 | 説明 |
|------|------|
| New | 新しい空のシートを作成 |
| Open... | .sspファイルを開く |
| Save | 現在のファイルに保存 |
| Save As... | 名前を付けて保存 |
| Import CSV... | CSVからインポート |
| Export CSV... | CSVにエクスポート |
| Exit | アプリケーションを終了 |

### Editメニュー

| 項目 | 説明 |
|------|------|
| Undo | 元に戻す (Ctrl+Z) |
| Redo | やり直し (Ctrl+Y) |
| Cut | 切り取り (Ctrl+X) |
| Copy | コピー (Ctrl+C) |
| Paste | 貼り付け (Ctrl+V) |
| Delete | 選択範囲を削除 |
| Insert Row | カーソル位置に行を挿入 |
| Insert Column | カーソル位置に列を挿入 |
| Delete Row | カーソル位置の行を削除 |
| Delete Column | カーソル位置の列を削除 |

### コンテキストメニュー（ヘッダー上で右クリック）

**列ヘッダー (A, B, C...) 上:**

| 項目 | 説明 |
|------|------|
| Insert Column | クリック位置に列を挿入 |
| Delete Column | クリック位置の列を削除 |
| Column Width... | 列幅を設定（ダイアログ） |

**行ヘッダー (1, 2, 3...) 上:**

| 項目 | 説明 |
|------|------|
| Insert Row | クリック位置に行を挿入 |
| Delete Row | クリック位置の行を削除 |
| Row Height... | 行高さを設定（ダイアログ） |

## 数式の例

### 基本的な算術演算

```lisp
=(+ A1 B1)           ; 加算
=(* A1 2)            ; 乗算
=(sqrt 2)            ; → 1.4142
=(expt 2 10)         ; → 1024
```

### 範囲操作

```lisp
=(sum (range A1 A10))   ; A1:A10の合計
=(avg (range A1 A26))   ; 平均
=(max (range B1 B26))   ; 最大値
```

### 相対参照

```lisp
=(rel -1 0)          ; 1行上
=(rel 0 -1)          ; 1列左
=(sum (rel-range -4 0 -1 0))  ; 上4行の合計
```

行・列の挿入・削除時、相対参照は同じ論理的な参照先を維持するように自動的に調整されます。

### ラムダ式

```lisp
=(mapcar (lambda (x) (* x x)) '(1 2 3 4 5))
; → (1 4 9 16 25)

=(remove-if-not (lambda (x) (> x 10)) '(5 15 8 20))
; → (15 20)
```

### 条件分岐

```lisp
=(if (> A1 10) :big :small)

=(cond
  ((< A1 0) :negative)
  ((= A1 0) :zero)
  (t :positive))
```

## スマート数式参照の自動更新

行・列の挿入・削除時、数式参照は自動的に更新されます：

### 絶対参照

| 操作 | 変更前 | 変更後 |
|------|--------|--------|
| 2行目に挿入 | =A3 | =A4 |
| 2行目を削除 | =A3 | =A2 |
| 2行目を削除 | =A2 | #REF! |
| B列に挿入 | =C1 | =D1 |

### 相対参照

| 操作 | セル | 変更前 | 変更後 |
|------|------|--------|--------|
| 2行目に挿入 | B3→B4 | =(rel -1 0) | =(rel -2 0) |
| B列に挿入 | C1→D1 | =(rel 0 -1) | =(rel 0 -2) |

相対オフセットは、数式が同じ元のセルを参照し続けるように調整されます。

## ファイル操作

### 保存と読み込み（.ssp形式）

```lisp
;; 現在のスプレッドシートを保存
(ss-gui:save "mydata.ssp")

;; ファイルから読み込み
(ss-gui:load-file "mydata.ssp")

;; 新しい空のシートを作成
(ss-gui:new-sheet)

;; 現在のファイルを確認
ss-gui:*current-file*
```

### .sspファイル形式

`.ssp`形式は人間が読めるS式形式です：

```lisp
(:spreadsheet
 :format-version 1
 :metadata (:created "2025-01-15T10:30:00"
            :app-version "0.5")
 :grid (:rows 26 :cols 14)
 :cells (("A1" 10)
         ("A2" 15 (+ a1 5))
         ("A3" 30 (* a2 2))
         ("B1" (1 2 3 4 5))))
```

### CSVエクスポート/インポート

```lisp
;; CSVにエクスポート（データのみ）
(ss-gui:export-csv "data.csv")

;; 列ヘッダー (A, B, C...) 付きでエクスポート
(ss-gui:export-csv "data.csv" :include-header t)

;; 数式付きでエクスポート
(ss-gui:export-csv "data.csv" :include-formulas t)

;; CSVからインポート
(ss-gui:import-csv "data.csv")

;; ヘッダー行をスキップしてインポート
(ss-gui:import-csv "data.csv" :has-header t)
```

## エラーメッセージ

| エラー | 意味 |
|--------|------|
| #構文: ... | 数式の構文エラー |
| #評価: ... | 評価エラー |
| #循環: A1 | 循環参照を検出 |
| #REF! | 削除されたセルへの参照 |

## バージョン履歴

### v0.5.1（現在）
- 入力枠のリサイズ可能（境界線をドラッグ）
- 入力枠に固定幅フォントを使用
- タブ幅を4文字に設定

### v0.5
- 行・列の挿入・削除時の**数式参照の自動更新**
  - 絶対参照 (A1, B2など)
  - 相対参照 (rel row-offset col-offset)
  - 相対範囲 (rel-range start-row start-col end-row end-col)
- 削除されたセルへの参照は **#REF! エラー**
- データ損失の可能性がある場合は**確認ダイアログ**を表示
- ヘッダー上での右クリックメニュー
- 列幅・行高さの数値指定ダイアログ
- 列幅・行高さの調整（ヘッダー境界をドラッグ）
- 行・列の挿入・削除（Editメニュー）

### v0.4
- Fileメニュー（保存/読み込み、.ssp形式）
- Editメニュー（Undo/Redo/Cut/Copy/Paste）
- ダイレクト入力モード（タイプして編集開始）
- F2で編集モード
- 複数のEnterキー動作
- CSVインポート/エクスポート
- キーボードショートカット (Ctrl+N/O/S/Z/Y/X/C/V)

### v0.3
- 依存関係追跡による自動再計算
- 範囲選択（ドラッグ、Shift+矢印）
- システムクリップボードとのコピー＆ペースト
- 位置参照関数 (rel, rel-range)
- ラムダ式
- 80以上のLisp関数

### v0.2
- 基本的な表計算機能
- S式数式

## ライセンス

MIT License

## 作者

Fukuyori
