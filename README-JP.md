# Spreadsheet GUI

Common Lisp + LTK で作るシンプルな表計算ソフト

![Version](https://img.shields.io/badge/version-0.4.1-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[English README](README.md)

## 概要

数式をLispのS式で記述するユニークなスプレッドシート。セルには数値、リスト、シンボル、文字列など任意のLisp値を格納できます。lambda式、自動再計算、Undo/Redo、ファイル保存/読み込み、CSVインポート/エクスポートに対応。

## スクリーンショット

```
┌─────────────────────────────────────────────────────────────────────┐
│ [ファイル(F)] [編集(E)]                                              │
├─────────────────────────────────────────────────────────────────────┤
│ =(+ A1 5)                                                           │
├────┬────────────┬────────────────────────┬──────────┬───────────────┤
│    │     A      │           B            │    C     │   ...         │
├────┼────────────┼────────────────────────┼──────────┼───────────────┤
│  1 │         10 │ (1 2 3 4 5)            │          │               │
│  2 │         15 │ ← 自動更新！           │          │               │
│  3 │         30 │ Hello                  │          │               │
│ .. │            │                        │          │               │
└────┴────────────┴────────────────────────┴──────────┴───────────────┘
         ↑ 数値は右寄せ         ↑ テキストは左寄せ
```

## 特徴

- **直接入力モード** - キー入力で即座に編集開始
- **数値の右寄せ表示** - 数値は右寄せ、テキストは左寄せで表示
- **Undo/Redo** - Ctrl+Z/Y で最大100操作まで履歴保持
- **ファイル保存/読み込み** - `.ssp`形式で数式と値を保存
- **CSVインポート/エクスポート** - Excel、Google Sheetsとの連携
- **自動再計算** - 依存セルが自動更新
- **S式による数式** - Lispの構文をそのまま活用
- **lambda式** - `(lambda (x) ...)` でカスタム変換
- **位置参照関数** - `this-row`, `this-col`, `rel`, `rel-range`
- **範囲選択** - ドラッグまたはShift+矢印で複数セルを選択
- **システムクリップボード** - 他アプリとコピー＆ペースト（TSV形式）
- **任意のLisp値** - 数値、リスト、シンボル、文字列、キーワード
- **80以上の純粋関数** - ホワイトリストされた非破壊関数
- **改善されたエラー処理** - 分かりやすいエラーメッセージ（#構文, #評価, #循環）

## 必要環境

- **SBCL** (Steel Bank Common Lisp)
- **Quicklisp**
- **LTK** (Lisp Toolkit)
- **Tcl/Tk** 8.5+

## インストール

```bash
# Ubuntu/Debian
sudo apt install sbcl tk

# macOS
brew install sbcl tcl-tk

# Quicklispをインストール
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
(ss-gui:start)                           ; デフォルト: 26行、14列
(ss-gui:start :rows 30 :cols 10)         ; カスタムサイズ
(ss-gui:start :input-lines 5)            ; 入力欄を大きく
```

## キーボードショートカット

### 基本操作

| キー | 動作 |
|------|------|
| 矢印キー | カーソル移動 |
| Shift+矢印 | 選択範囲を拡張 |
| クリック/ドラッグ | セル選択 |
| 文字キー | 直接入力開始（セルをクリア） |
| F2 | 編集モード（内容を保持） |
| Escape | 編集キャンセル |
| Delete/BackSpace | セル消去 |

### 入力確定

| キー | 動作 |
|------|------|
| Enter | 確定 → 下に移動 |
| Ctrl+Enter | 確定 → そのまま |
| Shift+Enter | 確定 → 右に移動 |
| Ctrl+Shift+Enter | 確定 → 左に移動 |
| Alt+Enter | 数式内で改行 |

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
| Ctrl+N | 新規作成 |
| Ctrl+O | ファイルを開く |
| Ctrl+S | 保存 |

## メニュー

### ファイルメニュー

| 項目 | 説明 |
|------|------|
| 新規作成(N) | 新しい空のシートを作成 |
| 開く(O)... | .sspファイルを開く |
| 保存(S) | 現在のファイルに保存 |
| 名前を付けて保存(A)... | 新しいファイルとして保存 |
| CSVインポート... | CSVから読み込み |
| CSVエクスポート... | CSVへ書き出し |
| 終了(X) | アプリケーションを終了 |

### 編集メニュー

| 項目 | 説明 |
|------|------|
| 元に戻す(U) | 直前の操作を取り消す (Ctrl+Z) |
| やり直し(R) | 取り消した操作をやり直す (Ctrl+Y) |
| 切り取り(X) | 選択範囲を切り取り (Ctrl+X) |
| コピー(C) | 選択範囲をコピー (Ctrl+C) |
| 貼り付け(V) | クリップボードから貼り付け (Ctrl+V) |
| 削除 | 選択範囲を削除 |

## ファイル操作

### 保存と読み込み（.ssp形式）

```lisp
;; スプレッドシートを保存
(ss-gui:save "mydata.ssp")

;; ファイルから読み込み
(ss-gui:load-file "mydata.ssp")

;; 新規シートを作成
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
            :app-version "0.4")
 :grid (:rows 26 :cols 14)
 :cells (("A1" 10)
         ("A2" 15 (+ A1 5))
         ("A3" 30 (* A2 2))
         ("B1" (1 2 3 4 5))))
```

### CSVエクスポート/インポート

```lisp
;; CSVにエクスポート（データのみ）
(ss-gui:export-csv "data.csv")

;; 列名ヘッダー付きでエクスポート（A, B, C...）
(ss-gui:export-csv "data.csv" :include-header t)

;; 数式も含めてエクスポート
(ss-gui:export-csv "data.csv" :include-formulas t)

;; CSVからインポート
(ss-gui:import-csv "data.csv")

;; ヘッダー行をスキップしてインポート
(ss-gui:import-csv "data.csv" :has-header t)
```

## 数式の例

### 基本的な算術

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

### lambda式

```lisp
=(mapcar (lambda (x) (* x x)) '(1 2 3 4 5))
; → (1 4 9 16 25)

=(remove-if-not (lambda (x) (> x 10)) '(5 15 8 20))
; → (15 20)
```

### 条件式

```lisp
=(if (> A1 10) :big :small)

=(cond
  ((< A1 0) :negative)
  ((= A1 0) :zero)
  (t :positive))
```

## エラーメッセージ

| エラー | 意味 |
|--------|------|
| #構文: ... | 数式の構文エラー |
| #評価: ... | 評価時エラー |
| #循環: A1 | 循環参照を検出 |

## Undo/Redo

- 最大100操作まで履歴保持
- セルの値と数式の変更を追跡
- 複数セルのペーストはグループ化

```lisp
;; プログラムからのアクセス
(ss-gui:undo)         ; 元に戻す
(ss-gui:redo)         ; やり直し
(ss-gui:clear-history) ; 履歴をクリア
```

## 自動再計算

セルの値が変更されると、依存するセルが自動的に再計算されます：

```
A1: 10
A2: =(+ A1 5)     → 15
A3: =(* A2 2)     → 30

A1を20に変更すると:
  A2 → 25
  A3 → 50
```

## バージョン履歴

### v0.4.1（現在）
- 数値の右寄せ表示
- ファイルメニュー（保存/読み込み、.ssp形式）
- 編集メニュー（Undo/Redo/切り取り/コピー/貼り付け）
- 直接入力モード（キー入力で編集開始）
- F2で編集モード
- 複数のEnterキー動作（下/そのまま/右/左）
- CSVインポート/エクスポート
- 改善されたエラー処理
- キーボードショートカット（Ctrl+N/O/S/Z/Y/X/C/V）

### v0.3.1
- 依存関係追跡による自動再計算
- 範囲選択（ドラッグ、Shift+矢印）
- システムクリップボードでコピー＆ペースト
- 位置参照関数（rel, rel-range）

### v0.3
- 起動パラメータ（rows, cols, input-lines）
- 複数行入力欄
- apply / funcall サポート

### v0.2
- lambda式
- テキストはみ出し表示
- 80以上のLisp関数

### v0.1
- 基本的な表計算機能

## ライセンス

MIT License

## 作者

Fukuyori
