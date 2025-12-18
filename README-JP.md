# Spreadsheet GUI

Common Lisp + LTK で作るシンプルな表計算ソフト

![Version](https://img.shields.io/badge/version-0.3.1-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[English README](README.md)

## 概要

数式をLispのS式で記述するユニークなスプレッドシート。セルには数値、リスト、シンボル、文字列など任意のLisp値を格納できます。lambda式、apply、funcallに加え、**自動再計算**機能により強力なデータ変換が可能です。

## v0.3.1 の新機能

- **自動再計算** - 参照元セルが変更されると依存セルが自動更新
- **依存関係追跡** - トポロジカルソートで正しい順序で再計算
- **範囲選択** - マウスドラッグまたはShift+矢印キーで複数セルを選択
- **コピー＆ペースト** - Ctrl+C/Vでシステムクリップボード対応（TSV形式）
- **セル消去** - Delete/BackSpaceでセル内容をクリア
- **位置参照関数** - `this-row`, `this-col`, `rel`, `rel-range`
- **相対参照** - 現在セルからの相対位置でセルを参照
- **デバッグ関数** - `show-dependencies`, `show-cell-deps`

## スクリーンショット

```
┌─────────────────────────────────────────────────────────────────────┐
│ =(+ A1 5)                                                           │
│                                                        [Enter: 確定]│
│                                                  [Shift+Enter: 改行]│
├────┬────────────┬────────────────────────┬──────────┬───────────────┤
│    │     A      │           B            │    C     │   ...         │
├────┼────────────┼────────────────────────┼──────────┼───────────────┤
│  1 │     10     │ (選択中)               │          │               │
│  2 │     15     │ ← 自動更新！           │          │               │
│  3 │     30     │ =(* A2 2)              │          │               │
│ .. │            │                        │          │               │
└────┴────────────┴────────────────────────┴──────────┴───────────────┘
       
A2 = =(+ A1 5) → A1を変更するとA2が自動再計算！
```

## 特徴

- **自動再計算** - 依存セルが自動更新
- **グリッドサイズ指定** - 起動時に行数、列数、入力欄を指定可能
- **S式による数式** - Lispの構文をそのまま活用
- **lambda式** - `(lambda (x) ...)` でカスタム変換
- **apply / funcall** - 関数適用をサポート
- **位置参照関数** - `this-row`, `this-col`, `rel`, `rel-range`
- **相対参照** - 現在セルからの相対位置でセルを参照
- **範囲選択** - ドラッグまたはShift+矢印で複数セルを選択
- **システムクリップボード** - Excel、Googleスプレッドシートとコピー＆ペースト（TSV形式）
- **複数行入力** - 数式入力欄が複数行に対応
- **任意のLisp値** - 数値、リスト、シンボル、文字列、キーワード
- **テキストはみ出し** - 長い内容は隣の空セルに拡張表示
- **80以上の純粋関数** - ホワイトリストされた非破壊関数

## 必要環境

- **SBCL** (Steel Bank Common Lisp)
- **Quicklisp**
- **LTK** (Lisp Toolkit)
- **Tcl/Tk** 8.5+

## インストール

```bash
# 1. SBCLをインストール
# Ubuntu/Debian
sudo apt install sbcl tk

# macOS
brew install sbcl tcl-tk

# 2. Quicklispをインストール
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
;; デフォルト: 26行、14列、3行入力欄
(ss-gui:start)

;; カスタムサイズ
(ss-gui:start :rows 30 :cols 10)

;; 入力欄を大きく
(ss-gui:start :rows 20 :cols 8 :input-lines 5)
```

| パラメータ | デフォルト | 説明 |
|-----------|---------|------|
| `:rows` | 26 | 行数 |
| `:cols` | 14 | 列数（最大26 = A-Z） |
| `:input-lines` | 3 | 入力欄の高さ |

### 操作方法

| キー | 動作 |
|------|------|
| 矢印キー | カーソル移動 |
| Shift+矢印 | 選択範囲を拡張 |
| クリック | セル選択 |
| ドラッグ | 範囲選択 |
| Ctrl+C | システムクリップボードにコピー（TSV） |
| Ctrl+V | システムクリップボードからペースト |
| Delete/BackSpace | セル消去（NILを設定） |
| Enter (グリッド) | 入力欄へ移動 |
| Enter (入力欄) | 入力確定＆自動再計算 |
| Shift+Enter | 入力欄内で改行 |

## 数式の例

数式は `=` で始めます

### 基本的な算術

```lisp
=(+ A1 B1)           ; 加算
=(* A1 2)            ; 乗算
=(sqrt 2)            ; → 1.4142
=(expt 2 10)         ; → 1024
=(* pi 2)            ; → 6.2831
```

### 範囲操作

```lisp
=(sum (range A1 A10))   ; A1:A10の合計
=(avg (range A1 A26))   ; 平均
=(+ (range A1 A5))      ; sumと同じ
=(max (range B1 B26))   ; 最大値
```

### 相対参照（v0.3.1 新機能）

```lisp
=(rel -1 0)          ; 1行上
=(rel 0 -1)          ; 1列左
=(rel 1 1)           ; 1行下、1列右
=(sum (rel-range -4 0 -1 0))  ; 上4行の合計
```

### 位置関数（v0.3.1 新機能）

```lisp
=(this-row)          ; 現在の行番号（1始まり）
=(this-col)          ; 現在の列番号（0始まり）
=(this-col-name)     ; 現在の列名（"A", "B"等）
=(this-cell-name)    ; 現在のセル名（"A1", "B2"等）
=(cell-at 1 0)       ; 行1、列0の値（=A1）
```

### リスト操作

```lisp
=(list 1 2 3 4 5)       ; → (1 2 3 4 5)
=(reverse '(a b c))     ; → (C B A)
=(append '(1 2) '(3 4)) ; → (1 2 3 4)
=(length '(1 2 3))      ; → 3
=(nth 2 '(a b c d))     ; → C
```

### lambda式

```lisp
; 各要素を2乗
=(mapcar (lambda (x) (* x x)) '(1 2 3 4 5))
; → (1 4 9 16 25)

; セル参照と組み合わせ（A1にリストが入っている場合）
=(mapcar (lambda (x) (* x x)) A1)

; 直接呼び出し
=((lambda (x) (+ x 1)) 10)
; → 11

; フィルタ: 10より大きい値を抽出
=(remove-if-not (lambda (x) (> x 10)) '(5 15 8 20))
; → (15 20)

; 偶数をカウント
=(count-if (lambda (x) (evenp x)) '(1 2 3 4 5 6))
; → 3
```

### apply と funcall

```lisp
; 関数をリストに適用
=(apply #'+ '(1 2 3 4 5))
; → 15

=(apply #'max '(3 1 4 1 5 9))
; → 9

; funcall
=(funcall #'+ 1 2 3)
; → 6

=(funcall (lambda (x) (* x x)) 5)
; → 25
```

### 条件式

```lisp
=(if (> A1 10) :big :small)

=(cond
  ((< A1 0) :negative)
  ((= A1 0) :zero)
  (t :positive))
```

## 自動再計算（v0.3.1 新機能）

セルの値が変更されると、そのセルに依存する全てのセルが正しい順序で自動的に再計算されます。

```
A1: 10
A2: =(+ A1 5)     → 15
A3: =(* A2 2)     → 30

A1を20に変更すると:
  A2は自動的に25になる
  A3は自動的に50になる
```

### 仕組み

1. **依存関係の抽出** - 数式入力時に全てのセル参照を抽出
2. **依存グラフ** - どのセルがどのセルを参照しているかを双方向グラフで追跡
3. **トポロジカルソート** - セル変更時、依存元を依存順にソート
4. **連鎖再計算** - 正しい値を保証するため、順番に再計算

自動再計算対応の参照形式:
- 絶対参照: `A1`, `B2`
- 相対参照: `(rel -1 0)`, `(rel 0 1)`
- 範囲: `(range A1 A5)`, `(rel-range -4 0 -1 0)`
- 位置指定: `(cell-at row col)`

## デバッグ（v0.3.1 新機能）

```lisp
;; 全ての依存関係を表示
(ss-gui:show-dependencies)

;; 特定セルの依存関係を表示
(ss-gui:show-cell-deps "A1")
```

コンソール出力:
- 参照抽出: `Extract refs from (+ A1 5): ("A1")`
- 依存更新: `Update deps: A2 refs ("A1")`
- 再計算: `Recalc A2: 15 → 25`

## 使用可能な関数

### 算術
`+` `-` `*` `/` `mod` `rem` `1+` `1-` `abs` `max` `min` `sqrt` `expt` `log` `exp` `sin` `cos` `tan` `floor` `ceiling` `round` `gcd` `lcm`

### リスト
`car` `cdr` `cons` `list` `first` `rest` `last` `append` `reverse` `length` `nth` `member` `assoc` `subseq` `butlast`

### 高階関数
`mapcar` `reduce` `remove-if` `remove-if-not` `count-if` `lambda` `apply` `funcall`

### 述語
`atom` `listp` `numberp` `stringp` `symbolp` `null` `zerop` `plusp` `minusp` `evenp` `oddp`

### 比較
`=` `/=` `<` `>` `<=` `>=` `equal`

### 文字列
`string-upcase` `string-downcase` `concatenate` `subseq`

### 特殊
`if` `cond` `and` `or` `quote` `range` `sum` `avg` `count`

### 位置参照（v0.3.1 新機能）
`this-row` `this-col` `this-col-name` `this-cell-name` `cell-at` `rel` `rel-range`

## セル値の型と色

| 色 | 型 | 例 |
|---|---|---|
| 白 | 数値 | `42` |
| 薄緑 | リスト | `(1 2 3)` |
| 薄赤 | シンボル | `HELLO` |
| 薄黄 | 文字列 | `"text"` |
| 明るい青 | 選択中/カーソル | 現在位置 |
| 淡い青 | 選択範囲 | 複数選択 |

## 仕様

| 項目 | デフォルト | 範囲 |
|------|---------|------|
| 行数 | 26 | 1-999 |
| 列数 | 14 | 1-26 (A-Z) |
| 入力欄行数 | 3 | 1-10 |
| セルサイズ | 100 × 24 px | - |
| 数式接頭辞 | `=` | - |

## ファイル構成

```
├── README.md
├── README-JP.md
├── LICENSE
├── .gitignore
├── spreadsheet-gui.lisp      ; 最新版 (v0.3.1)
├── spreadsheet-gui-v0.1.lisp
├── spreadsheet-gui-v0.2.lisp
├── spreadsheet-gui-v0.3.lisp
└── spreadsheet-gui-v0.3.1.lisp
```

## バージョン履歴

### v0.3.1（現在）
- **自動再計算** - 依存セルが自動更新
- **依存関係グラフ追跡** - トポロジカルソートで正しい順序で再計算
- **範囲選択** - マウスドラッグで複数セルを選択
- **Shift+矢印キーで範囲選択**
- **コピー＆ペースト** - Ctrl+C/V
- **システムクリップボード対応** - TSV形式でExcel/Sheetsと連携
- **セル消去** - Delete/BackSpaceでNILを設定
- **位置参照関数** - this-row, this-col, rel, rel-range
- **相対セル参照**
- **デバッグ関数** - show-dependencies, show-cell-deps

### v0.3
- 起動パラメータ: rows, cols, input-lines
- 複数行入力欄
- Enterで確定、Shift+Enterで改行
- apply / funcall サポート

### v0.2
- lambda式サポート
- lambda内でのセル参照
- テキストはみ出し表示
- 80以上のLisp関数

### v0.1
- 基本的な表計算機能
- 四則演算
- 範囲選択

## ライセンス

MIT License

## 作者

Fukuyori
