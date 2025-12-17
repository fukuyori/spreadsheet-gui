# Spreadsheet GUI

Common Lisp + LTK で作る Lisp パワード表計算ソフト

![Version](https://img.shields.io/badge/version-0.2-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[English README](README.md)

## 概要

数式をLispのS式で記述するユニークな表計算ソフト。セルには数値だけでなく、リスト、シンボル、文字列など任意のLisp値を格納可能。lambda式によるデータ変換もサポート。

## スクリーンショット

```
┌─────────────────────────────────────────────────────────────────────┐
│ =(mapcar (lambda (x) (* x x)) A1)                                   │
├────┬────────────┬────────────────────────┬──────────┬───────────────┤
│    │     A      │           B            │    C     │      D        │
├────┼────────────┼────────────────────────┼──────────┼───────────────┤
│  1 │ (1 2 3 4 5)│ (1 4 9 16 25)          │          │               │
│  2 │     10     │ Hello, World!                     │               │
│  3 │     20     │                        │          │               │
│  4 │     30     │                        │          │               │
└────┴────────────┴────────────────────────┴──────────┴───────────────┘
       (リスト/緑)  (リスト/緑)            
       
* 長いテキストは右側の空セルにはみ出して表示
* 背景色で値の型を識別
```

## 特徴

- **S式による数式** - Lispの構文をそのまま活用
- **lambda式** - `(lambda (x) ...)` でカスタム変換
- **任意のLisp値** - 数値、リスト、シンボル、文字列、キーワード
- **セル参照とlambda** - `(mapcar (lambda (x) (* x x)) A1)`
- **テキストはみ出し** - 長い内容は隣の空セルに拡張表示
- **型別の背景色** - 値の型を視覚的に識別
- **80以上の純粋関数** - ホワイトリストされた非破壊関数

## 必要環境

- **SBCL** (Steel Bank Common Lisp)
- **Quicklisp**
- **LTK** (Lisp Toolkit)
- **Tcl/Tk** 8.5以上

## インストール

```bash
# 1. SBCLのインストール
# Ubuntu/Debian
sudo apt install sbcl tk

# macOS
brew install sbcl tcl-tk

# 2. Quicklispのインストール
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

### 操作方法

| キー | 動作 |
|------|------|
| 矢印キー | カーソル移動 |
| クリック | セル選択 |
| Enter (グリッド) | 編集開始 |
| Enter (入力欄) | 確定 |

## 数式の例

すべての数式は `=` で始める

### 基本演算

```lisp
=(+ A1 B1)           ; 加算
=(* A1 2)            ; 乗算
=(sqrt 2)            ; → 1.4142
=(expt 2 10)         ; → 1024
=(* pi 2)            ; → 6.2831
```

### 範囲操作

```lisp
=(sum (range A1 A5))    ; A1:A5の合計
=(avg (range A1 A10))   ; 平均
=(+ (range A1 A5))      ; sumと同じ
=(max (range B1 B10))   ; 最大値
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
; 各要素を二乗
=(mapcar (lambda (x) (* x x)) '(1 2 3 4 5))
; → (1 4 9 16 25)

; セル参照と組み合わせ（A1にリストがある場合）
=(mapcar (lambda (x) (* x x)) A1)

; フィルタ：10より大きい値を抽出
=(remove-if-not (lambda (x) (> x 10)) '(5 15 8 20))
; → (15 20)

; 偶数をカウント
=(count-if (lambda (x) (evenp x)) '(1 2 3 4 5 6))
; → 3

; カスタムreduce
=(reduce (lambda (a b) (+ a b)) '(1 2 3 4 5))
; → 15
```

### 高階関数

```lisp
=(mapcar #'1+ '(1 2 3))              ; → (2 3 4)
=(reduce #'* '(1 2 3 4 5))           ; → 120
=(remove-if #'evenp '(1 2 3 4 5))    ; → (1 3 5)
=(remove-if-not #'plusp '(-1 0 1 2)) ; → (1 2)
```

### 条件分岐

```lisp
=(if (> A1 10) :big :small)

=(cond
  ((< A1 0) :negative)
  ((= A1 0) :zero)
  (t :positive))
```

### 文字列操作

```lisp
=(string-upcase "hello")               ; → "HELLO"
=(concatenate 'string "Hello" "World") ; → "HelloWorld"
```

## 使用可能な関数

### 算術
`+` `-` `*` `/` `mod` `rem` `1+` `1-` `abs` `max` `min` `sqrt` `expt` `log` `exp` `sin` `cos` `tan` `floor` `ceiling` `round` `gcd` `lcm`

### リスト
`car` `cdr` `cons` `list` `first` `rest` `last` `append` `reverse` `length` `nth` `member` `assoc` `subseq` `butlast`

### 高階関数
`mapcar` `reduce` `remove-if` `remove-if-not` `count-if` `lambda`

### 述語
`atom` `listp` `numberp` `stringp` `symbolp` `null` `zerop` `plusp` `minusp` `evenp` `oddp`

### 比較
`=` `/=` `<` `>` `<=` `>=` `equal`

### 文字列
`string-upcase` `string-downcase` `concatenate` `subseq`

### 特殊
`if` `cond` `and` `or` `quote` `range` `sum` `avg` `count`

## セルの値の型と背景色

| 色 | 型 | 例 |
|----|-----|-----|
| 白 | 数値 | `42` |
| 薄緑 | リスト | `(1 2 3)` |
| 薄赤 | シンボル | `HELLO` |
| 薄黄 | 文字列 | `"text"` |
| 水色 | 選択中 | 現在位置 |

## アーキテクチャ

```
┌──────────────────────────────────────────┐
│              GUI (LTK/Tk)                │
│  Entry (数式入力) + Canvas (グリッド)    │
├──────────────────────────────────────────┤
│             数式評価エンジン              │
│  - S式パーサー                           │
│  - lambda式とクロージャ                  │
│  - セル参照解決                          │
│  - 80以上の純粋関数                      │
├──────────────────────────────────────────┤
│              データモデル                 │
│  ハッシュテーブル: "A1" → (値 . 数式)    │
│  値: 任意のLispオブジェクト              │
└──────────────────────────────────────────┘
```

## ファイル構成

```
├── README.md              ; 英語
├── README-JP.md           ; 日本語
├── LICENSE                ; MIT
├── .gitignore
├── spreadsheet-gui.lisp   ; 最新版
└── spreadsheet-gui-v0.2.lisp
```

## バージョン履歴

### v0.2（現在）
- lambda式サポート
- lambdaでのセル参照
- テキストはみ出し表示
- 80以上のLisp関数
- 型別の背景色

### v0.1
- 基本的な表計算機能
- 四則演算
- 範囲指定
- sum, avg, max, min

## ライセンス

MIT License

## 作者

福寄
