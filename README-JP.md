# Spreadsheet GUI

Common Lisp + LTK で作るシンプルな表計算ソフト

![Version](https://img.shields.io/badge/version-0.1-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[English README](README.md)

## スクリーンショット

```
┌─────────────────────────────────────────────────────┐
│ =(sum (range A1 A4))                                │
├────┬────────┬────────┬────────┬────────┬────────────┤
│    │   A    │   B    │   C    │   D    │   ...      │
├────┼────────┼────────┼────────┼────────┼────────────┤
│  1 │  10    │  20    │        │        │            │
│  2 │  20    │        │        │        │            │
│  3 │  11    │        │        │        │            │
│  4 │  12    │        │        │        │            │
│  5 │ [53]   │        │        │        │            │
│... │        │        │        │        │            │
└────┴────────┴────────┴────────┴────────┴────────────┘
```

## 特徴

- **S式による数式入力** - Lispの構文をそのまま活用
- **セル参照** - `A1`, `B2` などシンボルで直接参照
- **範囲指定** - `(range A1 A5)` で範囲選択
- **集計関数** - `sum`, `avg`, `max`, `min`, `count`
- **GUI操作** - クリック・矢印キーでセル移動

## 必要環境

- **SBCL** (Steel Bank Common Lisp)
- **Quicklisp**
- **LTK** (Lisp Toolkit - Tk バインディング)
- **Tcl/Tk** 8.5以上

## インストール

### 1. SBCLのインストール

```bash
# Ubuntu/Debian
sudo apt install sbcl

# macOS (Homebrew)
brew install sbcl

# Windows
# https://www.sbcl.org/platform-table.html からダウンロード
```

### 2. Quicklispのインストール

```bash
curl -O https://beta.quicklisp.org/quicklisp.lisp
sbcl --load quicklisp.lisp
```

```lisp
(quicklisp-quickstart:install)
(ql:add-to-init-file)
(quit)
```

### 3. Tcl/Tkのインストール

```bash
# Ubuntu/Debian
sudo apt install tk

# macOS (Homebrew)
brew install tcl-tk

# Windows
# Tcl/Tkは通常プリインストール済み
```

## 使い方

### 起動

```lisp
(load "spreadsheet-gui.lisp")
(ss-gui:start)
```

### 操作方法

| 操作 | 説明 |
|------|------|
| `矢印キー` | セル移動 |
| `クリック` | セル選択 |
| `Enter` (グリッド上) | 入力欄へフォーカス |
| `Enter` (入力欄) | 入力確定 |

### 数式の入力

数式は `=` で始め、S式で記述します。

#### 基本演算

```lisp
=(+ A1 B1)        ; A1 + B1
=(- A1 B1)        ; A1 - B1
=(* A1 2)         ; A1 × 2
=(/ A1 B1)        ; A1 ÷ B1
```

#### 集計関数

```lisp
=(sum A1 B1 C1)   ; 合計
=(avg A1 B1 C1)   ; 平均
=(max A1 B1 C1)   ; 最大値
=(min A1 B1 C1)   ; 最小値
=(count A1 B1 C1) ; 個数
```

#### 範囲指定

```lisp
=(sum (range A1 A5))    ; A1〜A5の合計（縦）
=(sum (range A1 C1))    ; A1〜C1の合計（横）
=(sum (range A1 C3))    ; 矩形範囲の合計
=(avg (range A1 A10))   ; 範囲の平均
```

#### 入れ子

```lisp
=(+ (sum (range A1 A5)) B1)   ; 範囲合計 + 単一セル
=(* (avg (range A1 A3)) 2)    ; 平均 × 2
```

## アーキテクチャ

```
┌─────────────────────────────────────────────────────┐
│                    GUI (LTK/Tk)                     │
│  ┌──────────────┐  ┌─────────────────────────────┐  │
│  │    Entry     │  │          Canvas             │  │
│  │  (数式入力)   │  │   (セルグリッド描画)         │  │
│  └──────────────┘  └─────────────────────────────┘  │
├─────────────────────────────────────────────────────┤
│                  イベント処理                        │
│    キー入力 / マウスクリック / Enter確定            │
├─────────────────────────────────────────────────────┤
│                  数式評価エンジン                    │
│  ┌─────────────┐  ┌─────────────┐  ┌────────────┐  │
│  │ eval-formula│  │ expand-range│  │ cell-ref-p │  │
│  │ (S式評価)   │  │ (範囲展開)  │  │ (参照判定) │  │
│  └─────────────┘  └─────────────┘  └────────────┘  │
├─────────────────────────────────────────────────────┤
│                  データモデル                        │
│    *sheet* (ハッシュテーブル: セル名 → cell構造体)  │
│    cell: value (表示値) + formula (数式)            │
└─────────────────────────────────────────────────────┘
```

## 設定

`spreadsheet-gui.lisp` の先頭で変更可能：

```lisp
(defparameter *rows* 20)       ; 行数
(defparameter *cols* 10)       ; 列数
(defparameter *cell-w* 80)     ; セル幅（px）
(defparameter *cell-h* 24)     ; セル高さ（px）
```

## ファイル構成

```
.
├── README.md                 ; 英語ドキュメント
├── README-JP.md              ; 日本語ドキュメント
├── LICENSE
├── .gitignore
├── spreadsheet-gui.lisp      ; メインファイル（最新版）
└── spreadsheet-gui-v0.1.lisp ; v0.1 アーカイブ
```

## ロードマップ

### v0.2 (予定)

- [ ] ファイル保存/読み込み
- [ ] セルのコピー＆ペースト
- [ ] Undo/Redo

### v0.3 (予定)

- [ ] 列幅の調整
- [ ] セルの書式設定
- [ ] 条件付き書式

### 将来

- [ ] グラフ機能
- [ ] CSV インポート/エクスポート
- [ ] マクロ機能

## ライセンス

MIT License

## 作者

福寄

## 貢献

Issue、Pull Request 歓迎します。

## 参考

- [LTK Manual](http://www.peter-herth.de/ltk/ltkdoc/)
- [SBCL Manual](http://www.sbcl.org/manual/)
- [Common Lisp HyperSpec](http://www.lispworks.com/documentation/HyperSpec/Front/)
