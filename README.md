# Spreadsheet GUI

A Lisp-powered spreadsheet application built with Common Lisp and LTK

![Version](https://img.shields.io/badge/version-0.4.2-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[日本語版 README](README-JP.md)

## Overview

A unique spreadsheet where formulas are written in Lisp S-expressions. Cells can hold any Lisp value including numbers, lists, symbols, and strings. Supports lambda expressions, auto-recalculation, Undo/Redo, file save/load, and CSV import/export.

## Screenshot

```
┌─────────────────────────────────────────────────────────────────────┐
│ [ファイル(F)] [編集(E)]                                              │
├─────────────────────────────────────────────────────────────────────┤
│ =(+ A1 5)                                                           │
├────┬────────────┬────────────────────────┬──────────┬───────────────┤
│    │     A      │           B            │    C     │   ...         │
├────┼────────────┼────────────────────────┼──────────┼───────────────┤
│  1 │         10 │ (1 2 3 4 5)            │          │               │
│  2 │         15 │ ← Auto-updated!        │          │               │
│  3 │         30 │ Hello                  │          │               │
│ .. │            │                        │          │               │
└────┴────────────┴────────────────────────┴──────────┴───────────────┘
         ↑ Numbers right-aligned    ↑ Text left-aligned
```

## Features

- **Direct Input Mode** - Start typing to begin editing immediately
- **Number Right-Alignment** - Numbers displayed right-aligned, text left-aligned
- **Undo/Redo** - Ctrl+Z/Y with up to 100 operations history
- **File Save/Load** - Native `.ssp` format preserves formulas and values
- **CSV Import/Export** - Interoperability with Excel, Google Sheets
- **Auto-recalculation** - Dependent cells update automatically
- **S-expression formulas** - Full Lisp syntax for calculations
- **Lambda expressions** - `(lambda (x) ...)` for custom transformations
- **Position reference** - `this-row`, `this-col`, `rel`, `rel-range`
- **Range selection** - Drag or Shift+Arrow to select multiple cells
- **System clipboard** - Copy/paste with other applications (TSV format)
- **Any Lisp value** - Numbers, lists, symbols, strings, keywords in cells
- **80+ pure functions** - Whitelisted non-destructive Lisp functions
- **Improved error handling** - Clear error messages (#構文, #評価, #循環)

## Requirements

- **SBCL** (Steel Bank Common Lisp)
- **Quicklisp**
- **LTK** (Lisp Toolkit)
- **Tcl/Tk** 8.5+

## Installation

```bash
# Ubuntu/Debian
sudo apt install sbcl tk

# macOS
brew install sbcl tcl-tk

# Install Quicklisp
curl -O https://beta.quicklisp.org/quicklisp.lisp
sbcl --load quicklisp.lisp
```

```lisp
(quicklisp-quickstart:install)
(ql:add-to-init-file)
(quit)
```

## Usage

```lisp
(load "spreadsheet-gui.lisp")
(ss-gui:start)
```

### Startup Parameters

```lisp
(ss-gui:start)                           ; Default: 26 rows, 14 columns
(ss-gui:start :rows 30 :cols 10)         ; Custom size
(ss-gui:start :input-lines 5)            ; Larger input area
```

## Keyboard Shortcuts

### Basic Operations

| Key | Action |
|-----|--------|
| Arrow keys | Move cursor |
| Shift+Arrow | Extend selection |
| Click/Drag | Select cells |
| Any character | Start typing (clears cell) |
| F2 | Edit mode (keep content) |
| Escape | Cancel editing |
| Delete/BackSpace | Clear cells |

### Input Confirmation

| Key | Action |
|-----|--------|
| Enter | Confirm → Move down |
| Ctrl+Enter | Confirm → Stay |
| Shift+Enter | Newline in input |
| Alt+Enter | Confirm → Move right |
| Shift+Alt+Enter | Confirm → Move left |

### Edit Operations

| Key | Action |
|-----|--------|
| Ctrl+Z | Undo |
| Ctrl+Y | Redo |
| Ctrl+X | Cut |
| Ctrl+C | Copy |
| Ctrl+V | Paste |

### File Operations

| Key | Action |
|-----|--------|
| Ctrl+N | New sheet |
| Ctrl+O | Open file |
| Ctrl+S | Save file |

## Menus

### File Menu (ファイル)

| Item | Description |
|------|-------------|
| 新規作成(N) | Create new empty sheet |
| 開く(O)... | Open .ssp file |
| 保存(S) | Save to current file |
| 名前を付けて保存(A)... | Save as new file |
| CSVインポート... | Import from CSV |
| CSVエクスポート... | Export to CSV |
| 終了(X) | Exit application |

### Edit Menu (編集)

| Item | Description |
|------|-------------|
| 元に戻す(U) | Undo (Ctrl+Z) |
| やり直し(R) | Redo (Ctrl+Y) |
| 切り取り(X) | Cut (Ctrl+X) |
| コピー(C) | Copy (Ctrl+C) |
| 貼り付け(V) | Paste (Ctrl+V) |
| 削除 | Delete selection |

## File Operations

### Save and Load (.ssp format)

```lisp
;; Save current spreadsheet
(ss-gui:save "mydata.ssp")

;; Load from file
(ss-gui:load-file "mydata.ssp")

;; Create new empty sheet
(ss-gui:new-sheet)

;; Check current file
ss-gui:*current-file*
```

### .ssp File Format

The `.ssp` format is a human-readable S-expression format:

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

### CSV Export/Import

```lisp
;; Export to CSV (data only)
(ss-gui:export-csv "data.csv")

;; Export with column header (A, B, C...)
(ss-gui:export-csv "data.csv" :include-header t)

;; Export with formulas
(ss-gui:export-csv "data.csv" :include-formulas t)

;; Import from CSV
(ss-gui:import-csv "data.csv")

;; Import skipping header row
(ss-gui:import-csv "data.csv" :has-header t)
```

## Formula Examples

### Basic Arithmetic

```lisp
=(+ A1 B1)           ; Add
=(* A1 2)            ; Multiply
=(sqrt 2)            ; → 1.4142
=(expt 2 10)         ; → 1024
```

### Range Operations

```lisp
=(sum (range A1 A10))   ; Sum of A1:A10
=(avg (range A1 A26))   ; Average
=(max (range B1 B26))   ; Maximum
```

### Relative References

```lisp
=(rel -1 0)          ; One row up
=(rel 0 -1)          ; One column left
=(sum (rel-range -4 0 -1 0))  ; Sum of 4 rows above
```

### Lambda Expressions

```lisp
=(mapcar (lambda (x) (* x x)) '(1 2 3 4 5))
; → (1 4 9 16 25)

=(remove-if-not (lambda (x) (> x 10)) '(5 15 8 20))
; → (15 20)
```

### Conditionals

```lisp
=(if (> A1 10) :big :small)

=(cond
  ((< A1 0) :negative)
  ((= A1 0) :zero)
  (t :positive))
```

## Error Messages

| Error | Meaning |
|-------|---------|
| #構文: ... | Syntax error in formula |
| #評価: ... | Evaluation error |
| #循環: A1 | Circular reference detected |

## Undo/Redo

- Supports up to 100 operations
- Tracks cell value and formula changes
- Multi-cell paste operations are grouped

```lisp
;; Programmatic access
(ss-gui:undo)         ; Undo last operation
(ss-gui:redo)         ; Redo
(ss-gui:clear-history) ; Clear undo/redo history
```

## Auto-Recalculation

When a cell value changes, all dependent cells are automatically recalculated:

```
A1: 10
A2: =(+ A1 5)     → 15
A3: =(* A2 2)     → 30

When A1 changes to 20:
  A2 → 25
  A3 → 50
```

## Version History

### v0.4.2 (Current)
- Number right-alignment in cells
- File menu with save/load (.ssp format)
- Edit menu with Undo/Redo/Cut/Copy/Paste
- Direct input mode (type to start editing)
- F2 for edit mode
- Multiple Enter key behaviors (down/stay/right/left)
- CSV import/export
- Improved error handling
- Keyboard shortcuts (Ctrl+N/O/S/Z/Y/X/C/V)

### v0.3.1
- Auto-recalculation with dependency tracking
- Range selection (drag, Shift+Arrow)
- Copy & Paste with system clipboard
- Position reference functions (rel, rel-range)

### v0.3
- Startup parameters (rows, cols, input-lines)
- Multi-line formula input
- apply / funcall support

### v0.2
- Lambda expressions
- Text overflow display
- 80+ Lisp functions

### v0.1
- Basic spreadsheet functionality

## License

MIT License

## Author

Fukuyori
