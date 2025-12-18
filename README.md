# Spreadsheet GUI

A Lisp-powered spreadsheet application built with Common Lisp and LTK

![Version](https://img.shields.io/badge/version-0.3.1-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[日本語版 README](README-JP.md)

## Overview

A unique spreadsheet where formulas are written in Lisp S-expressions. Cells can hold any Lisp value including numbers, lists, symbols, and strings. Supports lambda expressions, apply, funcall, and **automatic recalculation** for powerful data transformations.

## What's New in v0.3.1

- **Auto-recalculation** - Dependent cells update automatically when referenced cells change
- **Dependency tracking** - Topological sort ensures correct recalculation order
- **Range selection** - Mouse drag or Shift+Arrow keys to select multiple cells
- **Copy & Paste** - Ctrl+C/V with system clipboard support (TSV format)
- **Delete cells** - Delete/BackSpace to clear cell contents
- **Position reference** - `this-row`, `this-col`, `rel`, `rel-range` functions
- **Relative references** - Reference cells relative to current position
- **Debug functions** - `show-dependencies`, `show-cell-deps` for troubleshooting

## Screenshot

```
┌─────────────────────────────────────────────────────────────────────┐
│ =(+ A1 5)                                                           │
│                                                        [Enter: 確定]│
│                                                  [Shift+Enter: 改行]│
├────┬────────────┬────────────────────────┬──────────┬───────────────┤
│    │     A      │           B            │    C     │   ...         │
├────┼────────────┼────────────────────────┼──────────┼───────────────┤
│  1 │     10     │ (selected)             │          │               │
│  2 │     15     │ ← Auto-updated!        │          │               │
│  3 │     30     │ =(* A2 2)              │          │               │
│ .. │            │                        │          │               │
└────┴────────────┴────────────────────────┴──────────┴───────────────┘
       
A2 = =(+ A1 5)  → When A1 changes, A2 automatically recalculates!
```

## Features

- **Auto recalculation** - Dependent cells update automatically
- **Configurable grid** - Specify rows, columns, input area at startup
- **S-expression formulas** - Full Lisp syntax for calculations
- **Lambda expressions** - `(lambda (x) ...)` for custom transformations
- **apply / funcall** - Function application support
- **Position reference** - `this-row`, `this-col`, `rel`, `rel-range`
- **Relative references** - Reference cells relative to current position
- **Range selection** - Drag or Shift+Arrow to select multiple cells
- **System clipboard** - Copy/paste with Excel, Google Sheets (TSV format)
- **Multi-line input** - Formula input area supports multiple lines
- **Any Lisp value** - Numbers, lists, symbols, strings, keywords in cells
- **Text overflow** - Long content extends into empty adjacent cells
- **80+ pure functions** - Whitelisted non-destructive Lisp functions

## Requirements

- **SBCL** (Steel Bank Common Lisp)
- **Quicklisp**
- **LTK** (Lisp Toolkit)
- **Tcl/Tk** 8.5+

## Installation

```bash
# 1. Install SBCL
# Ubuntu/Debian
sudo apt install sbcl tk

# macOS
brew install sbcl tcl-tk

# 2. Install Quicklisp
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
;; Default: 26 rows, 14 columns, 3-line input
(ss-gui:start)

;; Custom size
(ss-gui:start :rows 30 :cols 10)

;; With larger input area
(ss-gui:start :rows 20 :cols 8 :input-lines 5)
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `:rows` | 26 | Number of rows |
| `:cols` | 14 | Number of columns (max 26 = A-Z) |
| `:input-lines` | 3 | Height of formula input area |

### Controls

| Key | Action |
|-----|--------|
| Arrow keys | Move cursor |
| Shift+Arrow | Extend selection |
| Click | Select cell |
| Drag | Range selection |
| Ctrl+C | Copy to system clipboard (TSV) |
| Ctrl+V | Paste from system clipboard |
| Delete/BackSpace | Clear cells (set to NIL) |
| Enter (grid) | Focus input area |
| Enter (input) | Confirm input & auto-recalc |
| Shift+Enter | Newline in input |

## Formula Examples

All formulas start with `=`

### Basic Arithmetic

```lisp
=(+ A1 B1)           ; Add
=(* A1 2)            ; Multiply
=(sqrt 2)            ; → 1.4142
=(expt 2 10)         ; → 1024
=(* pi 2)            ; → 6.2831
```

### Range Operations

```lisp
=(sum (range A1 A10))   ; Sum of A1:A10
=(avg (range A1 A26))   ; Average
=(+ (range A1 A5))      ; Same as sum
=(max (range B1 B26))   ; Maximum
```

### Relative References (NEW in v0.3.1)

```lisp
=(rel -1 0)          ; One row up
=(rel 0 -1)          ; One column left
=(rel 1 1)           ; One row down, one column right
=(sum (rel-range -4 0 -1 0))  ; Sum of 4 rows above
```

### Position Functions (NEW in v0.3.1)

```lisp
=(this-row)          ; Current row number (1-based)
=(this-col)          ; Current column number (0-based)
=(this-col-name)     ; Current column name ("A", "B", etc.)
=(this-cell-name)    ; Current cell name ("A1", "B2", etc.)
=(cell-at 1 0)       ; Value at row 1, column 0 (=A1)
```

### List Operations

```lisp
=(list 1 2 3 4 5)       ; → (1 2 3 4 5)
=(reverse '(a b c))     ; → (C B A)
=(append '(1 2) '(3 4)) ; → (1 2 3 4)
=(length '(1 2 3))      ; → 3
=(nth 2 '(a b c d))     ; → C
```

### Lambda Expressions

```lisp
; Square each element
=(mapcar (lambda (x) (* x x)) '(1 2 3 4 5))
; → (1 4 9 16 25)

; With cell reference (A1 contains a list)
=(mapcar (lambda (x) (* x x)) A1)

; Direct invocation
=((lambda (x) (+ x 1)) 10)
; → 11

; Filter: keep values > 10
=(remove-if-not (lambda (x) (> x 10)) '(5 15 8 20))
; → (15 20)

; Count even numbers
=(count-if (lambda (x) (evenp x)) '(1 2 3 4 5 6))
; → 3
```

### Apply and Funcall

```lisp
; Apply function to list
=(apply #'+ '(1 2 3 4 5))
; → 15

=(apply #'max '(3 1 4 1 5 9))
; → 9

; Funcall
=(funcall #'+ 1 2 3)
; → 6

=(funcall (lambda (x) (* x x)) 5)
; → 25
```

### Conditionals

```lisp
=(if (> A1 10) :big :small)

=(cond
  ((< A1 0) :negative)
  ((= A1 0) :zero)
  (t :positive))
```

## Auto-Recalculation (NEW in v0.3.1)

When a cell value changes, all dependent cells are automatically recalculated in the correct order.

```
A1: 10
A2: =(+ A1 5)     → 15
A3: =(* A2 2)     → 30

When A1 is changed to 20:
  A2 automatically becomes 25
  A3 automatically becomes 50
```

### How It Works

1. **Dependency Extraction** - When a formula is entered, all cell references are extracted
2. **Dependency Graph** - A bidirectional graph tracks which cells reference which
3. **Topological Sort** - When a cell changes, dependents are sorted by dependency order
4. **Cascading Recalc** - Cells are recalculated in order, ensuring correct values

Supported reference types for auto-recalc:
- Absolute: `A1`, `B2`
- Relative: `(rel -1 0)`, `(rel 0 1)`
- Range: `(range A1 A5)`, `(rel-range -4 0 -1 0)`
- Position: `(cell-at row col)`

## Debugging (NEW in v0.3.1)

```lisp
;; Show all dependency relationships
(ss-gui:show-dependencies)

;; Show dependencies for a specific cell
(ss-gui:show-cell-deps "A1")
```

Console output shows:
- Reference extraction: `Extract refs from (+ A1 5): ("A1")`
- Dependency updates: `Update deps: A2 refs ("A1")`
- Recalculation: `Recalc A2: 15 → 25`

## Allowed Functions

### Arithmetic
`+` `-` `*` `/` `mod` `rem` `1+` `1-` `abs` `max` `min` `sqrt` `expt` `log` `exp` `sin` `cos` `tan` `floor` `ceiling` `round` `gcd` `lcm`

### List
`car` `cdr` `cons` `list` `first` `rest` `last` `append` `reverse` `length` `nth` `member` `assoc` `subseq` `butlast`

### Higher-Order
`mapcar` `reduce` `remove-if` `remove-if-not` `count-if` `lambda` `apply` `funcall`

### Predicates
`atom` `listp` `numberp` `stringp` `symbolp` `null` `zerop` `plusp` `minusp` `evenp` `oddp`

### Comparison
`=` `/=` `<` `>` `<=` `>=` `equal`

### String
`string-upcase` `string-downcase` `concatenate` `subseq`

### Special
`if` `cond` `and` `or` `quote` `range` `sum` `avg` `count`

### Position Reference (NEW in v0.3.1)
`this-row` `this-col` `this-col-name` `this-cell-name` `cell-at` `rel` `rel-range`

## Cell Value Types & Colors

| Color | Type | Example |
|-------|------|---------|
| White | Number | `42` |
| Light Green | List | `(1 2 3)` |
| Light Pink | Symbol | `HELLO` |
| Light Yellow | String | `"text"` |
| Light Blue | Selected/Cursor | current |
| Pale Blue | Selection Range | multi-select |

## Specifications

| Item | Default | Range |
|------|---------|-------|
| Rows | 26 | 1-999 |
| Columns | 14 | 1-26 (A-Z) |
| Input Lines | 3 | 1-10 |
| Cell Size | 100 × 24 px | - |
| Formula Prefix | `=` | - |

## File Structure

```
├── README.md
├── README-JP.md
├── LICENSE
├── .gitignore
├── spreadsheet-gui.lisp      ; Latest (v0.3.1)
├── spreadsheet-gui-v0.1.lisp
├── spreadsheet-gui-v0.2.lisp
├── spreadsheet-gui-v0.3.lisp
└── spreadsheet-gui-v0.3.1.lisp
```

## Version History

### v0.3.1 (Current)
- **Auto-recalculation** - Dependent cells update automatically
- **Dependency graph tracking** with topological sort
- **Range selection** - Mouse drag to select multiple cells
- **Shift+Arrow key selection**
- **Copy & Paste** - Ctrl+C/V with internal clipboard
- **System clipboard support** - TSV format for Excel/Sheets
- **Delete cells** - Delete/BackSpace sets NIL
- **Position reference functions** - this-row, this-col, rel, rel-range
- **Relative cell references**
- **Debug functions** - show-dependencies, show-cell-deps

### v0.3
- Startup parameters: rows, cols, input-lines
- Multi-line formula input area
- Enter to confirm, Shift+Enter for newline
- apply / funcall support

### v0.2
- Lambda expression support
- Cell references in lambda
- Text overflow display
- 80+ Lisp functions

### v0.1
- Basic spreadsheet
- Four arithmetic operations
- Range selection

## License

MIT License

## Author

Fukuyori
