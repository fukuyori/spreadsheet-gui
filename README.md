# Spreadsheet GUI

A Lisp-powered spreadsheet application built with Common Lisp and LTK

![Version](https://img.shields.io/badge/version-0.3-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[日本語版 README](README-JP.md)

## Overview

A unique spreadsheet where formulas are written in Lisp S-expressions. Cells can hold any Lisp value including numbers, lists, symbols, and strings. Supports lambda expressions, apply, and funcall for powerful data transformations.

## Screenshot

```
┌─────────────────────────────────────────────────────────────────────┐
│ =(mapcar (lambda (x) (* x x)) A1)                                   │
│                                                        [Enter: 確定]│
│                                                  [Shift+Enter: 改行]│
├────┬────────────┬────────────────────────┬──────────┬───────────────┤
│    │     A      │           B            │    C     │   ...         │
├────┼────────────┼────────────────────────┼──────────┼───────────────┤
│  1 │ (1 2 3 4 5)│ (1 4 9 16 25)          │          │               │
│  2 │     10     │ Hello, World!                     │               │
│ .. │            │                        │          │               │
└────┴────────────┴────────────────────────┴──────────┴───────────────┘
       
* Configurable grid size via startup parameters
* Multi-line formula input area
* Long text overflows into empty cells
```

## Features

- **Configurable grid** - Specify rows, columns, input area at startup
- **S-expression formulas** - Full Lisp syntax for calculations
- **Lambda expressions** - `(lambda (x) ...)` for custom transformations
- **apply / funcall** - Function application support
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
| Click | Select cell |
| Enter (grid) | Focus input area |
| Enter (input) | Confirm input |
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

; With lambda
=(apply (lambda (x y z) (+ x y z)) '(1 2 3))
; → 6

; Funcall
=(funcall #'+ 1 2 3)
; → 6

=(funcall (lambda (x) (* x x)) 5)
; → 25
```

### Higher-Order Functions

```lisp
=(mapcar #'1+ '(1 2 3))              ; → (2 3 4)
=(reduce #'* '(1 2 3 4 5))           ; → 120
=(remove-if #'evenp '(1 2 3 4 5))    ; → (1 3 5)
=(remove-if-not #'plusp '(-1 0 1 2)) ; → (1 2)
```

### Conditionals

```lisp
=(if (> A1 10) :big :small)

=(cond
  ((< A1 0) :negative)
  ((= A1 0) :zero)
  (t :positive))
```

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

## Cell Value Types & Colors

| Color | Type | Example |
|-------|------|---------|
| White | Number | `42` |
| Light Green | List | `(1 2 3)` |
| Light Pink | Symbol | `HELLO` |
| Light Yellow | String | `"text"` |
| Light Blue | Selected | current |

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
├── spreadsheet-gui.lisp      ; Latest
├── spreadsheet-gui-v0.1.lisp
├── spreadsheet-gui-v0.2.lisp
└── spreadsheet-gui-v0.3.lisp
```

## Version History

### v0.3 (Current)
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
