# Spreadsheet GUI

A Lisp-powered spreadsheet application built with Common Lisp and LTK

![Version](https://img.shields.io/badge/version-0.2-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Lisp](https://img.shields.io/badge/Lisp-Common%20Lisp-red.svg)

[日本語版 README](README-JP.md)

## Overview

A unique spreadsheet where formulas are written in Lisp S-expressions. Cells can hold any Lisp value including numbers, lists, symbols, and strings. Supports lambda expressions for powerful data transformations.

## Screenshot

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
       (list/green) (list/green)            
       
* Long text overflows into empty cells on the right
* Background colors indicate value types
```

## Features

- **S-expression formulas** - Full Lisp syntax for calculations
- **Lambda expressions** - `(lambda (x) ...)` for custom transformations
- **Any Lisp value** - Numbers, lists, symbols, strings, keywords in cells
- **Cell references in lambda** - `(mapcar (lambda (x) (* x x)) A1)`
- **Text overflow** - Long content extends into empty adjacent cells
- **Visual type hints** - Background colors by value type
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

### Controls

| Key | Action |
|-----|--------|
| Arrow keys | Move cursor |
| Click | Select cell |
| Enter (grid) | Edit cell |
| Enter (input) | Confirm |

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
=(sum (range A1 A5))    ; Sum of A1:A5
=(avg (range A1 A10))   ; Average
=(+ (range A1 A5))      ; Same as sum
=(max (range B1 B10))   ; Maximum
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

; Filter: keep values > 10
=(remove-if-not (lambda (x) (> x 10)) '(5 15 8 20))
; → (15 20)

; Count even numbers
=(count-if (lambda (x) (evenp x)) '(1 2 3 4 5 6))
; → 3

; Custom reduce
=(reduce (lambda (a b) (+ a b)) '(1 2 3 4 5))
; → 15
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

### String Operations

```lisp
=(string-upcase "hello")               ; → "HELLO"
=(concatenate 'string "Hello" "World") ; → "HelloWorld"
```

## Allowed Functions

### Arithmetic
`+` `-` `*` `/` `mod` `rem` `1+` `1-` `abs` `max` `min` `sqrt` `expt` `log` `exp` `sin` `cos` `tan` `floor` `ceiling` `round` `gcd` `lcm`

### List
`car` `cdr` `cons` `list` `first` `rest` `last` `append` `reverse` `length` `nth` `member` `assoc` `subseq` `butlast`

### Higher-Order
`mapcar` `reduce` `remove-if` `remove-if-not` `count-if` `lambda`

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

## Architecture

```
┌──────────────────────────────────────────┐
│              GUI (LTK/Tk)                │
│  Entry (formula) + Canvas (grid)         │
├──────────────────────────────────────────┤
│           Formula Evaluator              │
│  - S-expression parser                   │
│  - Lambda with closure support           │
│  - Cell reference resolution             │
│  - 80+ whitelisted pure functions        │
├──────────────────────────────────────────┤
│             Data Model                   │
│  Hash table: "A1" → (value . formula)    │
│  Values: Any Lisp object                 │
└──────────────────────────────────────────┘
```

## File Structure

```
├── README.md              ; English
├── README-JP.md           ; Japanese
├── LICENSE                ; MIT
├── .gitignore
├── spreadsheet-gui.lisp   ; Latest
└── spreadsheet-gui-v0.2.lisp
```

## Version History

### v0.2 (Current)
- Lambda expression support
- Cell references in lambda
- Text overflow display
- 80+ Lisp functions
- Type-based cell coloring

### v0.1
- Basic spreadsheet
- Four arithmetic operations
- Range selection
- sum, avg, max, min

## License

MIT License

## Author

Fukuyori
