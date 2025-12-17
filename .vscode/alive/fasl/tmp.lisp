;;;; =====================================================
;;;; spreadsheet-gui.lisp
;;;; Common Lisp + LTK で作るシンプルな表計算ソフト
;;;; 
;;;; Version: 0.3
;;;; Date: 2025-01-15
;;;; 
;;;; 機能:
;;;;   - セルへの値・数式入力
;;;;   - Lispの非破壊関数をサポート
;;;;   - lambda式、apply、funcall
;;;;   - 戻り値はリスト、シンボル等任意のLisp値
;;;;   - 範囲指定 (range A1 A5)
;;;;   - 条件分岐 (if ...) (cond ...)
;;;;   - 矢印キー・クリックでセル移動
;;;; 
;;;; 使い方: (load "spreadsheet-gui.lisp") (ss-gui:start)
;;;; =====================================================

(ql:quickload :ltk)

;; パッケージ再読み込み時のエラー回避
(when (find-package :ss-gui)
  (delete-package :ss-gui))

(defpackage :ss-gui
  (:use :cl :ltk)
  (:export :start))
(in-package :ss-gui)

;;;; =========================
;;;; 設定（グリッドサイズ・セル寸法）
;;;; =========================

(defparameter *rows* 26)          ; 行数
(defparameter *cols* 14)          ; 列数 (A-N)
(defparameter *cell-w* 100)       ; セル幅（ピクセル）※リスト表示用に拡大
(defparameter *cell-h* 24)        ; セル高さ（ピクセル）
(defparameter *header-h* 24)      ; 列ヘッダー(A,B,C...)の高さ
(defparameter *header-w* 40)      ; 行ヘッダー(1,2,3...)の幅
(defparameter *cur-x* 0)          ; カーソル位置（列）
(defparameter *cur-y* 0)          ; カーソル位置（行）

;;;; =========================
;;;; データモデル
;;;; =========================

;; セル構造体：値と数式を保持
(defstruct cell
  value      ; 表示値（任意のLisp値）
  formula)   ; 数式（S式）、なければnil

;; シート本体：セル名("A1"等)をキーとするハッシュテーブル
(defparameter *sheet* (make-hash-table :test #'equal))

(defun get-cell (name)
  "セルを取得。存在しなければ新規作成"
  (or (gethash name *sheet*)
      (setf (gethash name *sheet*) (make-cell))))

(defun cell-name (x y)
  "座標(x,y)からセル名を生成。(0,0)->\"A1\", (1,2)->\"B3\""
  (format nil "~a~d"
          (code-char (+ (char-code #\A) x))
          (1+ y)))

(defun current-cell ()
  "現在カーソル位置のセルを取得"
  (get-cell (cell-name *cur-x* *cur-y*)))

;;;; =========================
;;;; ユーティリティ
;;;; =========================

(defun flatten (lst)
  "ネストしたリストを平坦化"
  (cond ((null lst) nil)
        ((atom lst) (list lst))
        (t (append (flatten (car lst))
                   (flatten (cdr lst))))))

(defun try-parse-number (val)
  "文字列を数値に変換。失敗時は元の値を返す"
  (if (stringp val)
      (let ((parsed (ignore-errors (read-from-string val))))
        (if (numberp parsed) parsed val))
      val))

;;;; =========================
;;;; 許可された関数リスト
;;;; =========================

(defparameter *allowed-functions*
  '(;; 算術演算
    + - * / mod rem 1+ 1-
    floor ceiling round truncate
    abs max min signum
    sqrt expt log exp isqrt
    sin cos tan asin acos atan sinh cosh tanh
    gcd lcm
    ;; 比較
    = /= < > <= >=
    equal equalp eq eql
    ;; リスト操作
    car cdr cons list
    first second third fourth fifth sixth seventh eighth ninth tenth
    rest last butlast nthcdr
    append reverse length nth elt
    member assoc find position
    mapcar mapc maplist remove remove-if remove-if-not
    reduce count count-if
    subseq copy-list copy-seq
    list-length
    ;; リスト作成
    make-list iota
    ;; 文字列
    string-upcase string-downcase string-capitalize
    string-trim string-left-trim string-right-trim
    concatenate subseq char
    string= string/= string< string> string<= string>=
    string-equal
    parse-integer
    ;; 文字
    char-upcase char-downcase char-code code-char
    alpha-char-p digit-char-p upper-case-p lower-case-p
    ;; 論理
    not null and or
    ;; 述語
    atom listp consp numberp integerp floatp rationalp
    stringp symbolp characterp keywordp functionp
    zerop plusp minusp evenp oddp
    ;; 型変換
    float truncate round floor ceiling
    string coerce
    ;; ユーティリティ
    identity constantly values
    ;; 数学定数（変数として）
    pi
    ;; スプレッドシート専用
    sum avg count range ref
    ;; lambda / apply / funcall
    lambda apply funcall))

;;;; =========================
;;;; 数式評価エンジン（拡張版）
;;;; =========================

(defun cell-ref-p (sym)
  "シンボルがセル参照か判定。A1〜Z99形式を認識"
  (and (symbolp sym)
       (not (keywordp sym))
       (let ((name (symbol-name sym)))
         (and (>= (length name) 2)
              (<= (length name) 3)
              (alpha-char-p (char name 0))
              (every #'digit-char-p (subseq name 1))
              ;; 許可された関数名でないことを確認
              (not (member (string-upcase name) 
                          (mapcar #'symbol-name *allowed-functions*)
                          :test #'string=))))))

(defun parse-cell-coords (sym)
  "セル参照シンボルから座標(col row)を取得"
  (let* ((name (symbol-name sym))
         (col (- (char-code (char-upcase (char name 0))) (char-code #\A)))
         (row (1- (parse-integer (subseq name 1)))))
    (list col row)))

(defun get-cell-by-symbol (sym)
  "シンボル(A1等)からセル値を取得（生の値）"
  (let* ((coords (parse-cell-coords sym))
         (col (first coords))
         (row (second coords)))
    (cell-value (get-cell (cell-name col row)))))

(defun ref (name)
  "文字列でセル参照"
  (let* ((col (- (char-code (char-upcase (char name 0))) (char-code #\A)))
         (row (1- (parse-integer (subseq name 1)))))
    (cell-value (get-cell (cell-name col row)))))

(defun expand-range (start-sym end-sym)
  "範囲を展開してセル値のリストを返す"
  (let* ((start-coords (parse-cell-coords start-sym))
         (end-coords (parse-cell-coords end-sym))
         (col1 (first start-coords))
         (row1 (second start-coords))
         (col2 (first end-coords))
         (row2 (second end-coords))
         (min-col (min col1 col2))
         (max-col (max col1 col2))
         (min-row (min row1 row2))
         (max-row (max row1 row2))
         (values '()))
    (loop for r from min-row to max-row do
          (loop for c from min-col to max-col do
                (push (cell-value (get-cell (cell-name c r))) values)))
    (nreverse values)))

(defun allowed-function-p (sym)
  "関数が許可リストに含まれるか確認"
  (let ((name (string-upcase (symbol-name sym))))
    (member name (mapcar #'symbol-name *allowed-functions*)
            :test #'string=)))

(defun get-lisp-function (op-name)
  "関数名から実際の関数を取得。可変引数関数はリストを自動展開"
  (cond
    ;; スプレッドシート専用関数
    ((string-equal op-name "SUM")
     (lambda (&rest args)
       (apply #'+ (remove-if-not #'numberp (flatten args)))))
    ((string-equal op-name "AVG")
     (lambda (&rest args)
       (let ((nums (remove-if-not #'numberp (flatten args))))
         (if nums (float (/ (apply #'+ nums) (length nums))) 0))))
    ((string-equal op-name "COUNT")
     (lambda (&rest args)
       (length (flatten args))))
    ;; iota関数（範囲リスト生成）
    ((string-equal op-name "IOTA")
     (lambda (n &optional (start 0) (step 1))
       (loop for i from 0 below n collect (+ start (* i step)))))
    ;; 可変引数の算術演算（リスト自動展開）
    ((string-equal op-name "+")
     (lambda (&rest args)
       (apply #'+ (remove-if-not #'numberp (flatten args)))))
    ((string-equal op-name "-")
     (lambda (&rest args)
       (let ((nums (remove-if-not #'numberp (flatten args))))
         (if nums (apply #'- nums) 0))))
    ((string-equal op-name "*")
     (lambda (&rest args)
       (apply #'* (remove-if-not #'numberp (flatten args)))))
    ((string-equal op-name "/")
     (lambda (&rest args)
       (let ((nums (remove-if-not #'numberp (flatten args))))
         (if (and nums (not (member 0 (cdr nums))))
             (apply #'/ nums)
             "DIV/0!"))))
    ((string-equal op-name "MAX")
     (lambda (&rest args)
       (let ((nums (remove-if-not #'numberp (flatten args))))
         (if nums (apply #'max nums) 0))))
    ((string-equal op-name "MIN")
     (lambda (&rest args)
       (let ((nums (remove-if-not #'numberp (flatten args))))
         (if nums (apply #'min nums) 0))))
    ;; 標準Lisp関数
    (t 
     (let ((cl-sym (find-symbol op-name :cl)))
       (when (and cl-sym (fboundp cl-sym))
         (symbol-function cl-sym))))))

;;; lambda環境（動的束縛用）
(defparameter *lambda-env* nil)

(defun lookup-var (sym)
  "lambda環境から変数を探す"
  (let ((pair (assoc (symbol-name sym) *lambda-env* :test #'string-equal)))
    (if pair
        (values (cdr pair) t)
        (values nil nil))))

(defun eval-formula (expr)
  "S式の数式を評価。Lispの非破壊関数をサポート。
   戻り値: 数値、文字列、リスト、シンボルなど任意のLisp値"
  (cond
    ;; nil
    ((null expr) nil)
    ;; 数値・文字列はそのまま
    ((numberp expr) expr)
    ((stringp expr) expr)
    ;; キーワードシンボルはそのまま
    ((keywordp expr) expr)
    ;; クォートされた式
    ((and (listp expr) (eq (car expr) 'quote))
     (cadr expr))
    ;; シンボルの場合：lambda変数 > セル参照 > 定数 > そのまま
    ((symbolp expr)
     (multiple-value-bind (val found) (lookup-var expr)
       (cond
         ;; lambda変数に見つかった
         (found val)
         ;; セル参照
         ((cell-ref-p expr)
          (let ((val (get-cell-by-symbol expr)))
            (if (stringp val) (try-parse-number val) val)))
         ;; PI定数
         ((string-equal (symbol-name expr) "PI") pi)
         ;; その他のシンボルはそのまま
         (t expr))))
    ;; リスト（関数呼び出し）
    ((listp expr)
     (let* ((op (car expr))
            (op-name (if (symbolp op) (string-upcase (symbol-name op)) "")))
       (cond
         ;; ((lambda (x) ...) args...) - lambda式の直接呼び出し
         ((and (listp op)
               (symbolp (car op))
               (string-equal (symbol-name (car op)) "LAMBDA"))
          (let ((fn (eval-formula op))
                (args (mapcar #'eval-formula (cdr expr))))
            (if (functionp fn)
                (apply fn args)
                (format nil "ERR: not a function"))))
         ;; LAMBDA - 無名関数を作成
         ((string-equal op-name "LAMBDA")
          (let ((params (cadr expr))
                (body (caddr expr)))
            ;; クロージャとして現在の環境をキャプチャ
            (let ((captured-env *lambda-env*))
              (lambda (&rest args)
                (let ((*lambda-env* (append (mapcar #'cons 
                                                    (mapcar #'symbol-name params)
                                                    args)
                                            captured-env)))
                  (eval-formula body))))))
         ;; 範囲指定 (range A1 A5)
         ((string-equal op-name "RANGE")
          (if (and (>= (length expr) 3)
                   (cell-ref-p (cadr expr))
                   (cell-ref-p (caddr expr)))
              (expand-range (cadr expr) (caddr expr))
              :error))
         ;; IF式（短絡評価）
         ((string-equal op-name "IF")
          (if (eval-formula (cadr expr))
              (eval-formula (caddr expr))
              (eval-formula (cadddr expr))))
         ;; COND式
         ((string-equal op-name "COND")
          (loop for clause in (cdr expr)
                when (eval-formula (car clause))
                return (eval-formula (cadr clause))))
         ;; AND（短絡評価）
         ((string-equal op-name "AND")
          (loop for arg in (cdr expr)
                for val = (eval-formula arg)
                unless val return nil
                finally (return val)))
         ;; OR（短絡評価）
         ((string-equal op-name "OR")
          (loop for arg in (cdr expr)
                for val = (eval-formula arg)
                when val return val))
         ;; FUNCTION (#') - シンボルを関数として返す
         ((string-equal op-name "FUNCTION")
          (let ((fn-name (cadr expr)))
            (if (symbolp fn-name)
                (get-lisp-function (symbol-name fn-name))
                fn-name)))
         ;; MAPCAR（特別扱い：第一引数が関数）
         ((string-equal op-name "MAPCAR")
          (let* ((fn-expr (cadr expr))
                 (fn (cond
                       ;; #'func 形式
                       ((and (listp fn-expr) (eq (car fn-expr) 'function))
                        (get-lisp-function (symbol-name (cadr fn-expr))))
                       ;; lambda 形式
                       ((and (listp fn-expr) 
                             (string-equal (symbol-name (car fn-expr)) "LAMBDA"))
                        (eval-formula fn-expr))
                       ;; その他（評価）
                       (t (eval-formula fn-expr))))
                 (lists (mapcar #'eval-formula (cddr expr))))
            (if (functionp fn)
                (apply #'mapcar fn lists)
                (format nil "ERR: not a function"))))
         ;; REDUCE
         ((string-equal op-name "REDUCE")
          (let* ((fn-expr (cadr expr))
                 (fn (cond
                       ((and (listp fn-expr) (eq (car fn-expr) 'function))
                        (get-lisp-function (symbol-name (cadr fn-expr))))
                       ((and (listp fn-expr)
                             (string-equal (symbol-name (car fn-expr)) "LAMBDA"))
                        (eval-formula fn-expr))
                       (t (eval-formula fn-expr))))
                 (lst (eval-formula (caddr expr)))
                 (rest-args (cdddr expr)))
            (if (functionp fn)
                (apply #'reduce fn lst rest-args)
                (format nil "ERR: not a function"))))
         ;; APPLY
         ((string-equal op-name "APPLY")
          (let* ((fn-expr (cadr expr))
                 (fn (cond
                       ((and (listp fn-expr) (eq (car fn-expr) 'function))
                        (get-lisp-function (symbol-name (cadr fn-expr))))
                       ((and (listp fn-expr)
                             (string-equal (symbol-name (car fn-expr)) "LAMBDA"))
                        (eval-formula fn-expr))
                       (t (eval-formula fn-expr))))
                 (args (mapcar #'eval-formula (cddr expr))))
            (if (functionp fn)
                ;; 最後の引数がリストならspread、そうでなければそのまま
                (let* ((last-arg (car (last args)))
                       (init-args (butlast args))
                       (all-args (if (listp last-arg)
                                     (append init-args last-arg)
                                     args)))
                  (apply fn all-args))
                (format nil "ERR: not a function"))))
         ;; FUNCALL
         ((string-equal op-name "FUNCALL")
          (let* ((fn-expr (cadr expr))
                 (fn (cond
                       ((and (listp fn-expr) (eq (car fn-expr) 'function))
                        (get-lisp-function (symbol-name (cadr fn-expr))))
                       ((and (listp fn-expr)
                             (string-equal (symbol-name (car fn-expr)) "LAMBDA"))
                        (eval-formula fn-expr))
                       (t (eval-formula fn-expr))))
                 (args (mapcar #'eval-formula (cddr expr))))
            (if (functionp fn)
                (apply fn args)
                (format nil "ERR: not a function"))))
         ;; REMOVE-IF / REMOVE-IF-NOT / COUNT-IF
         ((or (string-equal op-name "REMOVE-IF")
              (string-equal op-name "REMOVE-IF-NOT")
              (string-equal op-name "COUNT-IF"))
          (let* ((fn-expr (cadr expr))
                 (fn (cond
                       ((and (listp fn-expr) (eq (car fn-expr) 'function))
                        (get-lisp-function (symbol-name (cadr fn-expr))))
                       ((and (listp fn-expr)
                             (string-equal (symbol-name (car fn-expr)) "LAMBDA"))
                        (eval-formula fn-expr))
                       (t (eval-formula fn-expr))))
                 (lst (eval-formula (caddr expr))))
            (if (functionp fn)
                (funcall (cond
                          ((string-equal op-name "REMOVE-IF") #'remove-if)
                          ((string-equal op-name "REMOVE-IF-NOT") #'remove-if-not)
                          (t #'count-if))
                         fn lst)
                (format nil "ERR: not a function"))))
         ;; LIST（特別扱い：リストを作成）
         ((string-equal op-name "LIST")
          (mapcar #'eval-formula (cdr expr)))
         ;; QUOTE
         ((string-equal op-name "QUOTE")
          (cadr expr))
         ;; CONCATENATE（型指定付き）
         ((string-equal op-name "CONCATENATE")
          (let ((type (eval-formula (cadr expr)))
                (args (mapcar #'eval-formula (cddr expr))))
            (apply #'concatenate type args)))
         ;; 許可された関数
         ((allowed-function-p op)
          (let ((func (get-lisp-function op-name))
                (args (mapcar #'eval-formula (cdr expr))))
            (if func
                (handler-case (apply func args)
                  (error (e) (format nil "ERR:~A" (type-of e))))
                (format nil "ERR:~A undefined" op-name))))
         ;; 許可されていない関数
         (t (format nil "ERR:~A not allowed" op-name)))))
    ;; その他
    (t expr)))

;;;; =========================
;;;; 値の表示
;;;; =========================

(defun format-value (val)
  "任意のLisp値を表示用文字列に変換"
  (cond
    ((null val) "")  ; nilは空表示
    ((eq val t) "T")
    ((stringp val) val)
    ((numberp val)
     (if (floatp val)
         (format nil "~,4G" val)  ; 小数は適度な精度で
         (princ-to-string val)))
    ((keywordp val) (format nil ":~A" (symbol-name val)))
    ((symbolp val) (symbol-name val))
    ((listp val) (princ-to-string val))  ; リストは省略せず表示
    (t (princ-to-string val))))

;;;; =========================
;;;; 描画（Canvas操作）
;;;; =========================

(defun draw-cell-background (canvas x y val selected)
  "セルの背景のみを描画"
  (let* ((px (+ *header-w* (* x *cell-w*)))
         (py (+ *header-h* (* y *cell-h*)))
         (px2 (+ px *cell-w*))
         (py2 (+ py *cell-h*))
         ;; 値の型によって背景色を変える
         (bg (cond
               (selected "#cce5ff")                    ; 選択中
               ((null val) "white")                    ; 空
               ((listp val) "#f0fff0")                 ; リストは薄緑
               ((and val (symbolp val)) "#fff0f0")    ; シンボルは薄赤
               ((stringp val) "#fffff0")              ; 文字列は薄黄
               (t "white")))
         (path (widget-path canvas)))
    (format-wish "~a create rectangle ~a ~a ~a ~a -fill {~a} -outline gray"
                 path px py px2 py2 bg)))

(defun count-overflow-cells (x y)
  "右側の空セルの数をカウント（はみ出し用）"
  (loop for i from (1+ x) below *cols*
        while (null (cell-value (get-cell (cell-name i y))))
        count t))

(defun draw-cell-text (canvas x y val)
  "セルのテキストのみを描画（右側の空セルにはみ出し可能）"
  (when val
    (let* ((px (+ *header-w* (* x *cell-w*)))
           (py (+ *header-h* (* y *cell-h*)))
           (display-text (format-value val))
           (path (widget-path canvas))
           ;; テキストの概算幅（1文字約7ピクセル）
           (text-width (* (length display-text) 7))
           ;; はみ出し可能なセル数
           (overflow-cells (count-overflow-cells x y))
           ;; 使用可能な幅
           (available-width (+ *cell-w* (* overflow-cells *cell-w*))))
      ;; テキストがセル幅を超える場合、クリップ領域を設定
      (if (> text-width (- *cell-w* 8))
          ;; はみ出し表示（空セルがある場合）
          (format-wish "~a create text ~a ~a -anchor w -text {~a} -font {Consolas 10}"
                       path (+ px 4) (+ py (floor *cell-h* 2)) display-text)
          ;; 通常表示
          (format-wish "~a create text ~a ~a -anchor w -text {~a} -font {Consolas 10}"
                       path (+ px 4) (+ py (floor *cell-h* 2)) display-text)))))

(defun draw-headers (canvas)
  "列名(A,B,C...)と行番号(1,2,3...)のヘッダーを描画"
  (let ((path (widget-path canvas)))
    ;; 左上隅の空白セル
    (format-wish "~a create rectangle 0 0 ~a ~a -fill {#e0e0e0} -outline gray"
                 path *header-w* *header-h*)
    ;; 列名ヘッダー
    (dotimes (x *cols*)
      (let* ((px (+ *header-w* (* x *cell-w*)))
             (px2 (+ px *cell-w*))
             (col-name (string (code-char (+ (char-code #\A) x)))))
        (format-wish "~a create rectangle ~a 0 ~a ~a -fill {#e0e0e0} -outline gray"
                     path px px2 *header-h*)
        (format-wish "~a create text ~a ~a -anchor center -text {~a} -font {Consolas 11 bold}"
                     path (+ px (floor *cell-w* 2)) (floor *header-h* 2) col-name)))
    ;; 行番号ヘッダー
    (dotimes (y *rows*)
      (let* ((py (+ *header-h* (* y *cell-h*)))
             (py2 (+ py *cell-h*))
             (row-num (1+ y)))
        (format-wish "~a create rectangle 0 ~a ~a ~a -fill {#e0e0e0} -outline gray"
                     path py *header-w* py2)
        (format-wish "~a create text ~a ~a -anchor center -text {~a} -font {Consolas 11 bold}"
                     path (floor *header-w* 2) (+ py (floor *cell-h* 2)) row-num)))))

(defun redraw (canvas)
  "画面全体を再描画（2パス：背景→テキスト）"
  (format-wish "~a delete all" (widget-path canvas))
  (draw-headers canvas)
  ;; パス1: 全セルの背景を描画
  (dotimes (y *rows*)
    (dotimes (x *cols*)
      (let ((cell (get-cell (cell-name x y))))
        (draw-cell-background canvas x y
                              (cell-value cell)
                              (and (= x *cur-x*) (= y *cur-y*))))))
  ;; パス2: 全セルのテキストを描画（背景の上に重ねる）
  (dotimes (y *rows*)
    (dotimes (x *cols*)
      (let ((cell (get-cell (cell-name x y))))
        (draw-cell-text canvas x y (cell-value cell))))))

;;;; =========================
;;;; 入力欄（Text）の操作
;;;; =========================

(defun get-text-content (text-widget)
  "Textウィジェットの内容を取得"
  (let ((content (text text-widget)))
    ;; 末尾の改行を除去
    (string-right-trim '(#\Newline #\Return) content)))

(defun set-text-content (text-widget s)
  "Textウィジェットに文字列を設定"
  (setf (text text-widget) (if s (princ-to-string s) "")))

(defun update-text-input (text-widget)
  "現在セルの内容をTextウィジェットに表示"
  (let ((c (current-cell)))
    (cond
      ;; 数式があれば =(...) 形式で表示
      ((cell-formula c)
       (set-text-content text-widget (format nil "=~S" (cell-formula c))))
      ;; 値があればそのまま表示
      ((cell-value c)
       (set-text-content text-widget (format nil "~S" (cell-value c))))
      (t
       (set-text-content text-widget "")))))

(defun commit-text-input (text-widget canvas)
  "Textウィジェットの内容をセルに確定"
  (let* ((raw (or (get-text-content text-widget) ""))
         (cell (current-cell)))
    (handler-case
        (if (and (> (length raw) 0)
                 (char= (char raw 0) #\=))
            ;; =で始まる → 数式として解析・評価
            (let* ((form (read-from-string (subseq raw 1)))
                   (value (eval-formula form)))
              (setf (cell-formula cell) form
                    (cell-value cell) value))
            ;; それ以外 → 通常の値として解釈
            (let ((parsed (ignore-errors (read-from-string raw))))
              (setf (cell-formula cell) nil
                    (cell-value cell) (or parsed raw))))
      (error (e)
        (setf (cell-value cell) (format nil "ERR: ~a" e))))
    (redraw canvas)))

;;; 後方互換性のため旧関数も残す
(defun entry-text (e) (get-text-content e))
(defun set-entry-text (e s) (set-text-content e s))
(defun update-entry (e) (update-text-input e))
(defun commit-entry (e c) (commit-text-input e c))

;;;; =========================
;;;; カーソル移動
;;;; =========================

(defun clamp (v lo hi)
  "値を範囲内に制限"
  (max lo (min hi v)))

(defun move-left (canvas text-widget)
  (setf *cur-x* (max 0 (1- *cur-x*)))
  (update-text-input text-widget)
  (redraw canvas))

(defun move-right (canvas text-widget)
  (setf *cur-x* (min (1- *cols*) (1+ *cur-x*)))
  (update-text-input text-widget)
  (redraw canvas))

(defun move-up (canvas text-widget)
  (setf *cur-y* (max 0 (1- *cur-y*)))
  (update-text-input text-widget)
  (redraw canvas))

(defun move-down (canvas text-widget)
  (setf *cur-y* (min (1- *rows*) (1+ *cur-y*)))
  (update-text-input text-widget)
  (redraw canvas))

;;;; =========================
;;;; メイン：GUIの構築と起動
;;;; =========================

(defun start (&key (rows 26) (cols 14) (input-lines 3))
  "スプレッドシートを起動
   :rows        行数（デフォルト26）
   :cols        列数（デフォルト14、最大26=A-Z）
   :input-lines 入力欄の行数（デフォルト3）"
  ;; パラメータ設定
  (setf *rows* rows)
  (setf *cols* (min cols 26))  ; 最大26列（A-Z）
  
  ;; 初期化
  (setf *sheet* (make-hash-table :test #'equal))
  (setf *cur-x* 0 *cur-y* 0)
  
  (with-ltk ()
    (wm-title *tk* (format nil "Spreadsheet v0.3 [~Dx~D]" *cols* *rows*))
    
    ;; ウィジェット作成
    (let* ((input-frame (make-instance 'frame))
           ;; 複数行入力用Textウィジェット
           (input-text (make-instance 'text
                                      :master input-frame
                                      :width 80
                                      :height input-lines
                                      :font "Consolas 11"))
           (input-scroll (make-instance 'scrollbar 
                                        :master input-frame
                                        :orientation :vertical))
           (canvas (make-instance 'canvas
                                  :width (+ *header-w* (* *cols* *cell-w*))
                                  :height (+ *header-h* (* *rows* *cell-h*)))))
      
      ;; スクロールバーとテキストを連携
      (configure input-scroll :command (format nil "~a yview" (widget-path input-text)))
      (configure input-text :yscrollcommand (format nil "~a set" (widget-path input-scroll)))
      
      ;; レイアウト
      (pack input-frame :fill :x :padx 2 :pady 2)
      (pack input-scroll :side :right :fill :y)
      (pack input-text :side :left :fill :both :expand t)
      (pack canvas)

      ;; サンプルデータ
      (setf (cell-value (get-cell "A1")) 10)
      (setf (cell-value (get-cell "A2")) 20)
      (setf (cell-value (get-cell "A3")) 30)
      (setf (cell-value (get-cell "B1")) '(1 2 3 4 5))

      ;; 初期描画
      (update-text-input input-text)
      (redraw canvas)

      ;;; --- イベントバインド ---

      ;; まずLTKのbindでコールバックを設定（内部でsendeventが設定される）
      (bind input-text "<Return>"
            (lambda (evt)
              (declare (ignore evt))
              (commit-text-input input-text canvas)
              (focus canvas)))
      
      ;; 次にTclレベルでバインディングを調整
      (let ((path (widget-path input-text)))
        ;; 現在のReturnバインディングを取得して、末尾にbreakを追加
        (format-wish "bind ~a <Return> \"[bind ~a <Return>]; break\"" path path)
        ;; Shift+Returnを追加（改行を挿入してbreak）
        (format-wish "bind ~a <Shift-Return> {~a insert insert \\n; break}" path path))

      ;; セルクリック → カーソル移動
      (bind canvas "<ButtonPress-1>"
            (lambda (evt)
              (let ((mx (and evt (slot-value evt 'ltk::x)))
                    (my (and evt (slot-value evt 'ltk::y))))
                (when (and mx my (numberp mx) (numberp my)
                           (> mx *header-w*) (> my *header-h*))
                  (setf *cur-x* (clamp (floor (/ (- mx *header-w*) *cell-w*)) 0 (1- *cols*))
                        *cur-y* (clamp (floor (/ (- my *header-h*) *cell-h*)) 0 (1- *rows*)))
                  (update-text-input input-text)
                  (redraw canvas)))
              (focus canvas)))

      ;; 矢印キー
      (bind canvas "<Left>"
            (lambda (evt) (declare (ignore evt)) (move-left canvas input-text)))
      (bind canvas "<Right>"
            (lambda (evt) (declare (ignore evt)) (move-right canvas input-text)))
      (bind canvas "<Up>"
            (lambda (evt) (declare (ignore evt)) (move-up canvas input-text)))
      (bind canvas "<Down>"
            (lambda (evt) (declare (ignore evt)) (move-down canvas input-text)))

      ;; Canvas上でEnter → Textにフォーカス
      (bind canvas "<Return>"
            (lambda (evt)
              (declare (ignore evt))
              (focus input-text)))

      (focus canvas))))

;;; ロード時メッセージ
(format t "~%=== Spreadsheet GUI v0.3 ===~%")
(format t "Lisp Powered Edition~%~%")
(format t "起動: (ss-gui:start)~%")
(format t "      (ss-gui:start :rows 30 :cols 10 :input-lines 5)~%")
(format t "~%パラメータ:~%")
(format t "  :rows        行数（デフォルト26）~%")
(format t "  :cols        列数（デフォルト14、最大26）~%")
(format t "  :input-lines 入力欄の行数（デフォルト3）~%")
(format t "~%基本操作:~%")
(format t "  矢印キー/クリック : セル移動~%")
(format t "  Enter            : 入力欄へ移動 / 入力確定~%")
(format t "  Shift+Enter      : 入力欄内で改行~%")
(format t "~%数式例（=で始める）:~%")
(format t "  =(+ A1 B1)~%")
(format t "  =(sum (range A1 A5))~%")
(format t "  =(mapcar (lambda (x) (* x x)) '(1 2 3 4))~%")
(format t "  =(apply #'+ '(1 2 3 4 5))~%")
