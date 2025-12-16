;;;; =====================================================
;;;; spreadsheet-gui.lisp
;;;; Common Lisp + LTK で作るシンプルな表計算ソフト
;;;; 
;;;; Version: 0.1
;;;; Date: 2025-01-15
;;;; 
;;;; 機能:
;;;;   - セルへの値・数式入力
;;;;   - 四則演算 (+, -, *, /)
;;;;   - 集計関数 (sum, avg, max, min, count)
;;;;   - 範囲指定 (range A1 A5)
;;;;   - 矢印キー・クリックでセル移動
;;;;   - 行番号・列名ヘッダー表示
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

(defparameter *rows* 20)          ; 行数
(defparameter *cols* 10)          ; 列数
(defparameter *cell-w* 80)        ; セル幅（ピクセル）
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
  value      ; 表示値（計算結果または入力値）
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
;;;; 数式評価エンジン
;;;; =========================

(defun try-parse-number (val)
  "文字列を数値に変換。失敗時は元の値を返す"
  (if (stringp val)
      (let ((parsed (ignore-errors (read-from-string val))))
        (if (numberp parsed) parsed 0))
      val))

(defun cell-ref-p (sym)
  "シンボルがセル参照か判定。A1〜J99形式を認識"
  (and (symbolp sym)
       (let ((name (symbol-name sym)))
         (and (>= (length name) 2)
              (<= (length name) 3)
              (alpha-char-p (char name 0))
              (every #'digit-char-p (subseq name 1))))))

(defun parse-cell-coords (sym)
  "セル参照シンボルから座標(col row)を取得。A1->(0 0), B3->(1 2)"
  (let* ((name (symbol-name sym))
         (col (- (char-code (char-upcase (char name 0))) (char-code #\A)))
         (row (1- (parse-integer (subseq name 1)))))
    (list col row)))

(defun get-cell-by-symbol (sym)
  "シンボル(A1等)からセル値を取得"
  (let* ((coords (parse-cell-coords sym))
         (col (first coords))
         (row (second coords)))
    (try-parse-number (cell-value (get-cell (cell-name col row))))))

(defun ref (name)
  "文字列でセル参照。(ref \"A1\") の形式"
  (let* ((col (- (char-code (char-upcase (char name 0))) (char-code #\A)))
         (row (1- (parse-integer (subseq name 1)))))
    (try-parse-number (cell-value (get-cell (cell-name col row))))))

(defun expand-range (start-sym end-sym)
  "範囲を展開してセル値のリストを返す
   (: A1 A5) → A1〜A5の値リスト
   (: A1 C1) → A1,B1,C1の値リスト
   (: A1 C3) → 矩形範囲の値リスト"
  (let* ((start-coords (parse-cell-coords start-sym))
         (end-coords (parse-cell-coords end-sym))
         (col1 (first start-coords))
         (row1 (second start-coords))
         (col2 (first end-coords))
         (row2 (second end-coords))
         ;; 開始・終了を正規化（小さい方を先に）
         (min-col (min col1 col2))
         (max-col (max col1 col2))
         (min-row (min row1 row2))
         (max-row (max row1 row2))
         (values '()))
    ;; 範囲内の全セルの値を収集
    (loop for r from min-row to max-row do
          (loop for c from min-col to max-col do
                (let ((val (try-parse-number 
                            (cell-value (get-cell (cell-name c r))))))
                  (push (if (numberp val) val 0) values))))
    (nreverse values)))

(defun eval-formula (expr)
  "S式の数式を再帰的に評価
   数値            -> そのまま
   セル参照        -> セルの値
   (range A1 A5)   -> 範囲内の値リスト
   リスト          -> 演算子を適用"
  (cond
    ((null expr) 0)
    ((numberp expr) expr)
    ((stringp expr) expr)
    ((cell-ref-p expr) (or (get-cell-by-symbol expr) 0))
    ((listp expr)
     (let* ((op (first expr))
            (op-name (if (symbolp op) (symbol-name op) "")))
       (cond
         ;; 範囲指定 (range A1 A5) → 値のリストを返す
         ((string-equal op-name "RANGE")
          (if (and (>= (length expr) 3)
                   (cell-ref-p (second expr))
                   (cell-ref-p (third expr)))
              (expand-range (second expr) (third expr))
              :error))
         ;; その他の演算
         (t
          (let ((args (mapcar #'eval-formula (rest expr))))
            ;; argsの中にリストがあれば展開（範囲指定の結果）
            (let ((flat-args (apply #'append 
                                    (mapcar (lambda (a) 
                                              (if (listp a) a (list a)))
                                            args))))
              (handler-case
                  (cond
                    ;; 基本演算
                    ((string= op-name "+") (apply #'+ flat-args))
                    ((string= op-name "-") (apply #'- flat-args))
                    ((string= op-name "*") (apply #'* flat-args))
                    ((string= op-name "/") (if (member 0 (cdr flat-args)) "DIV/0!" (apply #'/ flat-args)))
                    ;; 集計関数
                    ((string-equal op-name "SUM") (apply #'+ flat-args))
                    ((string-equal op-name "AVG") 
                     (if flat-args (float (/ (apply #'+ flat-args) (length flat-args))) 0))
                    ((string-equal op-name "MAX") (if flat-args (apply #'max flat-args) 0))
                    ((string-equal op-name "MIN") (if flat-args (apply #'min flat-args) 0))
                    ((string-equal op-name "COUNT") (length flat-args))
                    ((string-equal op-name "REF") (ref (first args)))
                    (t :error))
                (error () :error))))))))
    (t expr)))

;;;; =========================
;;;; 描画（Canvas操作）
;;;; =========================

(defun draw-cell (canvas x y text selected)
  "1つのセルを描画。selectedならハイライト"
  (let* ((px (+ *header-w* (* x *cell-w*)))   ; 描画X座標
         (py (+ *header-h* (* y *cell-h*)))   ; 描画Y座標
         (px2 (+ px *cell-w*))
         (py2 (+ py *cell-h*))
         (bg (if selected "#cce5ff" "white")) ; 選択時は青背景
         (display-text (if text (princ-to-string text) ""))
         (path (widget-path canvas)))
    ;; Tclコマンドで矩形とテキストを描画
    (format-wish "~a create rectangle ~a ~a ~a ~a -fill {~a} -outline gray"
                 path px py px2 py2 bg)
    (format-wish "~a create text ~a ~a -anchor w -text {~a} -font {Consolas 11}"
                 path (+ px 4) (+ py (floor *cell-h* 2)) display-text)))

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
  "画面全体を再描画"
  (format-wish "~a delete all" (widget-path canvas))
  (draw-headers canvas)
  (dotimes (y *rows*)
    (dotimes (x *cols*)
      (let ((cell (get-cell (cell-name x y))))
        (draw-cell canvas x y
                   (cell-value cell)
                   (and (= x *cur-x*) (= y *cur-y*)))))))

;;;; =========================
;;;; 入力欄（Entry）の操作
;;;; =========================

(defun entry-text (e)
  "Entryの内容を取得"
  (text e))

(defun set-entry-text (e s)
  "Entryに文字列を設定"
  (setf (text e) (if s (princ-to-string s) "")))

(defun update-entry (entry)
  "現在セルの内容をEntryに表示"
  (let ((c (current-cell)))
    (cond
      ;; 数式があれば =(...) 形式で表示
      ((cell-formula c)
       (set-entry-text entry (format nil "=~s" (cell-formula c))))
      ;; 値があればそのまま表示
      ((cell-value c)
       (set-entry-text entry (princ-to-string (cell-value c))))
      (t
       (set-entry-text entry "")))))

(defun commit-entry (entry canvas)
  "Entryの内容をセルに確定"
  (let* ((raw (or (entry-text entry) ""))
         (cell (current-cell)))
    (handler-case
        (if (and (> (length raw) 0)
                 (char= (char raw 0) #\=))
            ;; =で始まる → 数式として解析・評価
            (let* ((form (read-from-string (subseq raw 1)))
                   (value (eval-formula form)))
              (setf (cell-formula cell) form
                    (cell-value cell) value))
            ;; それ以外 → 通常の値
            (setf (cell-formula cell) nil
                  (cell-value cell) raw))
      (error (e)
        (setf (cell-value cell) (format nil "ERR: ~a" e))))
    (redraw canvas)))

;;;; =========================
;;;; カーソル移動
;;;; =========================

(defun clamp (v lo hi)
  "値を範囲内に制限"
  (max lo (min hi v)))

(defun move-left (canvas entry)
  (setf *cur-x* (max 0 (1- *cur-x*)))
  (update-entry entry)
  (redraw canvas))

(defun move-right (canvas entry)
  (setf *cur-x* (min (1- *cols*) (1+ *cur-x*)))
  (update-entry entry)
  (redraw canvas))

(defun move-up (canvas entry)
  (setf *cur-y* (max 0 (1- *cur-y*)))
  (update-entry entry)
  (redraw canvas))

(defun move-down (canvas entry)
  (setf *cur-y* (min (1- *rows*) (1+ *cur-y*)))
  (update-entry entry)
  (redraw canvas))

;;;; =========================
;;;; メイン：GUIの構築と起動
;;;; =========================

(defun start ()
  "スプレッドシートを起動"
  ;; 初期化
  (setf *sheet* (make-hash-table :test #'equal))
  (setf *cur-x* 0 *cur-y* 0)
  
  (with-ltk ()
    (wm-title *tk* "Spreadsheet v0.1")
    
    ;; ウィジェット作成
    (let* ((canvas (make-instance 'canvas
                                  :width (+ *header-w* (* *cols* *cell-w*))
                                  :height (+ *header-h* (* *rows* *cell-h*))))
           (entry (make-instance 'entry)))
      
      ;; レイアウト：Entry上部、Canvas下部
      (pack entry :fill :x)
      (pack canvas)

      ;; サンプルデータ
      (setf (cell-value (get-cell "A1")) "10")
      (setf (cell-value (get-cell "B1")) "20")

      ;; 初期描画
      (update-entry entry)
      (redraw canvas)

      ;;; --- イベントバインド ---

      ;; Entry上でEnter → 入力確定、Canvasにフォーカス
      (bind entry "<Return>"
            (lambda (evt)
              (declare (ignore evt))
              (commit-entry entry canvas)
              (focus canvas)))

      ;; セルクリック → カーソル移動
      (bind canvas "<ButtonPress-1>"
            (lambda (evt)
              (let ((mx (and evt (slot-value evt 'ltk::x)))
                    (my (and evt (slot-value evt 'ltk::y))))
                ;; ヘッダー部分は無視
                (when (and mx my (numberp mx) (numberp my)
                           (> mx *header-w*) (> my *header-h*))
                  (setf *cur-x* (clamp (floor (/ (- mx *header-w*) *cell-w*)) 0 (1- *cols*))
                        *cur-y* (clamp (floor (/ (- my *header-h*) *cell-h*)) 0 (1- *rows*)))
                  (update-entry entry)
                  (redraw canvas)))
              (focus canvas)))

      ;; 矢印キー → カーソル移動
      (bind canvas "<Left>"
            (lambda (evt) (declare (ignore evt)) (move-left canvas entry)))
      (bind canvas "<Right>"
            (lambda (evt) (declare (ignore evt)) (move-right canvas entry)))
      (bind canvas "<Up>"
            (lambda (evt) (declare (ignore evt)) (move-up canvas entry)))
      (bind canvas "<Down>"
            (lambda (evt) (declare (ignore evt)) (move-down canvas entry)))

      ;; Canvas上でEnter → Entryにフォーカス（入力開始）
      (bind canvas "<Return>"
            (lambda (evt)
              (declare (ignore evt))
              (focus entry)))

      ;; 起動時はCanvasにフォーカス
      (focus canvas))))

;;; ロード時メッセージ
(format t "~%=== Spreadsheet GUI v0.1 ===~%")
(format t "起動: (ss-gui:start)~%")
(format t "~%操作方法:~%")
(format t "  矢印キー   : セル移動~%")
(format t "  クリック   : セル選択~%")
(format t "  Enter      : 入力開始/確定~%")
(format t "~%数式例:~%")
(format t "  =(+ A1 B1)         : 加算~%")
(format t "  =(* A1 2)          : 乗算~%")
(format t "  =(sum A1 B1 C1)    : 合計~%")
(format t "  =(avg A1 B1 C1)    : 平均~%")
(format t "~%範囲指定:~%")
(format t "  =(sum (range A1 A5))   : A1〜A5の合計~%")
(format t "  =(avg (range A1 C3))   : 矩形範囲の平均~%")
(format t "  =(count (range A1 A10)): セル数カウント~%")
