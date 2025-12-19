;;;; =====================================================
;;;; spreadsheet-gui.lisp
;;;; Common Lisp + LTK で作るシンプルな表計算ソフト
;;;; 
;;;; Version: 0.5.2
;;;; Date: 2025-01-15
;;;; 
;;;; 機能:
;;;;   - セルへの値・数式入力
;;;;   - Lispの非破壊関数をサポート
;;;;   - lambda式、apply、funcall
;;;;   - 戻り値はリスト、シンボル等任意のLisp値
;;;;   - 範囲指定 (range A1 A5)
;;;;   - 相対参照 (rel -1 0), (rel-range -4 0 -1 0)
;;;;   - 自動再計算（依存関係追跡）
;;;;   - 範囲選択、コピー＆ペースト
;;;;   - システムクリップボード対応（TSV形式）
;;;;   - ファイル保存/読み込み (.ssp形式)
;;;;   - CSV エクスポート/インポート
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
  (:export :start 
           :show-dependencies :show-cell-deps
           ;; ファイル操作
           :save :load-file :new-sheet
           :export-csv :import-csv
           ;; Undo/Redo
           :undo :redo :clear-history
           ;; 現在のファイル
           :*current-file*))
(in-package :ss-gui)

;;;; =========================
;;;; 設定（グリッドサイズ・セル寸法）
;;;; =========================

(defparameter *rows* 26)          ; 行数
(defparameter *cols* 14)          ; 列数 (A-N)
(defparameter *default-cell-w* 100) ; デフォルトセル幅
(defparameter *default-cell-h* 24)  ; デフォルトセル高さ
(defparameter *min-cell-w* 30)    ; 最小セル幅
(defparameter *min-cell-h* 16)    ; 最小セル高さ
(defparameter *header-h* 24)      ; 列ヘッダー(A,B,C...)の高さ
(defparameter *header-w* 40)      ; 行ヘッダー(1,2,3...)の幅
(defparameter *cur-x* 0)          ; カーソル位置（列）
(defparameter *cur-y* 0)          ; カーソル位置（行）

;; 各列の幅、各行の高さ（個別設定用）
(defparameter *col-widths* nil)   ; 列幅の配列
(defparameter *row-heights* nil)  ; 行高さの配列

;; リサイズ用
(defparameter *resize-mode* nil)  ; :col または :row
(defparameter *resize-index* nil) ; リサイズ中の列/行インデックス
(defparameter *resize-start* nil) ; ドラッグ開始位置

;; 範囲選択用
(defparameter *sel-start-x* nil)  ; 選択開始列
(defparameter *sel-start-y* nil)  ; 選択開始行
(defparameter *sel-end-x* nil)    ; 選択終了列
(defparameter *sel-end-y* nil)    ; 選択終了行
(defparameter *selecting* nil)    ; ドラッグ中フラグ

;; クリップボード（2次元リスト: ((val formula) ...)）
(defparameter *clipboard* nil)
(defparameter *clipboard-rows* 0)
(defparameter *clipboard-cols* 0)

;; 数式評価時の現在セル位置（動的束縛用）
(defvar *eval-row* 0)    ; 評価中の行（0始まり）
(defvar *eval-col* 0)    ; 評価中の列（0始まり）
(defvar *eval-stack* nil) ; 循環参照検出用スタック

;; 依存関係グラフ
(defparameter *refs* (make-hash-table :test 'equal))       ; セル→参照先リスト
(defparameter *dependents* (make-hash-table :test 'equal)) ; セル→依存元リスト

;; Undo/Redo スタック
(defparameter *undo-stack* nil)       ; Undo用スタック
(defparameter *redo-stack* nil)       ; Redo用スタック
(defparameter *max-undo-history* 100) ; 最大履歴数

;; 現在のファイル
(defparameter *current-file* nil)  ; 現在開いているファイルのパス

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
;;;; 列幅・行高さ管理
;;;; =========================

(defun init-sizes ()
  "列幅・行高さを初期化"
  (setf *col-widths* (make-array *cols* :initial-element *default-cell-w*))
  (setf *row-heights* (make-array *rows* :initial-element *default-cell-h*)))

(defun col-width (x)
  "列xの幅を取得"
  (if (and *col-widths* (< x (length *col-widths*)))
      (aref *col-widths* x)
      *default-cell-w*))

(defun row-height (y)
  "行yの高さを取得"
  (if (and *row-heights* (< y (length *row-heights*)))
      (aref *row-heights* y)
      *default-cell-h*))

(defun set-col-width (x w)
  "列xの幅を設定"
  (when (and *col-widths* (< x (length *col-widths*)))
    (setf (aref *col-widths* x) (max *min-cell-w* w))))

(defun set-row-height (y h)
  "行yの高さを設定"
  (when (and *row-heights* (< y (length *row-heights*)))
    (setf (aref *row-heights* y) (max *min-cell-h* h))))

(defun col-left (x)
  "列xの左端X座標を取得"
  (let ((pos *header-w*))
    (dotimes (i x pos)
      (incf pos (col-width i)))))

(defun row-top (y)
  "行yの上端Y座標を取得"
  (let ((pos *header-h*))
    (dotimes (i y pos)
      (incf pos (row-height i)))))

(defun total-width ()
  "全列の合計幅"
  (let ((w *header-w*))
    (dotimes (x *cols* w)
      (incf w (col-width x)))))

(defun total-height ()
  "全行の合計高さ"
  (let ((h *header-h*))
    (dotimes (y *rows* h)
      (incf h (row-height y)))))

(defun find-col-at (px)
  "X座標pxから列インデックスを取得（ヘッダー内なら-1）"
  (if (< px *header-w*)
      -1
      (let ((x 0)
            (pos *header-w*))
        (loop while (and (< x *cols*) (>= px pos))
              do (incf pos (col-width x))
                 (incf x))
        (1- x))))

(defun find-row-at (py)
  "Y座標pyから行インデックスを取得（ヘッダー内なら-1）"
  (if (< py *header-h*)
      -1
      (let ((y 0)
            (pos *header-h*))
        (loop while (and (< y *rows*) (>= py pos))
              do (incf pos (row-height y))
                 (incf y))
        (1- y))))

(defun near-col-border-p (px tolerance)
  "X座標pxが列境界の近くにあるか判定。境界の列インデックスを返す（なければnil）"
  (let ((pos *header-w*))
    (dotimes (x *cols*)
      (incf pos (col-width x))
      (when (<= (- pos tolerance) px (+ pos tolerance))
        (return-from near-col-border-p x))))
  nil)

(defun near-row-border-p (py tolerance)
  "Y座標pyが行境界の近くにあるか判定。境界の行インデックスを返す（なければnil）"
  (let ((pos *header-h*))
    (dotimes (y *rows*)
      (incf pos (row-height y))
      (when (<= (- pos tolerance) py (+ pos tolerance))
        (return-from near-row-border-p y))))
  nil)

;;;; =========================
;;;; 範囲選択
;;;; =========================

(defun clear-selection ()
  "選択をクリア"
  (setf *sel-start-x* nil
        *sel-start-y* nil
        *sel-end-x* nil
        *sel-end-y* nil
        *selecting* nil))

(defun has-selection-p ()
  "範囲選択があるか"
  (and *sel-start-x* *sel-start-y* *sel-end-x* *sel-end-y*))

(defun selection-bounds ()
  "選択範囲の境界を返す (min-x min-y max-x max-y)"
  (when (has-selection-p)
    (values (min *sel-start-x* *sel-end-x*)
            (min *sel-start-y* *sel-end-y*)
            (max *sel-start-x* *sel-end-x*)
            (max *sel-start-y* *sel-end-y*))))

(defun cell-in-selection-p (x y)
  "セル(x,y)が選択範囲内にあるか"
  (when (has-selection-p)
    (multiple-value-bind (min-x min-y max-x max-y) (selection-bounds)
      (and (<= min-x x max-x)
           (<= min-y y max-y)))))

;;;; =========================
;;;; 行・列の挿入・削除
;;;; =========================

(defun shift-cell-name-row (name delta)
  "セル名の行番号をdeltaだけシフト。範囲外ならnilを返す"
  (let* ((col-char (char name 0))
         (row-num (parse-integer (subseq name 1)))
         (new-row (+ row-num delta)))
    (if (and (>= new-row 1) (<= new-row *rows*))
        (format nil "~a~d" col-char new-row)
        nil)))

(defun shift-cell-name-col (name delta)
  "セル名の列をdeltaだけシフト。範囲外ならnilを返す"
  (let* ((col-char (char name 0))
         (col-idx (- (char-code (char-upcase col-char)) (char-code #\A)))
         (new-col (+ col-idx delta))
         (row-str (subseq name 1)))
    (if (and (>= new-col 0) (< new-col *cols*))
        (format nil "~a~a" (code-char (+ (char-code #\A) new-col)) row-str)
        nil)))

;;; データ存在チェック
(defun row-has-data-p (row-idx)
  "指定行にデータ（値または数式）があるかチェック"
  (loop for x from 0 below *cols*
        for cell = (gethash (cell-name x row-idx) *sheet*)
        thereis (and cell (or (cell-value cell) (cell-formula cell)))))

(defun col-has-data-p (col-idx)
  "指定列にデータ（値または数式）があるかチェック"
  (loop for y from 0 below *rows*
        for cell = (gethash (cell-name col-idx y) *sheet*)
        thereis (and cell (or (cell-value cell) (cell-formula cell)))))

;;; 確認ダイアログ
(defun confirm-dialog (title message)
  "確認ダイアログを表示。OKならt、キャンセルならnilを返す"
  ;; tk_messageBoxの結果を直接取得
  (format-wish "senddatastring [tk_messageBox -type okcancel -icon warning -title {~a} -message {~a}]"
               title message)
  (string= (ltk::read-data) "ok"))

;;; 数式参照の更新（全セル対象）
;;; セル位置（cell-row, cell-col）は移動後の位置
;;; 移動前の位置を逆算して補正を行う

(defun rel-symbol-p (sym)
  "シンボルがREL（パッケージ不問）かチェック"
  (and (symbolp sym)
       (string-equal (symbol-name sym) "REL")))

(defun rel-range-symbol-p (sym)
  "シンボルがREL-RANGE（パッケージ不問）かチェック"
  (and (symbolp sym)
       (string-equal (symbol-name sym) "REL-RANGE")))

(defun update-formula-ref-for-row-insert (formula at-row cell-row cell-col)
  "行挿入時の数式参照更新。
   - 絶対参照: at-row以降の行参照を+1
   - 相対参照(rel): 移動前の参照先がat-rowより前なら補正
   - 相対範囲(rel-range): 同様に補正
   注: cell-rowは移動後の位置"
  (cond
    ((null formula) nil)
    ((numberp formula) formula)
    ((stringp formula) formula)
    ((keywordp formula) formula)
    ((symbolp formula)
     (let ((name (symbol-name formula)))
       (if (and (>= (length name) 2)
                (alpha-char-p (char name 0))
                (every #'digit-char-p (subseq name 1)))
           (let* ((row-num (parse-integer (subseq name 1)))
                  (row-idx (1- row-num)))
             (if (>= row-idx at-row)
                 ;; 挿入位置以降なら+1
                 (let ((new-name (shift-cell-name-row name 1)))
                   (if new-name (intern new-name :ss-gui) formula))
                 formula))
           formula)))
    ((listp formula)
     (cond
       ;; REL の特別処理
       ((and (rel-symbol-p (car formula))
             (= (length formula) 3)
             (numberp (second formula))
             (numberp (third formula)))
        (let ((row-offset (second formula))
              (col-offset (third formula)))
          (cond
            ;; セルが移動した（現在at-rowより下にある）
            ((> cell-row at-row)
             (let* ((orig-cell-row (1- cell-row))  ; 移動前の位置
                    (orig-target-row (+ orig-cell-row row-offset)))
               (if (< orig-target-row at-row)
                   ;; 参照先は移動しなかった、オフセットを調整
                   (list 'rel (1- row-offset) col-offset)
                   ;; 参照先も移動した、オフセットはそのまま
                   (list 'rel row-offset col-offset))))
            ;; セルが移動していない
            ((< cell-row at-row)
             (let ((target-row (+ cell-row row-offset)))
               (if (>= target-row at-row)
                   ;; 参照先が移動した、オフセットを調整
                   (list 'rel (1+ row-offset) col-offset)
                   ;; 参照先も移動しなかった、オフセットはそのまま
                   (list 'rel row-offset col-offset))))
            ;; cell-row == at-row は挿入された空セル
            (t formula))))
       ;; REL-RANGE の特別処理
       ((and (rel-range-symbol-p (car formula))
             (= (length formula) 5)
             (every #'numberp (cdr formula)))
        (let ((start-row-off (second formula))
              (start-col-off (third formula))
              (end-row-off (fourth formula))
              (end-col-off (fifth formula)))
          (cond
            ((> cell-row at-row)
             (let* ((orig-cell-row (1- cell-row))
                    (orig-start-row (+ orig-cell-row start-row-off))
                    (orig-end-row (+ orig-cell-row end-row-off)))
               (list 'rel-range
                     (if (< orig-start-row at-row) (1- start-row-off) start-row-off)
                     start-col-off
                     (if (< orig-end-row at-row) (1- end-row-off) end-row-off)
                     end-col-off)))
            ((< cell-row at-row)
             (let ((start-target (+ cell-row start-row-off))
                   (end-target (+ cell-row end-row-off)))
               (list 'rel-range
                     (if (>= start-target at-row) (1+ start-row-off) start-row-off)
                     start-col-off
                     (if (>= end-target at-row) (1+ end-row-off) end-row-off)
                     end-col-off)))
            (t formula))))
       ;; 通常のリスト処理
       (t (mapcar (lambda (x) (update-formula-ref-for-row-insert x at-row cell-row cell-col)) formula))))
    (t formula)))

(defun update-formula-ref-for-row-delete (formula at-row cell-row cell-col)
  "行削除時の数式参照更新。
   - 絶対参照: at-rowへの参照は#REF!、at-rowより後は-1
   - 相対参照(rel): 参照先が削除行なら#REF!、移動前の参照先がat-rowより前なら補正
   - 相対範囲(rel-range): 同様に補正
   注: cell-rowは移動後の位置"
  (cond
    ((null formula) nil)
    ((numberp formula) formula)
    ((stringp formula) formula)
    ((keywordp formula) formula)
    ((symbolp formula)
     (let ((name (symbol-name formula)))
       (if (and (>= (length name) 2)
                (alpha-char-p (char name 0))
                (every #'digit-char-p (subseq name 1)))
           (let* ((row-num (parse-integer (subseq name 1)))
                  (row-idx (1- row-num)))
             (cond
               ((= row-idx at-row)
                (intern "#REF!" :ss-gui))
               ((> row-idx at-row)
                (let ((new-name (shift-cell-name-row name -1)))
                  (if new-name (intern new-name :ss-gui) formula)))
               (t formula)))
           formula)))
    ((listp formula)
     (cond
       ;; REL の特別処理
       ((and (rel-symbol-p (car formula))
             (= (length formula) 3)
             (numberp (second formula))
             (numberp (third formula)))
        (let ((row-offset (second formula))
              (col-offset (third formula)))
          (cond
            ;; セルが移動した（削除後、at-row以降にある = 移動前はat-row+1以降）
            ((>= cell-row at-row)
             (let* ((orig-cell-row (1+ cell-row))  ; 移動前の位置
                    (orig-target-row (+ orig-cell-row row-offset)))
               (cond
                 ((= orig-target-row at-row)
                  '(quote |#REF!|))
                 ((< orig-target-row at-row)
                  ;; 参照先は移動しなかった、オフセットを調整
                  (list 'rel (1+ row-offset) col-offset))
                 ;; 参照先も移動した、オフセットはそのまま
                 (t (list 'rel row-offset col-offset)))))
            ;; セルが移動していない
            (t
             (let ((target-row (+ cell-row row-offset)))
               (cond
                 ((= target-row at-row)
                  '(quote |#REF!|))
                 ((> target-row at-row)
                  ;; 参照先が移動した、オフセットを調整
                  (list 'rel (1- row-offset) col-offset))
                 (t (list 'rel row-offset col-offset))))))))
       ;; REL-RANGE の特別処理
       ((and (rel-range-symbol-p (car formula))
             (= (length formula) 5)
             (every #'numberp (cdr formula)))
        (let ((start-row-off (second formula))
              (start-col-off (third formula))
              (end-row-off (fourth formula))
              (end-col-off (fifth formula)))
          (cond
            ((>= cell-row at-row)
             (let* ((orig-cell-row (1+ cell-row))
                    (orig-start-row (+ orig-cell-row start-row-off))
                    (orig-end-row (+ orig-cell-row end-row-off)))
               (if (or (= orig-start-row at-row) (= orig-end-row at-row))
                   '(quote |#REF!|)
                   (list 'rel-range
                         (if (< orig-start-row at-row) (1+ start-row-off) start-row-off)
                         start-col-off
                         (if (< orig-end-row at-row) (1+ end-row-off) end-row-off)
                         end-col-off))))
            (t
             (let ((start-target (+ cell-row start-row-off))
                   (end-target (+ cell-row end-row-off)))
               (if (or (= start-target at-row) (= end-target at-row))
                   '(quote |#REF!|)
                   (list 'rel-range
                         (if (> start-target at-row) (1- start-row-off) start-row-off)
                         start-col-off
                         (if (> end-target at-row) (1- end-row-off) end-row-off)
                         end-col-off)))))))
       ;; 通常のリスト処理
       (t (mapcar (lambda (x) (update-formula-ref-for-row-delete x at-row cell-row cell-col)) formula))))
    (t formula)))

(defun update-formula-ref-for-col-insert (formula at-col cell-row cell-col)
  "列挿入時の数式参照更新。
   - 絶対参照: at-col以降の列参照を+1
   - 相対参照(rel): 移動前の参照先がat-colより前なら補正
   - 相対範囲(rel-range): 同様に補正
   注: cell-colは移動後の位置"
  (cond
    ((null formula) nil)
    ((numberp formula) formula)
    ((stringp formula) formula)
    ((keywordp formula) formula)
    ((symbolp formula)
     (let ((name (symbol-name formula)))
       (if (and (>= (length name) 2)
                (alpha-char-p (char name 0))
                (every #'digit-char-p (subseq name 1)))
           (let* ((col-idx (- (char-code (char-upcase (char name 0))) (char-code #\A))))
             (if (>= col-idx at-col)
                 ;; 挿入位置以降なら+1
                 (let ((new-name (shift-cell-name-col name 1)))
                   (if new-name (intern new-name :ss-gui) formula))
                 formula))
           formula)))
    ((listp formula)
     (cond
       ;; REL の特別処理
       ((and (rel-symbol-p (car formula))
             (= (length formula) 3)
             (numberp (second formula))
             (numberp (third formula)))
        (let ((row-offset (second formula))
              (col-offset (third formula)))
          (cond
            ;; セルが移動した（現在at-colより右にある）
            ((> cell-col at-col)
             (let* ((orig-cell-col (1- cell-col))  ; 移動前の位置
                    (orig-target-col (+ orig-cell-col col-offset)))
               (if (< orig-target-col at-col)
                   ;; 参照先は移動しなかった、オフセットを調整
                   (list 'rel row-offset (1- col-offset))
                   ;; 参照先も移動した、オフセットはそのまま
                   (list 'rel row-offset col-offset))))
            ;; セルが移動していない
            ((< cell-col at-col)
             (let ((target-col (+ cell-col col-offset)))
               (if (>= target-col at-col)
                   ;; 参照先が移動した、オフセットを調整
                   (list 'rel row-offset (1+ col-offset))
                   ;; 参照先も移動しなかった、オフセットはそのまま
                   (list 'rel row-offset col-offset))))
            ;; cell-col == at-col は挿入された空セル
            (t formula))))
       ;; REL-RANGE の特別処理
       ((and (rel-range-symbol-p (car formula))
             (= (length formula) 5)
             (every #'numberp (cdr formula)))
        (let ((start-row-off (second formula))
              (start-col-off (third formula))
              (end-row-off (fourth formula))
              (end-col-off (fifth formula)))
          (cond
            ((> cell-col at-col)
             (let* ((orig-cell-col (1- cell-col))
                    (orig-start-col (+ orig-cell-col start-col-off))
                    (orig-end-col (+ orig-cell-col end-col-off)))
               (list 'rel-range
                     start-row-off
                     (if (< orig-start-col at-col) (1- start-col-off) start-col-off)
                     end-row-off
                     (if (< orig-end-col at-col) (1- end-col-off) end-col-off))))
            ((< cell-col at-col)
             (let ((start-target (+ cell-col start-col-off))
                   (end-target (+ cell-col end-col-off)))
               (list 'rel-range
                     start-row-off
                     (if (>= start-target at-col) (1+ start-col-off) start-col-off)
                     end-row-off
                     (if (>= end-target at-col) (1+ end-col-off) end-col-off))))
            (t formula))))
       ;; 通常のリスト処理
       (t (mapcar (lambda (x) (update-formula-ref-for-col-insert x at-col cell-row cell-col)) formula))))
    (t formula)))

(defun update-formula-ref-for-col-delete (formula at-col cell-row cell-col)
  "列削除時の数式参照更新。
   - 絶対参照: at-colへの参照は#REF!、at-colより後は-1
   - 相対参照(rel): 参照先が削除列なら#REF!、移動前の参照先がat-colより前なら補正
   - 相対範囲(rel-range): 同様に補正
   注: cell-colは移動後の位置"
  (cond
    ((null formula) nil)
    ((numberp formula) formula)
    ((stringp formula) formula)
    ((keywordp formula) formula)
    ((symbolp formula)
     (let ((name (symbol-name formula)))
       (if (and (>= (length name) 2)
                (alpha-char-p (char name 0))
                (every #'digit-char-p (subseq name 1)))
           (let* ((col-idx (- (char-code (char-upcase (char name 0))) (char-code #\A))))
             (cond
               ((= col-idx at-col)
                (intern "#REF!" :ss-gui))
               ((> col-idx at-col)
                (let ((new-name (shift-cell-name-col name -1)))
                  (if new-name (intern new-name :ss-gui) formula)))
               (t formula)))
           formula)))
    ((listp formula)
     (cond
       ;; REL の特別処理
       ((and (rel-symbol-p (car formula))
             (= (length formula) 3)
             (numberp (second formula))
             (numberp (third formula)))
        (let ((row-offset (second formula))
              (col-offset (third formula)))
          (cond
            ;; セルが移動した（削除後、at-col以降にある = 移動前はat-col+1以降）
            ((>= cell-col at-col)
             (let* ((orig-cell-col (1+ cell-col))  ; 移動前の位置
                    (orig-target-col (+ orig-cell-col col-offset)))
               (cond
                 ((= orig-target-col at-col)
                  '(quote |#REF!|))
                 ((< orig-target-col at-col)
                  ;; 参照先は移動しなかった、オフセットを調整
                  (list 'rel row-offset (1+ col-offset)))
                 ;; 参照先も移動した、オフセットはそのまま
                 (t (list 'rel row-offset col-offset)))))
            ;; セルが移動していない
            (t
             (let ((target-col (+ cell-col col-offset)))
               (cond
                 ((= target-col at-col)
                  '(quote |#REF!|))
                 ((> target-col at-col)
                  ;; 参照先が移動した、オフセットを調整
                  (list 'rel row-offset (1- col-offset)))
                 (t (list 'rel row-offset col-offset))))))))
       ;; REL-RANGE の特別処理
       ((and (rel-range-symbol-p (car formula))
             (= (length formula) 5)
             (every #'numberp (cdr formula)))
        (let ((start-row-off (second formula))
              (start-col-off (third formula))
              (end-row-off (fourth formula))
              (end-col-off (fifth formula)))
          (cond
            ((>= cell-col at-col)
             (let* ((orig-cell-col (1+ cell-col))
                    (orig-start-col (+ orig-cell-col start-col-off))
                    (orig-end-col (+ orig-cell-col end-col-off)))
               (if (or (= orig-start-col at-col) (= orig-end-col at-col))
                   '(quote |#REF!|)
                   (list 'rel-range
                         start-row-off
                         (if (< orig-start-col at-col) (1+ start-col-off) start-col-off)
                         end-row-off
                         (if (< orig-end-col at-col) (1+ end-col-off) end-col-off)))))
            (t
             (let ((start-target (+ cell-col start-col-off))
                   (end-target (+ cell-col end-col-off)))
               (if (or (= start-target at-col) (= end-target at-col))
                   '(quote |#REF!|)
                   (list 'rel-range
                         start-row-off
                         (if (> start-target at-col) (1- start-col-off) start-col-off)
                         end-row-off
                         (if (> end-target at-col) (1- end-col-off) end-col-off))))))))
       ;; 通常のリスト処理
       (t (mapcar (lambda (x) (update-formula-ref-for-col-delete x at-col cell-row cell-col)) formula))))
    (t formula)))

;;; 全セルの数式参照を更新
(defun update-all-formulas-with-position (update-fn)
  "全セルの数式を更新関数で更新。update-fnは (formula row col) を受け取る"
  (maphash (lambda (name cell)
             (when (cell-formula cell)
               (let* ((coords (parse-cell-name name))
                      (col (first coords))
                      (row (second coords)))
                 (setf (cell-formula cell)
                       (funcall update-fn (cell-formula cell) row col)))))
           *sheet*))

(defun insert-row (at-row &optional force)
  "指定行に空行を挿入（at-row以降を下にシフト、最終行は破棄）
   force=tの場合は確認なしで実行"
  (when (and (>= at-row 0) (< at-row *rows*))
    ;; 最終行にデータがあるかチェック
    (when (and (not force) (row-has-data-p (1- *rows*)))
      (unless (confirm-dialog "Insert Row" 
                              (format nil "Row ~a contains data and will be deleted. Continue?" *rows*))
        (return-from insert-row nil)))
    ;; 最終行のセルを削除
    (loop for x from 0 below *cols* do
      (remhash (cell-name x (1- *rows*)) *sheet*))
    ;; 下から上に向かってセルを移動（最終行-1から挿入行まで）
    (loop for y from (- *rows* 2) downto at-row do
      (loop for x from 0 below *cols* do
        (let* ((src-name (cell-name x y))
               (dst-name (cell-name x (1+ y)))
               (src-cell (gethash src-name *sheet*)))
          (when src-cell
            ;; セルを移動（数式はまだ更新しない）
            (setf (gethash dst-name *sheet*) src-cell)
            (remhash src-name *sheet*)))))
    ;; 全セルの数式参照を更新（挿入位置以降の参照を+1）
    (update-all-formulas-with-position 
     (lambda (f row col) (update-formula-ref-for-row-insert f at-row row col)))
    ;; 行高さ配列を更新（シフト）
    (when *row-heights*
      (loop for i from (- *rows* 2) downto at-row do
        (setf (aref *row-heights* (1+ i)) (aref *row-heights* i)))
      (setf (aref *row-heights* at-row) *default-cell-h*))
    ;; 全セルを再評価
    (recalculate-all)
    ;; 依存関係を再構築
    (rebuild-all-dependencies)
    t))

(defun delete-row (at-row)
  "指定行を削除（at-row以降を上にシフト、最終行は空になる）"
  (when (and (>= at-row 0) (< at-row *rows*))
    ;; 削除行のセルをクリア
    (loop for x from 0 below *cols* do
      (remhash (cell-name x at-row) *sheet*))
    ;; 上にシフト
    (loop for y from (1+ at-row) below *rows* do
      (loop for x from 0 below *cols* do
        (let* ((src-name (cell-name x y))
               (dst-name (cell-name x (1- y)))
               (src-cell (gethash src-name *sheet*)))
          (when src-cell
            ;; セルを移動（数式はまだ更新しない）
            (setf (gethash dst-name *sheet*) src-cell)
            (remhash src-name *sheet*)))))
    ;; 全セルの数式参照を更新（削除行への参照は#REF!、それより後は-1）
    (update-all-formulas-with-position 
     (lambda (f row col) (update-formula-ref-for-row-delete f at-row row col)))
    ;; 行高さ配列を更新（シフト）
    (when *row-heights*
      (loop for i from at-row below (1- *rows*) do
        (setf (aref *row-heights* i) (aref *row-heights* (1+ i))))
      (setf (aref *row-heights* (1- *rows*)) *default-cell-h*))
    ;; 全セルを再評価
    (recalculate-all)
    ;; 依存関係を再構築
    (rebuild-all-dependencies)
    t))

(defun insert-col (at-col &optional force)
  "指定列に空列を挿入（at-col以降を右にシフト、最終列は破棄）
   force=tの場合は確認なしで実行"
  (when (and (>= at-col 0) (< at-col *cols*))
    ;; 最終列にデータがあるかチェック
    (when (and (not force) (col-has-data-p (1- *cols*)))
      (let ((col-name (string (code-char (+ (char-code #\A) (1- *cols*))))))
        (unless (confirm-dialog "Insert Column"
                                (format nil "Column ~a contains data and will be deleted. Continue?" col-name))
          (return-from insert-col nil))))
    ;; 最終列のセルを削除
    (loop for y from 0 below *rows* do
      (remhash (cell-name (1- *cols*) y) *sheet*))
    ;; 右から左に向かってセルを移動（最終列-1から挿入列まで）
    (loop for x from (- *cols* 2) downto at-col do
      (loop for y from 0 below *rows* do
        (let* ((src-name (cell-name x y))
               (dst-name (cell-name (1+ x) y))
               (src-cell (gethash src-name *sheet*)))
          (when src-cell
            ;; セルを移動（数式はまだ更新しない）
            (setf (gethash dst-name *sheet*) src-cell)
            (remhash src-name *sheet*)))))
    ;; 全セルの数式参照を更新（挿入位置以降の参照を+1）
    (update-all-formulas-with-position 
     (lambda (f row col) (update-formula-ref-for-col-insert f at-col row col)))
    ;; 列幅配列を更新（シフト）
    (when *col-widths*
      (loop for i from (- *cols* 2) downto at-col do
        (setf (aref *col-widths* (1+ i)) (aref *col-widths* i)))
      (setf (aref *col-widths* at-col) *default-cell-w*))
    ;; 全セルを再評価
    (recalculate-all)
    ;; 依存関係を再構築
    (rebuild-all-dependencies)
    t))

(defun delete-col (at-col)
  "指定列を削除（at-col以降を左にシフト、最終列は空になる）"
  (when (and (>= at-col 0) (< at-col *cols*))
    ;; 削除列のセルをクリア
    (loop for y from 0 below *rows* do
      (remhash (cell-name at-col y) *sheet*))
    ;; 左にシフト
    (loop for x from (1+ at-col) below *cols* do
      (loop for y from 0 below *rows* do
        (let* ((src-name (cell-name x y))
               (dst-name (cell-name (1- x) y))
               (src-cell (gethash src-name *sheet*)))
          (when src-cell
            ;; セルを移動（数式はまだ更新しない）
            (setf (gethash dst-name *sheet*) src-cell)
            (remhash src-name *sheet*)))))
    ;; 全セルの数式参照を更新（削除列への参照は#REF!、それより後は-1）
    (update-all-formulas-with-position 
     (lambda (f row col) (update-formula-ref-for-col-delete f at-col row col)))
    ;; 列幅配列を更新（シフト）
    (when *col-widths*
      (loop for i from at-col below (1- *cols*) do
        (setf (aref *col-widths* i) (aref *col-widths* (1+ i))))
      (setf (aref *col-widths* (1- *cols*)) *default-cell-w*))
    ;; 全セルを再評価
    (recalculate-all)
    ;; 依存関係を再構築
    (rebuild-all-dependencies)
    t))

(defun rebuild-all-dependencies ()
  "全セルの依存関係を再構築"
  (clear-dependencies)
  (maphash (lambda (name cell)
             (when (cell-formula cell)
               (let* ((coords (parse-cell-name name))
                      (col (first coords))
                      (row (second coords))
                      (refs (extract-references (cell-formula cell) row col)))
                 (update-dependencies name refs))))
           *sheet*))

(defun recalculate-all ()
  "全セルの数式を再評価"
  (maphash (lambda (name cell)
             (when (cell-formula cell)
               (let* ((coords (parse-cell-name name))
                      (col (first coords))
                      (row (second coords)))
                 ;; 動的変数を設定して評価
                 (let ((*eval-col* col)
                       (*eval-row* row)
                       (*eval-stack* (list name)))
                   (setf (cell-value cell)
                         (handler-case
                             (eval-formula (cell-formula cell))
                           (error (e)
                             (format nil "#評価:~a" (type-of e)))))))))
           *sheet*))

;;;; =========================
;;;; コピー＆ペースト
;;;; ===========================

(defun copy-selection ()
  "選択範囲をクリップボードにコピー"
  (when (has-selection-p)
    (multiple-value-bind (min-x min-y max-x max-y) (selection-bounds)
      (setf *clipboard-cols* (1+ (- max-x min-x))
            *clipboard-rows* (1+ (- max-y min-y))
            *clipboard* nil)
      ;; セルデータを収集
      (loop for y from min-y to max-y do
        (loop for x from min-x to max-x do
          (let ((cell (get-cell (cell-name x y))))
            (push (list (cell-value cell) (cell-formula cell)) *clipboard*))))
      (setf *clipboard* (nreverse *clipboard*)))))

(defun paste-clipboard ()
  "クリップボードの内容をカーソル位置にペースト"
  (when *clipboard*
    (let ((idx 0)
          (pasted-cells nil)
          (before-snapshots nil)
          (after-snapshots nil))
      ;; まず全てのセルに値と数式を設定
      (loop for dy from 0 below *clipboard-rows* do
        (loop for dx from 0 below *clipboard-cols* do
          (let* ((x (+ *cur-x* dx))
                 (y (+ *cur-y* dy)))
            (when (and (< x *cols*) (< y *rows*))
              (let* ((name (cell-name x y))
                     (cell (get-cell name))
                     (data (nth idx *clipboard*))
                     (formula (second data)))
                ;; 変更前の状態を保存
                (push (make-cell-snapshot name) before-snapshots)
                ;; 数式がある場合は再評価
                (if formula
                    (progn
                      (setf *eval-col* x *eval-row* y)
                      (let ((*eval-stack* (list name)))
                        (handler-case
                            (setf (cell-value cell) (eval-formula formula))
                          (error (e)
                            (setf (cell-value cell) (format nil "ERR: ~a" e)))))
                      (setf (cell-formula cell) formula)
                      (update-dependencies name (extract-references formula y x)))
                    (progn
                      (setf (cell-value cell) (first data)
                            (cell-formula cell) nil)
                      (update-dependencies name nil)))
                ;; 変更後の状態を保存
                (push (make-cell-snapshot name) after-snapshots)
                (push name pasted-cells))))
          (incf idx)))
      ;; Undo履歴に記録
      (when pasted-cells
        (record-multi-change (nreverse before-snapshots) (nreverse after-snapshots)))
      ;; 貼り付けたセルの依存元を再計算
      (dolist (name (nreverse pasted-cells))
        (recalc-dependents name)))))

(defun clear-selection-cells ()
  "選択範囲のセルをクリア（NILを設定）し、依存元を再計算"
  (let ((cleared-cells nil)
        (before-snapshots nil)
        (after-snapshots nil))
    (if (has-selection-p)
        ;; 範囲選択がある場合
        (multiple-value-bind (min-x min-y max-x max-y) (selection-bounds)
          (loop for y from min-y to max-y do
            (loop for x from min-x to max-x do
              (let* ((name (cell-name x y))
                     (cell (get-cell name)))
                ;; 変更前の状態を保存
                (push (make-cell-snapshot name) before-snapshots)
                (setf (cell-value cell) nil
                      (cell-formula cell) nil)
                (update-dependencies name nil)
                ;; 変更後の状態を保存
                (push (make-cell-snapshot name) after-snapshots)
                (push name cleared-cells)))))
        ;; 範囲選択がない場合はカーソル位置のみ
        (let* ((name (cell-name *cur-x* *cur-y*))
               (cell (get-cell name)))
          (push (make-cell-snapshot name) before-snapshots)
          (setf (cell-value cell) nil
                (cell-formula cell) nil)
          (update-dependencies name nil)
          (push (make-cell-snapshot name) after-snapshots)
          (push name cleared-cells)))
    ;; Undo履歴に記録
    (when cleared-cells
      (record-multi-change (nreverse before-snapshots) (nreverse after-snapshots)))
    ;; 全ての削除されたセルの依存元を再計算
    (dolist (name cleared-cells)
      (recalc-dependents name))))

;;;; =========================
;;;; システムクリップボード
;;;; =========================

(defun get-system-clipboard ()
  "システムクリップボードからテキストを取得"
  (handler-case
      (ltk::clipboard-get)
    (error () nil)))

(defun set-system-clipboard (text)
  "システムクリップボードにテキストを設定"
  (format-wish "clipboard clear")
  (format-wish "clipboard append {~a}" text))

(defun split-string (string separator)
  "文字列を区切り文字で分割"
  (loop for start = 0 then (1+ pos)
        for pos = (position separator string :start start)
        collect (subseq string start (or pos (length string)))
        while pos))

(defun format-cell-for-clipboard (val formula)
  "セル値をクリップボード用文字列に変換"
  (let ((*package* (find-package :ss-gui)))  ; パッケージプレフィックスなしで表示
    (cond
      ;; 数式があればそれを優先
      (formula (format nil "=~S" formula))
      ;; NILは空文字列
      ((null val) "")
      ;; 文字列はそのまま
      ((stringp val) val)
      ;; その他はprinc形式
      (t (princ-to-string val)))))

(defun selection-to-tsv ()
  "選択範囲をTSV文字列に変換"
  (if (has-selection-p)
      (multiple-value-bind (min-x min-y max-x max-y) (selection-bounds)
        (with-output-to-string (s)
          (loop for y from min-y to max-y do
            (loop for x from min-x to max-x do
              (let ((cell (get-cell (cell-name x y))))
                (when (> x min-x) (write-char #\Tab s))
                (write-string (format-cell-for-clipboard 
                               (cell-value cell) 
                               (cell-formula cell)) s)))
            (when (< y max-y) (terpri s)))))
      ;; 選択範囲がない場合はカーソル位置のセル
      (let ((cell (current-cell)))
        (format-cell-for-clipboard (cell-value cell) (cell-formula cell)))))

(defun copy-to-system-clipboard ()
  "選択範囲をシステムクリップボードにコピー"
  (let ((text (selection-to-tsv)))
    (set-system-clipboard text)
    ;; 内部クリップボードにもコピー
    (copy-selection)))

(defun parse-clipboard-value (text)
  "クリップボードのテキストをセル値に変換"
  (let ((trimmed (string-trim '(#\Space #\Tab #\Return) text)))
    (cond
      ;; 空文字列 → NIL
      ((string= trimmed "") nil)
      ;; =で始まる → 数式として処理
      ((and (> (length trimmed) 0) (char= (char trimmed 0) #\=))
       (handler-case
           (let* ((*package* (find-package :ss-gui))  ; SS-GUIパッケージで読み込み
                  (form (read-from-string (subseq trimmed 1)))
                  (value (eval-formula form)))
             (values value form))
         (error () (values trimmed nil))))
      ;; 数値を試す
      (t (let* ((*package* (find-package :ss-gui))
                (num (ignore-errors (read-from-string trimmed))))
           (if (numberp num)
               num
               trimmed))))))

(defun paste-from-system-clipboard ()
  "システムクリップボードから貼り付け"
  (let ((text (get-system-clipboard))
        (pasted-cells nil)
        (before-snapshots nil)
        (after-snapshots nil))
    (when (and text (> (length text) 0))
      ;; 行で分割
      (let* ((lines (split-string text #\Newline))
             ;; 空行を末尾から除去
             (lines (loop for l in lines
                         for i from 0
                         while (or (< i (1- (length lines)))
                                  (> (length l) 0))
                         collect l)))
        (if (and (= (length lines) 1)
                 (not (find #\Tab (first lines))))
            ;; 単一値の場合
            (let* ((name (cell-name *cur-x* *cur-y*))
                   (cell (get-cell name)))
              ;; 変更前の状態を保存
              (push (make-cell-snapshot name) before-snapshots)
              ;; 評価位置を設定
              (setf *eval-col* *cur-x* *eval-row* *cur-y*)
              (multiple-value-bind (val form) 
                  (parse-clipboard-value (first lines))
                (setf (cell-value cell) val
                      (cell-formula cell) form)
                (if form
                    (update-dependencies name (extract-references form *cur-y* *cur-x*))
                    (update-dependencies name nil))
                ;; 変更後の状態を保存
                (push (make-cell-snapshot name) after-snapshots)
                (push name pasted-cells)))
            ;; 複数セル（TSV形式）の場合
            (loop for line in lines
                  for dy from 0 do
              (loop for col-text in (split-string line #\Tab)
                    for dx from 0 do
                (let ((x (+ *cur-x* dx))
                      (y (+ *cur-y* dy)))
                  (when (and (< x *cols*) (< y *rows*))
                    (let* ((name (cell-name x y))
                           (cell (get-cell name)))
                      ;; 変更前の状態を保存
                      (push (make-cell-snapshot name) before-snapshots)
                      ;; 評価位置を設定
                      (setf *eval-col* x *eval-row* y)
                      (multiple-value-bind (val form)
                          (parse-clipboard-value col-text)
                        (setf (cell-value cell) val
                              (cell-formula cell) form)
                        (if form
                            (update-dependencies name (extract-references form y x))
                            (update-dependencies name nil))
                        ;; 変更後の状態を保存
                        (push (make-cell-snapshot name) after-snapshots)
                        (push name pasted-cells))))))))))
    ;; Undo履歴に記録
    (when pasted-cells
      (record-multi-change (nreverse before-snapshots) (nreverse after-snapshots)))
    ;; 貼り付けたセルの依存元を再計算
    (dolist (name (nreverse pasted-cells))
      (recalc-dependents name))))

;;;; =========================
;;;; 位置参照関数
;;;; =========================

(defun this-row ()
  "現在のセルの行番号を返す（1始まり）"
  (1+ *eval-row*))

(defun this-col ()
  "現在のセルの列番号を返す（0始まり）"
  *eval-col*)

(defun this-col-name ()
  "現在のセルの列名を返す"
  (string (code-char (+ (char-code #\A) *eval-col*))))

(defun this-cell-name ()
  "現在のセル名を返す"
  (cell-name *eval-col* *eval-row*))

(defun cell-at (row col)
  "行列番号でセルの値を取得（row:1始まり, col:0始まりまたは文字列）"
  (let* ((actual-col (if (stringp col)
                         (- (char-code (char (string-upcase col) 0)) (char-code #\A))
                         col))
         (actual-row (1- row))  ; 1始まり→0始まり
         (name (cell-name actual-col actual-row)))
    (when (and (>= actual-col 0) (< actual-col *cols*)
               (>= actual-row 0) (< actual-row *rows*))
      ;; 循環参照チェック
      (if (member name *eval-stack* :test #'equal)
          :CIRCULAR-REF
          (let ((cell (get-cell name)))
            (cell-value cell))))))

(defun rel (drow dcol)
  "現在のセルからの相対位置のセル値を取得"
  (let* ((new-col (+ *eval-col* dcol))
         (new-row (+ *eval-row* drow))
         (name (cell-name new-col new-row)))
    (if (and (>= new-col 0) (< new-col *cols*)
             (>= new-row 0) (< new-row *rows*))
        ;; 循環参照チェック
        (if (member name *eval-stack* :test #'equal)
            :CIRCULAR-REF
            (let ((cell (get-cell name)))
              (cell-value cell)))
        ;; 範囲外はNIL
        nil)))

(defun rel-range (dr1 dc1 dr2 dc2)
  "相対位置で範囲を指定して値のリストを取得"
  (let* ((r1 (+ *eval-row* dr1))
         (c1 (+ *eval-col* dc1))
         (r2 (+ *eval-row* dr2))
         (c2 (+ *eval-col* dc2))
         ;; 正規化
         (min-r (min r1 r2))
         (max-r (max r1 r2))
         (min-c (min c1 c2))
         (max-c (max c1 c2))
         (result nil))
    (loop for r from min-r to max-r do
      (loop for c from min-c to max-c do
        (when (and (>= c 0) (< c *cols*)
                   (>= r 0) (< r *rows*))
          (let ((name (cell-name c r)))
            (unless (member name *eval-stack* :test #'equal)
              (let ((val (cell-value (get-cell name))))
                (when val (push val result))))))))
    (nreverse result)))

;;;; =========================
;;;; 依存関係管理と再計算
;;;; =========================

(defun extract-references (formula row col)
  "数式から参照しているセル名のリストを抽出"
  (let ((refs nil))
    (labels ((sym-eq (sym name)
               "シンボル名を文字列比較"
               (and (symbolp sym)
                    (string-equal (symbol-name sym) name)))
             (walk (expr)
               (cond
                 ;; セル参照シンボル（A1形式）
                 ((and (symbolp expr)
                       (not (keywordp expr))
                       (let ((name (symbol-name expr)))
                         (and (>= (length name) 2)
                              (<= (length name) 3)
                              (alpha-char-p (char name 0))
                              (every #'digit-char-p (subseq name 1)))))
                  (pushnew (string-upcase (symbol-name expr)) refs :test #'string-equal))
                 ;; (rel drow dcol) - 相対参照
                 ((and (listp expr)
                       (sym-eq (car expr) "REL")
                       (= (length expr) 3)
                       (numberp (second expr))
                       (numberp (third expr)))
                  (let* ((drow (second expr))
                         (dcol (third expr))
                         (new-row (+ row drow))
                         (new-col (+ col dcol)))
                    (when (and (>= new-row 0) (< new-row *rows*)
                               (>= new-col 0) (< new-col *cols*))
                      (pushnew (cell-name new-col new-row) refs :test #'string-equal))))
                 ;; (rel-range dr1 dc1 dr2 dc2) - 相対範囲
                 ((and (listp expr)
                       (sym-eq (car expr) "REL-RANGE")
                       (= (length expr) 5))
                  (let* ((dr1 (second expr))
                         (dc1 (third expr))
                         (dr2 (fourth expr))
                         (dc2 (fifth expr)))
                    (when (and (numberp dr1) (numberp dc1)
                               (numberp dr2) (numberp dc2))
                      (let ((r1 (+ row dr1))
                            (c1 (+ col dc1))
                            (r2 (+ row dr2))
                            (c2 (+ col dc2)))
                        (loop for r from (min r1 r2) to (max r1 r2) do
                          (loop for c from (min c1 c2) to (max c1 c2) do
                            (when (and (>= r 0) (< r *rows*)
                                       (>= c 0) (< c *cols*))
                              (pushnew (cell-name c r) refs :test #'string-equal))))))))
                 ;; (range start end) - 絶対範囲
                 ((and (listp expr)
                       (sym-eq (car expr) "RANGE")
                       (= (length expr) 3))
                  (let ((start (second expr))
                        (end (third expr)))
                    (when (and (symbolp start) (symbolp end))
                      (let* ((start-name (symbol-name start))
                             (end-name (symbol-name end))
                             (c1 (- (char-code (char-upcase (char start-name 0))) (char-code #\A)))
                             (r1 (1- (parse-integer (subseq start-name 1))))
                             (c2 (- (char-code (char-upcase (char end-name 0))) (char-code #\A)))
                             (r2 (1- (parse-integer (subseq end-name 1)))))
                        (loop for r from (min r1 r2) to (max r1 r2) do
                          (loop for c from (min c1 c2) to (max c1 c2) do
                            (pushnew (cell-name c r) refs :test #'string-equal)))))))
                 ;; (cell-at row col) - 行列指定
                 ((and (listp expr)
                       (sym-eq (car expr) "CELL-AT")
                       (>= (length expr) 3))
                  (let ((r (second expr))
                        (c (third expr)))
                    (when (and (numberp r) (or (numberp c) (stringp c)))
                      (let ((actual-col (if (stringp c)
                                           (- (char-code (char-upcase (char c 0))) (char-code #\A))
                                           c))
                            (actual-row (1- r)))
                        (when (and (>= actual-row 0) (< actual-row *rows*)
                                   (>= actual-col 0) (< actual-col *cols*))
                          (pushnew (cell-name actual-col actual-row) refs :test #'string-equal))))))
                 ;; リストの場合は再帰
                 ((listp expr)
                  (dolist (e expr)
                    (walk e))))))
      (walk formula))
    refs))

(defun update-dependencies (cell-name new-refs)
  "セルの依存関係を更新"
  (let ((old-refs (gethash cell-name *refs*)))
    ;; 古い参照先から自分を削除
    (dolist (ref old-refs)
      (let ((deps (gethash ref *dependents*)))
        (setf (gethash ref *dependents*)
              (remove cell-name deps :test #'string-equal))))
    ;; 新しい参照先を設定
    (setf (gethash cell-name *refs*) new-refs)
    ;; 新しい参照先に自分を追加
    (dolist (ref new-refs)
      (pushnew cell-name (gethash ref *dependents*) :test #'string-equal))))

(defun collect-all-dependents (cell-name)
  "セルに依存する全てのセルを収集（再帰的に）"
  (let ((visited (make-hash-table :test 'equal))
        (result nil))
    (labels ((collect (name)
               (unless (gethash name visited)
                 (setf (gethash name visited) t)
                 (dolist (dep (gethash name *dependents*))
                   (push dep result)
                   (collect dep)))))
      (collect cell-name))
    (nreverse result)))

(defun topological-sort-cells (cells)
  "セルをトポロジカルソート（依存順）"
  (let ((in-degree (make-hash-table :test 'equal))
        (graph (make-hash-table :test 'equal))
        (result nil)
        (queue nil))
    ;; 初期化
    (dolist (c cells)
      (setf (gethash c in-degree) 0)
      (setf (gethash c graph) nil))
    ;; グラフ構築
    (dolist (c cells)
      (dolist (ref (gethash c *refs*))
        (when (member ref cells :test #'string-equal)
          (push c (gethash ref graph))
          (incf (gethash c in-degree)))))
    ;; 入次数0のセルをキューに
    (dolist (c cells)
      (when (zerop (gethash c in-degree))
        (push c queue)))
    ;; BFS
    (loop while queue do
      (let ((current (pop queue)))
        (push current result)
        (dolist (dep (gethash current graph))
          (decf (gethash dep in-degree))
          (when (zerop (gethash dep in-degree))
            (push dep queue)))))
    ;; 結果（依存順）
    (nreverse result)))

(defun recalc-cell (cell-name)
  "単一セルを再計算"
  (let* ((cell (get-cell cell-name))
         (formula (cell-formula cell)))
    (when formula
      (let* ((coords (parse-cell-name cell-name))
             (col (first coords))
             (row (second coords)))
        ;; 評価位置を設定
        (setf *eval-col* col *eval-row* row)
        (let ((*eval-stack* (list cell-name)))
          (handler-case
              (setf (cell-value cell) (eval-formula formula))
            (error (e)
              (let ((msg (princ-to-string e)))
                (if (search "循環参照" msg)
                    (setf (cell-value cell) (format-error-message :circular cell-name))
                    (setf (cell-value cell) (format-error-message :eval msg)))))))))))

(defun parse-cell-name (name)
  "セル名から座標(col row)を取得"
  (let* ((col (- (char-code (char-upcase (char name 0))) (char-code #\A)))
         (row (1- (parse-integer (subseq name 1)))))
    (list col row)))

(defun recalc-dependents (cell-name)
  "セルに依存する全てのセルを再計算"
  (let* ((deps (collect-all-dependents cell-name))
         (sorted (topological-sort-cells deps)))
    (dolist (dep sorted)
      (recalc-cell dep))))

(defun clear-dependencies ()
  "依存関係をクリア"
  (clrhash *refs*)
  (clrhash *dependents*))

(defun show-dependencies ()
  "依存関係をデバッグ表示"
  (format t "~%=== Refs (セル→参照先) ===~%")
  (maphash (lambda (k v) (format t "  ~a → ~a~%" k v)) *refs*)
  (format t "~%=== Dependents (セル→依存元) ===~%")
  (maphash (lambda (k v) (format t "  ~a ← ~a~%" k v)) *dependents*)
  (values))

(defun show-cell-deps (cell-name)
  "特定セルの依存関係を表示"
  (format t "~%セル ~a:~%" cell-name)
  (format t "  参照先: ~a~%" (gethash cell-name *refs*))
  (format t "  依存元: ~a~%" (gethash cell-name *dependents*))
  (values))

;;;; =========================
;;;; Undo/Redo 機能
;;;; =========================

;;; 操作タイプ:
;;;   :cell-change  - 単一セルの変更
;;;   :multi-change - 複数セルの変更（ペースト、削除等）

(defun make-cell-snapshot (cell-name)
  "セルの現在状態をスナップショットとして取得"
  (let ((cell (get-cell cell-name)))
    (list :name cell-name
          :value (cell-value cell)
          :formula (cell-formula cell))))

(defun restore-cell-snapshot (snapshot)
  "スナップショットからセルを復元"
  (let* ((name (getf snapshot :name))
         (value (getf snapshot :value))
         (formula (getf snapshot :formula))
         (cell (get-cell name)))
    (setf (cell-value cell) value
          (cell-formula cell) formula)
    ;; 依存関係を再構築
    (let ((coords (parse-cell-name name)))
      (if formula
          (update-dependencies name 
                               (extract-references formula 
                                                   (second coords) 
                                                   (first coords)))
          ;; 数式がなくなった場合は依存関係をクリア
          (update-dependencies name nil)))))

(defun record-cell-change (cell-name old-value old-formula)
  "単一セルの変更を記録"
  (let ((action (list :type :cell-change
                      :before (list :name cell-name
                                    :value old-value
                                    :formula old-formula)
                      :after (make-cell-snapshot cell-name))))
    (push action *undo-stack*)
    ;; Redo履歴をクリア
    (setf *redo-stack* nil)
    ;; 履歴数制限
    (when (> (length *undo-stack*) *max-undo-history*)
      (setf *undo-stack* (butlast *undo-stack*)))))

(defun record-multi-change (before-snapshots after-snapshots)
  "複数セルの変更を記録"
  (let ((action (list :type :multi-change
                      :before before-snapshots
                      :after after-snapshots)))
    (push action *undo-stack*)
    (setf *redo-stack* nil)
    (when (> (length *undo-stack*) *max-undo-history*)
      (setf *undo-stack* (butlast *undo-stack*)))))

(defun undo ()
  "直前の操作を取り消す"
  (if (null *undo-stack*)
      (progn
        (format t "Undo: 履歴がありません~%")
        nil)
      (let ((action (pop *undo-stack*)))
        (push action *redo-stack*)
        (case (getf action :type)
          (:cell-change
           (restore-cell-snapshot (getf action :before))
           (let ((name (getf (getf action :before) :name)))
             (recalc-dependents name)))
          (:multi-change
           (dolist (snapshot (getf action :before))
             (restore-cell-snapshot snapshot))
           ;; 全ての変更セルの依存先を再計算
           (dolist (snapshot (getf action :before))
             (recalc-dependents (getf snapshot :name)))))
        t)))

(defun redo ()
  "取り消した操作をやり直す"
  (if (null *redo-stack*)
      (progn
        (format t "Redo: 履歴がありません~%")
        nil)
      (let ((action (pop *redo-stack*)))
        (push action *undo-stack*)
        (case (getf action :type)
          (:cell-change
           (restore-cell-snapshot (getf action :after))
           (let ((name (getf (getf action :after) :name)))
             (recalc-dependents name)))
          (:multi-change
           (dolist (snapshot (getf action :after))
             (restore-cell-snapshot snapshot))
           (dolist (snapshot (getf action :after))
             (recalc-dependents (getf snapshot :name)))))
        t)))

(defun clear-history ()
  "Undo/Redo履歴をクリア"
  (setf *undo-stack* nil
        *redo-stack* nil)
  (format t "履歴をクリアしました~%"))

;;;; =========================
;;;; ファイル保存/読み込み
;;;; =========================

(defun iso-timestamp ()
  "ISO 8601形式のタイムスタンプを生成"
  (multiple-value-bind (sec min hour day month year)
      (get-decoded-time)
    (format nil "~4,'0d-~2,'0d-~2,'0dT~2,'0d:~2,'0d:~2,'0d"
            year month day hour min sec)))

(defun collect-cells-data ()
  "保存用にセルデータを収集"
  (let ((cells-data nil))
    (maphash (lambda (name cell)
               (when (or (cell-value cell) (cell-formula cell))
                 (push (if (cell-formula cell)
                           (list name (cell-value cell) (cell-formula cell))
                           (list name (cell-value cell)))
                       cells-data)))
             *sheet*)
    (sort cells-data #'string< :key #'first)))

(defun save (filename)
  "スプレッドシートを.sspファイルに保存"
  (let ((cells-data (collect-cells-data)))
    (with-open-file (out filename 
                         :direction :output 
                         :if-exists :supersede
                         :external-format :utf-8)
      (let ((*print-pretty* t)
            (*print-right-margin* 80)
            (*print-case* :downcase)
            (*package* (find-package :ss-gui)))  ; パッケージプレフィックスなしで保存
        ;; ヘッダーコメント
        (format out ";;;; -*- mode: lisp; coding: utf-8 -*-~%")
        (format out ";;;; Spreadsheet GUI File (.ssp)~%")
        (format out ";;;; Created: ~a~%~%" (iso-timestamp))
        ;; データ
        (prin1 
         `(:spreadsheet
           :format-version 1
           :metadata (:created ,(iso-timestamp)
                      :modified ,(iso-timestamp)
                      :app-version "0.5.2")
           :grid (:rows ,*rows* :cols ,*cols*)
           :cells ,cells-data)
         out)
        (terpri out)))
    (setf *current-file* filename)
    (format t "~%Saved: ~a (~d cells)~%" filename (length cells-data))
    filename))

(defun load-file (filename)
  "ファイルからスプレッドシートを読み込み"
  (with-open-file (in filename 
                      :direction :input
                      :external-format :utf-8)
    (let* ((*package* (find-package :ss-gui))  ; SS-GUIパッケージで読み込み
           (data (read in)))
      ;; 形式チェック
      (unless (and (listp data) (eq (car data) :spreadsheet))
        (error "無効なファイル形式: ~a" filename))
      (let ((version (getf (cdr data) :format-version))
            (grid (getf (cdr data) :grid))
            (cells (getf (cdr data) :cells)))
        ;; バージョンチェック
        (when (and version (> version 1))
          (warn "ファイルバージョン ~a は完全にはサポートされていません" version))
        ;; グリッド設定
        (when grid
          (setf *rows* (or (getf grid :rows) *rows*)
                *cols* (or (getf grid :cols) *cols*)))
        ;; シートをクリア
        (setf *sheet* (make-hash-table :test #'equal))
        (clear-dependencies)
        ;; Undo/Redo履歴をクリア
        (setf *undo-stack* nil
              *redo-stack* nil)
        ;; セルデータを復元
        (dolist (cell-data cells)
          (let* ((name (first cell-data))
                 (value (second cell-data))
                 (formula (third cell-data))
                 (cell (get-cell name)))
            (setf (cell-value cell) value
                  (cell-formula cell) formula)
            ;; 依存関係を再構築
            (when formula
              (let ((coords (parse-cell-name name)))
                (update-dependencies 
                 name 
                 (extract-references formula 
                                    (second coords) 
                                    (first coords)))))))
        (setf *current-file* filename)
        (format t "~%Loaded: ~a (~d cells)~%" filename (length cells))
        filename))))

(defun new-sheet ()
  "新規シートを作成（現在のデータをクリア）"
  (setf *sheet* (make-hash-table :test #'equal))
  (setf *cur-x* 0 *cur-y* 0)
  (clear-selection)
  (setf *clipboard* nil)
  (clear-dependencies)
  (setf *current-file* nil)
  ;; Undo/Redo履歴もクリア
  (setf *undo-stack* nil
        *redo-stack* nil)
  (format t "~%New sheet created~%")
  t)

;;;; =========================
;;;; CSV エクスポート/インポート
;;;; =========================

(defun get-used-range ()
  "使用されているセル範囲を取得 (min-col min-row max-col max-row)"
  (let ((min-col (1- *cols*)) (min-row (1- *rows*))
        (max-col 0) (max-row 0)
        (has-data nil))
    (maphash (lambda (name cell)
               (when (or (cell-value cell) (cell-formula cell))
                 (setf has-data t)
                 (let* ((coords (parse-cell-name name))
                        (col (first coords))
                        (row (second coords)))
                   (setf min-col (min min-col col)
                         min-row (min min-row row)
                         max-col (max max-col col)
                         max-row (max max-row row)))))
             *sheet*)
    (if has-data
        (values min-col min-row max-col max-row)
        (values 0 0 0 0))))

(defun escape-csv-field (str)
  "CSV用にフィールドをエスケープ"
  (let ((s (if (stringp str) str (princ-to-string str))))
    (if (or (find #\, s) 
            (find #\" s) 
            (find #\Newline s)
            (find #\Return s))
        ;; クォートが必要
        (format nil "\"~a\"" 
                (with-output-to-string (out)
                  (loop for c across s do
                    (when (char= c #\") (write-char #\" out))
                    (write-char c out))))
        s)))

(defun cell-value-for-csv (cell)
  "セル値をCSV出力用文字列に変換"
  (let ((val (cell-value cell)))
    (cond
      ((null val) "")
      ((stringp val) val)
      ((numberp val) (princ-to-string val))
      ((listp val) (princ-to-string val))
      ((symbolp val) (symbol-name val))
      (t (princ-to-string val)))))

(defun export-csv (filename &key (include-header nil) (include-formulas nil))
  "CSVファイルにエクスポート
   :include-header   列名ヘッダーを含める（デフォルトnil）
   :include-formulas 数式を含める（デフォルトnil、値のみ）"
  (multiple-value-bind (min-col min-row max-col max-row) (get-used-range)
    (with-open-file (out filename 
                         :direction :output 
                         :if-exists :supersede
                         :external-format :utf-8)
      ;; ヘッダー行（列名）
      (when include-header
        (loop for c from min-col to max-col
              for first = t then nil do
          (unless first (write-char #\, out))
          (write-string (string (code-char (+ (char-code #\A) c))) out))
        (terpri out))
      ;; データ行
      (let ((*package* (find-package :ss-gui)))  ; パッケージプレフィックスなしで出力
        (loop for r from min-row to max-row do
          (loop for c from min-col to max-col
                for first = t then nil do
            (unless first (write-char #\, out))
            (let* ((cell (get-cell (cell-name c r)))
                   (text (if (and include-formulas (cell-formula cell))
                             (format nil "=~S" (cell-formula cell))
                             (cell-value-for-csv cell))))
              (write-string (escape-csv-field text) out)))
          (terpri out))))
    (format t "~%Exported CSV: ~a~%" filename)
    filename))

(defun parse-csv-line (line)
  "CSV行をフィールドのリストにパース"
  (let ((fields nil)
        (current (make-array 0 :element-type 'character :adjustable t :fill-pointer 0))
        (in-quotes nil)
        (i 0)
        (len (length line)))
    (loop while (< i len) do
      (let ((c (char line i)))
        (cond
          ;; クォート開始/終了
          ((char= c #\")
           (if in-quotes
               ;; クォート内で次もクォートならエスケープ
               (if (and (< (1+ i) len) (char= (char line (1+ i)) #\"))
                   (progn
                     (vector-push-extend #\" current)
                     (incf i))
                   (setf in-quotes nil))
               (setf in-quotes t)))
          ;; カンマ（クォート外）
          ((and (char= c #\,) (not in-quotes))
           (push (copy-seq current) fields)
           (setf (fill-pointer current) 0))
          ;; その他の文字
          (t (vector-push-extend c current))))
      (incf i))
    ;; 最後のフィールド
    (push (copy-seq current) fields)
    (nreverse fields)))

(defun parse-csv-value (text)
  "CSVフィールドをセル値に変換"
  (let ((trimmed (string-trim '(#\Space #\Tab) text)))
    (cond
      ;; 空文字列
      ((string= trimmed "") nil)
      ;; =で始まる → 数式として処理
      ((and (> (length trimmed) 0) (char= (char trimmed 0) #\=))
       (handler-case
           (let* ((*package* (find-package :ss-gui))
                  (form (read-from-string (subseq trimmed 1))))
             (values (eval-formula form) form))
         (error () (values trimmed nil))))
      ;; 数値を試す
      (t (let* ((*package* (find-package :ss-gui))
                (num (ignore-errors (read-from-string trimmed))))
           (if (numberp num)
               num
               trimmed))))))

(defun import-csv (filename &key (has-header nil) (start-col 0) (start-row 0))
  "CSVファイルからインポート
   :has-header 最初の行をヘッダーとしてスキップ（デフォルトnil）
   :start-col  開始列（デフォルト0=A列）
   :start-row  開始行（デフォルト0=1行目）"
  (with-open-file (in filename 
                      :direction :input
                      :external-format :utf-8)
    (let ((row-idx start-row)
          (cell-count 0)
          (first-line t))
      ;; シートをクリア
      (setf *sheet* (make-hash-table :test #'equal))
      (clear-dependencies)
      ;; 各行を処理
      (loop for line = (read-line in nil nil)
            while line do
        ;; ヘッダー行をスキップ
        (if (and first-line has-header)
            (setf first-line nil)
            (progn
              (setf first-line nil)
              (let ((fields (parse-csv-line line))
                    (col-idx start-col))
                (dolist (field fields)
                  (when (< col-idx *cols*)
                    (multiple-value-bind (val formula) (parse-csv-value field)
                      (when val
                        (let* ((name (cell-name col-idx row-idx))
                               (cell (get-cell name)))
                          (setf (cell-value cell) val
                                (cell-formula cell) formula)
                          (when formula
                            (update-dependencies 
                             name 
                             (extract-references formula row-idx col-idx)))
                          (incf cell-count)))))
                  (incf col-idx)))
              (incf row-idx))))
      (setf *current-file* nil)  ; CSVなのでsspではない
      (format t "~%Imported CSV: ~a (~d cells)~%" filename cell-count)
      filename)))

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
    ;; 乱数
    random
    ;; ビット演算
    logand logior logxor lognot ash logcount
    ;; 比較
    = /= < > <= >=
    equal equalp eq eql
    ;; リスト操作
    car cdr cons list
    first second third fourth fifth sixth seventh eighth ninth tenth
    rest last butlast nthcdr
    append reverse length nth elt
    member assoc assoc-if rassoc rassoc-if
    find find-if position position-if
    getf get  ; plistアクセス
    mapcar mapc maplist mapcon mapcan
    remove remove-if remove-if-not remove-duplicates
    reduce count count-if count-if-not
    substitute substitute-if substitute-if-not
    subseq copy-list copy-seq copy-tree copy-alist
    list-length
    tree-equal sublis
    ;; ソート（非破壊的に実装）
    sort stable-sort
    ;; 条件チェック
    every some notevery notany
    ;; 集合演算
    intersection union set-difference set-exclusive-or subsetp
    ;; 検索
    search mismatch
    ;; リスト作成
    make-list iota pairlis acons
    ;; 文字列
    string-upcase string-downcase string-capitalize
    string-trim string-left-trim string-right-trim
    concatenate subseq char
    string= string/= string< string> string<= string>=
    string-equal string-not-equal
    string-lessp string-greaterp string-not-greaterp string-not-lessp
    parse-integer
    ;; 文字
    char-upcase char-downcase char-code code-char digit-char
    char= char/= char< char> char<= char>=
    char-equal char-not-equal char-lessp char-greaterp
    alpha-char-p digit-char-p upper-case-p lower-case-p
    alphanumericp graphic-char-p
    ;; 論理
    not null and or
    ;; 述語
    atom listp consp numberp integerp floatp rationalp realp complexp
    stringp symbolp characterp keywordp functionp
    zerop plusp minusp evenp oddp
    ;; 型
    type-of typep
    ;; 型変換
    float truncate round floor ceiling
    string coerce
    ;; ユーティリティ
    identity constantly values
    ;; 数学定数（変数として）
    pi
    ;; スプレッドシート専用
    sum avg cell-count range ref
    ;; 位置参照
    this-row this-col this-col-name this-cell-name
    cell-at rel rel-range
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
    ((string-equal op-name "CELL-COUNT")
     ;; スプレッドシート専用: 範囲内のセル数を数える
     (lambda (&rest args)
       (length (flatten args))))
    ;; COUNT は CL の count を使用（get-function-by-name で処理）
    ;; 位置参照関数
    ((string-equal op-name "THIS-ROW") #'this-row)
    ((string-equal op-name "THIS-COL") #'this-col)
    ((string-equal op-name "THIS-COL-NAME") #'this-col-name)
    ((string-equal op-name "THIS-CELL-NAME") #'this-cell-name)
    ((string-equal op-name "CELL-AT") #'cell-at)
    ((string-equal op-name "REL") #'rel)
    ((string-equal op-name "REL-RANGE") #'rel-range)
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
         ;; FIND-IF / POSITION-IF
         ((or (string-equal op-name "FIND-IF")
              (string-equal op-name "POSITION-IF"))
          (let* ((fn-expr (cadr expr))
                 (fn (cond
                       ((and (listp fn-expr) (eq (car fn-expr) 'function))
                        (get-lisp-function (symbol-name (cadr fn-expr))))
                       ((and (listp fn-expr)
                             (string-equal (symbol-name (car fn-expr)) "LAMBDA"))
                        (eval-formula fn-expr))
                       (t (eval-formula fn-expr))))
                 (seq (eval-formula (caddr expr))))
            (if (functionp fn)
                (funcall (if (string-equal op-name "FIND-IF")
                             #'find-if
                             #'position-if)
                         fn seq)
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
         ;; REMOVE-IF / REMOVE-IF-NOT / COUNT-IF / COUNT-IF-NOT
         ((or (string-equal op-name "REMOVE-IF")
              (string-equal op-name "REMOVE-IF-NOT")
              (string-equal op-name "COUNT-IF")
              (string-equal op-name "COUNT-IF-NOT"))
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
                          ((string-equal op-name "COUNT-IF") #'count-if)
                          (t #'count-if-not))
                         fn lst)
                (format nil "ERR: not a function"))))
         ;; SUBSTITUTE / SUBSTITUTE-IF / SUBSTITUTE-IF-NOT
         ((or (string-equal op-name "SUBSTITUTE")
              (string-equal op-name "SUBSTITUTE-IF")
              (string-equal op-name "SUBSTITUTE-IF-NOT"))
          (if (string-equal op-name "SUBSTITUTE")
              ;; (substitute new old seq)
              (let ((new (eval-formula (cadr expr)))
                    (old (eval-formula (caddr expr)))
                    (seq (eval-formula (cadddr expr))))
                (substitute new old seq))
              ;; (substitute-if new pred seq) / (substitute-if-not new pred seq)
              (let* ((new (eval-formula (cadr expr)))
                     (fn-expr (caddr expr))
                     (fn (cond
                           ((and (listp fn-expr) (eq (car fn-expr) 'function))
                            (get-lisp-function (symbol-name (cadr fn-expr))))
                           ((and (listp fn-expr)
                                 (string-equal (symbol-name (car fn-expr)) "LAMBDA"))
                            (eval-formula fn-expr))
                           (t (eval-formula fn-expr))))
                     (seq (eval-formula (cadddr expr))))
                (if (functionp fn)
                    (funcall (if (string-equal op-name "SUBSTITUTE-IF")
                                 #'substitute-if
                                 #'substitute-if-not)
                             new fn seq)
                    (format nil "ERR: not a function")))))
         ;; EVERY / SOME / NOTEVERY / NOTANY
         ((or (string-equal op-name "EVERY")
              (string-equal op-name "SOME")
              (string-equal op-name "NOTEVERY")
              (string-equal op-name "NOTANY"))
          (let* ((fn-expr (cadr expr))
                 (fn (cond
                       ((and (listp fn-expr) (eq (car fn-expr) 'function))
                        (get-lisp-function (symbol-name (cadr fn-expr))))
                       ((and (listp fn-expr)
                             (string-equal (symbol-name (car fn-expr)) "LAMBDA"))
                        (eval-formula fn-expr))
                       (t (eval-formula fn-expr))))
                 (seqs (mapcar #'eval-formula (cddr expr))))
            (if (functionp fn)
                (apply (cond
                         ((string-equal op-name "EVERY") #'every)
                         ((string-equal op-name "SOME") #'some)
                         ((string-equal op-name "NOTEVERY") #'notevery)
                         (t #'notany))
                       fn seqs)
                (format nil "ERR: not a function"))))
         ;; ASSOC-IF / RASSOC-IF
         ((or (string-equal op-name "ASSOC-IF")
              (string-equal op-name "RASSOC-IF"))
          (let* ((fn-expr (cadr expr))
                 (fn (cond
                       ((and (listp fn-expr) (eq (car fn-expr) 'function))
                        (get-lisp-function (symbol-name (cadr fn-expr))))
                       ((and (listp fn-expr)
                             (string-equal (symbol-name (car fn-expr)) "LAMBDA"))
                        (eval-formula fn-expr))
                       (t (eval-formula fn-expr))))
                 (alist (eval-formula (caddr expr))))
            (if (functionp fn)
                (funcall (if (string-equal op-name "ASSOC-IF")
                             #'assoc-if
                             #'rassoc-if)
                         fn alist)
                (format nil "ERR: not a function"))))
         ;; SORT（非破壊的：コピーしてからソート、:key対応）
         ((string-equal op-name "SORT")
          (let* ((args (cdr expr))
                 (seq (eval-formula (first args)))
                 (pred-expr (second args))
                 (pred (cond
                         ((null pred-expr) #'<)
                         ((and (listp pred-expr) (eq (car pred-expr) 'function))
                          (get-lisp-function (symbol-name (cadr pred-expr))))
                         ((and (listp pred-expr)
                               (string-equal (symbol-name (car pred-expr)) "LAMBDA"))
                          (eval-formula pred-expr))
                         (t (eval-formula pred-expr))))
                 ;; :key キーワード引数を探す
                 (key-pos (position :key args :test #'(lambda (k x) 
                                                        (and (symbolp x)
                                                             (string-equal (symbol-name x) "KEY")))))
                 (key-fn (when key-pos
                           (let ((key-expr (nth (1+ key-pos) args)))
                             (cond
                               ((and (listp key-expr) (eq (car key-expr) 'function))
                                (get-lisp-function (symbol-name (cadr key-expr))))
                               ((and (listp key-expr)
                                     (string-equal (symbol-name (car key-expr)) "LAMBDA"))
                                (eval-formula key-expr))
                               ;; :key :age のようなキーワードアクセス
                               ((keywordp key-expr)
                                (lambda (x) (getf x key-expr)))
                               (t (eval-formula key-expr)))))))
            (if (functionp pred)
                (if key-fn
                    (sort (copy-seq seq) pred :key key-fn)
                    (sort (copy-seq seq) pred))
                (format nil "ERR: not a function"))))
         ;; STABLE-SORT（非破壊的、:key対応）
         ((string-equal op-name "STABLE-SORT")
          (let* ((args (cdr expr))
                 (seq (eval-formula (first args)))
                 (pred-expr (second args))
                 (pred (cond
                         ((null pred-expr) #'<)
                         ((and (listp pred-expr) (eq (car pred-expr) 'function))
                          (get-lisp-function (symbol-name (cadr pred-expr))))
                         ((and (listp pred-expr)
                               (string-equal (symbol-name (car pred-expr)) "LAMBDA"))
                          (eval-formula pred-expr))
                         (t (eval-formula pred-expr))))
                 (key-pos (position :key args :test #'(lambda (k x)
                                                        (and (symbolp x)
                                                             (string-equal (symbol-name x) "KEY")))))
                 (key-fn (when key-pos
                           (let ((key-expr (nth (1+ key-pos) args)))
                             (cond
                               ((and (listp key-expr) (eq (car key-expr) 'function))
                                (get-lisp-function (symbol-name (cadr key-expr))))
                               ((and (listp key-expr)
                                     (string-equal (symbol-name (car key-expr)) "LAMBDA"))
                                (eval-formula key-expr))
                               ((keywordp key-expr)
                                (lambda (x) (getf x key-expr)))
                               (t (eval-formula key-expr)))))))
            (if (functionp pred)
                (if key-fn
                    (stable-sort (copy-seq seq) pred :key key-fn)
                    (stable-sort (copy-seq seq) pred))
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

(defun draw-cell-background (canvas x y val selected in-selection)
  "セルの背景のみを描画"
  (let* ((px (col-left x))
         (py (row-top y))
         (w (col-width x))
         (h (row-height y))
         (px2 (+ px w))
         (py2 (+ py h))
         ;; 値の型によって背景色を変える
         (bg (cond
               (selected "#cce5ff")                    ; カーソル位置
               (in-selection "#d0e8ff")                ; 選択範囲
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
  "セルのテキストのみを描画（数値は右寄せ、それ以外は左寄せ）"
  (when val
    (let* ((px (col-left x))
           (py (row-top y))
           (w (col-width x))
           (h (row-height y))
           (display-text (format-value val))
           (path (widget-path canvas))
           ;; 数値かどうか
           (is-number (numberp val)))
      ;; 数値は右寄せ、それ以外は左寄せ
      (if is-number
          ;; 右寄せ（anchor: e）
          (format-wish "~a create text ~a ~a -anchor e -text {~a} -font {Consolas 10}"
                       path (+ px w -4) (+ py (floor h 2)) display-text)
          ;; 左寄せ（anchor: w）
          (format-wish "~a create text ~a ~a -anchor w -text {~a} -font {Consolas 10}"
                       path (+ px 4) (+ py (floor h 2)) display-text)))))

(defun draw-headers (canvas)
  "列名(A,B,C...)と行番号(1,2,3...)のヘッダーを描画"
  (let ((path (widget-path canvas)))
    ;; 左上隅の空白セル
    (format-wish "~a create rectangle 0 0 ~a ~a -fill {#e0e0e0} -outline gray"
                 path *header-w* *header-h*)
    ;; 列名ヘッダー
    (dotimes (x *cols*)
      (let* ((px (col-left x))
             (w (col-width x))
             (px2 (+ px w))
             (col-name (string (code-char (+ (char-code #\A) x)))))
        (format-wish "~a create rectangle ~a 0 ~a ~a -fill {#e0e0e0} -outline gray"
                     path px px2 *header-h*)
        (format-wish "~a create text ~a ~a -anchor center -text {~a} -font {Consolas 11 bold}"
                     path (+ px (floor w 2)) (floor *header-h* 2) col-name)))
    ;; 行番号ヘッダー
    (dotimes (y *rows*)
      (let* ((py (row-top y))
             (h (row-height y))
             (py2 (+ py h))
             (row-num (1+ y)))
        (format-wish "~a create rectangle 0 ~a ~a ~a -fill {#e0e0e0} -outline gray"
                     path py *header-w* py2)
        (format-wish "~a create text ~a ~a -anchor center -text {~a} -font {Consolas 11 bold}"
                     path (floor *header-w* 2) (+ py (floor h 2)) row-num)))))

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
                              (and (= x *cur-x*) (= y *cur-y*))
                              (cell-in-selection-p x y)))))
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

;;;; =========================
;;;; Syntax Highlighting
;;;; =========================

(defparameter *rainbow-colors*
  '("#E00000"    ; 深さ0: 赤
    "#0000FF"    ; 深さ1: 青
    "#008800"    ; 深さ2: 緑
    "#DD6600"    ; 深さ3: オレンジ
    "#AA00AA"    ; 深さ4: 紫
    "#007777")   ; 深さ5: シアン
  "Rainbow括弧の色リスト（深さに応じて循環）")

(defparameter *syntax-colors*
  '((:string    . "#B8860B")    ; 文字列: ダークゴールデンロッド
    (:number    . "#008888")    ; 数値: ティール
    (:keyword   . "#9932CC")    ; キーワード: ダークオーキッド
    (:cell-ref  . "#228B22")    ; セル参照: フォレストグリーン
    (:function  . "#0000CD")    ; 関数名: ミディアムブルー
    (:special   . "#DC143C")    ; 特殊シンボル: クリムゾン
    (:rel-ref   . "#006400"))   ; 相対参照: ダークグリーン
  "シンタックスハイライトの色定義")

(defparameter *known-functions*
  '("+" "-" "*" "/" "=" "/=" "<" ">" "<=" ">=" 
    "if" "cond" "and" "or" "not"
    "sum" "avg" "count" "cell-count" "max" "min"
    "mapcar" "mapc" "maplist" "mapcan" "mapcon" "reduce"
    "remove-if" "remove-if-not" "remove-duplicates"
    "count-if" "count-if-not" "find" "find-if" "position" "position-if"
    "substitute" "substitute-if" "substitute-if-not"
    "every" "some" "notevery" "notany"
    "intersection" "union" "set-difference" "set-exclusive-or" "subsetp"
    "search" "mismatch"
    "assoc" "assoc-if" "rassoc" "rassoc-if" "pairlis" "acons" "getf"
    "apply" "funcall" "lambda" "function"
    "list" "cons" "car" "cdr" "first" "rest" "nth" "length" "elt"
    "append" "reverse" "sort" "stable-sort" "concatenate" "member" "subseq"
    "copy-list" "copy-seq" "copy-tree" "copy-alist"
    "abs" "sqrt" "expt" "exp" "log" "sin" "cos" "tan"
    "floor" "ceiling" "round" "truncate" "mod" "rem"
    "random" "logand" "logior" "logxor" "lognot" "ash"
    "string" "string-upcase" "string-downcase" "char"
    "char=" "char-equal" "digit-char"
    "format" "parse-integer" "type-of" "typep" "coerce"
    "range" "iota" "this-row" "this-col" "this-col-name" "this-cell-name" "cell-at"
    "quote")
  "ハイライト対象の関数名リスト")

(defparameter *special-symbols*
  '("t" "nil")
  "特殊シンボルリスト")

(defparameter *rel-symbols*
  '("rel" "rel-range")
  "相対参照シンボルリスト")

(defun setup-syntax-tags (text-widget)
  "Textウィジェットにシンタックスハイライト用タグを設定"
  ;; Rainbow括弧のタグ
  (loop for color in *rainbow-colors*
        for i from 0
        do (format-wish "~a tag configure paren~d -foreground {~a}"
                        (widget-path text-widget) i color))
  ;; 不一致括弧用のタグ（赤背景）
  (format-wish "~a tag configure paren-error -foreground white -background red"
               (widget-path text-widget))
  ;; シンタックス要素のタグ
  (format-wish "~a tag configure syn-string -foreground {~a}"
               (widget-path text-widget) (cdr (assoc :string *syntax-colors*)))
  (format-wish "~a tag configure syn-number -foreground {~a}"
               (widget-path text-widget) (cdr (assoc :number *syntax-colors*)))
  (format-wish "~a tag configure syn-keyword -foreground {~a}"
               (widget-path text-widget) (cdr (assoc :keyword *syntax-colors*)))
  (format-wish "~a tag configure syn-cell-ref -foreground {~a}"
               (widget-path text-widget) (cdr (assoc :cell-ref *syntax-colors*)))
  (format-wish "~a tag configure syn-function -foreground {~a}"
               (widget-path text-widget) (cdr (assoc :function *syntax-colors*)))
  (format-wish "~a tag configure syn-special -foreground {~a}"
               (widget-path text-widget) (cdr (assoc :special *syntax-colors*)))
  (format-wish "~a tag configure syn-rel-ref -foreground {~a}"
               (widget-path text-widget) (cdr (assoc :rel-ref *syntax-colors*))))

(defun pos-to-tk-index (content pos)
  "文字位置をTkのline.col形式に変換"
  (let* ((before (subseq content 0 pos))
         (line (1+ (count #\Newline before)))
         (last-nl (position #\Newline before :from-end t))
         (col (if last-nl (- pos last-nl 1) pos)))
    (format nil "~d.~d" line col)))

(defun add-syntax-tag (text-widget content tag start end)
  "指定範囲にシンタックスタグを適用"
  (let ((tk-start (pos-to-tk-index content start))
        (tk-end (pos-to-tk-index content end)))
    (format-wish "~a tag add ~a ~a ~a"
                 (widget-path text-widget) tag tk-start tk-end)))

(defun symbol-char-p (char)
  "シンボルを構成できる文字か判定"
  (or (alphanumericp char)
      (find char "+-*/<>=!?_")))

(defun syntax-highlight (text-widget)
  "Textウィジェット内をシンタックスハイライトする"
  (let* ((content (text text-widget))
         (len (length content))
         (num-colors (length *rainbow-colors*))
         ;; 解析用状態
         (i 0)
         (depth 0)
         (paren-positions nil)
         (tokens nil))  ; ((type start end) ...)
    
    ;; 既存のタグを削除
    (loop for j from 0 below num-colors
          do (format-wish "~a tag remove paren~d 1.0 end"
                          (widget-path text-widget) j))
    (format-wish "~a tag remove paren-error 1.0 end" (widget-path text-widget))
    (dolist (tag '("syn-string" "syn-number" "syn-keyword" 
                   "syn-cell-ref" "syn-function" "syn-special" "syn-rel-ref"))
      (format-wish "~a tag remove ~a 1.0 end" (widget-path text-widget) tag))
    
    ;; トークン解析
    (loop while (< i len) do
      (let ((char (char content i)))
        (cond
          ;; 文字列
          ((char= char #\")
           (let ((start i))
             (incf i)
             (loop while (< i len)
                   for c = (char content i)
                   do (cond
                        ((char= c #\\) (incf i 2))  ; エスケープをスキップ
                        ((char= c #\") (incf i) (return))
                        (t (incf i))))
             (push (list :string start i) tokens)))
          
          ;; キーワード
          ((char= char #\:)
           (let ((start i))
             (incf i)
             (loop while (and (< i len) (symbol-char-p (char content i)))
                   do (incf i))
             (when (> i (1+ start))
               (push (list :keyword start i) tokens))))
          
          ;; 開き括弧と関数名
          ((char= char #\()
           (push (cons i depth) paren-positions)
           (incf depth)
           (incf i)
           ;; 括弧直後のシンボルを関数名として検出
           (loop while (and (< i len) (find (char content i) " ~%	"))
                 do (incf i))
           (when (and (< i len) (symbol-char-p (char content i)))
             (let ((start i))
               (loop while (and (< i len) (symbol-char-p (char content i)))
                     do (incf i))
               (let ((sym (subseq content start i)))
                 (cond
                   ;; 相対参照
                   ((member sym *rel-symbols* :test #'string-equal)
                    (push (list :rel-ref start i) tokens))
                   ;; 既知の関数
                   ((member sym *known-functions* :test #'string-equal)
                    (push (list :function start i) tokens))
                   ;; セル参照 (関数位置でも)
                   ((and (>= (length sym) 2)
                         (<= (length sym) 4)
                         (alpha-char-p (char sym 0))
                         (every #'digit-char-p (subseq sym 1)))
                    (push (list :cell-ref start i) tokens)))))))
          
          ;; 閉じ括弧
          ((char= char #\))
           (decf depth)
           (if (< depth 0)
               (progn
                 (push (cons i -1) paren-positions)
                 (setf depth 0))
               (push (cons i depth) paren-positions))
           (incf i))
          
          ;; 数値（符号付きも対応）
          ((or (digit-char-p char)
               (and (or (char= char #\-) (char= char #\+))
                    (< (1+ i) len)
                    (digit-char-p (char content (1+ i)))))
           (let ((start i))
             (when (or (char= char #\-) (char= char #\+))
               (incf i))
             (loop while (and (< i len) 
                              (or (digit-char-p (char content i))
                                  (char= (char content i) #\.)))
                   do (incf i))
             ;; 指数表記
             (when (and (< i len) (find (char content i) "eE"))
               (incf i)
               (when (and (< i len) (find (char content i) "+-"))
                 (incf i))
               (loop while (and (< i len) (digit-char-p (char content i)))
                     do (incf i)))
             (push (list :number start i) tokens)))
          
          ;; シンボル（セル参照、特殊シンボル）
          ((alpha-char-p char)
           (let ((start i))
             (loop while (and (< i len) (symbol-char-p (char content i)))
                   do (incf i))
             (let ((sym (subseq content start i)))
               (cond
                 ;; 相対参照
                 ((member sym *rel-symbols* :test #'string-equal)
                  (push (list :rel-ref start i) tokens))
                 ;; 特殊シンボル
                 ((member sym *special-symbols* :test #'string-equal)
                  (push (list :special start i) tokens))
                 ;; セル参照 (A1-Z99, AA1-ZZ99形式)
                 ((and (>= (length sym) 2)
                       (<= (length sym) 4)
                       (alpha-char-p (char sym 0))
                       (or (and (= (length sym) 2)
                                (digit-char-p (char sym 1)))
                           (and (>= (length sym) 2)
                                (let ((first-digit-pos 
                                       (position-if #'digit-char-p sym)))
                                  (and first-digit-pos
                                       (> first-digit-pos 0)
                                       (every #'alpha-char-p (subseq sym 0 first-digit-pos))
                                       (every #'digit-char-p (subseq sym first-digit-pos)))))))
                  (push (list :cell-ref start i) tokens))))))
          
          ;; その他
          (t (incf i)))))
    
    ;; タグを適用
    ;; 括弧（Rainbow）
    (dolist (pp paren-positions)
      (let* ((pos (car pp))
             (d (cdr pp))
             (tk-start (pos-to-tk-index content pos))
             (tk-end (pos-to-tk-index content (1+ pos))))
        (if (= d -1)
            (format-wish "~a tag add paren-error ~a ~a"
                         (widget-path text-widget) tk-start tk-end)
            (format-wish "~a tag add paren~d ~a ~a"
                         (widget-path text-widget) 
                         (mod d num-colors) tk-start tk-end))))
    
    ;; その他のトークン
    (dolist (tok tokens)
      (let ((type (first tok))
            (start (second tok))
            (end (third tok)))
        (add-syntax-tag text-widget content
                        (case type
                          (:string "syn-string")
                          (:number "syn-number")
                          (:keyword "syn-keyword")
                          (:cell-ref "syn-cell-ref")
                          (:function "syn-function")
                          (:special "syn-special")
                          (:rel-ref "syn-rel-ref"))
                        start end)))))

;; 後方互換性のためのエイリアス
(defun setup-rainbow-tags (text-widget)
  (setup-syntax-tags text-widget))

(defun colorize-parentheses (text-widget)
  (syntax-highlight text-widget))

;;;; =========================
;;;; S式フォーマッター
;;;; =========================

(defun pprint-to-string (form)
  "S式を整形して文字列に変換"
  (let ((*print-pretty* t)
        (*print-right-margin* 60)
        (*print-miser-width* 40)
        (*package* (find-package :ss-gui)))
    (with-output-to-string (s)
      (pprint form s))))

(defun format-sexp (text-widget)
  "入力枠のS式を整形する"
  (let* ((content (get-text-content text-widget))
         (trimmed (string-trim '(#\Space #\Tab #\Newline) content)))
    (when (and (> (length trimmed) 0)
               (char= (char trimmed 0) #\=))
      ;; =で始まる数式の場合
      (handler-case
          (let* ((*package* (find-package :ss-gui))
                 (form (read-from-string (subseq trimmed 1)))
                 (formatted (string-trim '(#\Newline) (pprint-to-string form))))
            (set-text-content text-widget (format nil "=~a" formatted))
            (syntax-highlight text-widget))
        (error (e)
          ;; パースエラーの場合はメッセージ表示
          (format t "Format error: ~a~%" e))))))

(defun update-text-input (text-widget)
  "現在セルの内容をTextウィジェットに表示"
  (let ((c (current-cell))
        (*package* (find-package :ss-gui)))  ; パッケージプレフィックスなしで表示
    (cond
      ;; 数式があれば =(...) 形式で表示
      ((cell-formula c)
       (set-text-content text-widget (format nil "=~S" (cell-formula c))))
      ;; 値があればそのまま表示
      ((cell-value c)
       (set-text-content text-widget (format nil "~S" (cell-value c))))
      (t
       (set-text-content text-widget ""))))
  ;; Rainbow括弧の色分けを更新
  (colorize-parentheses text-widget))

(defun format-error-message (error-type details)
  "エラーメッセージを整形"
  (case error-type
    (:syntax (format nil "#構文: ~a" details))
    (:eval   (format nil "#評価: ~a" details))
    (:ref    (format nil "#参照: ~a" details))
    (:circular (format nil "#循環: ~a" details))
    (t (format nil "#ERR: ~a" details))))

(defun commit-text-input (text-widget canvas)
  "Textウィジェットの内容をセルに確定"
  (let* ((raw (or (get-text-content text-widget) ""))
         (cell (current-cell))
         (cell-nm (cell-name *cur-x* *cur-y*))
         ;; Undo用に変更前の状態を保存
         (old-value (cell-value cell))
         (old-formula (cell-formula cell))
         (error-occurred nil))
    ;; 評価位置を設定
    (setf *eval-col* *cur-x*
          *eval-row* *cur-y*)
    (if (and (> (length raw) 0)
             (char= (char raw 0) #\=))
        ;; =で始まる → 数式として解析・評価
        (let ((form nil))
          ;; 構文解析
          (handler-case
              (let ((*package* (find-package :ss-gui)))
                (setf form (read-from-string (subseq raw 1))))
            (end-of-file ()
              (setf (cell-value cell) (format-error-message :syntax "式が不完全です")
                    (cell-formula cell) nil
                    error-occurred t))
            (error (e)
              (setf (cell-value cell) (format-error-message :syntax (princ-to-string e))
                    (cell-formula cell) nil
                    error-occurred t)))
          ;; 評価（構文解析が成功した場合のみ）
          (unless error-occurred
            (handler-case
                (let* ((*eval-stack* (list cell-nm))  ; 循環参照検出用
                       (value (eval-formula form))
                       (refs (extract-references form *cur-y* *cur-x*)))
                  (setf (cell-formula cell) form
                        (cell-value cell) value)
                  ;; 依存関係を更新
                  (update-dependencies cell-nm refs))
              (error (e)
                (let ((msg (princ-to-string e)))
                  ;; 循環参照エラーを検出
                  (if (search "循環参照" msg)
                      (setf (cell-value cell) (format-error-message :circular cell-nm))
                      (setf (cell-value cell) (format-error-message :eval msg)))
                  ;; エラー時も数式を保持（再編集可能に）
                  (setf (cell-formula cell) form)
                  (update-dependencies cell-nm nil))))))
        ;; それ以外 → 通常の値として解釈
        (let ((parsed (handler-case 
                          (let ((*package* (find-package :ss-gui)))
                            (read-from-string raw))
                        (end-of-file () nil)
                        (error () raw))))
          (setf (cell-formula cell) nil
                (cell-value cell) (if (string= raw "") nil (or parsed raw)))
          ;; 数式がないので依存関係をクリア
          (update-dependencies cell-nm nil)))
    ;; 変更があった場合のみUndo履歴に記録
    (unless (and (equal old-value (cell-value cell))
                 (equal old-formula (cell-formula cell)))
      (record-cell-change cell-nm old-value old-formula))
    ;; 依存元を再計算
    (recalc-dependents cell-nm)
    (redraw canvas)))

(defun commit-and-move (text-widget canvas direction)
  "入力を確定して指定方向に移動
   direction: :down, :up, :right, :left, :stay"
  (commit-text-input text-widget canvas)
  (clear-selection)
  (case direction
    (:down  (when (< *cur-y* (1- *rows*)) (incf *cur-y*)))
    (:up    (when (> *cur-y* 0) (decf *cur-y*)))
    (:right (when (< *cur-x* (1- *cols*)) (incf *cur-x*)))
    (:left  (when (> *cur-x* 0) (decf *cur-x*)))
    (:stay  nil))
  (setf *sel-start-x* *cur-x*
        *sel-start-y* *cur-y*
        *sel-end-x* *cur-x*
        *sel-end-y* *cur-y*)
  (update-text-input text-widget)
  (redraw canvas))

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
;;;; ダイアログ
;;;; =========================

;; 入力ダイアログの結果を保持する変数
(defparameter *input-dialog-result* nil)

(defun show-size-dialog (title prompt default-value callback)
  "サイズ入力ダイアログを表示。callbackは入力値を受け取る関数"
  ;; ダイアログウィンドウを作成
  (let* ((dlg (make-instance 'toplevel))
         (lbl (make-instance 'label :master dlg :text prompt))
         (ent (make-instance 'entry :master dlg :width 15))
         (btn-frame (make-instance 'frame :master dlg))
         (ok-btn (make-instance 'button :master btn-frame :text "OK" :width 8))
         (cancel-btn (make-instance 'button :master btn-frame :text "Cancel" :width 8)))
    
    ;; ウィンドウ設定
    (wm-title dlg title)
    (format-wish "wm transient ~a ." (widget-path dlg))
    (format-wish "wm resizable ~a 0 0" (widget-path dlg))
    
    ;; 初期値を設定
    (setf (text ent) (princ-to-string default-value))
    
    ;; OKボタンの動作
    (setf (command ok-btn)
          (lambda ()
            (let* ((input (text ent))
                   (value (ignore-errors (parse-integer input))))
              (destroy dlg)
              (when (and value (> value 0))
                (funcall callback value)))))
    
    ;; キャンセルボタンの動作
    (setf (command cancel-btn)
          (lambda ()
            (destroy dlg)))
    
    ;; Enterキーでも確定
    (bind ent "<Return>"
          (lambda (evt)
            (declare (ignore evt))
            (let* ((input (text ent))
                   (value (ignore-errors (parse-integer input))))
              (destroy dlg)
              (when (and value (> value 0))
                (funcall callback value)))))
    
    ;; Escapeキーでキャンセル
    (bind dlg "<Escape>"
          (lambda (evt)
            (declare (ignore evt))
            (destroy dlg)))
    
    ;; レイアウト
    (pack lbl :padx 10 :pady 5)
    (pack ent :padx 10 :pady 5)
    (pack btn-frame :pady 10)
    (pack ok-btn :side :left :padx 5)
    (pack cancel-btn :side :left :padx 5)
    
    ;; フォーカスを入力欄に
    (focus ent)
    (format-wish "~a selection range 0 end" (widget-path ent))
    
    ;; モーダルにする
    (format-wish "grab ~a" (widget-path dlg))))

;;;; =========================
;;;; メイン：GUIの構築と起動
;;;; =========================

(defun update-window-title ()
  "ウィンドウタイトルを更新"
  (wm-title *tk* (format nil "Spreadsheet v0.5.2 [~Dx~D]~a" 
                         *cols* *rows*
                         (if *current-file* 
                             (format nil " - ~a" (file-namestring *current-file*))
                             ""))))

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
  (clear-selection)
  (setf *clipboard* nil)
  (clear-dependencies)
  (setf *current-file* nil)
  (init-sizes)  ; 列幅・行高さを初期化
  
  (with-ltk ()
    (update-window-title)
    
    ;; 垂直分割用のPanedWindowをTclコマンドで作成
    (format-wish "ttk::panedwindow .paned -orient vertical")
    
    ;; ウィジェット作成
    (let* (;; 入力エリア用フレーム（.paned配下）
           (input-frame (make-instance 'frame))
           ;; 複数行入力用Textウィジェット（固定幅フォント）
           (input-text (make-instance 'text
                                      :master input-frame
                                      :width 80
                                      :height input-lines
                                      :font "TkFixedFont"))
           (input-scroll (make-instance 'scrollbar 
                                        :master input-frame
                                        :orientation :vertical))
           ;; スプレッドシート用フレーム
           (canvas-frame (make-instance 'frame))
           (canvas (make-instance 'canvas
                                  :master canvas-frame
                                  :width (total-width)
                                  :height (total-height)))
           ;; メニューバー
           (mb (make-menubar))
           (file-menu (make-menu mb "File"))
           (edit-menu (make-menu mb "Edit")))
      
      ;; ファイルメニュー項目
      (make-menubutton file-menu "New             Ctrl+N"
                       (lambda ()
                         (new-sheet)
                         (update-window-title)
                         (update-text-input input-text)
                         (redraw canvas)))
      
      (make-menubutton file-menu "Open...         Ctrl+O"
                       (lambda ()
                         (let ((filename (get-open-file 
                                          :filetypes '(("Spreadsheet" "*.ssp")
                                                       ("All files" "*")))))
                           (when (and filename (> (length filename) 0))
                             (handler-case
                                 (progn
                                   (load-file filename)
                                   (update-window-title)
                                   (update-text-input input-text)
                                   (redraw canvas))
                               (error (e)
                                 (do-msg (format nil "Load Error: ~a" e))))))))
      
      (make-menubutton file-menu "Save            Ctrl+S"
                       (lambda ()
                         (if *current-file*
                             (progn
                               (save *current-file*)
                               (do-msg (format nil "Saved: ~a" 
                                              (file-namestring *current-file*))))
                             ;; ファイルがない場合は名前を付けて保存
                             (let ((filename (get-save-file 
                                              :filetypes '(("Spreadsheet" "*.ssp")
                                                           ("All files" "*")))))
                               (when (and filename (> (length filename) 0))
                                 ;; 拡張子がなければ追加
                                 (unless (search ".ssp" filename :test #'char-equal)
                                   (setf filename (concatenate 'string filename ".ssp")))
                                 (save filename)
                                 (update-window-title)
                                 (do-msg (format nil "Saved: ~a" 
                                                (file-namestring filename))))))))
      
      (make-menubutton file-menu "Save As..."
                       (lambda ()
                         (let ((filename (get-save-file 
                                          :filetypes '(("Spreadsheet" "*.ssp")
                                                       ("All files" "*")))))
                           (when (and filename (> (length filename) 0))
                             ;; 拡張子がなければ追加
                             (unless (search ".ssp" filename :test #'char-equal)
                               (setf filename (concatenate 'string filename ".ssp")))
                             (save filename)
                             (update-window-title)
                             (do-msg (format nil "Saved: ~a" 
                                            (file-namestring filename)))))))
      
      (add-separator file-menu)
      
      (make-menubutton file-menu "Import CSV..."
                       (lambda ()
                         (let ((filename (get-open-file 
                                          :filetypes '(("CSV files" "*.csv")
                                                       ("TSV files" "*.tsv")
                                                       ("Text files" "*.txt")
                                                       ("All files" "*")))))
                           (when (and filename (> (length filename) 0))
                             (handler-case
                                 (progn
                                   (import-csv filename)
                                   (update-window-title)
                                   (update-text-input input-text)
                                   (redraw canvas)
                                   (do-msg (format nil "Imported: ~a" 
                                                  (file-namestring filename))))
                               (error (e)
                                 (do-msg (format nil "Import Error: ~a" e))))))))
      
      (make-menubutton file-menu "Export CSV..."
                       (lambda ()
                         (let ((filename (get-save-file 
                                          :filetypes '(("CSV files" "*.csv")
                                                       ("All files" "*")))))
                           (when (and filename (> (length filename) 0))
                             ;; 拡張子がなければ追加
                             (unless (search ".csv" filename :test #'char-equal)
                               (setf filename (concatenate 'string filename ".csv")))
                             (handler-case
                                 (progn
                                   (export-csv filename)
                                   (do-msg (format nil "Exported: ~a" 
                                                  (file-namestring filename))))
                               (error (e)
                                 (do-msg (format nil "Export Error: ~a" e))))))))
      
      (add-separator file-menu)
      
      (make-menubutton file-menu "Exit"
                       (lambda ()
                         (setf *exit-mainloop* t)))
      
      ;; 編集メニュー項目
      (make-menubutton edit-menu "Undo            Ctrl+Z"
                       (lambda ()
                         (when (undo)
                           (update-text-input input-text)
                           (redraw canvas))))
      
      (make-menubutton edit-menu "Redo            Ctrl+Y"
                       (lambda ()
                         (when (redo)
                           (update-text-input input-text)
                           (redraw canvas))))
      
      (add-separator edit-menu)
      
      (make-menubutton edit-menu "Cut             Ctrl+X"
                       (lambda ()
                         (copy-to-system-clipboard)
                         (clear-selection-cells)
                         (clear-selection)
                         (redraw canvas)
                         (update-text-input input-text)))
      
      (make-menubutton edit-menu "Copy            Ctrl+C"
                       (lambda ()
                         (copy-to-system-clipboard)))
      
      (make-menubutton edit-menu "Paste           Ctrl+V"
                       (lambda ()
                         (let ((sys-clip (get-system-clipboard)))
                           (if (and sys-clip (> (length sys-clip) 0))
                               (paste-from-system-clipboard)
                               (paste-clipboard)))
                         (clear-selection)
                         (redraw canvas)
                         (update-text-input input-text)))
      
      (add-separator edit-menu)
      
      (make-menubutton edit-menu "Delete          Delete"
                       (lambda ()
                         (clear-selection-cells)
                         (clear-selection)
                         (redraw canvas)
                         (update-text-input input-text)))
      
      (add-separator edit-menu)
      
      (make-menubutton edit-menu "Insert Row"
                       (lambda ()
                         (insert-row *cur-y*)
                         (configure canvas :height (total-height))
                         (redraw canvas)))
      
      (make-menubutton edit-menu "Insert Column"
                       (lambda ()
                         (insert-col *cur-x*)
                         (configure canvas :width (total-width))
                         (redraw canvas)))
      
      (make-menubutton edit-menu "Delete Row"
                       (lambda ()
                         (delete-row *cur-y*)
                         (configure canvas :height (total-height))
                         (redraw canvas)
                         (update-text-input input-text)))
      
      (make-menubutton edit-menu "Delete Column"
                       (lambda ()
                         (delete-col *cur-x*)
                         (configure canvas :width (total-width))
                         (redraw canvas)
                         (update-text-input input-text)))
      
      ;; セパレーター
      (add-separator edit-menu)
      
      ;; Format Expression
      (make-menubutton edit-menu "Format Expression   Ctrl+Shift+F"
                       (lambda ()
                         (format-sexp input-text)))
      
      ;; キーボードショートカット (Ctrl+N, Ctrl+O, Ctrl+S)
      (bind *tk* "<Control-n>"
            (lambda (evt)
              (declare (ignore evt))
              (new-sheet)
              (update-window-title)
              (update-text-input input-text)
              (redraw canvas)))
      
      (bind *tk* "<Control-o>"
            (lambda (evt)
              (declare (ignore evt))
              (let ((filename (get-open-file 
                               :filetypes '(("Spreadsheet" "*.ssp")
                                            ("All files" "*")))))
                (when (and filename (> (length filename) 0))
                  (handler-case
                      (progn
                        (load-file filename)
                        (update-window-title)
                        (update-text-input input-text)
                        (redraw canvas))
                    (error (e)
                      (do-msg (format nil "読み込みエラー: ~a" e))))))))
      
      (bind *tk* "<Control-s>"
            (lambda (evt)
              (declare (ignore evt))
              (if *current-file*
                  (progn
                    (save *current-file*)
                    (do-msg (format nil "保存しました: ~a" 
                                   (file-namestring *current-file*))))
                  ;; ファイルがない場合は名前を付けて保存
                  (let ((filename (get-save-file 
                                   :filetypes '(("Spreadsheet" "*.ssp")
                                                ("All files" "*")))))
                    (when (and filename (> (length filename) 0))
                      (unless (search ".ssp" filename :test #'char-equal)
                        (setf filename (concatenate 'string filename ".ssp")))
                      (save filename)
                      (update-window-title)
                      (do-msg (format nil "保存しました: ~a" 
                                     (file-namestring filename))))))))
      
      ;; Ctrl+Z → Undo
      (bind *tk* "<Control-z>"
            (lambda (evt)
              (declare (ignore evt))
              (when (undo)
                (update-text-input input-text)
                (redraw canvas))))
      
      ;; Ctrl+Y → Redo
      (bind *tk* "<Control-y>"
            (lambda (evt)
              (declare (ignore evt))
              (when (redo)
                (update-text-input input-text)
                (redraw canvas))))
      
      ;; Ctrl+X → 切り取り
      (bind *tk* "<Control-x>"
            (lambda (evt)
              (declare (ignore evt))
              (copy-to-system-clipboard)
              (clear-selection-cells)
              (clear-selection)
              (redraw canvas)
              (update-text-input input-text)))
      
      ;; スクロールバーとテキストを連携
      (configure input-scroll :command (format nil "~a yview" (widget-path input-text)))
      (configure input-text :yscrollcommand (format nil "~a set" (widget-path input-scroll)))
      ;; タブ幅を4文字分に設定
      (format-wish "~a configure -tabs [list [expr {[font measure TkFixedFont 0] * 4}]]" 
                   (widget-path input-text))
      
      ;; Rainbow括弧のタグをセットアップ
      (setup-rainbow-tags input-text)
      ;; キー入力時に括弧を色分け
      (bind input-text "<KeyRelease>"
            (lambda (evt)
              (declare (ignore evt))
              (colorize-parentheses input-text)))
      
      ;; レイアウト - 入力フレーム内
      (pack input-scroll :side :right :fill :y)
      (pack input-text :side :left :fill :both :expand t)
      ;; レイアウト - キャンバスフレーム内
      (pack canvas :fill :both :expand t)
      
      ;; PanedWindowにペインを追加
      (format-wish ".paned add ~a -weight 0" (widget-path input-frame))
      (format-wish ".paned add ~a -weight 1" (widget-path canvas-frame))
      ;; PanedWindowをパック
      (format-wish "pack .paned -fill both -expand true -padx 2 -pady 2")

      ;; 初期描画
      (update-text-input input-text)
      (redraw canvas)

      ;;; --- イベントバインド ---

      ;; Enter → 確定して下に移動
      (bind input-text "<Return>"
            (lambda (evt)
              (declare (ignore evt))
              (commit-and-move input-text canvas :down)
              (focus canvas)))
      
      ;; Ctrl+Enter → 確定してそのまま
      (bind input-text "<Control-Return>"
            (lambda (evt)
              (declare (ignore evt))
              (commit-and-move input-text canvas :stay)
              (focus canvas)))
      
      ;; Alt+Enter → 確定して右に移動
      (bind input-text "<Alt-Return>"
            (lambda (evt)
              (declare (ignore evt))
              (commit-and-move input-text canvas :right)
              (focus canvas)))
      
      ;; Shift+Alt+Enter → 確定して左に移動
      (bind input-text "<Shift-Alt-Return>"
            (lambda (evt)
              (declare (ignore evt))
              (commit-and-move input-text canvas :left)
              (focus canvas)))
      
      ;; Escape → 編集キャンセル（元に戻す）
      (bind input-text "<Escape>"
            (lambda (evt)
              (declare (ignore evt))
              (update-text-input input-text)
              (focus canvas)))
      
      ;; Ctrl+Shift+F → S式を整形
      (bind input-text "<Control-Shift-f>"
            (lambda (evt)
              (declare (ignore evt))
              (format-sexp input-text)))
      
      ;; Tclレベルでバインディングを調整（breakでデフォルト動作を抑制）
      (let ((path (widget-path input-text)))
        (format-wish "bind ~a <Return> \"[bind ~a <Return>]; break\"" path path)
        (format-wish "bind ~a <Control-Return> \"[bind ~a <Control-Return>]; break\"" path path)
        (format-wish "bind ~a <Alt-Return> \"[bind ~a <Alt-Return>]; break\"" path path)
        (format-wish "bind ~a <Shift-Alt-Return> \"[bind ~a <Shift-Alt-Return>]; break\"" path path)
        (format-wish "bind ~a <Escape> \"[bind ~a <Escape>]; break\"" path path)
        ;; Shift+Enter で改行
        (format-wish "bind ~a <Shift-Return> {~a insert insert \\n; break}" path path))

      ;; セルクリック → 選択開始 または リサイズ開始
      (bind canvas "<ButtonPress-1>"
            (lambda (evt)
              (let ((mx (and evt (slot-value evt 'ltk::x)))
                    (my (and evt (slot-value evt 'ltk::y))))
                (when (and mx my (numberp mx) (numberp my))
                  ;; ヘッダー領域でのリサイズチェック
                  (let ((col-border (and (< my *header-h*) 
                                         (near-col-border-p mx 4)))
                        (row-border (and (< mx *header-w*) 
                                         (near-row-border-p my 4))))
                    (cond
                      ;; 列幅リサイズ開始
                      (col-border
                       (setf *resize-mode* :col
                             *resize-index* col-border
                             *resize-start* mx))
                      ;; 行高さリサイズ開始
                      (row-border
                       (setf *resize-mode* :row
                             *resize-index* row-border
                             *resize-start* my))
                      ;; 通常のセル選択
                      ((and (> mx *header-w*) (> my *header-h*))
                       (let ((x (clamp (find-col-at mx) 0 (1- *cols*)))
                             (y (clamp (find-row-at my) 0 (1- *rows*))))
                         (setf *cur-x* x
                               *cur-y* y
                               *sel-start-x* x
                               *sel-start-y* y
                               *sel-end-x* x
                               *sel-end-y* y
                               *selecting* t)
                         (update-text-input input-text)
                         (redraw canvas)))))))
              (focus canvas)))

      ;; ドラッグ → 選択範囲拡張 または リサイズ
      (bind canvas "<B1-Motion>"
            (lambda (evt)
              (let ((mx (and evt (slot-value evt 'ltk::x)))
                    (my (and evt (slot-value evt 'ltk::y))))
                (when (and mx my (numberp mx) (numberp my))
                  (cond
                    ;; 列幅リサイズ中
                    ((eq *resize-mode* :col)
                     (let* ((old-w (col-width *resize-index*))
                            (delta (- mx *resize-start*))
                            (new-w (max *min-cell-w* (+ old-w delta))))
                       (set-col-width *resize-index* new-w)
                       (setf *resize-start* mx)
                       ;; キャンバスサイズ更新
                       (configure canvas :width (total-width))
                       (redraw canvas)))
                    ;; 行高さリサイズ中
                    ((eq *resize-mode* :row)
                     (let* ((old-h (row-height *resize-index*))
                            (delta (- my *resize-start*))
                            (new-h (max *min-cell-h* (+ old-h delta))))
                       (set-row-height *resize-index* new-h)
                       (setf *resize-start* my)
                       ;; キャンバスサイズ更新
                       (configure canvas :height (total-height))
                       (redraw canvas)))
                    ;; 通常の選択
                    (*selecting*
                     (let ((x (find-col-at mx))
                           (y (find-row-at my)))
                       (when (>= x 0)
                         (setf *sel-end-x* (clamp x 0 (1- *cols*))))
                       (when (>= y 0)
                         (setf *sel-end-y* (clamp y 0 (1- *rows*)))))
                     (redraw canvas)))))))

      ;; ドラッグ終了
      (bind canvas "<ButtonRelease-1>"
            (lambda (evt)
              (declare (ignore evt))
              (setf *selecting* nil
                    *resize-mode* nil
                    *resize-index* nil)))

      ;; 右クリック用の変数
      (let ((context-col nil)   ; 右クリックされた列
            (context-row nil))  ; 右クリックされた行
        
        ;; 列ヘッダー用コンテキストメニュー作成
        (format-wish "menu .colmenu -tearoff 0")
        (format-wish ".colmenu add command -label {Insert Column} -command {}")
        (format-wish ".colmenu add command -label {Delete Column} -command {}")
        (format-wish ".colmenu add separator")
        (format-wish ".colmenu add command -label {Column Width...} -command {}")
        
        ;; 行ヘッダー用コンテキストメニュー作成
        (format-wish "menu .rowmenu -tearoff 0")
        (format-wish ".rowmenu add command -label {Insert Row} -command {}")
        (format-wish ".rowmenu add command -label {Delete Row} -command {}")
        (format-wish ".rowmenu add separator")
        (format-wish ".rowmenu add command -label {Row Height...} -command {}")
        
        ;; 右クリック → コンテキストメニュー表示
        (bind canvas "<ButtonPress-3>"
              (lambda (evt)
                (let ((mx (and evt (slot-value evt 'ltk::x)))
                      (my (and evt (slot-value evt 'ltk::y))))
                  (when (and mx my (numberp mx) (numberp my))
                    (cond
                      ;; 列ヘッダー上で右クリック
                      ((and (< my *header-h*) (>= mx *header-w*))
                       (setf context-col (find-col-at mx))
                       (when (and context-col (>= context-col 0))
                         ;; メニューコマンドを更新
                         (format-wish ".colmenu entryconfigure 0 -command {event generate . <<ColInsert>>}")
                         (format-wish ".colmenu entryconfigure 1 -command {event generate . <<ColDelete>>}")
                         (format-wish ".colmenu entryconfigure 3 -command {event generate . <<ColWidth>>}")
                         ;; メニュー表示（画面座標を取得）
                         (format-wish "tk_popup .colmenu [expr [winfo rootx ~a] + ~a] [expr [winfo rooty ~a] + ~a]"
                                      (widget-path canvas) mx (widget-path canvas) my)))
                      
                      ;; 行ヘッダー上で右クリック
                      ((and (< mx *header-w*) (>= my *header-h*))
                       (setf context-row (find-row-at my))
                       (when (and context-row (>= context-row 0))
                         ;; メニューコマンドを更新
                         (format-wish ".rowmenu entryconfigure 0 -command {event generate . <<RowInsert>>}")
                         (format-wish ".rowmenu entryconfigure 1 -command {event generate . <<RowDelete>>}")
                         (format-wish ".rowmenu entryconfigure 3 -command {event generate . <<RowHeight>>}")
                         ;; メニュー表示
                         (format-wish "tk_popup .rowmenu [expr [winfo rootx ~a] + ~a] [expr [winfo rooty ~a] + ~a]"
                                      (widget-path canvas) mx (widget-path canvas) my))))))))
        
        ;; 列挿入イベント
        (bind *tk* "<<ColInsert>>"
              (lambda (evt)
                (declare (ignore evt))
                (when context-col
                  (insert-col context-col)
                  (configure canvas :width (total-width))
                  (redraw canvas))))
        
        ;; 列削除イベント
        (bind *tk* "<<ColDelete>>"
              (lambda (evt)
                (declare (ignore evt))
                (when context-col
                  (delete-col context-col)
                  (configure canvas :width (total-width))
                  (redraw canvas)
                  (update-text-input input-text))))
        
        ;; 列幅設定イベント
        (bind *tk* "<<ColWidth>>"
              (lambda (evt)
                (declare (ignore evt))
                (when context-col
                  (let* ((current-w (col-width context-col))
                         (col-name (string (code-char (+ (char-code #\A) context-col))))
                         (target-col context-col))
                    (show-size-dialog "Column Width" 
                                      (format nil "Width for column ~a:" col-name)
                                      current-w
                                      (lambda (w)
                                        (set-col-width target-col w)
                                        (configure canvas :width (total-width))
                                        (redraw canvas)))))))
        
        ;; 行挿入イベント
        (bind *tk* "<<RowInsert>>"
              (lambda (evt)
                (declare (ignore evt))
                (when context-row
                  (insert-row context-row)
                  (configure canvas :height (total-height))
                  (redraw canvas))))
        
        ;; 行削除イベント
        (bind *tk* "<<RowDelete>>"
              (lambda (evt)
                (declare (ignore evt))
                (when context-row
                  (delete-row context-row)
                  (configure canvas :height (total-height))
                  (redraw canvas)
                  (update-text-input input-text))))
        
        ;; 行高さ設定イベント
        (bind *tk* "<<RowHeight>>"
              (lambda (evt)
                (declare (ignore evt))
                (when context-row
                  (let* ((current-h (row-height context-row))
                         (row-num (1+ context-row))
                         (target-row context-row))
                    (show-size-dialog "Row Height"
                                      (format nil "Height for row ~a:" row-num)
                                      current-h
                                      (lambda (h)
                                        (set-row-height target-row h)
                                        (configure canvas :height (total-height))
                                        (redraw canvas))))))))

      ;; Ctrl+C → システムクリップボードにコピー
      (bind canvas "<Control-c>"
            (lambda (evt)
              (declare (ignore evt))
              (copy-to-system-clipboard)))

      ;; Ctrl+V → システムクリップボードからペースト
      (bind canvas "<Control-v>"
            (lambda (evt)
              (declare (ignore evt))
              (let ((sys-clip (get-system-clipboard)))
                (if (and sys-clip (> (length sys-clip) 0))
                    ;; システムクリップボードにデータあり
                    (paste-from-system-clipboard)
                    ;; なければ内部クリップボード
                    (paste-clipboard)))
              (clear-selection)
              (redraw canvas)
              (update-text-input input-text)))

      ;; Delete → 選択範囲をクリア
      (bind canvas "<Delete>"
            (lambda (evt)
              (declare (ignore evt))
              (clear-selection-cells)
              (clear-selection)
              (redraw canvas)
              (update-text-input input-text)))

      ;; BackSpace → 選択範囲をクリア
      (bind canvas "<BackSpace>"
            (lambda (evt)
              (declare (ignore evt))
              (clear-selection-cells)
              (clear-selection)
              (redraw canvas)
              (update-text-input input-text)))

      ;; 矢印キー（選択解除してから移動）
      (bind canvas "<Left>"
            (lambda (evt) 
              (declare (ignore evt)) 
              (clear-selection)
              (move-left canvas input-text)))
      (bind canvas "<Right>"
            (lambda (evt) 
              (declare (ignore evt)) 
              (clear-selection)
              (move-right canvas input-text)))
      (bind canvas "<Up>"
            (lambda (evt) 
              (declare (ignore evt)) 
              (clear-selection)
              (move-up canvas input-text)))
      (bind canvas "<Down>"
            (lambda (evt) 
              (declare (ignore evt)) 
              (clear-selection)
              (move-down canvas input-text)))

      ;; Shift+矢印キー（範囲選択）
      (bind canvas "<Shift-Left>"
            (lambda (evt)
              (declare (ignore evt))
              ;; 選択開始していなければ現在位置から開始
              (unless (has-selection-p)
                (setf *sel-start-x* *cur-x*
                      *sel-start-y* *cur-y*
                      *sel-end-x* *cur-x*
                      *sel-end-y* *cur-y*))
              ;; カーソルを移動し、選択終端も移動
              (when (> *cur-x* 0)
                (decf *cur-x*)
                (setf *sel-end-x* *cur-x*
                      *sel-end-y* *cur-y*))
              (redraw canvas)
              (update-text-input input-text)))
      (bind canvas "<Shift-Right>"
            (lambda (evt)
              (declare (ignore evt))
              (unless (has-selection-p)
                (setf *sel-start-x* *cur-x*
                      *sel-start-y* *cur-y*
                      *sel-end-x* *cur-x*
                      *sel-end-y* *cur-y*))
              (when (< *cur-x* (1- *cols*))
                (incf *cur-x*)
                (setf *sel-end-x* *cur-x*
                      *sel-end-y* *cur-y*))
              (redraw canvas)
              (update-text-input input-text)))
      (bind canvas "<Shift-Up>"
            (lambda (evt)
              (declare (ignore evt))
              (unless (has-selection-p)
                (setf *sel-start-x* *cur-x*
                      *sel-start-y* *cur-y*
                      *sel-end-x* *cur-x*
                      *sel-end-y* *cur-y*))
              (when (> *cur-y* 0)
                (decf *cur-y*)
                (setf *sel-end-x* *cur-x*
                      *sel-end-y* *cur-y*))
              (redraw canvas)
              (update-text-input input-text)))
      (bind canvas "<Shift-Down>"
            (lambda (evt)
              (declare (ignore evt))
              (unless (has-selection-p)
                (setf *sel-start-x* *cur-x*
                      *sel-start-y* *cur-y*
                      *sel-end-x* *cur-x*
                      *sel-end-y* *cur-y*))
              (when (< *cur-y* (1- *rows*))
                (incf *cur-y*)
                (setf *sel-end-x* *cur-x*
                      *sel-end-y* *cur-y*))
              (redraw canvas)
              (update-text-input input-text)))

      ;; Canvas上でEnter → Textにフォーカス
      (bind canvas "<Return>"
            (lambda (evt)
              (declare (ignore evt))
              (focus input-text)))
      
      ;; F2 → 編集モード（既存内容を編集）
      (bind canvas "<F2>"
            (lambda (evt)
              (declare (ignore evt))
              (focus input-text)
              ;; カーソルを末尾に移動
              (format-wish "~a mark set insert end" (widget-path input-text))))
      
      ;; Canvas上でキー入力 → 入力欄をクリアしてフォーカス移動（Tclレベルで処理）
      (let ((canvas-path (widget-path canvas))
            (text-path (widget-path input-text)))
        ;; 印刷可能文字のみを処理するTclバインディング
        (format-wish "bind ~a <Key> {
          set k %K
          set c %A
          # 特殊キーを除外
          if {$c ne {} && [string length $c] == 1 && [string is print $c]} {
            # 制御キーやファンクションキーでない場合
            if {![string match *Control* $k] && 
                ![string match *Shift* $k] && 
                ![string match *Alt* $k] && 
                ![string match *Meta* $k] &&
                ![string match *Super* $k] &&
                ![string match F?* $k] &&
                $k ne {Up} && $k ne {Down} && $k ne {Left} && $k ne {Right} &&
                $k ne {Return} && $k ne {Escape} && $k ne {Tab} &&
                $k ne {Delete} && $k ne {BackSpace} &&
                $k ne {Home} && $k ne {End} && $k ne {Prior} && $k ne {Next} &&
                $k ne {Insert} && $k ne {Caps_Lock} && $k ne {Num_Lock}} {
              # テキストをクリアして文字を挿入
              ~a delete 1.0 end
              ~a insert end $c
              focus ~a
            }
          }
        }" canvas-path text-path text-path text-path))

      (focus canvas))))

;;; ロード時メッセージ
(format t "~%=== Spreadsheet GUI v0.5.2 ===~%")
(format t "Lisp Powered Edition~%~%")
(format t "起動: (ss-gui:start)~%")
(format t "~%基本操作:~%")
(format t "  矢印キー          : セル移動~%")
(format t "  Shift+矢印        : 範囲選択~%")
(format t "  ドラッグ          : 範囲選択~%")
(format t "  文字キー          : 直接入力開始~%")
(format t "  F2                : 編集モード~%")
(format t "  Ctrl+X/C/V        : 切り取り/コピー/ペースト~%")
(format t "  Ctrl+Z/Y          : 元に戻す/やり直し~%")
(format t "  Delete/BackSpace  : セル消去~%")
(format t "  Escape            : 編集キャンセル~%")
(format t "~%入力確定:~%")
(format t "  Enter             : 確定→下に移動~%")
(format t "  Ctrl+Enter        : 確定→そのまま~%")
(format t "  Shift+Enter       : 入力欄で改行~%")
(format t "  Alt+Enter         : 確定→右に移動~%")
(format t "  Shift+Alt+Enter   : 確定→左に移動~%")
(format t "~%ファイル操作:~%")
(format t "  Ctrl+N            : 新規作成~%")
(format t "  Ctrl+O            : ファイルを開く~%")
(format t "  Ctrl+S            : 保存~%")
(format t "~%v0.5.2 新機能: 関数150以上、sort :key対応、シンタックスハイライト、S式整形~%")
