;;;; =====================================================
;;;; spreadsheet-gui.lisp
;;;; Common Lisp + LTK で作るシンプルな表計算ソフト
;;;; 
;;;; Version: 0.3.1
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
  (:export :start :show-dependencies :show-cell-deps))
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
;;;; コピー＆ペースト
;;;; =========================

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
          (pasted-cells nil))
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
                (push name pasted-cells))))
          (incf idx)))
      ;; 貼り付けたセルの依存元を再計算
      (dolist (name (nreverse pasted-cells))
        (recalc-dependents name)))))

(defun clear-selection-cells ()
  "選択範囲のセルをクリア（NILを設定）し、依存元を再計算"
  (let ((cleared-cells nil))
    (if (has-selection-p)
        ;; 範囲選択がある場合
        (multiple-value-bind (min-x min-y max-x max-y) (selection-bounds)
          (loop for y from min-y to max-y do
            (loop for x from min-x to max-x do
              (let* ((name (cell-name x y))
                     (cell (get-cell name)))
                (setf (cell-value cell) nil
                      (cell-formula cell) nil)
                (update-dependencies name nil)
                (push name cleared-cells)))))
        ;; 範囲選択がない場合はカーソル位置のみ
        (let* ((name (cell-name *cur-x* *cur-y*))
               (cell (get-cell name)))
          (setf (cell-value cell) nil
                (cell-formula cell) nil)
          (update-dependencies name nil)
          (push name cleared-cells)))
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
  (cond
    ;; 数式があればそれを優先
    (formula (format nil "=~S" formula))
    ;; NILは空文字列
    ((null val) "")
    ;; 文字列はそのまま
    ((stringp val) val)
    ;; その他はprinc形式
    (t (princ-to-string val))))

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
           (let* ((form (read-from-string (subseq trimmed 1)))
                  (value (eval-formula form)))
             (values value form))
         (error () (values trimmed nil))))
      ;; 数値を試す
      (t (let ((num (ignore-errors (read-from-string trimmed))))
           (if (numberp num)
               num
               trimmed))))))

(defun paste-from-system-clipboard ()
  "システムクリップボードから貼り付け"
  (let ((text (get-system-clipboard))
        (pasted-cells nil))
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
              ;; 評価位置を設定
              (setf *eval-col* *cur-x* *eval-row* *cur-y*)
              (multiple-value-bind (val form) 
                  (parse-clipboard-value (first lines))
                (setf (cell-value cell) val
                      (cell-formula cell) form)
                (if form
                    (update-dependencies name (extract-references form *cur-y* *cur-x*))
                    (update-dependencies name nil))
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
                      ;; 評価位置を設定
                      (setf *eval-col* x *eval-row* y)
                      (multiple-value-bind (val form)
                          (parse-clipboard-value col-text)
                        (setf (cell-value cell) val
                              (cell-formula cell) form)
                        (if form
                            (update-dependencies name (extract-references form y x))
                            (update-dependencies name nil))
                        (push name pasted-cells)))))))))
      ;; 貼り付けたセルの依存元を再計算
      (dolist (name (nreverse pasted-cells))
        (recalc-dependents name)))))

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
    (when refs
      (format t "  Extract refs from ~S: ~a~%" formula refs))
    refs))

(defun update-dependencies (cell-name new-refs)
  "セルの依存関係を更新"
  (when new-refs
    (format t "  Update deps: ~a refs ~a~%" cell-name new-refs))
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
              (let ((new-val (eval-formula formula)))
                (format t "  Recalc ~a: ~a → ~a~%" cell-name (cell-value cell) new-val)
                (setf (cell-value cell) new-val))
            (error (e)
              (setf (cell-value cell) (format nil "ERR: ~a" e)))))))))

(defun parse-cell-name (name)
  "セル名から座標(col row)を取得"
  (let* ((col (- (char-code (char-upcase (char name 0))) (char-code #\A)))
         (row (1- (parse-integer (subseq name 1)))))
    (list col row)))

(defun recalc-dependents (cell-name)
  "セルに依存する全てのセルを再計算"
  (let* ((deps (collect-all-dependents cell-name))
         (sorted (topological-sort-cells deps)))
    (when deps
      (format t "~%Recalc triggered by ~a~%" cell-name)
      (format t "  Dependents: ~a~%" deps)
      (format t "  Sorted: ~a~%" sorted))
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
    ((string-equal op-name "COUNT")
     (lambda (&rest args)
       (length (flatten args))))
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

(defun draw-cell-background (canvas x y val selected in-selection)
  "セルの背景のみを描画"
  (let* ((px (+ *header-w* (* x *cell-w*)))
         (py (+ *header-h* (* y *cell-h*)))
         (px2 (+ px *cell-w*))
         (py2 (+ py *cell-h*))
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
         (cell (current-cell))
         (cell-nm (cell-name *cur-x* *cur-y*)))
    ;; 評価位置を設定
    (setf *eval-col* *cur-x*
          *eval-row* *cur-y*)
    (handler-case
        (if (and (> (length raw) 0)
                 (char= (char raw 0) #\=))
            ;; =で始まる → 数式として解析・評価
            (let* ((form (read-from-string (subseq raw 1)))
                   (*eval-stack* (list cell-nm))  ; 循環参照検出用
                   (value (eval-formula form))
                   (refs (extract-references form *cur-y* *cur-x*)))
              (setf (cell-formula cell) form
                    (cell-value cell) value)
              ;; 依存関係を更新
              (update-dependencies cell-nm refs))
            ;; それ以外 → 通常の値として解釈
            (let ((parsed (ignore-errors (read-from-string raw))))
              (setf (cell-formula cell) nil
                    (cell-value cell) (or parsed raw))
              ;; 数式がないので依存関係をクリア
              (update-dependencies cell-nm nil)))
      (error (e)
        (setf (cell-value cell) (format nil "ERR: ~a" e))))
    ;; 依存元を再計算
    (recalc-dependents cell-nm)
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
  (clear-selection)
  (setf *clipboard* nil)
  (clear-dependencies)
  
  (with-ltk ()
    (wm-title *tk* (format nil "Spreadsheet v0.3.1 [~Dx~D]" *cols* *rows*))
    
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

      ;; セルクリック → 選択開始
      (bind canvas "<ButtonPress-1>"
            (lambda (evt)
              (let ((mx (and evt (slot-value evt 'ltk::x)))
                    (my (and evt (slot-value evt 'ltk::y))))
                (when (and mx my (numberp mx) (numberp my)
                           (> mx *header-w*) (> my *header-h*))
                  (let ((x (clamp (floor (/ (- mx *header-w*) *cell-w*)) 0 (1- *cols*)))
                        (y (clamp (floor (/ (- my *header-h*) *cell-h*)) 0 (1- *rows*))))
                    (setf *cur-x* x
                          *cur-y* y
                          *sel-start-x* x
                          *sel-start-y* y
                          *sel-end-x* x
                          *sel-end-y* y
                          *selecting* t)
                    (update-text-input input-text)
                    (redraw canvas))))
              (focus canvas)))

      ;; ドラッグ → 選択範囲拡張
      (bind canvas "<B1-Motion>"
            (lambda (evt)
              (when *selecting*
                (let ((mx (and evt (slot-value evt 'ltk::x)))
                      (my (and evt (slot-value evt 'ltk::y))))
                  (when (and mx my (numberp mx) (numberp my))
                    (setf *sel-end-x* (clamp (floor (/ (- mx *header-w*) *cell-w*)) 0 (1- *cols*))
                          *sel-end-y* (clamp (floor (/ (- my *header-h*) *cell-h*)) 0 (1- *rows*)))
                    (redraw canvas))))))

      ;; ドラッグ終了
      (bind canvas "<ButtonRelease-1>"
            (lambda (evt)
              (declare (ignore evt))
              (setf *selecting* nil)))

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

      (focus canvas))))

;;; ロード時メッセージ
(format t "~%=== Spreadsheet GUI v0.3.1 ===~%")
(format t "Lisp Powered Edition (Auto-Recalc)~%~%")
(format t "起動: (ss-gui:start)~%")
(format t "~%基本操作:~%")
(format t "  矢印キー          : セル移動~%")
(format t "  Shift+矢印        : 範囲選択~%")
(format t "  ドラッグ          : 範囲選択~%")
(format t "  Ctrl+C/V          : コピー/ペースト~%")
(format t "  Delete/BackSpace  : セル消去~%")
(format t "  Enter             : 入力確定（自動再計算）~%")
(format t "~%デバッグ:~%")
(format t "  (ss-gui:show-dependencies)  ; 依存関係を表示~%")
(format t "  (ss-gui:show-cell-deps \"A1\") ; セルの依存関係~%")
(format t "~%数式例:~%")
(format t "  A1に10を入力、A2に =(+ A1 5) を入力~%")
(format t "  → A1を変更するとA2が自動更新~%")
(format t "~%v0.3.1 新機能: 自動再計算、範囲選択、クリップボード~%")
