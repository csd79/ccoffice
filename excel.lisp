;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Workbooks, worksheets


(defconstant +xl-calculation-automatic+ -4105)
(defconstant +xl-calculation-manual+    -4135)

(defmacro excellerate ((excel) &body body)
  "Execute BODY using faster Excel interaction."
  `(let ((su (?screenupdating ,excel))
         (da (?displayalerts  ,excel))
         (c  (?calculation    ,excel))
         (ee (?enableevents   ,excel)))
     (unwind-protect 
         (progn
           (setf (?screenupdating ,excel) nil
                 (?displayalerts  ,excel) nil
                 (?calculation    ,excel) +xl-calculation-manual+
                 (?enableevents   ,excel) nil)
           ,@body)
       (progn
         (setf (?calculation    ,excel) c
               (?displayalerts  ,excel) da
               (?screenupdating ,excel) su
               (?enableevents   ,excel) ee)))))

(defun column->row (array &key (element-type t) (getter #'identity))
  "Transpose a column (an (n 0) array) into a vector."
  (let* ((height (array-dimension array 0))
         (result (make-array height :element-type element-type)))
    (loop for i from 0 below height doing
          (setf (svref result i)
                (funcall getter
                         (aref array i 0))))
    result))

(defun excel-value-as-number (value) ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
  "Try to interpret VALUE as a number."
  (let ((not-string (if (stringp value)
                      (read-from-string value)
                      value)))
    (assert (numberp not-string)
        (not-string)
      "Value ~a is not readable as number." not-string)
    not-string))

(defun parse-string (value &key (numproc #'identity) (ignore-error t))
  "Turn VALUE (a string or a number) into a string."
  (typecase value
    (string value)
    (number (format nil "~d" (funcall numproc value)))
    (t      (unless ignore-error (error "Value should be of type NUMBER or STRING.")))))

(defun parse-number (value &key (strproc #'read-from-string) (ignore-error t))
  "Turn VALUE (a string or a number) into a number."
  (let ((parsed (typecase value
                  (string (unless (string= value "") (funcall strproc value)))
                  (number value)
                  (t      (unless ignore-error (error "Value should be of type NUMBER or STRING."))))))
    (when (numberp parsed)
      parsed)))


(defun freeze-panes (wsheet &key (after-column nil) (after-row nil))
  "Freeze panes after given column and/or row. When no parameter is supplied, unfreeze panes."
  (cclet* ((app (?application wsheet))
           (win (?activewindow app)))
    (if (or after-column after-row)
      (progn
        (setf (?freezepanes win) nil)
        (when after-column
          (setf (?splitcolumn win) after-column))
        (when after-row
          (setf (?splitrow win) after-row))
        (setf (?freezepanes win) t))
      (setf (?freezepanes win) nil))))


;;; ----------------------------------------------------------------------
;;; Ranges, values


(defconstant +xl-to-left+          -4159)
(defconstant +xl-up+               -4162)
(defconstant +xl-to-right+         -4161)
(defconstant +xl-down+             -4121)
(defconstant +xl-values+           -4163)
(defconstant +xl-by-rows+              1)
(defconstant +xl-by-columns+           2)
(defconstant +xl-whole+                1)
(defconstant +xl-previous+             2)

(defun range (worksheet &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  "Create range object within given 'coordinates'."
  (let ((upper-left (cell-index x1 y1)))
    (if (and x2 y2)
      (?range worksheet upper-left (cell-index x2 y2))
      (?range worksheet upper-left))))

(defmacro with-used-range ((worksheet left top right bottom) &body body)
  "Edges of the used range of WORKSHEET."
  (let ((whole   (gensym))
        (from    (gensym))
        (endup   (gensym))
        (endleft (gensym))
        (findbtm (gensym))
        (findrgt (gensym)))
    `(let ((,left nil) (,top nil) (,right nil) (,bottom nil))
       (cclet* ((,whole (?cells ,worksheet)))
         (cclet* ((,endup (?end ,whole +xl-up+)))
           (setf ,top (?row ,endup)))
         (cclet* ((,endleft (?end ,whole +xl-to-left+)))
           (setf ,left (?column ,endleft)))
         (cclet* ((,from (range ,worksheet ,left ,top)))
           (cclet* ((,findbtm (!find ,whole "*" ,from +xl-values+
                                      +xl-whole+ +xl-by-rows+ +xl-previous+)))
             (setf ,bottom (?row ,findbtm)))
           (cclet* ((,findrgt (!find ,whole "*" ,from +xl-values+
                                      +xl-whole+ +xl-by-columns+
                                      +xl-previous+)))
             (setf ,right (?column ,findrgt)))))
       ,@body)))

(defun used-range (worksheet)
  "The range of all content within WORKSHEET."
  (with-used-range (worksheet left top right bottom)
    (range worksheet left top right bottom)))

(defun last-row (worksheet)
  "The number of last row (& last column) of WORKSHEET."
  (with-used-range (worksheet left top right bottom)
    (values bottom right)))

(defmacro with-range ((range left top right bottom) &body body)
  "Edges of RANGE."
  (let ((cols (gensym))
        (rows (gensym)))
  `(let ((,left   (?column ,range))
         (,top    (?row ,range))
         (,right  nil)
         (,bottom nil))
     (cclet* ((,cols (?columns ,range)))
       (setf ,right (1- (+ ,left (?count ,cols)))))
     (cclet* ((,rows (?rows ,range)))
       (setf ,bottom (1- (+ ,top (?count ,rows)))))
     ,@body)))

(defun delete-rows (wsheet start end)
  "Delete rows between START and END."
  (!delete (?entirerow (range wsheet 1 start 1 end))))

(defun delete-columns (wsheet start end)
  "Delete columns between START and END."
  (!delete (?entirecolumn (range wsheet start 1 end 1))))

(defun empty-cell-p (value)
  "Does VALUE indicate and empty cell?"
  (or (and (stringp value)
           (string= value ""))
      (and (symbolp value)
           (member value '(empty :empty)))
      (null value)))


;;; ----------------------------------------------------------------------
;;; Cell formatting


(defconstant +xl-paste-formats+    -4122)
(defconstant +xl-paste-comments+   -4144)
(defconstant +xl-paste-validation+     6)
(defconstant +xl-align-center+     -4108)
(defconstant +rgb-black+               0)

(defun copy-formatting (src dst)
  "Copy cell formatting between ranges."
  (!copy src)
  (!pastespecial dst +xl-paste-formats+)
  (!pastespecial dst +xl-paste-comments+)
  (!pastespecial dst +xl-paste-validation+))

(defun autofit-cols (worksheet)
  "Set auto width for all columns."
  (cclet* ((range (?columns worksheet)))
    (!autofit range)))

(defun font (range &optional first last)
  "Target range for font formatting."
  (?font (if (and first last)
            (?characters range first last)
            range)))

(defparameter *style-elements*
  `(:wrap-text     (:prop   "WrapText"            ,t)
    :halign-center (:prop   "HorizontalAlignment" ,+xl-align-center+)
    :valign-center (:prop   "VerticalAlignment"   ,+xl-align-center+)
    :border        (:method "BorderAround"        ,:null ,:null ,:null ,+rgb-black+)
    :bold          (:prop   "Bold"                ,t))
  "Style definitions for APPLY-STYLE.")

(defun apply-style (range style)
  "Apply any style defined in *STYLE-ELEMENTS*."
  (dolist (style-element style)
    (destructuring-bind (type name &rest values)
        (getf *style-elements* style-element)
      (case type
        (:prop   (setf (com::invoke-dispatch-get-property range name)
                       (first values)))
        (:method (apply #'com::invoke-dispatch-method range name values))))))


;;; ----------------------------------------------------------------------
;;; Opening & closing


(defun close-workbook (workbook)
  "Close WORKBOOK & quit application when there are no more workbooks open."
  (cclet* ((application (?application workbook))
           (workbooks   (?workbooks application)))
    (!close workbook)
    (when (zerop (?count workbooks))
      (setf (?displayalerts application) nil)
      (!quit application))))

(defun window-visible (workbook &optional (window 1) (visible t))
  "Change window visibility."
  (cclet* ((windows (?windows workbook))
           (window  (?item windows window)))
    (setf (?visible window) visible)))

(defmacro with-workbook ((&key (app       nil)
                               (wbook     'wbook)
                               (use       nil)
                               (open      nil)
                               (wsvars    '())
                               (read-only nil)
                               (save      nil)
                               (close     t))
                         &body body)
  "Create bindings for an Excel workbook with optional method calls (open, save at the end, close)."
  (let ((app-obj (gensym))
        (wbooks  (gensym))
        (doc-obj (gensym))
        (wsheets (gensym)))
    `(cclet* ((,app-obj (or ,app 
                            (cclet* ((global (com:create-object :progid "Excel.Application")))
                              (?application global))))
              (,wbooks  (?workbooks ,app-obj))
              (,doc-obj (or (when (typep ,use 'com::com-interface) ,use)
                            (when ,open (!open ,wbooks ,open nil ,read-only))
                            (!add ,wbooks)))
              (,wbook   ,doc-obj)
              (,wsheets (?worksheets ,doc-obj))
              ,@(loop for n from 1
                      for var in wsvars
                      when var collect
                      (list var `(?item ,wsheets ,n))))
       (unwind-protect
           (excellerate (,app-obj)
             (when (eq ,use ,doc-obj)
               (print "·hh·····")
               (com:add-ref ,doc-obj))
             (window-visible ,doc-obj 1 t)
             ,@body)
         (progn 
           (when (and ,save (not ,read-only))
             (!save ,doc-obj))
           (when ,close
             (setf (?saved ,doc-obj) t)
             (setf (?displayalerts ,app-obj) nil)
             (!close ,doc-obj)
             (unless ,app
               (!quit ,app-obj))))))))


;;; ----------------------------------------------------------------------
;;; Cell reference


(defparameter *header-row* 1)

(defun title-column (worksheet title &key (key #'identity) (header-row *header-row*) (test #'equalp)
                         (from-end nil) (start nil) (end nil))
  "Find the column that has TITLE in its header."
  (with-used-range (worksheet left top right bottom)
    (let ((header (?value2 (range worksheet left header-row right header-row))))
      (if (arrayp header)
        (if from-end
          (loop for i from (1- (or end right)) downto (1- (or start left))
                for v = (funcall key (aref header (1- header-row) i))
                thereis (and (funcall test title v) (1+ i)))
          (loop for i from (1- (or start left)) upto (1- (or end right))
                for v = (funcall key (aref header (1- header-row) i))
                thereis (and (funcall test title v) (1+ i))))
        (and (funcall test title header) left)))))

(defun column-designator-p (designator)
  "Column designator type and resolution."
  (or (and (integerp designator)
           (<= 1 designator 16384))
      (and (keywordp designator)
           (<= (length (symbol-name designator)) 3))
      (stringp designator)))

(deftype column-designator ()
  "Integer, keyword (:a) or string (header in row 1)."
  '(satisfies column-designator-p))

(defun resolve-column-designator (designator &optional worksheet)
  "Convert column designator to 1-base column index."
  (assert (typep designator 'column-designator)
      (designator)
    "Invalid column designator ~a - should be az integer, a keyword or a string" designator)
  (typecase designator
    (integer designator)
    (keyword (letters-column designator))
    (string  (title-column worksheet designator))))


(defun locate-row (worksheet column-designator value function)
  "Find row in WORKSHEET where designated column equals VALUE according to FUNCTION."
  (let* ((column-idx (resolve-column-designator column-designator worksheet))
         (column     (column->row (?value2 (range worksheet column-idx 2 column-idx (last-row worksheet))))))
    (loop for row from 0 below (length column)
          for v = (svref column row)
          thereis (and (funcall function v value)
                       (+ row 2)))))

(defun row-designator-p (designator)
  "Row designator type and resolution"
  (or (and (integerp designator)
           (<= 1 designator 1048576))
      (and (listp designator)
           (typep (first designator) 'column-designator)
           (if (third designator)
             (functionp (third designator))
             t))))

(deftype row-designator ()
  "Integer or a list of parameters for LOCATE-ROW."
  '(satisfies row-designator-p))

(defun resolve-row-designator (designator &optional worksheet)
  "Convert row designator to 1-base row index."
  (assert (typep designator 'row-designator)
      (designator)
    "Invalid row designator ~a - should be az integer or a list of a column designator, a value and a predicate function" designator)
  (typecase designator
    (integer designator)
    (list    (destructuring-bind (column-designator value &optional (function #'equalp))
                 designator
               (locate-row worksheet column-designator value function)))))

(defun xcell-core (worksheet column row &key (value nil value-supplied-p) (prop "value2"))
  "Main body of XCELL and SET-XCELL."
  (let ((row-final    (resolve-row-designator row worksheet))
        (column-final (resolve-column-designator column worksheet)))
    (if value-supplied-p
      (setf (com:invoke-dispatch-get-property (range worksheet column-final row-final) prop) value)
      (com:invoke-dispatch-get-property (range worksheet column-final row-final) prop))))
#|      (setf (funcall prop (range worksheet column-final row-final)) value)
      (funcall prop (range worksheet column-final row-final)))))|#

(defun xcell (worksheet column row &key (prop "value2"))
  "Simplified cell reference."
  (xcell-core worksheet column row :prop prop))

(defun (setf xcell) (value worksheet column row &key (prop "value2"))
  "Simplified cell assignment."
  (xcell-core worksheet column row :value value :prop prop))


;;; ----------------------------------------------------------------------
;;; Range reference


(defun 2d-array-column-p (value)
  "Single-column 2D array predicate."
  (and (arrayp value)
       (= (length (array-dimensions value)) 2)
       (= (array-dimension value 1) 1)))

(deftype 2d-array-column ()
  "Single-columns 2D array type."
  '(satisfies 2d-array-column-p))

(defun 2d-array-row-p (value)
  "Single-row 2D array predicate."
  (and (arrayp value)
       (= (length (array-dimensions value)) 2)
       (= (array-dimension value 0) 1)))

(deftype 2d-array-row ()
  "Single-row 2D array type."
  '(satisfies 2d-array-row-p))

(defun 2d-array-p (value)
  "Multi-column, multi-row 2D array predicate."
  (and (arrayp value)
       (= (length (array-dimensions value)) 2)
       (> (array-dimension value 0) 1)
       (> (array-dimension value 1) 1)))

(deftype 2d-array ()
  "Multi-column, multi-row 2D array type."
  '(satisfies 2d-array-p))

(defparameter *xrange-default-value*    "" "Used when *XRANGE-TARGET-TOO-LARGE* = :FULFILL.")
(defparameter *xrange-target-too-small* :fill-to-limit "Possible values: :ERROR, :FILL-TO-LIMIT")
(defparameter *xrange-target-too-large* :restrict "Possible values: :ERROR, :FULFILL, :RESTRICT")

(defun set-xrange-worker (worksheet c1 r1 c2 r2 value prop)
  "Multiple cell range assignment."
  (let ((range-width  (1+ (abs (- c1 c2))))
        (range-height (1+ (abs (- r1 r2))))
        (value-width  (array-dimension value 1))
        (value-height (array-dimension value 0)))
    ;; If target range size and value size is not equal and either control variables are set to :ERROR
    (when (and (or (/= range-width  value-width)
                   (/= range-height value-height))
               (or (eq *xrange-target-too-small* :error)
                   (eq *xrange-target-too-large* :error)))
      (error "Cannot fit column ~a into range ~a,~a - ~a,~a." value c1 r1 c2 r2))
    ;; Resize target area according to VALUE size.
    (let ((c2r (if (= value-width 1)
                 c2
                 (min c2 (1- (+ c1 value-width)))))
          (r2r (if (= value-height 1)
                 r2
                 (min r2 (1- (+ r1 value-height))))))
      ;; Assign VALUE to a downsized target range.
      (setf (com:invoke-dispatch-get-property (range worksheet c1 r1 c2r r2r) prop) value)
;      (setf (funcall prop (range worksheet c1 r1 c2r r2r)) value)
      ;; Add fill when needed
      (when (eq *xrange-target-too-large* :fulfill)
        (let ((rightp (and (not (2d-array-column-p value))
                           (> range-width value-width)))
              (lowerp (and (not (2d-array-row-p value))
                           (> range-height value-height))))
          ;; Right side fill
          (when rightp
            (setf (com:invoke-dispatch-get-property (range worksheet (1+ c2r) r1 c2 r2r) prop)
;            (setf (funcall prop (range worksheet (1+ c2r) r1 c2 r2r))
                  *xrange-default-value*))
          ;; Lower fill
          (when lowerp
            (setf (com:invoke-dispatch-get-property (range worksheet c1 (1+ r2r) c2r r2) prop)
;            (setf (funcall prop (range worksheet c1 (1+ r2r) c2r r2))
                  *xrange-default-value*))
          ;; Lower right fill
          (when (and rightp lowerp)
            (setf (com:invoke-dispatch-get-property (range worksheet (1+ c2r) (1+ r2r) c2 r2) prop)
;            (setf (funcall prop (range worksheet (1+ c2r) (1+ r2r) c2 r2))
                  *xrange-default-value*)))))))

(defun precide-xrange (worksheet c1 r1 c2 r2 value prop)
  "Prefilter arg combinations that yield simple assignment or no assignment."
  ;; Prevent invalid keywords in control vars.
  (when (not (member *xrange-target-too-small* '(:error :fill-to-limit)))
    (error "Unknown keyword: ~a." *xrange-target-too-small*))
  (when (not (member *xrange-target-too-large* '(:error :fulfill :restrict)))
    (error "Unknown keyword: ~a." *xrange-target-too-large*))
  (cond
   ;; Target is a single cell.
   ((and (= c1 c2)
         (= r1 r2))
    (if (atom value)
      (setf (xcell worksheet c1 r1) value)
      (error "Cannot assign ~a to a single cell." value)))
   ;; Value is an atom, will be assigned to every cell in range.
   ((and (atom value)
         (not (arrayp value)))
;    (setf (funcall prop (range worksheet c1 r1 c2 r2)) value))
    (setf (com:invoke-dispatch-get-property (range worksheet c1 r1 c2 r2) prop) value))
   ;; Value is a list or a vector: convert value to a 2D array row.
   ((or (listp   value)
        (vectorp value))
    (precide-xrange
     worksheet c1 r1 c2 r2
     (make-array (list 1 (length value))
                 :initial-contents (list (coerce value 'list)))
     prop))
   ;; Range is a column but value is not: error.
   ((and (= c1 c2)
         (not (2d-array-column-p value)))
    (error "Cannot assign ~a to a column of cells." value))
   ;; Range is a row but value is not: error.
   ((and (= r1 r2)
         (or (2d-array-column-p value) (2d-array-p value)))
    (error "Cannot assign ~a to a row of cells." value))
   ;; Otherwise, set range to values in collection.
   (t (set-xrange-worker worksheet c1 r1 c2 r2 value prop))))

(defun xrange-core (worksheet c1 r1 c2 r2 &key (value nil value-supplied-p) (prop "value2"))
  "Main body of XRANGE and SET-XRANGE."
  (let ((c1final (resolve-column-designator c1 worksheet))
        (c2final (resolve-column-designator c2 worksheet))
        (r1final (resolve-row-designator r1 worksheet))
        (r2final (resolve-row-designator r2 worksheet)))
    (if value-supplied-p
      (precide-xrange worksheet c1final r1final c2final r2final value prop)
      (com:invoke-dispatch-get-property (range worksheet c1final r1final c2final r2final) prop))))

(defun xrange (worksheet c1 r1 c2 r2 &key (prop "value2"))
  "Simplified cell range reference."
  (xrange-core worksheet c1 r1 c2 r2 :prop prop))

(defun (setf xrange) (value worksheet c1 r1 c2 r2 &key (prop "value2"))
  "Simplified cell range assingment."
  (xrange-core worksheet c1 r1 c2 r2 :value value :prop prop))


;;; ----------------------------------------------------------------------
;;; Filtering


(defconstant +xl-celltype-visible+   12)
(defconstant +xl-paste-values+    -4163)
(defconstant +xl-paste-all+       -4104)
(defconstant +xl-paste-column-widths+ 8)
(defconstant +and+                    1)
(defconstant +or+                     2)

;;; (xselect> ws `(("TK" "Eger" +or+ "Baja") ("SZTSZ" ">5" +and+ "<999999")))
(defun xselect> (worksheet subscripts &key (paste-all nil) (paste-column-widths nil))
  "Select rows from WORKSHEET using AutoFilter, copy the results into a new temporary workbook."
  (cclet* ((cellidx (cell-index 1 1))
           (range   (used-range worksheet))
           (topleft (range worksheet 1 1))
           (excel   (?application worksheet))
           (wbooks  (?workbooks excel))
           (rbook   (!add wbooks))
           (rsheets (?worksheets rbook))
           (rsheet  (?item rsheets 1))
           (rrange  (?range rsheet cellidx))
           (corsubs (mapcar #'(lambda (sub)
                                (append (list (resolve-column-designator (first sub) worksheet))
                                        (rest sub)))
                            subscripts)))
    ;; Apply filters.
    (dolist (subscript corsubs)
      (apply #'com::invoke-dispatch-method topleft "autofilter" subscript))
    ;; Copy selected data to new worksheet
    (!copy (!specialcells range +xl-celltype-visible+))
    (when paste-column-widths
      (!pastespecial rrange +xl-paste-column-widths+))
    (if paste-all
      (!pastespecial rrange +xl-paste-all+)
      (!pastespecial rrange +xl-paste-values+))
    ;; Turn autofilter off
    (!autofilter range)
    ;; Return worksheet containing results
    (setf (?saved (?parent worksheet)) t
          (?saved rbook) t)
    rsheet))

(defmacro with-xselection ((selected source subscripts &key (paste-all nil) (paste-column-widths nil)) &body body)
  "Select rows from SOURCE into SELECTED according to SUBSCRIPTS."
  (let ((wbook (gensym)))
    `(cclet* ((,selected (xselect> ,source ,subscripts :paste-all ,paste-all
                                   :paste-column-widths ,paste-column-widths)))
       (unwind-protect
           (progn
             ,@body)
         (cclet* ((,wbook (!parent ,selected)))
           (close-workbook ,wbook))))))


;;; ----------------------------------------------------------------------
;;; Dates


(defun excel-date (n)
  "Convert Excel date value to a list of year, month and day."
  (let* ((a    (+ n 2483588))
         (b    (truncate (/ (* a 4) 146097)))
         (c    (- a (truncate (/ (+ (* 146097 b) 3) 4))))
         (d    (truncate (/ (* 4000 (+ c 1)) 1461001)))
         (e    (+ (- c (+ (truncate (/ (* 1461 d) 4)))) 31))
         (f    (truncate (/ (* 80 e) 2447)))
         (day  (round (- e (truncate (/ (* 2447 f) 80)))))
         (g    (truncate (/ f 11)))
         (mon  (- (+ f 2) (* 12 g)))
         (year (+ (* 100 (- b 49)) d g)))
    (list year mon day)))

(defun excel-date-string (n &key (words nil))
  "Convert Excel date value to a hu formated string."
  (destructuring-bind (year mon day)
      (excel-date (truncate n))
    (if words
      (let ((months '("janu·r" "febru·r" "m·rcius" "·prilis" "m·jus" "j˙nius" "j˙lius"
                      "augusztus" "szeptember" "oktÛber" "november" "december")))
        (format nil "~4d. ~a ~d." year (nth (1- mon) months) day))
    (format nil "~4d.~2,'0d.~2,'0d." year mon day))))


;;; ----------------------------------------------------------------------
;;; Worksheet as xarray


(defclass xarray ()
  ((head  :initarg  :head
          :accessor head)
   (body  :initarg  :body
          :accessor body)
   (index :initarg  :index
          :accessor index))
  (:documentation "Class for representing Excel worksheet as Lisp array(s)."))

(defun array-row->vector (array row)
  "Extract header into a simple vector."
  (let* ((width  (array-dimension array 1))
         (result (make-array width)))
    (loop for c from 0 below width doing
          (setf (aref result c)
                (aref array row c)))
    result))

(defun raw-index (body)
  "Create new raw index for BODY."
  (let ((height (array-dimension body 0)))
    (coerce (loop for i from 0 below height collecting i)
            'simple-vector)))

(defun read-xarray (range &key (accessor #'?value2))
  "Craete XARRAY obj containing values from RANGE."
  (with-range (range left top right bottom)
    (cclet* ((wsheet (?worksheet range))
             (headr  (range wsheet left top right top))
             (bodyr  (range wsheet left (1+ top) right bottom))
             (body   (funcall accessor bodyr))
             (index  (raw-index body)))
    (make-instance 'xarray
                   :head  (array-row->vector (funcall accessor headr) 0)
                   :body  body
                   :index index))))

(defun make-xarray (headers number-of-rows)
  "Create an empty xarray."
  (let ((body (make-array (list number-of-rows (length headers)))))
    (make-instance 'xarray
                   :head  (coerce headers 'simple-vector)
                   :body  body
                   :index (raw-index body))))

(defun rearrange (xarray)
  "Rearrange XARRAY into a new xarray according to its index."
  (with-slots ((head head) (body body) (index index))
      xarray
    (let* ((width    (length head))
           (new-body (make-array (list (length index) width))))
      (loop for idx across index
            for drow from 0 doing
            (loop for col from 0 below width doing
                  (setf (aref new-body drow col)
                        (aref body idx col))))
      (make-instance 'xarray
                     :head head
                     :body new-body
                     :index (raw-index new-body)))))

(defun write-xarray (xarray range &key (accessor "value2"))
  "Copy contents of XARRAY into Excel RANGE."
  (let ((rearranged (rearrange xarray)))
    (with-range (range left top right bottom)
      (cclet* ((wsheet (?worksheet range)))
        (setf
         (com:invoke-dispatch-get-property (range wsheet left top right top) accessor) (head rearranged)
         (com:invoke-dispatch-get-property (range wsheet left (1+ top) right bottom) accessor) (body rearranged))))))
#|        (setf (funcall accessor (range wsheet left top right top)) (head rearranged)
              (funcall accessor (range wsheet left (1+ top) right bottom)) (body rearranged))))))|#

(defmethod title-xacolumn ((obj xarray) title &key (test #'equalp))
  "Find xarray column by header title. In order to be able for columns to be identified by header, all headers must be non-numerical!"
  (with-slots ((head head))
      obj
    (loop for c from 0 below (length head)
          for v = (svref head c)
          thereis (and (funcall test title v) c))))

(defmethod resolve-xacolumn-designator ((obj xarray) designator)
  "Find xarray column by number, title or Excel letter."
  (assert (or (typep designator 'column-designator)
              (zerop designator))
      (designator)
    "Invalid column designator ~a - should be az integer, a keyword or a string" designator)
  (typecase designator
    (integer designator)
    (keyword (1- (letters-column designator)))
    (string  (title-xacolumn obj designator))))
  

(defmethod xaref* ((obj xarray) idx-row column-designator
                   &key (value nil value-supplied-p) (if-column-exceeds-limit :ignore))
  "General xarray cell accessor."
  (block xaref-body
    (with-slots ((body body) (index index)) obj
      (let ((column (resolve-xacolumn-designator obj column-designator))
            (row    (svref index idx-row)))
        (when column
          (when (>= column (array-dimension body 1))
            (case if-column-exceeds-limit
              (:error  (error "Column index ~a exceeds body width of xarray ~a." column obj))
              (:ignore (return-from xaref-body))))
          (if value-supplied-p
            (setf (aref body row column) value)
            (aref body row column)))))))

(defun xaref (xarray idx-row column-designator)
  "Generalized xarray access."
  (xaref* xarray idx-row column-designator))

(defun (setf xaref) (value xarray idx-row column-designator)
  "Generalized xarray assignment."
  (xaref* xarray idx-row column-designator :value value))

#|(defmethod xcref ((obj xarray) column-designator)
  "Column accessor for 1-row xarray."
  (xaref* obj 0 column-designator))|#

(defun xcref (xarray column-designator)
  "Column accessor for 1-row xarray."
  (xaref* xarray 0 column-designator))

(defun (setf xcref) (value xarray column-designator)
  "Column assignment  for 1-row xarray."
  (xaref* xarray 0 column-designator :value value))

(defmethod xarows ((obj xarray) (rows vector))
  "Create new xarray containing only selected rows from XARRAY."
  (with-slots ((index index)) obj
    (let ((new-index (loop for i across rows collecting
                           (svref index i))))
      (make-instance 'xarray
                     :head (head obj)
                     :body (body obj)
                     :index (coerce new-index 'simple-vector)))))

(defmethod xarows ((obj xarray) (row integer))
  "Create new xarray containing only selected row from XARRAY."
  (make-instance 'xarray
                 :head (head obj)
                 :body (body obj)
                 :index (make-array 1 :initial-contents (list (svref (index obj) row)))))

(defmacro do-xarows ((row row-number xarray) &body body)
  "Iterate over the indexed rows of XARRAY."
  (let ((height (gensym)))
    `(let ((,height (length (index ,xarray))))
       (loop for ,row-number from 0 below ,height
             for ,row = (xarows ,xarray ,row-number) doing
             ,@body))))

(defmethod xaselect ((obj xarray) selector-fn)
  "Return a copy of OBJ indexed according to SELECTOR-FN."
  (with-slots ((head head) (body body) (index index)) obj
    (let ((new-index '()))
      (do-xarows (row r obj)
        (when (funcall selector-fn row)
          (push (svref index r) new-index)))
      (make-instance 'xarray
                     :head  head
                     :body  body
                     :index (coerce (nreverse new-index) 'simple-vector)))))

(defmethod xauniques ((obj xarray) column-designator &key (test #'equalp))
  "Extract unique values from designated column."
  (let ((accumulator '()))
    ;; Collecting all column values
    (do-xarows (row r obj)
      (push (xcref row column-designator) accumulator))
    ;; Dropping duplicates
    (when accumulator
      (coerce (nreverse
               (remove-duplicates accumulator :test test))
              'simple-vector))))

(defmacro xadouniques ((val xarray column-designator &optional (test 'equalp)) &body body)
  "Iterate over unique values in designated column."
  (let ((uniques (gensym)))
    `(let ((,uniques (xauniques ,xarray ,column-designator :test #',test)))
       (loop for ,val across ,uniques doing
             ,@body))))

(defmethod xarray-width ((obj xarray))
  "Number of columns in OBJ (an xarray)."
  (with-slots ((body body)) obj
    (array-dimension body 1)))

(defmethod xarray-actual-height ((obj xarray))
  "Number of rows of OBJ's (an xarray) body."
  (with-slots ((body body)) obj
    (array-dimension body 0)))

(defmethod xarray-indexed-height ((obj xarray))
  "Number of rows in OBJ's (an xarray) index."
  (with-slots ((index index)) obj
    (length index)))

(defun xarray-zero-index-p (xarray)
  "Is indexed height 0?"
  (zerop (xarray-indexed-height xarray)))

;;; Helper functions for XAPRED:
(defun xapred-prev (record)
  (destructuring-bind (column-designator sorting equality)
      record
    (declare (ignore sorting))
    `(funcall ,equality
              (xcref a ,column-designator)
              (xcref b ,column-designator))))

(defun xapred-curr (record)
  (destructuring-bind (column-designator sorting &optional (equality nil))
      record
    (declare (ignore equality))
    `(funcall ,sorting
              (xcref a ,column-designator)
              (xcref b ,column-designator))))

(defun xapred-and (records)
  (if (second records)
    (append '(and) (mapcar #'xapred-prev (butlast records))
            `(,(xapred-curr (first (last records)))))
    (xapred-curr (first records))))

(defun xapred-or (records &optional (result '()))
  (if records
    (xapred-or (butlast records)
               (cons (xapred-and records)
                     result))
    (append '(or) result)))

;;; (xapred ("TK"    #'string<  #'string=)
;;;         ("SZK"   #'string<  #'string=)
;;;         ("SZTSZ" #'string<=))
(defmacro xapred (&rest records)
  "Predicate generator for XASORT, e.g."
  `(lambda (a b)
     ,(xapred-or records)))

(defmethod xasort ((obj xarray) predicate)
  "Return a copy of OBJ with a sorted index according to PREDICATE."
  (with-slots ((body body) (head head)) obj
    (let* ((new-index    (raw-index body))
           (sorted-index (sort new-index #'(lambda (a b)
                                             (funcall predicate
                                                      (xarows obj a)
                                                      (xarows obj b))))))
      (make-instance 'xarray
                     :head  head
                     :body  body
                     :index sorted-index))))
