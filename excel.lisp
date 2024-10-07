;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Workbooks, worksheets


(defconstant +xl-calculation-automatic+ -4105)
(defconstant +xl-calculation-manual+    -4135)


;;; Execute BODY using faster Excel interaction.
(defmacro excellerate ((excel) &body body)
  `(let ((su #~('screenupdating ,excel))
         (da #~('displayalerts ,excel))
         (c  #~('calculation ,excel)))
     (unwind-protect 
         (progn
           (setf #~('screenupdating ,excel) nil
                 #~('displayalerts ,excel)  nil)
           (when #~('visible ,excel)
             (setf #~('calculation ,excel) +xl-calculation-manual+))
           ,@body)
       (progn
         (when #~('visible ,excel)
           (setf #~('calculation ,excel) c))
         (setf #~('displayalerts  ,excel) da
               #~('screenupdating ,excel) su)))))


;;; Transpose a column (an (n 0) array) into a vector.
(defun column->row (array &key (element-type t) (getter #'identity))
  (let* ((height (array-dimension array 0))
         (result (make-array height :element-type element-type)))
    (loop for i from 0 below height doing
          (setf (svref result i)
                (funcall getter
                         (aref array i 0))))
    result))


;;; Try to interpret VALUE as a number.
(defun excel-value-as-number (value)
  (let ((not-string (if (stringp value)
                      (read-from-string value)
                      value)))
    (assert (numberp not-string)
        (not-string)
      "Value ~a is not readable as number." not-string)
    not-string))


;;; ----------------------------------------------------------------------
;;; Ranges, values


;; Create range object within given 'coordinates'.
(defun range (worksheet &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (let ((upper-left (cell-index x1 y1)))
    (if (and x2 y2)
      #~('range worksheet upper-left (cell-index x2 y2))
      #~('range worksheet upper-left))))


(defconstant +xl-to-left+          -4159)
(defconstant +xl-up+               -4162)
(defconstant +xl-values+           -4163)
(defconstant +xl-by-rows+              1)
(defconstant +xl-by-columns+           2)
(defconstant +xl-whole+                1)
(defconstant +xl-previous+             2)

(defmacro with-used-range ((worksheet left top right bottom) &body body)
  (let ((whole   (gensym))
        (from    (gensym))
        (endup   (gensym))
        (endleft (gensym))
        (findbtm (gensym))
        (findrgt (gensym)))
    `(let ((,left nil) (,top nil) (,right nil) (,bottom nil))
       (cclet* ((,whole #~('cells ,worksheet)))
         (cclet* ((,endup #~('end ,whole +xl-up+)))
           (setf ,top #~('row ,endup)))
         (cclet* ((,endleft #~('end ,whole +xl-to-left+)))
           (setf ,left #~('column ,endleft)))
         (cclet* ((,from (range ,worksheet ,left ,top)))
           (cclet* ((,findbtm (#_find ,whole "*" ,from +xl-values+
                                      +xl-whole+ +xl-by-rows+ +xl-previous+)))
             (setf ,bottom #~('row ,findbtm)))
           (cclet* ((,findrgt (#_find ,whole "*" ,from +xl-values+
                                      +xl-whole+ +xl-by-columns+
                                      +xl-previous+)))
             (setf ,right #~('column ,findrgt)))))
       ,@body)))


;;; Locate used range of WORKSHEET.
(defun used-range (worksheet)
  (with-used-range (worksheet left top right bottom)
    (range worksheet left top right bottom)))


;;; Return number of last row (& last column) of WORKSHEET.
(defun last-row (worksheet)
  (with-used-range (worksheet left top right bottom)
    (values bottom right)))


;;; ----------------------------------------------------------------------
;;; Cell formatting


;;; Constants used to copy cell formatting.
(defconstant +xl-paste-formats+    -4122)
(defconstant +xl-paste-comments+   -4144)
(defconstant +xl-paste-validation+     6)


;;; Copy cell formatting between ranges.
(defun copy-formatting (src dst)
  (#_copy src)
  (#_pastespecial dst +xl-paste-formats+)
  (#_pastespecial dst +xl-paste-comments+)
  (#_pastespecial dst +xl-paste-validation+))


;;; Set auto width for all columns.
(defun autofit-cols (worksheet)
  (cclet* ((range #~('columns worksheet)))
    (#_autofit range)))


;;; Target range for font formatting.
(defun font (range &optional first last)
  #~('font (if (and first last)
            #~('characters range first last)
            range)))


(defconstant +xl-align-center+     -4108)
(defconstant +rgb-black+               0)

(defparameter *style-elements*
  `(:wrap-text     (:prop   "WrapText"            ,t)
    :halign-center (:prop   "HorizontalAlignment" ,+xl-align-center+)
    :valign-center (:prop   "VerticalAlignment"   ,+xl-align-center+)
    :border        (:method "BorderAround"        ,:null ,:null ,:null ,+rgb-black+)
    :bold          (:prop   "Bold"                ,t)))


;;; Apply any style defined in *STYLE-ELEMENTS*.
(defun apply-style (range style)
  (dolist (style-element style)
    (destructuring-bind (type name &rest values)
        (getf *style-elements* style-element)
      (case type
        (:prop   (setf (com::invoke-dispatch-get-property range name)
                       (first values)))
        (:method (apply #'com::invoke-dispatch-method range name values))))))


;;; ----------------------------------------------------------------------
;;; Opening & closing


;;; Close WORKBOOK & quit application when there are no more workbooks open.
(defun close-workbook (workbook)
  (cclet* ((application #~('application workbook))
           (workbooks   #~('workbooks application)))
;    (setf #~('saved workbook) t)
    (#_close workbook)
    (when (zerop #~('count workbooks))
      (setf #~('displayalerts application) nil)
      (#_quit application))))


;;; Change window visibility.
(defun window-visible (workbook &optional (window 1) (visible t))
  (cclet* ((windows #~('windows workbook))
           (window  #~('item windows window)))
    (setf #~('visible window) visible)))


;;; Create bindings for an Excel workbook with optional method calls
;;; (open, save at the end, close).
;; MAYBE OK:
(defmacro with-workbook ((&key (wbook     'wbook)
                               (open-file nil)
                               (app       nil)
                               (read-only nil)
                               (wsheets   'wsheets)
                               (wsvars    '())
                               (close     t)
                               (save      nil))
                         &body body)
  (let ((wsheets-clauses (loop for n from 1
                               for var in wsvars collecting
                               (list var `#~('item ,wsheets ,n))))
        (app2    (gensym)))
    `(cclet* ((,app2    (if ,app
                          ,app
                          (cclet* ((global (com:create-object :progid "Excel.Application")))
                            #~('application global))))
              (wbooks   #~('workbooks ,app2))
              (,wbook   (if ,open-file
                          (#_open wbooks ,open-file nil ,read-only)
                          (#_add wbooks)))
              (,wsheets #~('worksheets ,wbook))
              ,@wsheets-clauses)
       (unwind-protect
           (excellerate (,app2)
             (window-visible ,wbook 1 t)
             ,@body)
         (progn 
           (when (and ,save (not ,read-only))
             (#_save ,wbook))
           (when ,close
             (setf #~('saved ,wbook) t)
             (setf #~('displayalerts ,app2) nil)
             (#_close ,wbook)
             (unless ,app
               (#_quit ,app2))))))))
;             (close-workbook ,wbook)))))))


;;; ----------------------------------------------------------------------
;;; Cell reference


(defparameter *header-row* 1)

;;; Find the column that has TITLE in its header.
(defun title-column (worksheet title &key (key #'identity) (header-row *header-row*) (test #'equalp)
                         (from-end nil) (start nil) (end nil))
  (with-used-range (worksheet left top right bottom)
    (let ((header #~('value2 (range worksheet left header-row right header-row))))
      (if (arrayp header)
        (if from-end
          (loop for i from (1- (or end right)) downto (1- (or start left))
                for v = (funcall key (aref header header-row i))
                thereis (and (funcall test title v) (1+ i)))
          (loop for i from (1- (or start left)) upto (1- (or end right))
                for v = (funcall key (aref header header-row i))
                thereis (and (funcall test title v) (1+ i))))
        (and (funcall test title header) left)))))


;;; Column designator type and resolution
(defun column-designator-p (designator)
  (or (and (integerp designator)
           (<= 1 designator 16384))
      (and (keywordp designator)
           (<= (length (symbol-name designator)) 3))
      (stringp designator)))

(deftype column-designator ()
  '(satisfies column-designator-p))

(defun resolve-column-designator (designator &optional worksheet)
  (assert (typep designator 'column-designator)
      (designator)
    "Invalid column designator ~a - should be az integer, a keyword or a string" designator)
  (typecase designator
    (integer designator)
    (keyword (letters-column designator))
    (string  (title-column worksheet designator))))


;;; Determine row index by search for value in a column.
#|(defun locate-row (worksheet title value function)
  (let ((column (title-column worksheet title)))
    (loop for row from 1
          for v = #~('value2 (range worksheet column row))
          until (null v)
          thereis (and (funcall function v value)
                       row))))|#

(defun locate-row (worksheet column-designator value function)
  (let* ((column-idx (resolve-column-designator column-designator worksheet))
         (column     (column->row #~('value2 (range worksheet column-idx 2 column-idx (last-row worksheet))))))
    (loop for row from 0 below (length column)
          for v = (svref column row)
;          until (null v)
          thereis (and (funcall function v value)
                       (+ row 2)))))


;;; Row designator type and resolution
(defun row-designator-p (designator)
  (or (and (integerp designator)
           (<= 1 designator 1048576))
      (and (listp designator)
           (typep (first designator) 'column-designator)
           (if (third designator)
             (functionp (third designator))
             t))))

(deftype row-designator ()
  '(satisfies row-designator-p))

#|(defun resolve-row-designator (designator &optional worksheet)
  (assert (typep designator 'row-designator)
      (designator)
    "Invalid row designator ~a - should be az integer or a list of a column designator, a value and a predicate function" designator)
  (typecase designator
    (integer designator)
    (list    (destructuring-bind (title value &optional (function #'equalp))
                 designator
               (locate-row worksheet title value function)))))|#

(defun resolve-row-designator (designator &optional worksheet)
  (assert (typep designator 'row-designator)
      (designator)
    "Invalid row designator ~a - should be az integer or a list of a column designator, a value and a predicate function" designator)
  (typecase designator
    (integer designator)
    (list    (destructuring-bind (column-designator value &optional (function #'equalp))
                 designator
               (locate-row worksheet column-designator value function)))))


;;; Main body of XCELL and SET-XCELL.
(defun xcell-core (worksheet column row &key (value nil value-supplied-p) (prop 'value2))
  (let ((row-final    (resolve-row-designator row worksheet))
        (column-final (resolve-column-designator column worksheet)))
    (if value-supplied-p
#|      (setf #~('value2 (range worksheet column-final row-final)) value)
      #~('value2 (range worksheet column-final row-final)))))|#
      (setf #~(prop (range worksheet column-final row-final)) value)
      #~(prop (range worksheet column-final row-final)))))


;;; Simplified cell reference.
(defun xcell (worksheet column row &key (prop 'value2))
  (xcell-core worksheet column row :prop prop))

(defun (setf xcell) (value worksheet column row &key (prop 'value2))
  (xcell-core worksheet column row :value value :prop prop))


;(defun set-xcell (worksheet column row value &key (prop 'value2))
;  (xcell-core worksheet column row :value value :prop prop))

;(defsetf xcell set-xcell)

(defun (setf xrange) (value worksheet c1 r1 c2 r2 &key (prop 'value2))
  (xrange-core worksheet c1 r1 c2 r2 :value value :prop prop))

;;; ----------------------------------------------------------------------
;;; Range reference


;;; Types for possible values to assign to a range
(defun 2d-array-column-p (value)
  (and (arrayp value)
       (= (length (array-dimensions value)) 2)
       (= (array-dimension value 1) 1)))

(deftype 2d-array-column ()
  '(satisfies 2d-array-column-p))


(defun 2d-array-row-p (value)
  (and (arrayp value)
       (= (length (array-dimensions value)) 2)
       (= (array-dimension value 0) 1)))

(deftype 2d-array-row ()
  '(satisfies 2d-array-row-p))


(defun 2d-array-p (value)
  (and (arrayp value)
       (= (length (array-dimensions value)) 2)
       (> (array-dimension value 0) 1)
       (> (array-dimension value 1) 1)))

(deftype 2d-array ()
  '(satisfies 2d-array-p))


(defparameter *xrange-default-value*    "")             ; used when *XRANGE-TARGET-TOO-LARGE* = :FULFILL
(defparameter *xrange-target-too-small* :fill-to-limit) ; :ERROR  :FILL-TO-LIMIT
(defparameter *xrange-target-too-large* :restrict)      ; :ERROR  :FULFILL  :RESTRICT

;;; Multiple cell range assignment.
(defun set-xrange-worker (worksheet c1 r1 c2 r2 value prop)
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
      (setf #~(prop (range worksheet c1 r1 c2r r2r)) value)
      ;; Add fill when needed
      (when (eq *xrange-target-too-large* :fulfill)
        (let ((rightp (and (not (2d-array-column-p value))
                           (> range-width value-width)))
              (lowerp (and (not (2d-array-row-p value))
                           (> range-height value-height))))
          ;; Right side fill
          (when rightp
            (setf #~(prop (range worksheet (1+ c2r) r1 c2 r2r))
                  *xrange-default-value*))
          ;; Lower fill
          (when lowerp
            (setf #~(prop (range worksheet c1 (1+ r2r) c2r r2))
                  *xrange-default-value*))
          ;; Lower right fill
          (when (and rightp lowerp)
            (setf #~(prop (range worksheet (1+ c2r) (1+ r2r) c2 r2))
                  *xrange-default-value*)))))))


;;; Prefilter arg combinations that yield simple assignment or no assignment.
(defun precide-xrange (worksheet c1 r1 c2 r2 value prop)
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
    (setf #~(prop (range worksheet c1 r1 c2 r2)) value))
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

;;; Main body of XRANGE and SET-XRANGE.
(defun xrange-core (worksheet c1 r1 c2 r2 &key (value nil value-supplied-p) (prop 'value2))
  (let ((c1final (resolve-column-designator c1 worksheet))
        (c2final (resolve-column-designator c2 worksheet))
        (r1final (resolve-row-designator r1 worksheet))
        (r2final (resolve-row-designator r2 worksheet)))
    (if value-supplied-p
      (precide-xrange worksheet c1final r1final c2final r2final value prop)
      #~(prop (range worksheet c1final r1final c2final r2final)))))


;;; Simplified cell reference.
(defun xrange (worksheet c1 r1 c2 r2 &key (prop 'value2))
  (xrange-core worksheet c1 r1 c2 r2 :prop prop))

(defun (setf xrange) (value worksheet c1 r1 c2 r2 &key (prop 'value2))
  (xrange-core worksheet c1 r1 c2 r2 :value value :prop prop))

;(defun set-xrange (worksheet c1 r1 c2 r2 value &key (prop 'value2))
;(defsetf xrange set-xrange)


;;; ----------------------------------------------------------------------
;;; Filtering


(defconstant +xl-celltype-visible+   12)
(defconstant +xl-paste-values+    -4163)
(defconstant +xl-paste-all+       -4104)
(defconstant +xl-paste-column-widths+ 8)
(defconstant +and+                    1)
(defconstant +or+                     2)

;;; Select rows from WORKSHEET using AutoFilter,
;;; copy the results into a new temporary workbook.
;;;     (xselect> ws `(("TK" "Eger" +or+ "Baja") ("SZTSZ" ">5" +and+ "<999999")))
(defun xselect> (worksheet subscripts &key (paste-all nil) (paste-column-widths nil))
  (cclet* ((cellidx (cell-index 1 1))
           (range   (used-range worksheet))
           (topleft (range worksheet 1 1))
           (excel   #~('application worksheet))
           (wbooks  #~('workbooks excel))
           (rbook   (#_add wbooks))
           (rsheets #~('worksheets rbook))
           (rsheet  #~('item rsheets 1))
           (rrange  #~('range rsheet cellidx))
           (corsubs (mapcar #'(lambda (sub)
                                (append (list (resolve-column-designator (first sub) worksheet))
                                        (rest sub)))
                            subscripts)))
    ;; Apply filters.
    (dolist (subscript corsubs)
      (apply #'com::invoke-dispatch-method topleft "autofilter" subscript))
    ;; Copy selected data to new worksheet
    (#_copy (#_specialcells range +xl-celltype-visible+))
    (when paste-column-widths
      (#_pastespecial rrange +xl-paste-column-widths+))
    (if paste-all
      (#_pastespecial rrange +xl-paste-all+)
      (#_pastespecial rrange +xl-paste-values+))
    ;; Turn autofilter off
    (#_autofilter range)
    ;; Return worksheet containing results
    (setf #~('saved #~('parent worksheet)) t
          #~('saved rbook) t)
    rsheet))


(defmacro with-xselection ((selected source subscripts &key (paste-all nil) (paste-column-widths nil)) &body body)
  (let ((wbook (gensym)))
    `(cclet* ((,selected (xselect> ,source ,subscripts :paste-all ,paste-all
                                   :paste-column-widths ,paste-column-widths)))
       (unwind-protect
           (progn
             ,@body)
         (cclet* ((,wbook (#_parent ,selected)))
           (close-workbook ,wbook))))))


;;; ----------------------------------------------------------------------
;;; Dates


;;; Convert Excel date value to a list of year, month and day.
(defun excel-date (n)
  (let* ((a    (+ n 2483588))
         (b    (truncate (/ (* a 4) 146097)))
         (c    (- a (truncate (/ (+ (* 146097 b) 3) 4))))
         (d    (truncate (/ (* 4000 (+ c 1)) 1461001)))
         (e    (+ (- c (+ (truncate (/ (* 1461 d) 4)))) 31))
         (f    (truncate (/ (* 80 e) 2447)))
         (day  (- e (truncate (/ (* 2447 f) 80))))
         (g    (truncate (/ f 11)))
         (mon  (- (+ f 2) (* 12 g)))
         (year (+ (* 100 (- b 49)) d g)))
    (list year mon day)))


;;; Convert Excel date value to a hu formated string.
(defun excel-date-string (n &key (words nil))
  (destructuring-bind (year mon day)
      (excel-date (truncate n))
    (if words
      (let ((months '("január" "február" "március" "április" "május" "június" "július"
                      "augusztus" "szeptember" "október" "november" "december")))
        (format nil "~4d. ~a ~d." year (nth (1- mon) months) day))
    (format nil "~4d.~2,'0d.~2,'0d." year mon day))))


;;; ----------------------------------------------------------------------
;;; Worksheet as xarray


;;; Find xarray column by title.
(defun title-xacolumn (xarray title &key (test #'equalp) (header-row (1- *header-row*)) (from-end nil))
  (let ((width (array-dimension xarray 1)))
    (if from-end
      (loop for c from (1- width) downto 0
            for v = (aref xarray header-row c)
            thereis (and (funcall test title v) c))
      (loop for c from 0 below width
            for v = (aref xarray header-row c)
            thereis (and (funcall test title v) c)))))


;;; Find xarray column by number, title or Excel letter.
(defun resolve-xacolumn-designator (designator xarray)
  (assert (or (typep designator 'column-designator)
              (zerop designator))
      (designator)
    "Invalid column designator ~a - should be az integer, a keyword or a string" designator)
  (typecase designator
    (integer designator)
    (keyword (1- (letters-column designator)))
    (string  (title-xacolumn xarray designator))))


;;; Extract (indexed) ROW from XARRAY.
(defun xarow (xarray row &optional (index nil))
  (let* ((width  (array-dimension xarray 1))
         (result (make-array (list 2 width)))
         (r      (if index (svref index row) row)))
    (loop for c from 0 below width doing
          (setf (aref result 0 c)
                (aref xarray 0 c)
                (aref result 1 c)
                (aref xarray r c)))
    result))


;;; Indexes xarray reference (reading only).
(defun xacell (xarray column-designator &optional (row 1) (index nil))
  (let ((r (if index (svref index (1- row)) row))
        (c (resolve-xacolumn-designator column-designator xarray)))
    (aref xarray r c)))


;;; Iterate over (indexed) xarray rows.
(defmacro do-xarows ((current row xarray &optional (index nil)) &body body)
  (let ((width (gensym))
        (i     (gensym)))
    `(let* ((,width   (array-dimension ,xarray 1))
            (,current (make-array (list 2 ,width))))
       ;; Filling in the header
       (loop for c from 0 below ,width doing
             (setf (aref ,current 0 c)
                   (aref ,xarray 0 c)))
       ;; Helper function for filling CURRENT with the current row's values 
       (flet ((fill-in (,i)
                (loop for c from 0 below ,width doing
                      (setf (aref ,current 1 c)
                            (aref ,xarray ,i c)))))
         (if ,index
           ;; If INDEX provided, iterate over it
           (loop for ,row across ,index doing
                 (fill-in ,row)
                 ,@body)
           ;; Otherwise, iterate over every row
           (loop for ,row from 1 below (array-dimension ,xarray 0) doing
                 (fill-in ,row)
                 ,@body))))))


;;; Create an index vector containing row selected by SELECTOR-FN.
(defun xaselect (xarray selector-fn &optional (previous-index nil))
  (let ((rowlist '()))
    (do-xarows (current row xarray previous-index)
      (when (funcall selector-fn current)
        (push row rowlist)))
    (when rowlist
      (coerce (nreverse rowlist) 'simple-vector))))


#|(defmacro with-xaselection ((selection xarray selector-fn &optional previous-selection) &body body)
  `(let ((,selection (select-xarows ,xarray ,selector-fn ,previous-selection)))
     ,@body))|#


;;; Extract unique values from designated column.
(defun xauniques (xarray column-designator &key (test #'equalp) (index nil))
  (let ((values '())
        (column (resolve-xacolumn-designator column-designator xarray)))
    ;; Collecting all column values
    (do-xarows (current row xarray index)
      (push (xacell current column) values))
    ;; Dropping duplicates
    (when values
      (coerce (nreverse
               (remove-duplicates values :test test))
              'simple-vector))))
           

;;; Iterate over unique values in designated column.
(defmacro xadouniques ((val xarray column-designator &key (test #'equalp) (selection nil)) &body body)
  (let ((uniques (gensym)))
    `(let ((,uniques (xauniques ,xarray ,column-designator :test ,test :selection ,selection)))
       (loop for ,val across ,uniques doing
             ,@body))))


;;; Helper functions for XAPRED:
(defun xapred-prev (record)
  (destructuring-bind (column-designator sorting equality)
      record
    (declare (ignore sorting))
    `(funcall ,equality
              (xacell a ,column-designator)
              (xacell b ,column-designator))))

(defun xapred-curr (record)
  (destructuring-bind (column-designator sorting &optional (equality nil))
      record
    (declare (ignore equality))
    `(funcall ,sorting
              (xacell a ,column-designator)
              (xacell b ,column-designator))))

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
  
;;; Predicate generator for XASORT.
(defmacro xapred (&rest records)
  `(lambda (a b)
     ,(xapred-or records)))

;;; Create new sorted index according to PREDICATE.
(defun xasort (xarray predicate &optional (previous-index nil))
  (let ((seq (or (copy-seq previous-index)
                 (coerce (loop for i from 0 below (1- (array-dimension xarray 0))
                               collecting i)
                         'simple-vector))))
    (sort seq #'(lambda (a b)
                  (funcall predicate
                           (xarow xarray a previous-index)
                           (xarow xarray b previous-index))))))

















