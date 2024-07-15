;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Excel workbooks, worksheets


;;; Get Excell app.
(defun get-excel (&optional object)
  (if object
    ;; Parrent instance of OBJECT.
    #p(application object)
    ;; Any running Excel instance.
    (com::get-active-object :progid "Excel.Application"
                            :riid   'i-dispatch
                            :errorp nil)))


(defconstant +xl-calculation-automatic+ -4105)
(defconstant +xl-calculation-manual+    -4135)


;;; Execute BODY using faster Excel interaction.
(defmacro excellerate ((excel) &body body)
  `(unwind-protect 
       (progn 
         (setf #p(screenupdating ,excel) nil
               #p(displayalerts ,excel)  nil)
         (when #p(visible ,excel)
           (setf #p(calculation ,excel) +xl-calculation-manual+))
         ,@body)
     (setf #p(screenupdating ,excel) t
           #p(displayalerts ,excel)  t)
     (when #p(visible ,excel)
       (setf #p(calculation ,excel) +xl-calculation-automatic+))))


;;; Open given document with default application, hopefully Excel.
(defun get-document (fullname)
  (com::get-object fullname :riid 'com::i-dispatch))


;;; List pathname of every open Excel workbook.
(defun open-workbooks (excel)
  (cclet* ((workbooks #p(workbooks excel))
           (count     #p(count workbooks)))
    (loop for i from 1 upto count collecting
          #p(fullname #p(item workbooks i)))))


;;; List every open worksheet.
(defun open-worksheets (excel)
  (let ((workbooks (open-workbooks excel))
        (results   '()))
    (dolist (workbook workbooks)
      (cclet* ((worksheets #p(worksheets (get-document workbook)))
               (count      #p(count worksheets)))
        (push workbook results)
        (push (loop for i from 1 upto count collecting
                    #p(name #p(item worksheets i)))
              results)))
    (nreverse results)))


;;; ----------------------------------------------------------------------
;;; Excel ranges, values


;; Create range object within given 'coordinates'.
(defun range (worksheet &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (let ((upper-left (cell-index x1 y1)))
    (if (and x2 y2)
      #p(range worksheet upper-left (cell-index x2 y2))
      #p(range worksheet upper-left))))


(defconstant +xl-to-left+          -4159)
(defconstant +xl-up+               -4162)
(defconstant +xl-values+           -4163)
(defconstant +xl-by-rows+              1)
(defconstant +xl-by-columns+           2)
(defconstant +xl-whole+                1)
(defconstant +xl-previous+             2) 

(defmacro with-used-edges ((worksheet left top right bottom) &body body)
  (let ((whole (gensym))
        (from  (gensym)))
    `(cclet* ((,whole  #p(cells ,worksheet))
              (,top    #p(row #p(end ,whole +xl-up+)))
              (,left   #p(column #p(end ,whole +xl-to-left+)))
              (,from   (range ,worksheet ,left ,top))
              (,bottom #p(row #m(find ,whole "*" ,from +xl-values+
                                     +xl-whole+ +xl-by-rows+ +xl-previous+)))
              (,right  #p(column #m(find ,whole "*" ,from +xl-values+
                                        +xl-whole+ +xl-by-columns+
                                        +xl-previous+))))
         ,@body)))


;;; Locate used range of WORKSHEET.
(defun used-range (worksheet)
  (with-used-edges (worksheet left top right bottom)
    (range worksheet left top right bottom)))


;;; Return number of last row (& last column) of WORKSHEET.
(defun last-row (worksheet)
  (with-used-edges (worksheet left top right bottom)
    (values bottom right)))


;;; ----------------------------------------------------------------------
;;; Excel cell formatting


;;; Constants used to copy cell formatting.
(defconstant +xl-paste-formats+    -4122)
(defconstant +xl-paste-comments+   -4144)
(defconstant +xl-paste-validation+     6)


;;; Copy cell formatting between ranges.
(defun copy-formatting (source destination)
  #m(copy source)
  #m(pastespecial destination +xl-paste-formats+)
  #m(pastespecial destination +xl-paste-comments+)
  #m(pastespecial destination +xl-paste-validation+))


;;; Set auto width for all columns.
(defun autofit-cols (worksheet)
  (cclet* ((range #p(columns worksheet)))
    #m(auto-fit range)))


;;; Target range for font formatting.
(defun font (range &optional first last)
  #p(font (if (and first last)
            #p(characters range first last)
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
;;; Excel indexing & searching basics


(defparameter *header-row* 1)

;;; Find the column that has TITLE in its header.
(defun title-column (worksheet title &key (key #'identity) (header-row *header-row*) (test #'equalp)
                         (from-end nil) (start nil) (end nil))
  (with-used-edges (worksheet left top right bottom)
;    (declare (ignore top bottom))
    (cclet* ((header #p(value2 (range worksheet left header-row right header-row))))
      (if from-end
        (loop for i from (1- (or end right)) downto (1- (or start left))
              for v = (funcall key (aref header 0 i))
              thereis (and (funcall test title v) (1+ i)))
        (loop for i from (1- (or start left)) upto (1- (or end right))
              for v = (funcall key (aref header 0 i))
              thereis (and (funcall test title v) (1+ i)))))))


;;; Create bindings for an Excel workbook with optional method calls
;;; (open, save at the end, close).
(defmacro with-workbook ((wbook &key (open-file   nil)
                                     (wsheets     'wsheets)
                                     (wsvars      '())
                                     (app         'excel)
                                     (excellerate t)
                                     (close       nil)
                                     (save        nil))
                         &body body)
  (let* ((open-clauses    `((,app     (com:create-object :progid "Excel.Application"))
                            (wbooks   #p(workbooks ,app))
                            (,wbook   #m(open wbooks ,open-file))
                            (,wsheets #p(worksheets ,wbook))))
        (existing-clauses `((,app     #p(application ,wbook))
                            (,wsheets #p(worksheets ,wbook))))
        (init-clauses     (if open-file
                            open-clauses
                            existing-clauses))
        (wsheets-clauses  (loop for n from 1
                                for var in wsvars collecting
                                (list var `#p(item ,wsheets ,n))))
        (body-wrap        (if excellerate
                            `(excellerate (,app)
                               ,@body)
                            `(progn
                               ,@body))))
    `(ccom:cclet* ,(append init-clauses wsheets-clauses)
       (unwind-protect
           ,body-wrap
         (progn
           (when ,save
             #m(save ,wbook))
           (when ,close
             #m(close ,wbook)))))))


(defconstant +xl-celltype-visible+ 12)
(defconstant +xl-paste-values+  -4163)

;;; Select rows from WORKSHEET using AutoFilter,
;;; copy the results into a new temporary workbook.
(defun wsselect (worksheet &rest subscripts)
  (format t "wsselect: ws: ~a; subs: ~a~%" worksheet subscripts)
  (cclet* ((cellidx (cell-index 1 1))
           (range   (used-range worksheet))
           (excel   #p(application worksheet))
           (wbooks  #p(workbooks excel))
           (rbook   #m(add wbooks))
           (rsheets #p(worksheets rbook))
           (rsheet  #p(item rsheets 1))
           (rrange  #p(range rsheet cellidx)))
    ;; Filter source range
    (dolist (subscript subscripts)
      (destructuring-bind (title value)
          subscript
        (format t "wsselect: type of val: ~a~%" (type-of value))
        (format t "wsselect: titlecol: ~a~%" (title-column worksheet title))
        #m(autofilter range
                      (title-column worksheet title)
                      value)))
    ;; Copy selected data to new worksheet
    #m(copy #m(specialcells range +xl-celltype-visible+))
    #m(pastespecial rrange +xl-paste-values+)
    ;; Turn autofilter off
    #m(autofilter range)
    ;; Return worksheet containing results
    (setf #p(saved #p(parent worksheet)) t
          #p(saved rbook) t)
    rsheet))


;;; Transpose a column (an (n 0) array) into a vector.
(defun column->row (array &key (element-type t) (getter #'identity))
  (let* ((height (array-dimension array 0))
         (result (make-array height :element-type element-type)))
    (loop for i from 0 below height doing
          (setf (svref result i)
                (funcall getter
                         (aref array i 0))))
    result))


;;; ----------------------------------------------------------------------
;;; Excel cell reference


#|;;; XCELL column index assertions.
(defun assert-xcol (column)
  (assert (or (integerp column)
              (stringp  column))
      (column)
    "Invalid column for worksheet: ~a" column)
  column)


;;; XCELL row index assertions.
(defun assert-xrow (row)
  (assert (or (integerp row)
              (and (listp row)
                   (= (length row) 2)))
      (row)
    "Invalid row for worksheet: ~a" row)
  row)|#


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
(defun locate-row (worksheet title value function)
  (let ((column (title-column worksheet title)))
    (loop for row from 1
          for v = #p(value2 (range worksheet column row))
          until (null v)
;          thereis (and (equalp v value)
          thereis (and (funcall function v value)
                       row))))


;;; Row designator type and resolution
(defun row-designator-p (designator)
  (or (and (integerp designator)
           (<= 1 designator 1048576))
      (and (listp designator)
           (typep (first designator) 'column-designator)
           (functionp (third designator)))))

(deftype row-designator ()
  '(satisfies row-designator-p))

(defun resolve-row-designator (designator &optional worksheet)
  (assert (typep designator 'row-designator)
      (designator)
    "Invalid row designator ~a - should be az integer or a list of a column designator, a value and a predicate function" designator)
  (typecase designator
    (integer designator)
    (list    (destructuring-bind (title value function)
                 designator
               (locate-row worksheet title value function)))))


#|;;; Main body of XCELL and SET-XCELL.
(defun xcell-core (worksheet column row &optional value)
  (let* ((col-ok (assert-xcol column))
         (row-ok (assert-xrow row))
         (col-ac (if (stringp col-ok)
                   (title-column worksheet col-ok)
                   col-ok))
         (row-ac (if (listp row-ok)
                   (apply #'locate-row worksheet row-ok)
                   row-ok)))
    (if value
      (setf #p(value2 (range worksheet col-ac row-ac)) value)
      #p(value2 (range worksheet col-ac row-ac)))))|#


;;; Main body of XCELL and SET-XCELL.
(defun xcell-core (worksheet column row &optional value)
  (let* ((row-mid      (if (and (listp row)
                                (<= (length row) 2))
                         (append row (list #'equalp))
                         row))
         (row-final    (resolve-row-designator row-mid worksheet))
         (column-final (resolve-column-designator column worksheet)))
    (if value
      (setf #p(value2 (range worksheet column-final row-final)) value)
      #p(value2 (range worksheet column-final row-final)))))


;;; Simplified cell reference.
(defun xcell (worksheet column row)
  (xcell-core worksheet column row))

(defun set-xcell (worksheet column row value)
  (xcell-core worksheet column row value))

(defsetf xcell set-xcell)


;;; ----------------------------------------------------------------------
;;; Word


(defconstant +wd-find-continue+ 1)

;;; Replace text in Word doc.
(defun word-replace-text (document orig-text new-text)
  (cclet* ((find #p(find #p(content document))))
    #m(execute find orig-text nil nil nil nil nil t
               +wd-find-continue+ nil new-text)))


;;; ----------------------------------------------------------------------
;;; Sandbox


(defun test1 ()
  (cclet* ((file    "c:\\Users\\cselovszkid\\Desktop\\lisp-book-structure-comparison.xlsx")
           (wbook   (get-document file))
           (wsheets #p(worksheets wbook))
           (wsheet  #p(item wsheets 1))
           (source  (range wsheet 1 1 1 5))
           (dest    (range wsheet 3 1 3 5))
           (write   (range wsheet 1 100 1 105))
           (used    (used-range wsheet))
           (app     (get-excel wbook)))
    (excellerate (app)
      ;; Copy formatting.
      (copy-formatting source dest)
      ;; Set values.
      (setf #p(value2 write) "Grr")
      ;; Read values.
      #p(value2 used)
      ;; Add borders.
      (apply-style (range wsheet 1 9 1 14) '(:border))
      ;; Bold text.
      (apply-style (font (range wsheet 1 11 1 14)) '(:bold)))))


(defparameter *root* "c:\\Users\\cselovszkid\\Downloads\\2024.06.27. Újabb kinevezések elõkészület\\Lekérdezés\\")
(defparameter *alap* "B8_0008IT_20240701.XLSX")
(defparameter *rendszeres* "B8_0014IT_20240701.XLSX")

(defun workfile (file)
  (probe-file
   (pathname
    (concatenate 'string (namestring *root*)
                 (namestring file)))))


