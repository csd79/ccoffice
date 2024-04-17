;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;; ----------------------------------------------------------------------
;; Utilities


;; Provide an alias for an existing function.
(defmacro define-fn-nickname (function symbol &key (setf-able nil))
  `(progn
     (intern (symbol-name ',symbol))
     (setf (fdefinition ',symbol) ,function)
     (when ,setf-able
       (setf (fdefinition '(setf ,symbol))
             #'(lambda (new-value &rest args)
                 (setf (apply ,function args)
                       new-value))))))


;; Convert Excel column number to letter-based address.
(defun column (column)
  (let* ((third  (if (>= column 703)
                     (floor (- column 27) 676)
                   0))
         (n2     (- column (* third 676)))
         (second (if (>= n2 27)
                     (floor (1- n2) 26)
                   0))
         (first  (- n2 (* second 26)))
         (chars  (mapcar #'(lambda (code)
                             (if (zerop code)
                                 #\Space
                               (code-char (+ code 64))))
                         (list third second first))))
    (remove #\Space (format nil "~{~C~}" chars))))


;; Create Excel-style cell address from numerical address.
(defun cell-index (column row)
  (format nil "~a~d" (column column) row))


;; ----------------------------------------------------------------------
;; Basic COM crapper


;; Is COM initialized?
(defparameter *com-initialized-p* nil)


;; Binding context for COM operations.
(defmacro with-com-initialized (&body body)
  `(unwind-protect
       (progn
         (unless *com-initialized-p*
           (co-initialize))
         (let ((*com-initialized-p* t))
           ,@body))
     (unless *com-initialized-p*
       (co-uninitialize))))


;; Alieses for uncomfortably long function names.
(define-fn-nickname #'invoke-dispatch-get-property com-property :setf-able t)
(define-fn-nickname #'invoke-dispatch-method com-method)


;; Generate binding forms for COMLET+.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun comhelper (params body)
    (if params
        (let* ((current (first params))
               (sym     (first current))
               (val     (second current)))
          `(let ((,sym ,val))
             (if (typep ,sym 'com-interface)
               (with-temp-interface (,sym) ,val
                 ,(comhelper (rest params) body))
               ,(comhelper (rest params) body))))
      `(progn ,@body))))


;; Locally bind both variables and interface pointers.
(defmacro comlet* (parameter-list &body body)
  (comhelper parameter-list body))


;; Set visibility for an application.
(defun set-app-visibility (var vis)
  (comlet* ((app (com-property var "Application")))
           (setf (com-property app "Visible") vis)))


;; Binding context to use given application's objects.
(defmacro with-app ((variable application &key (visible t) (observe-running nil)) &body body)
  (with-gensyms (running)
    `(with-com-initialized
       (comlet* ((,running  (get-active-object :progid ,application
                                               :riid   'i-dispatch
                                               :errorp nil))
                 (,variable (or ,running
                                (create-object :progid ,application))))
         (unwind-protect
             (progn
               (set-app-visibility ,variable ,visible)
               ,@body)
           (when ,observe-running
             (unless ,running
               (com-method ,variable "Quit"))))))))


;; ----------------------------------------------------------------------
;; Excel


;; Describing Excel region of interest.
(defclass xlsx ()
  ((fullname  :initarg :fullname  :accessor fullname)
   (name      :initarg :name      :accessor name)
   (sheetname :initarg :sheetname :accessor sheetname)
   (x1        :initarg :x1        :accessor x1        :initform nil)
   (y1        :initarg :y1        :accessor y1        :initform nil)
   (x2        :initarg :x2        :accessor x2        :initform nil)
   (y2        :initarg :y2        :accessor y2        :initform nil)))


(defconstant +xl-calculation-automatic+ -4105)
(defconstant +xl-calculation-manual+    -4135)


;; Binding context for talking to Excel.
(defmacro excel ((variable &key (visible t) (observe-running nil)) &body body)
  `(with-app (,variable "Excel.Application" :visible ,visible :observe-running ,observe-running)
     (unwind-protect
         (progn 
           ;; Faster transactions.
           (setf (com-property ,variable "ScreenUpdating") nil
                 (com-property ,variable "Calculation")    +xl-calculation-manual+
                 (com-property ,variable "DisplayAlerts")  nil)
           ,@body)
       (progn
         ;; Back to normal Excel behaviour.
         (setf (com-property ,variable "DisplayAlerts")  nil
               (com-property ,variable "Calculation")    +xl-calculation-automatic+
               (com-property ,variable "ScreenUpdating") t)))))


;; Create a list of XLSX objects, each representing an open Excel workbook.
(defun open-workbooks ()
  (excel (xl)
    (comlet* ((workbooks (com-property xl "Workbooks"))
              (count     (com-property workbooks "Count")))
      (loop for i from 1 upto count collecting
            (comlet* ((workbook    (com-property workbooks   "Item" i))
                      (fullname    (com-property workbook    "FullName"))
                      (name        (com-property workbook    "Name"))
                      (worksheets  (com-property workbook    "Worksheets"))
                      (first-sheet (com-property worksheets  "Item" 1))
                      (sheetname   (com-property first-sheet "Name")))
              (make-instance 'xlsx
                             :fullname  fullname
                             :name      name
                             :sheetname sheetname))))))


;; Return a list of the names of every worksheet in an open Excel file.
(defun worksheet-names (xlsx)
  (excel (xl)
    (comlet* ((workbooks (com-property xl "Workbooks"))
              (workbook  (com-property workbooks "Item" (name xlsx))))
      (when workbook
        (comlet* ((worksheets (com-property workbook "Worksheets"))
                  (count      (com-property worksheets "Count")))
          (loop for i from 1 upto count collecting
                (comlet* ((worksheet (com-property worksheets "Item" i)))
                  (com-property worksheet "Name"))))))))


;; Create a copy of XLSX.
(defun copy-xlsx (xlsx)
  (make-instance 'xlsx
                 :fullname  (fullname xlsx)
                 :name      (name xlsx)
                 :sheetname (sheetname xlsx)
                 :x1        (x1 xlsx)
                 :y1        (y1 xlsx)
                 :x2        (x2 xlsx)
                 :y2        (y2 xlsx)))


;; Return a list of every worksheet in every open Excel file.
(defun open-worksheets ()
  (let ((workbooks (open-workbooks))
        (results   '()))
    (dolist (workbook workbooks)
      (dolist (sheetname (worksheet-names workbook))
        (let ((new-xlsx (copy-xlsx workbook)))
          (setf (sheetname new-xlsx) sheetname)
          (push new-xlsx results))))
    (nreverse results)))


;; Get interface pointer for workbook described by XLSX.
(defun grab-workbook (xlsx)
  (excel (xl)
    (comlet* ((workbooks  (com-property xl "Workbooks")))
      (com-property workbooks "Item" (name xlsx)))))


;; Get interface pointer for worksheet described by XLSX.
(defun grab-worksheet (xlsx)
  (excel (xl)
    (comlet* ((workbook   (grab-workbook xlsx))
              (worksheets (com-property workbook "Worksheets")))
      (com-property worksheets "Item" (sheetname xlsx)))))


(defun range (worksheet &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (excel (xl)
    (apply #'com-property worksheet "Range"
           ;; Single cell...
           (append (list (cell-index x1 y1))
                   ;; or if X2 & Y1 provided, a range.
                   (when (and x2 y2)
                     (list (cell-index x2 y2)))))))


#|;; Create new XLSX with added range.
(defun add-range (xlsx &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (let ((new (copy-xlsx xlsx)))
    (setf (x1 new) x1
          (y1 new) y1
          (x2 new) x2
          (y2 new) y2)
    new))


;; Get interface pointer for range described by XLSX.
(defun grab-range (xlsx)
  (with-slots (x1 y1 x2 y2) xlsx
    (when (and x1 y1)
      (excel (xl)
        (comlet* ((worksheet (grab-worksheet xlsx)))
          (range worksheet x1 y1 x2 y2))))))|#


(defun let-range-helper (bindings)
  (when bindings
    (destructuring-bind (var (xlsx &optional x1 y1 x2 y2))
        (first bindings)
      (cons `(,var (grab-range (add-range ,xlsx ,x1 ,y1 ,x2 ,y2)))
            (with-range-helper (rest bindings))))))


(defmacro let-range (bindings &body body)
  `(comlet* ,(let-range-helper bindings)
     ,@body))

           










;; Constants used to copy cell formatting.
(defconstant +xl-paste-formats+    -4122)
(defconstant +xl-paste-comments+   -4144)
(defconstant +xl-paste-validation+     6)


;; Copy cell formatting between ranges.
(defun copy-formatting (from x y into &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (excel (xl)
    (let ((from-range (add-range from x y))
          (into-range (add-range into x1 y1 x2 y2)))
      (comlet* ((source (grab-range from-range))
                (target (grab-range into-range)))
        (com-method source "Copy")
        (com-method target "PasteSpecial" +xl-paste-formats+)
        (com-method target "PasteSpecial" +xl-paste-comments+)
        (com-method target "PasteSpecial" +xl-paste-validation+)))))


(defun get-range (worksheet x1 y1 &optional (x2 nil) (y2 nil))
  (comlet* ((range (range worksheet x1 y1 x2 y2)))
    (com-property range "Value2")))


(defun set-range (worksheet value x1 y1 &optional (x2 nil) (y2 nil))
  (comlet* ((range (range worksheet x1 y1 x2 y2)))
    (setf (com-property range "Value2") value)))








      
    


(defun test6 (xlsx)
  (excel (xl)
    (comlet* ((worksheet (grab-worksheet xlsx))
              (data      (get-range worksheet 1 1 2 30)))
      (set-range worksheet data 6 1 7 30))))


(defun test5 (xlsx)
  (excel (xl)
    (with-no-screen-updating (xl)
      (comlet* ((worksheet (grab-worksheet xlsx))
                (data      (get-range worksheet 1 1 2 30)))
        (set-range worksheet data 6 1 7 30)))))


(defun test4 (xlsx)
  (excel (xl)
    (comlet* ((worksheet (grab-worksheet xlsx)))
      (get-range worksheet 2 1 2 30)))))














#|
(defconstant +xl-workbook-default+ 51)


(defun save-workbook (workbook file file-exists-p)
  (if file-exists-p
      (com-method workbook "Save")
    (com-method workbook "SaveAs" file +xl-workbook-default+)))


(defun close-workbook (workbook)
  (setf (com-property workbook "Saved") t)
  (com-method workbook "Close" nil))


(defmacro with-xlsx ((xlsx &key (file nil) (save t)
                           (close nil) (visible t) (observe-running nil)) &body body)
  (with-gensyms (file-exists-p wbs wb shs sh)
    `(excel (xl :visible ,visible :observe-running ,observe-running)
       (let ((,file-exists-p (and ,file (probe-file ,file))))
         (comlet* ((,wbs (com-property xl "Workbooks"))
                   (,wb  (if ,file-exists-p
                             (com-method ,wbs "Open" (namestring ,file))
                           (com-method ,wbs "Add" +xl-worksheet+)))
                   (,shs (com-property ,wb "Worksheets"))
                   (,sh  (com-property ,shs "Item" 1)))
           (let ((,xlsx (make-instance 'xlsx :app xl :file ,file :workbook ,wb
                                       :worksheet ,sh)))
             (unwind-protect
                 (progn ,@body)
               (when ,file
                 (when ,save  (save-workbook ,wb ,file ,file-exists-p))
                 (when ,close (close-workbook ,wb))))))))))


(defmacro with-range ((xlsx range x y &optional x2 y2) &body body)
  `(comlet* ((,range (apply #'com-property (worksheet ,xlsx) "Range"
                             (append (list (cell-index ,x ,y))
                                     (when (and ,x2 ,y2)
                                       (list (cell-index ,x2 ,y2)))))))
     ,@body))


(defun getcell (xlsx x y)
  (with-range (xlsx range x y)
    (com-property range "Value2")))


(defun setcell (xlsx x y value)
  (with-range (xlsx range x y)
    (setf (com-property range "Value2") value)))


(defun locate-string (string list partial)
  (let ((index (if partial
                   (position-if #'(lambda (item)
                       (search (string-downcase string)
                               (string-downcase item)))
                   list)
                 (position string list :test #'string=))))
    (if index
        (1+ index)
      (error (if partial
                 "Worksheet not found: ~a (partial name)"
               "Worksheet not found: ~a (full name)")
             string))))


(defun worksheet-name (worksheet)
  (com-property worksheet "Name"))


(defun list-sheet-names (worksheets)
  (let ((count (com-property worksheets "Count")))
    (loop for i from 1 upto count collecting
          (worksheet-name (com-property worksheets "Item" i)))))


(defun get-sheet (xlsx sheet &optional partial)
  (comlet* ((worksheets (com-property (workbook xlsx) "Worksheets")))
    (let ((number (typecase sheet
                    (number sheet)
                    (string (locate-string
                             sheet (list-sheet-names worksheets) partial)))))
      (values
       (com-property worksheets "Item" number)
       number))))


(defgeneric select-sheet (xlsx sheet &key partial))

(defmethod select-sheet ((xlsx xlsx) sheet &key partial)
  (setf (worksheet xlsx)
        (get-sheet xlsx sheet partial))
  xlsx)


(defgeneric rename-sheet (xlsx new-name &key sheet partial))

(defmethod rename-sheet ((xlsx xlsx) (new-name string) &key sheet partial)
  (when sheet
    (select-sheet xlsx sheet partial))
  (let ((worksheet (worksheet xlsx)))
    (setf (com-property worksheet "Name") new-name))
  xlsx)


(defgeneric create-sheet (xlsx new-sheet-name &key previous partial))

(defmethod create-sheet ((xlsx xlsx) (new-sheet-name string) &key previous partial)
  (multiple-value-bind (previous-sheet previous-sheet-number)
      (get-sheet xlsx (if previous
                          previous
                        (worksheet-name (worksheet xlsx)))
                 partial)
    (comlet* ((worksheets (com-property (workbook xlsx) "Worksheets")))
      (com-method worksheets "Add" :null previous-sheet))
    (select-sheet xlsx (1+ previous-sheet-number))
    (rename-sheet xlsx new-sheet-name))
  xlsx)


(defun used-range (worksheet)
  (comlet* ((range (com-property worksheet "Cells"))
            (up    (com-property range "End" +xl-up+))
            (left  (com-property range "End" +xl-to-left+))
            (from  (com-property worksheet "Range"
                                 (cell-index (com-property left "Column")
                                             (com-property up "Row"))))
            (down  (com-method range "Find" "*" from +xl-values+
                               +xl-whole+ +xl-by-rows+ +xl-previous+))
            (right (com-method range "Find" "*" from +xl-values+
                               +xl-whole+ +xl-by-columns+ +xl-previous+)))
    (values (com-property right "Column")
            (com-property down "Row"))))


(defmacro with-used-range ((xlsx last-column last-row) &body body)
  `(multiple-value-bind (,last-column ,last-row)
       (used-range (worksheet ,xlsx))
     ,@body))|#


;; ----------------------------------------------------------------------
;; Excel styles


#|
(defconstant +rgb-black+               0)
(defconstant +rgb-light-grey+   13882323)
(defconstant +rgb-yellow+          65535)
(defconstant +rgb-light-yellow+ 14745599)
(defconstant +xl-worksheet+        -4167)
(defconstant +xl-to-left+          -4159)
(defconstant +xl-up+               -4162)
(defconstant +xl-formulas+         -4123)
(defconstant +xl-values+           -4163)
(defconstant +xl-by-rows+              1)
(defconstant +xl-by-columns+           2)
(defconstant +xl-whole+                1)
(defconstant +xl-previous+             2)
(defconstant +xl-shift-down+       -4121)
(defconstant +xl-align-center+     -4108)


(defparameter *style-elements*
  `(:wrap-text     (:prop   "WrapText"            ,t)
    :halign-center (:prop   "HorizontalAlignment" ,+xl-align-center+)
    :valign-center (:prop   "VerticalAlignment"   ,+xl-align-center+)
    :border        (:method "BorderAround"        ,:null ,:null ,:null ,+rgb-black+)
    :bold          (:prop   "Bold"                ,t)))


(defun char-range (range &optional first last)
  (if (and first last)
      (comlet* ((chars (com-property range "Characters" first last)))
        (com-property chars "Font"))
    (com-property range "Font")))


(defun apply-style (range style)
  (dolist (style-element style)
    (destructuring-bind (type name &rest values)
        (getf *style-elements* style-element)
      (case type
        (:prop   (setf (com-property range name) (first values)))
        (:method (apply #'com-method range name values))))))


(defun insert-rows (xlsx row how-many)
  (comlet* ((range (com-property (worksheet xlsx) "Range"
                                 (cell-index 1 row)
                                 (cell-index 3 row))))
    (dotimes (i how-many)
      (com-method range "Insert" +xl-shift-down+))))


(defun autofit-cols (xlsx)
  (comlet* ((range (com-property (worksheet xlsx) "Columns")))
    (com-method range "AutoFit")))
|#
