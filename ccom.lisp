;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;; ----------------------------------------------------------------------
;; Utilities


;; Provide an alias for an existing function.
#|(defmacro define-fn-nickname (function symbol &key (setf-able nil))
  `(progn
     (intern (symbol-name ',symbol))
     (setf (fdefinition ',symbol) ,function)
     (when ,setf-able
       (setf (fdefinition '(setf ,symbol))
             #'(lambda (new-value &rest args)
                 (setf (apply ,function args)
                       new-value))))))|#


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
;; COM syntactic sugar


  ;; <(PROPERTY OBJECT)"

(defmacro comprop (property object)
  (let ((symbol-string (symbol-name property)))
    `(invoke-dispatch-get-property ,object ,symbol-string)))

#|(eval-when (:load-toplevel)
  (defun comprop-reader (stream char)
    (declare (ignore char))
    (nconc (list 'comprop)
           (read stream)))

  (set-macro-character #\< 'comprop-reader))
|#

;; >(METHOD OBJECT)"
(defmacro commethod (method object &rest args)
  (let ((symbol-string (symbol-name method)))
    `(invoke-dispatch-method ,object ,symbol-string ,@args)))

#|(eval-when (:load-toplevel)
  (defun commethod-reader (stream char)
    (declare (ignore char))
    (nconc (list 'commethod)
           (read stream)))

  (set-macro-character #\> 'commethod-reader))
|#

(defparameter *list-reader* (get-macro-character #\( nil))

(set-dispatch-macro-character
 #\# #\<
 (lambda (stream sub-char infix)
   (declare (ignore sub-char infix))
   (nconc (list 'comprop)
          (read stream))))

(set-dispatch-macro-character
 #\# #\>
 (lambda (stream sub-char infix)
   (declare (ignore sub-char infix))
   (nconc (list 'commethod)
          (read stream))))






;; ----------------------------------------------------------------------
;; Basic COM wrapper


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
#|(define-fn-nickname #'invoke-dispatch-get-property com-property :setf-able t)
(define-fn-nickname #'invoke-dispatch-method com-method)|#


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
;; Workbooks, worksheets


;; Describing Excel region of interest.
(defclass xlsx-handle ()
  ((fullname  :initarg :fullname  :accessor fullname)
   (name      :initarg :name      :accessor name)
   (sheetname :initarg :sheetname :accessor sheetname)))


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


;; Create a list of XLSX-HANDLE objects, each representing an open Excel workbook.
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
              (make-instance 'xlsx-handle
                             :fullname  fullname
                             :name      name
                             :sheetname sheetname))))))


;; Return a list of the names of every worksheet in an open Excel file.
(defun worksheet-names (handle)
  (excel (xl)
    (comlet* ((workbooks (com-property xl "Workbooks"))
              (workbook  (com-property workbooks "Item" (name handle))))
      (when workbook
        (comlet* ((worksheets (com-property workbook "Worksheets"))
                  (count      (com-property worksheets "Count")))
          (loop for i from 1 upto count collecting
                (comlet* ((worksheet (com-property worksheets "Item" i)))
                  (com-property worksheet "Name"))))))))


;; Create a copy of HANDLE.
(defun copy-xlsx-handle (handle)
  (make-instance 'xlsx-handle
                 :fullname  (fullname handle)
                 :name      (name handle)
                 :sheetname (sheetname handle)))


;; Return a list of every worksheet in every open Excel file.
(defun open-worksheets ()
  (let ((workbooks (open-workbooks))
        (results   '()))
    (dolist (handle workbooks)
      (dolist (sheetname (worksheet-names handle))
        (let ((new-handle (copy-xlsx-handle handle)))
          (setf (sheetname new-handle) sheetname)
          (push new-handle results))))
    (nreverse results)))


;; Get interface pointer for workbook described by XLSX.
(defun grab-workbook (handle)
  (excel (xl)
    (comlet* ((workbooks  (com-property xl "Workbooks")))
      (com-property workbooks "Item" (name handle)))))


;; Get interface pointer for worksheet described by XLSX.
(defun grab-worksheet (handle)
  (excel (xl)
    (comlet* ((workbook   (grab-workbook handle))
              (worksheets (com-property workbook "Worksheets")))
      (com-property worksheets "Item" (sheetname handle)))))


;; ----------------------------------------------------------------------
;; Ranges, values


;; Create range object within given 'coordinates'.
(defun range (worksheet &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (excel (xl)
    (apply #'com-property worksheet "Range"
           ;; Single cell...
           (append (list (cell-index x1 y1))
                   ;; or if X2 & Y1 provided, a range.
                   (when (and x2 y2)
                     (list (cell-index x2 y2)))))))


;; Helper function to create binding clauses for LET-RANGE.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun let-range-helper (bindings)
    (when bindings
      (destructuring-bind (var (handle &optional x1 y1 x2 y2))
          (first bindings)
        (cons `(,var (range (grab-worksheet ,handle) ,x1 ,y1 ,x2 ,y2))
              (let-range-helper (rest bindings)))))))


;; Local range bindings.
(defmacro let-range (bindings &body body)
  `(comlet* ,(let-range-helper bindings)
     ,@body))

           
;; Get value(s) of given range.
(defun get-range (worksheet x1 y1 &optional (x2 nil) (y2 nil))
  (comlet* ((range (range worksheet x1 y1 x2 y2)))
    (com-property range "Value2")))


;; Set value(s) of given range.
(defun set-range (worksheet value x1 y1 &optional (x2 nil) (y2 nil))
  (comlet* ((range (range worksheet x1 y1 x2 y2)))
    (setf (com-property range "Value2") value)))


;; ----------------------------------------------------------------------
;; Cell formatting


;; Constants used to copy cell formatting.
(defconstant +xl-paste-formats+    -4122)
(defconstant +xl-paste-comments+   -4144)
(defconstant +xl-paste-validation+     6)


;; Copy cell formatting between ranges.
(defun copy-formatting (range-from range-into)
  (excel (xl)
;    (com-method range-from "Copy")
    #>(copy range-from)
    #>(pastespecial range-into +xl-paste-formats+)
    #>(pastespecial range-into +xl-paste-comments+)
    #>(pastespecial range-into +xl-paste-validation+)))
#|    (com-method range-into "PasteSpecial" +xl-paste-formats+)
    (com-method range-into "PasteSpecial" +xl-paste-comments+)
    (com-method range-into "PasteSpecial" +xl-paste-validation+)))|#







(defun test8 (handle)
  (excel (xl)
    (comlet* ((sheet (grab-worksheet handle))
              (from  (range sheet 1 1 1 6))
              (into  (range sheet 3 1 3 6)))
      (copy-formatting from into))))
      
    
#|(defun test7 (handle)
  (excel (xl)
    (let-range ((from (handle 2 1 2 30))
                (onto (handle 8 1 8 30)))
      (copy-formatting from onto))))|#


#|(defun test6 (handle)
  (excel (xl)
    (comlet* ((worksheet (grab-worksheet handle))
              (data      (get-range worksheet 1 1 2 30)))
      (set-range worksheet data 6 1 7 30))))


(defun test5 (handle)
  (excel (xl)
    (comlet* ((worksheet (grab-worksheet handle))
              (data      (get-range worksheet 1 1 2 30)))
      (set-range worksheet data 6 1 7 30))))


(defun test4 (handle)
  (excel (xl)
    (comlet* ((worksheet (grab-worksheet handle)))
      (get-range worksheet 2 1 2 30))))|#














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
