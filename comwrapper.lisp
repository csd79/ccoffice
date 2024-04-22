;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:csd79-lispworks-comwrapper)


;; ----------------------------------------------------------------------
;; Constants


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
(defconstant +xl-workbook-default+    51)
(defconstant +ol-mail-item+            0)
(defconstant +ol-by-value+             1)


;; ----------------------------------------------------------------------
;; Basic COM crapper


(defparameter *com* nil)


(defmacro define-fn-nickname (function symbol &key (setf-able nil))
  `(progn
     (intern (symbol-name ',symbol))
     (setf (fdefinition ',symbol) ,function)
     (when ,setf-able
       (setf (fdefinition '(setf ,symbol))
             #'(lambda (new-value &rest args)
                 (setf (apply ,function args)
                       new-value))))))


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


(defun cell-index (column row)
  (format nil "~a~d" (column column) row))

(defmacro with-com-initialized (&body body)
  `(unwind-protect
       (progn
         (unless *com*
           (co-initialize))
         (let ((*com* t))
           ,@body))
     (unless *com*
       (co-uninitialize))))


(define-fn-nickname #'invoke-dispatch-get-property com-property :setf-able t)
(define-fn-nickname #'invoke-dispatch-method com-method)


#|(eval-when (:load-toplevel :compile-toplevel)
  (defun comlet*-helper (parameter-list &optional (result '()))
    (if parameter-list
        (let ((current (first parameter-list)))
          (comlet*-helper (rest parameter-list)
                          `(with-temp-interface (,(first current))
                               ,(second current)
                               ,result)))
      result)))


(defmacro comlet* (parameter-list &body body)
  (comlet*-helper (nreverse parameter-list)
                  (append '(progn) body)))|#


(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun comhelper (params body)
    (if params
        (let ((current (first params)))
          `(with-temp-interface (,(first current))
                                ,(second current)
                                ,(comhelper (rest params) body)))
      `(progn ,@body))))


(defmacro comlet* (parameter-list &body body)
  (comhelper parameter-list body))


(defun set-app-visibility (var vis)
  (comlet* ((app (com-property var "Application")))
    (setf (com-property app "Visible") vis)))


(defparameter *observe-app-running* nil)


(defmacro with-app ((variable application) &body body)
  (with-gensyms (running)
    `(with-com-initialized
       (comlet* ((,running  (get-active-object :progid ,application
                                               :riid   'i-dispatch
                                               :errorp nil))
                  (,variable (or ,running
                                 (create-object :progid ,application))))
         (progn ,@body)))))
#|         (unwind-protect
             (progn
               (set-app-visibility ,variable t)
               ,@body)
           (when *observe-app-running*
             (unless ,running
               (com-method ,variable "Quit"))))))))|#


(defmacro excel ((variable) &body body)
  `(with-app (,variable "Excel.Application")
     ,@body))


(defmacro outlook ((variable) &body body)
  `(with-app (,variable "Outlook.Application")
     ,@body))


;; ----------------------------------------------------------------------
;; XLSX class


(defclass xlsx ()
  ((app       :initarg :app       :Accessor app)
   (file      :initarg :file      :accessor file)
   (workbook  :initarg :workbook  :accessor workbook)
   (worksheet :initarg :worksheet :accessor worksheet)))


(defun close-workbook (workbook)
  (setf (com-property workbook "Saved") t)
  (com-method workbook "Close" nil))


(defun save-and-close-workbook (workbook file file-exists-p)
  (if file-exists-p
      (com-method workbook "Save")
    (com-method workbook "SaveAs" file +xl-workbook-default+))
  (com-method workbook "Close" nil))


(defmacro with-xlsx ((xlsx &key file (read-only nil)) &body body)
  (with-gensyms (file-exists-p wbs wb shs sh)
    `(excel (xl)
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
               (if ,read-only
                   (close-workbook ,wb)
                 (save-and-close-workbook ,wb ,file ,file-exists-p)))))))))


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
     ,@body))


;; ----------------------------------------------------------------------
;; Styles


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


;; ----------------------------------------------------------------------
;; Sending mail


(defun body-from-file (file)
  (read-file-into-string
   file :external-format :utf-8))


(defun new-mail (from to cc subj body attch &optional (command :view))
  (outlook (ol)
    (comlet* ((msg (com-method ol "CreateItem" +ol-mail-item+))
               (att (com-property msg "Attachments")))
      (mapc #'(lambda (pair) (setf (com-property msg (first pair)) (second pair)))
            `(("SentOnBehalfOfName" ,from)
              ("To"                 ,to)
              ("CC"                 ,cc)
              ("Subject"            ,subj)
              ("HTMLBody"           ,(body-from-file body))))
      (let ((attachments (if (atom attch)
                             (list attch)
                           attch)))
        (dolist (a attachments)
          (com-method att "Add" a +ol-by-value+)))
      (cond ((eq command :view) (com-method msg "Display"))
            ((eq command :send) (com-method msg "Send"))
            (t msg)))))
