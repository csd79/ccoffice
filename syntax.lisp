;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;; ----------------------------------------------------------------------
;; COM syntactic sugar


#|
(com:with-temp-interface (workbook)
    (com:with-temp-interface (workbooks)
        (com:invoke-dispatch-get-property global "Workbooks")
      (com:invoke-dispatch-method workbooks "Add" +xl-worksheet+))
  (com:with-temp-interface (worksheet)
      (com:with-temp-interface (sheets)
          (com:invoke-dispatch-get-property workbook "Worksheets")
        (com:invoke-dispatch-get-property sheets "Item" 1))
    (let ((address "A1"))
      (get-value worksheet address))
    (setf (com:invoke-dispatch-get-property worksheet "Name") "New name")))

                |   |   |
                ¡   ¡   ¡

(comlet* ((workbooks #<(workbooks global))
          (workbook  #>(add workbooks +xl-worksheet+))
          (sheets    #<(worksheets workbook))
          (worksheet #<(item sheets 1))
          (address   "A1"))
  (get-value worksheet address)
  (setf #<(name worksheet) "New name"))
|#


;; #<(PROPERTY OBJECT)"
(defmacro comprop (property object &rest args)
  (let ((symbol-string (symbol-name property)))
    `(invoke-dispatch-get-property ,object ,symbol-string ,@args)))

(set-dispatch-macro-character
 #\# #\<
 (lambda (stream sub-char infix)
   (declare (ignore sub-char infix))
   (nconc (list 'comprop)
          (read stream))))


  ;; #>(METHOD OBJECT)"
(defmacro commethod (method object &rest args)
  (let ((symbol-string (symbol-name method)))
    `(invoke-dispatch-method ,object ,symbol-string ,@args)))

(set-dispatch-macro-character
 #\# #\>
 (lambda (stream sub-char infix)
   (declare (ignore sub-char infix))
   (nconc (list 'commethod)
          (read stream))))


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
