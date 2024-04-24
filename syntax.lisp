;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; COM syntactic sugar


#|
Comparison:
 
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

             ||     ||     ||
             \/     \/     \/

(cclet* ((workbooks #<(workbooks global))             ; value of property 'workbooks' of object 'global'
         (workbook  #>(add workbooks +xl-worksheet+)) ; method 'add' on object 'workbooks' with arg '+xl-worksheet+'
         (sheets    #<(worksheets workbook))
         (worksheet #<(item sheets 1))                ; value of property 'item' of object 'sheets' with arg '1'
         (address   "A1")                             ; plain string
         (range     #<(range worksheet address)))
  (print #<(value2 range))
  (setf #<(name worksheet) "New name"))               ; modify value of property 'name' of object 'worksheet'
|#


;;; Create dispatch functions for '#<' and '#>'.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun com-dispatch-function (invoke-function)
    (lambda (stream sub-char infix)
      (declare (ignore sub-char infix))
      (destructuring-bind (thing object &rest args)
          (read stream)
        `(,invoke-function ,object (symbol-name ',thing) ,@args)))))


;;; COM 'get property' read macro.
(set-dispatch-macro-character
 #\# #\<
 (com-dispatch-function 'invoke-dispatch-get-property))


;; COM 'method call' read macro.
(set-dispatch-macro-character
 #\# #\>
 (com-dispatch-function 'invoke-dispatch-method))


;;; Generate binding clauses for CCLET*.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun cclet-bindings (bindings body)
    (if bindings
      (destructuring-bind (var expr)
          (first bindings)
        `(let ((,var ,expr))
           (if (typep ,var 'com-interface)
             ;; If EXPR returned an interface pointer, increase her reference count.
             (with-temp-interface (,var) ,expr
               ,(cclet-bindings (rest bindings) body))
             ;; If EXPR returned an object of a different type, just carry on.
             ,(cclet-bindings (rest bindings) body))))
      body)))


;;; Local binding context for both interface pointers and plain Lisp values.
(defmacro cclet* (bindings &body body)
  (let ((enclosed `(progn ,@body)))
    (cclet-bindings bindings enclosed)))

