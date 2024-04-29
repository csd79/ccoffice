;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; COM property and method calls


;;; Create dispatch function for '#p' and '#m'.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun com-dispatch-reader (invoke-function)
    (lambda (stream sub-char infix)
      (declare (ignore sub-char infix))
      (destructuring-bind (thing object &rest args)
          (read stream t nil t)
        `(,invoke-function ,object (symbol-name ',thing) ,@args)))))


;;; COM 'get property' read macro.
(set-dispatch-macro-character
 #\# #\p
 (com-dispatch-reader 'com::invoke-dispatch-get-property))


;; COM 'method call' read macro.
(set-dispatch-macro-character
 #\# #\m
 (com-dispatch-reader 'com::invoke-dispatch-method))


;;; ----------------------------------------------------------------------
;;; Binding interface pointers


;;; Generate binding clauses for CCLET*.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun cclet-bindings (bindings body)
    (if bindings
      (destructuring-bind (var expr)
          (first bindings)
        `(let ((,var ,expr))
           (if (typep ,var 'com::com-interface)
             ;; If EXPR returned an interface pointer, increase her reference count.
             (com::with-temp-interface (,var) ,var
               ,(cclet-bindings (rest bindings) body))
             ;; If EXPR returned an object of a different type, just carry on.
             ,(cclet-bindings (rest bindings) body))))
      body)))


;;; Local binding context for both interface pointers and plain Lisp values.
(defmacro cclet* (bindings &body body)
  (let ((enclosed `(progn ,@body)))
    (cclet-bindings bindings '() enclosed)))
