;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;; ----------------------------------------------------------------------
;; COM syntactic sugar


;; #<(PROPERTY OBJECT)"
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defmacro comprop (property object &rest args)
    (let ((symbol-string (symbol-name property)))
      `(invoke-dispatch-get-property ,object ,symbol-string ,@args))))


(eval-when (:load-toplevel :compile-toplevel :execute)
  (set-dispatch-macro-character
   #\# #\<
   (lambda (stream sub-char infix)
     (declare (ignore sub-char infix))
     (nconc (list 'comprop)
            (read stream)))))


#|(defmacro comprop (property object &rest args)
  (let ((symbol-string (symbol-name property)))
    `(invoke-dispatch-get-property ,object ,symbol-string ,@args)))


(set-dispatch-macro-character
 #\# #\<
 (lambda (stream sub-char infix)
   (declare (ignore sub-char infix))
   (nconc (list 'comprop)
          (read stream))))|#


  ;; #>(METHOD OBJECT)"
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defmacro commethod (method object &rest args)
    (let ((symbol-string (symbol-name method)))
      `(invoke-dispatch-method ,object ,symbol-string ,@args))))


(eval-when (:load-toplevel :compile-toplevel :execute)
  (set-dispatch-macro-character
   #\# #\>
   (lambda (stream sub-char infix)
     (declare (ignore sub-char infix))
     (nconc (list 'commethod)
            (read stream)))))


#|(defmacro commethod (method object &rest args)
  (let ((symbol-string (symbol-name method)))
    `(invoke-dispatch-method ,object ,symbol-string ,@args)))


(set-dispatch-macro-character
 #\# #\>
 (lambda (stream sub-char infix)
   (declare (ignore sub-char infix))
   (nconc (list 'commethod)
          (read stream))))|#
