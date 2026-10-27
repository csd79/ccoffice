;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccoffice)
#.(enable-ccom-syntax)


;;; ----------------------------------------------------------------------
;;; Sandbox



(defparameter *t1* "c:\\Users\\cselovszkid\\common-lisp\\ccoffice\\teszt.xlsx")


(defun t1 ()
  (with-workbook (:open (namestring *t1*) :wsvars (ws1))
    (time (loop for r from 2 upto 10001 doing
                (print (xcell ws1 :b r :prop 'value))))))

(defun t2 ()
  (with-workbook (:open (namestring *t1*) :wsvars (ws1))
    (xrange ws1 2 2 2 6 :prop 'value)))

(defun t3 ()
  (with-workbook (:open (namestring *t1*) :wsvars (ws1) :save t)
    (setf (xrange ws1 3 2 3 6)
          #2A(("1") ("2") ("3") ("4") ("5")))))

(defun t4 ()
  (with-workbook (:open (namestring *t1*) :wsvars (ws1) :save t)
    (setf (xrange ws1 3 2 3 8)
          "Betmeeeeen")))

(defun t5 ()
  (with-workbook (:open (namestring *t1*) :wsvars (ws1))
    (xrange ws1 4 2 6 11 :prop 'value)))

(defun t6 ()
  (with-workbook (:open (namestring *t1*) :wsvars (ws1))
    (xcell ws1 4 2 :prop 'value)))




#|
(destructuring-bind (a b &optional c &rest args) expr
  (let ((op-name (symbol-name a)))
    (cond (;; !SYMBOL
           (and (> (length op-name) 1)
                (member (char op-name 0) '(#\! #\?) :test #'char=))
           (dispatch-expr (char op-name 0) b
                          (intern-by-name (subseq op-name 1) *package*)
                          (if c (cons c args) args)))
          (;; !'SYMBOL
           (and (member op-name '("!" "?") :test #'string=)
                (listp b)
                (eq (first b) 'quote))
           (dispatch-expr (char op-name 0) c b args))
          (;; Otherwise, return the expression.
           t expr))))
|#


;(defun tt ()
  






#.(disable-ccom-syntax)
