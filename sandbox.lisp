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




#.(disable-ccom-syntax)
