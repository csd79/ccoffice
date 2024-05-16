;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Utilities


;;; Convert Excel column number to letter-based address.
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


;;; Create Excel-style cell address from numerical address.
(defun cell-index (column row)
  (format nil "~a~d" (column column) row))


;;; Is COM initialized?
(defparameter *com-init-count* 0)


;;; Binding context for COM operations.
(defmacro with-com-initialized (&body body)
  `(unwind-protect
       (progn
         (when (zerop *com-init-count*)
           (com::co-initialize))
         (incf *com-init-count*)
         ,@body)
     (progn
       (decf *com-init-count*)
       (when (zerop *com-init-count*)
         (com::co-uninitialize)))))
