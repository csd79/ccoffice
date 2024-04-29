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


;;; Return s-expression read from INPUT as a string.
(defun preread-sexp (input &key (opening-char #\() (closing-char #\)))
  (with-output-to-string (output)
    (let ((depth 0))
      (loop for current = (read-char input) doing
            (progn
              (write-char current output)
              (when (char= current opening-char)
                (incf depth))
              (when (char= current closing-char)
                (decf depth)
                (when (< depth 1)
                  (loop-finish))))))))
