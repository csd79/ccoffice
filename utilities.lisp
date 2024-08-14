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


;;; Convert Excel column letters (given as keyword symbol) to number.
(defun letters-column (keyword)
  (let* ((string     (symbol-name keyword))
         (length     (length string))
         (substring  (if (> length 3)
                       (subseq string (- length 3))
                       string))
         (char-codes (loop for c across substring collecting
                           (- (char-code c) 64)))
         (char-values (mapcar #'* (nreverse char-codes)
                              '(1 26 676)))) 
    (apply #'+ char-values)))


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


;;; Construct new pathname based on DIRECTORY and FILE.
(defun file-in-dir (directory file &key (namestring nil))
  (let* ((filename (pathname-name file))
         (filetype (pathname-type file))
         (new      (make-pathname :name filename :type filetype
                                  :defaults directory)))
    (if namestring
      (namestring new)
      new)))


;;; ...
(defmacro flitmp (&body body)
  `(progn
     (fli:start-collecting-template-info)
     ,@body
     (fli:print-collected-template-info)))
