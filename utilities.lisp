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


;;; Convert Excel column letters (given as keyword symbol) to 1-based numerical index.
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
  `(let ((orig *com-init-count*))
     (unwind-protect
         (progn 
           (when (zerop *com-init-count*)
             (com::co-initialize))
           (incf *com-init-count*)
           ,@body)
       (progn
         (setf *com-init-count* orig)
         (when (zerop *com-init-count*)
           (com::co-uninitialize))))))


;;; Construct new pathname based on DIRECTORY and FILE.
(defun file-in-dir (directory file &key (namestring nil))
  (let* ((filename (pathname-name file))
         (filetype (pathname-type file))
         (new      (make-pathname :name filename :type filetype
                                  :defaults directory)))
    (if namestring
      (namestring new)
      new)))


;;; After #\# #\| has been found by caller, burn through STREAM until #\| #\# is found.
;;; Leave the next character in STREAM.
(defun burn-block-comment (stream)
  ;; Find closing #\|.
  (peek-char #\| stream t nil t)
  (read-char stream t nil t)
  (let ((after (peek-char nil stream t nil t)))
    ;; If the char after is #\#, read it and return.
    (if (char= after #\#)
      (read-char stream t nil t)
      (progn
        ;; If not, continue searching for closing #\| #\#.
        (when (char= after #\|)
          (read-char stream t nil t))
        (burn-block-comment stream)))))


;;; Peek the next, non-comment character from STREAM.
(defun peek-code-char (stream type)
  (let ((next (peek-char type stream t nil t)))
    (cond ((char= next #\;)
           ;; Comment until the end of the line.
           (read-line stream t nil t)
           (peek-code-char stream type))
          ((char= next #\#)
           ;; Maybe block comment.
           (read-char stream t nil t)
           (let ((second-next (peek-char type stream t nil t)))
             (if (char= second-next #\|)
               ;; Confirmed block comment.
               (progn
                 (read-char stream t nil t)
                 (burn-block-comment stream)
                 (peek-code-char stream type))
               ;; Not a comment.
               (progn
                 (unread-char #\# stream)
                 next))))
          (t
           ;; Not a comment.
           next))))


;;; Keep poping chars from STREAM while they correspond to STRING.
(defun peek-string= (stream string)
  (let* ((length (length string))
         ;; Read characters while they correspond with STRING's characters.
         (same   (loop for i from 0 below length
                       for c = (if (zerop i)
                                 (peek-code-char stream t) ; To reach first char, skip any whitespace
                                 (peek-code-char stream nil))
                       while (char= c (char string i))
                       collecting c
                       doing (read-char stream t nil t))))
    ;; Put all characters read back into STREAM.
    (unless (zerop (length same))
      (loop for i from (1- length) downto 0 doing
            (unread-char (elt same i) stream)))
    ;; Were all chars the same?
    (string= string
             (coerce same 'string))))


;;; Collect and print the FLI templates needed to evaluate BODY.
(defmacro print-used-fli-templates (&body body)
  `(progn
     (fli:start-collecting-template-info)
     ,@body
     (fli:print-collected-template-info)))
