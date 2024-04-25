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




#|(defun replace-subseq-all (old new sequence &key (stack '()) (test #'string=) (start 0))
  (let ((position (search old sequence :test test :start2 start)))
    (if position
      (replace-subseq-all
       old new sequence :test test :start (+ position (length old))
       :stack (cons new (cons (subseq sequence start position) stack)))
      (let ((full (cons (subseq sequence start) stack)))
        (apply #'concatenate (type-of (first stack)) (nreverse full))))))
(defun subseq-replacements (sequence oldlist newlist &key (test #'string=))
  (let ((current sequence))
    (mapc #'(lambda (old new)
              (setf current (replace-subseq-all old new current :test test)))
          oldlist newlist)
    current))
(defun string-replace-all (string old new &optional (testfn #'string-equal))
  (let ((position (search old string :test testfn)))
    (if position
      (string-replace-all (concatenate 'string (subseq string 0 position)
                                       new (subseq string (+ position (length old))))
                          old new testfn)
      string)))
(defun replace-substrings (string olds news)
  (if olds
    (replace-substrings
     (string-replace-all string (first olds) (first news))
     (rest olds) (rest news))
    string))|#


(defun replace-substrings (string olds news)
  (if olds
    (replace-substrings (cl-ppcre:regex-replace-all (first old) string (first new))
                        (rest olds)
                        (rest news))
    string))


(defun preread-sexp (input &key (opening-char #\() (closing-char #\)))
  (with-output-to-string (output)
    (let ((depth 0))
      (loop for next = (peek-char nil input t nil t) doing
            (progn
              (when (char= next opening-char)
                (incf depth))
              (when (char= next closing-char)
                (decf depth)
                (when (< depth 1)
                  (loop-finish))))))))


(defun random-string (length)
  (concatenate 'string (loop for i from 0 below length collecting
                             (code-char (+ (random 26) 97)))))


(defun unique-prefix (exclusion-string &key (length 3) (retry-limit 1000))
  (if (zerop retry-limit)
    (error "UNIQUE-PREFIX: Too many retries.")
    (let ((random-string (random-string length)))
      (if (search random-string exclusion-string)
        (unique-prefix exclusion-string :length length :retry-limit (1- retry-limit))
        random-string))))


#|(defun whitespace-char-p (x)
  (or (char= #\space x)
      (not (graphic-char-p x))))
(defun remove-trailing-whitespaces (string substring &optional (result ""))
  (let ((substring-length (length substring))
        (substring-from   (search substring string)))
    (if substring-from
      (let* ((substring-upto (+ substring-from substring-length))
             (next-valid     (position-if-not #'whitespace-char-p string :start substring-upto))
             (ending         (if next-valid
                               (subseq string next-valid)
                               "")))
        (format t "~a~%" ending)
        (remove-trailing-whitespaces 
         ending
         substring
         (concatenate 'string result
                      (subseq string 0 substring-upto)))
      result))))|#


(defun remove-trailing-whitespaces (string subs)
  (if subs
    (let ((current (first subs)))
      (remove-trailing-whitespaces
       (cl-ppcre:regex-replace-all (format nil "~a\\s+\\S" current) string current)
       (rest subs)))
    string))



(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun com-dispatch-function (invoke-fn)
    (lambda (input sub-char infix)
      (declare (ignore sub-char infix))
      (let* ((raw-input  (read-delimited-string input :opening-char #\( :closing-char #\)))
             (ignorables (list "," "`"))
             (prefixes   (loop for i from 0 below (length ignorables) collecting
                               (unique-prefix raw-input)))
#|              (mapcar #'(lambda (ignore)
                                     (declare (ignore ignore))
                                     (unique-prefix raw-input))
                                 ignorables))|#
             (safe-input (replace-substrings raw-input ignorables prefixes)))
        (destructuring-bind (thing object &rest args)
            (read-from-string safe-input)
          (let* ((transformed   (write-to-string `(,invoke-fn ,object (symbol-name ',thing) ,@args)))
                 (unsafe-output (replace-substrings transformed prefixes ignorables))
                 (safe-output   (remove-trailing-whitespaces unsafe-output ignorables))
#|                  (replace-substrings unsafe-output
                                                    (mapcar #'(lambda (sub)
                                                                (concatenate 'string sub " "))
                                                            ignorables)
                                                    ignorables))|#
;                 (reader-fn    (get-macro-character #\`))
                 (proper-input (make-string-input-stream safe-output)))
            (print safe-output)
;            (eval (funcall reader-fn proper-input nil))))))))
;            (funcall reader-fn proper-input nil)))))))
            (read proper-input)))))))


;;; Generate dispatch function for '#<' and '#>'.
#|(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun com-dispatch-function (invoke-function)
    (lambda (stream sub-char infix)
      (declare (ignore sub-char infix))
      (destructuring-bind (thing object &rest args)
          (read stream)
        `(,invoke-function ,object (symbol-name ',thing) ,@args)))))|#


(defun tt ()
  (let ((obj 'grr))
    `(outer (prop ,obj 1))))



;;; COM 'get property' read macro.
#|(set-dispatch-macro-character
 #\# #\<
 (com-dispatch-function 'com::invoke-dispatch-get-property))|#
(set-dispatch-macro-character
 #\# #\p
 (com-dispatch-function 'com::invoke-dispatch-get-property))


;; COM 'method call' read macro.
#|(set-dispatch-macro-character
 #\# #\>
 (com-dispatch-function 'com::invoke-dispatch-method))|#
(set-dispatch-macro-character
 #\# #\m
 (com-dispatch-function 'com::invoke-dispatch-method))


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
    (cclet-bindings bindings enclosed)))

