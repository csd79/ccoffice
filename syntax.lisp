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

(cclet* ((workbooks #p(workbooks global))             ; value of property 'workbooks' of object 'global'
         (workbook  #m(add workbooks +xl-worksheet+)) ; method 'add' on object 'workbooks' with arg '+xl-worksheet+'
         (sheets    #p(worksheets workbook))
         (worksheet #p(item sheets 1))                ; value of property 'item' of object 'sheets' with arg '1'
         (address   "A1")                             ; plain string
         (range     #p(range worksheet address)))
  (print #p(value2 range))
  (setf #p(name worksheet) "New name"))               ; modify value of property 'name' of object 'worksheet'
|#


#|(eval-when (:load-toplevel :compile-toplevel :execute)
  ;;; Random string of LENGTH made up of lowercase letters.
  (defun random-string (length)
    (concatenate 'string (loop for i from 0 below length collecting
                               (code-char (+ (random 26) 97)))))

  ;;; Random string of length LENGTH that doesn't appear in EXCLUSION-STRING.
  (defun unique-prefix (exclusion-string &key (length 3) (retry-limit 1000))
    (if (zerop retry-limit)
      (error "UNIQUE-PREFIX: Too many retries.")
      (let ((random-string (random-string length)))
        (if (search random-string exclusion-string)
          (unique-prefix exclusion-string :length length :retry-limit (1- retry-limit))
          random-string))))

  ;;; Find all occurences of SUBs in STRING, and remove any trailing whitespaces.
  (defun remove-trailing-whitespaces (string subs)
    (if subs
      (let ((current (first subs)))
        (remove-trailing-whitespaces
         (cl-ppcre:regex-replace-all (format nil "(?i)~a\\s+\\S" current) string current)
         (rest subs)))
      string))

  ;;; Generate reader function for #\p and #\m.
  (defun com-dispatch-function (invoke-macro)
    (lambda (input sub-char infix)
      (declare (ignore sub-char infix))
      (let* ((raw-input  (preread-sexp input))
             (ignorables (list "," "`"))
             (prefixes   (loop for i from 0 below (length ignorables) collecting
                               (unique-prefix raw-input)))
             (safe-input (replace-substrings raw-input ignorables prefixes)))



        (destructuring-bind (thing object &rest args)
            (read-from-string safe-input)
          (let* ((transformed   (write-to-string `(,invoke-fn ,object (symbol-name ',thing) ,@args)))
                 (unsafe-output (replace-substrings transformed prefixes ignorables))
                 (safe-output   (remove-trailing-whitespaces unsafe-output ignorables))
                 (proper-input  (make-string-input-stream safe-output)))
;            (print safe-output)
            (read proper-input)))))))|#


(defmacro com-method (method obj &rest args)
  `(com::invoke-dispatch-method ,obj (symbol-name ,method) ,@args))

(defmacro com-property (property obj &rest args)
  `(com::invoke-dispatch-get-property ,obj (symbol-name ,property) ,@args))


(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun whitespace-char-p (x)
    (or (char= #\space x)
        (not (graphic-char-p x))))

  (defun com-dispatch-reader (augment)
    (lambda (input subchar infix)
      (declare (ignore subchar infix))
      (let* ((raw     (preread-sexp input))
             (no-lead (subseq raw (position-if-not #'whitespace-char-p raw :start 1)))
             (rewrite (format nil "(~a '~a" augment no-lead))
             )
;             (read-fn (get-macro-character #\()))
        (with-input-from-string (stream rewrite)
;          (funcall read-fn stream nil))))))
;        (read-from-string rewrite))))))
          (print (get-macro-character #\`))
          (print (get-macro-character #\,))
          (read stream))))))
    


#|(set-macro-character #\< (lambda (input char)
                           (declare (ignore char))
                           (let* ((raw     (preread-sexp input))
                                  (rewrite (format nil "(com-property ~a" (subseq raw 1))))
                             (read-from-string rewrite))))
(set-macro-character #\> (lambda (input char)
                           (declare (ignore char))
                           (let* ((raw     (preread-sexp input))
                                  (rewrite (format nil "(com-method ~a" (subseq raw 1))))
                             (read-from-string rewrite))))|#


#|(set-macro-character #\< (com-dispatch-reader "com-property"))
(set-macro-character #\> (com-dispatch-reader "com-method"))|#


;;; COM 'get property' read macro.
(set-dispatch-macro-character
 #\# #\p
 (com-dispatch-reader "com-property"))

;; COM 'method call' read macro.
(set-dispatch-macro-character
 #\# #\m
 (com-dispatch-reader "com-method"))







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

