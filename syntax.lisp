;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


#|
VBA:
Workbooks = Global.Workbooks
Workbook  = Workbooks.Add(template := xlWorksheet)
Worksheet = Workbook.Worksheets(1)
Print(Worksheet.Range("A1"))
Worksheet.Name = "New Name"


LW:
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

(cclet* ((workbooks (?workbooks global))
         (workbook  (!add workbooks +xl-worksheet+))
         (worksheet (?item (?worksheets workbook) 1))
         (address   "A1"))
  (print (?value2 (?range worksheet address)))
  (setf (?name worksheet) "New name"))
|#


;;; ----------------------------------------------------------------------
;;; COM property and method calls


;(make-package :ccom-accessors :use nil)

;;; COM 'get property' read macro.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defparameter *ccom-accessor-refix* "CCOM-ACCESSOR-")

  (defun intern-ccom-accessor (property-name package)
    (let ((symbol-name (string-upcase (concatenate 'string *ccom-accessor-refix* property-name))))
      (shadow symbol-name package)
      (find-symbol symbol-name package)))

  (defun decompose-lw-setf-symbol (string)
    (let* ((start 0)
           (quotes (loop for i from 0 below 4
                         for new = (position #\" string :test #'char= :start start)
                         when new do (setf start (1+ new))
                         when new collect new)))
      (when (= (length quotes) 4)
        (list (subseq string (1+ (first quotes)) (second quotes))
              (subseq string (1+ (third quotes)) (fourth quotes))))))
  
  (defun error-symbol (error)
    (let ((datum (cell-error-name error)))
      (cond ((symbolp datum)
             (let ((symbol-name (symbol-name datum)))
               (if (find #\" symbol-name :test #'char=)
                 ;; 'SETF::\"CCOM\"\ \"VALUE2\"
                 (append (decompose-lw-setf-symbol symbol-name)
                         (list 'setf))
                 ;; 'VALUE2
                 (list (package-name (symbol-package datum))
                       (symbol-name datum)
                       nil))))
            ;; '(SETF VALUE2)
            ((and (listp datum)
                  (= (length datum) 2)
                  (eql (first datum) 'setf))
             (list (package-name (symbol-package (second datum)))
                   (symbol-name (second datum))
                   'setf))
            (t (list nil nil)))))

  (set-macro-character
   #\?
   (lambda (stream char)
     (declare (ignore char))
     (let* ((prop      (read stream t nil t))
            (prop-name (symbol-name prop)))
       (intern-ccom-accessor prop-name *package*)))))


(defparameter *property-accessors-on* nil)

(defun property-accessors-on ()
  *property-accessors-on*)

(defun (setf property-accessors-on) (value)
  (setf *property-accessors-on* (not (not value))))

(defmacro with-property-accessors (&body body)
  `(handler-bind
       ((undefined-function
         #'(lambda (error)
             (when *property-accessors-on*
               (destructuring-bind (package-name symbol-name type)
                   (error-symbol error)
                 (when (and symbol-name
                            (> (length symbol-name) 15)
                            (string= (subseq symbol-name 0 14) *ccom-accessor-refix*))
                   (let* ((property-name   (subseq symbol-name 14))
                          (property-setter (concatenate 'string "SET-" property-name))
                          (getter-symbol   (intern-ccom-accessor property-name package-name))
                          (setter-symbol   (intern-ccom-accessor property-setter package-name))
                          (getter-func     #'(lambda (obj &rest args)
                                               (apply #'com:invoke-dispatch-get-property
                                                      obj property-name args)))
                          (setter-func     #'(lambda (value obj &rest args)
                                               (apply #'com:invoke-dispatch-put-property
                                                      obj property-name
                                                      (nconc args (list value))))))
                     (setf (fdefinition getter-symbol) getter-func
                           (fdefinition setter-symbol) setter-func
                           (fdefinition `(setf ,getter-symbol)) (fdefinition setter-symbol))
                     (invoke-restart (find-restart 'use-value error)
                                     (if (eq type 'setf)
                                       setter-func
                                       getter-func)))))))))
     ,@body))


;; COM 'method call' read macro.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (set-macro-character
   #\!
   (lambda (stream char)
     (declare (ignore char))
     (let ((method (read stream t nil t)))
       `(lambda (obj &rest rest)
          (apply #'com:invoke-dispatch-method obj (symbol-name ',method) rest))))))


;;; ----------------------------------------------------------------------
;;; Binding interface pointers


;;; Is THING an interface pointer?
(defun live-iptr-p (thing)
  (and thing
       (not (eq thing :nothing))
       (typep thing 'com::com-interface)))


;;; Local binding context for both interface pointers and plain Lisp values.
(defmacro cclet* (bindings &body body)
  (if bindings
    (let ((vars (mapcar #'first bindings)))
      `(with-com-initialized
         (let* ,bindings
           (flet ((do-iptrs (fn &optional (order-fn #'identity))
                    (dolist (var (funcall order-fn ',vars))
                      (when (live-iptr-p var)
                        (funcall fn var)))))
             (unwind-protect
                 (progn
                   ,@body)
               (do-iptrs #'com:release #'reverse))))))
    `(progn
       ,@body)))
