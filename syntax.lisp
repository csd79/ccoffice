;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


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

(cclet* ((workbooks #~('workbooks global))            ; value of property 'workbooks' of object 'global'
         (workbook  (#_add workbooks +xl-worksheet+)) ; method 'add' on object 'workbooks' with arg '+xl-worksheet+'
         (sheets    #~('worksheets workbook))
         (worksheet #~('item sheets 1))               ; value of property 'item' of object 'sheets' with arg '1'
         (address   "A1")                             ; plain string
         (range     #~('range worksheet address)))
  (print #~('value2 range))
  (setf #~('name worksheet) "New name"))              ; modify value of property 'name' of object 'worksheet'
|#


;;; ----------------------------------------------------------------------
;;; COM property and method calls


;;; COM 'get property' read macro.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun com-property-reader (stream sub-char infix)
    (declare (ignore sub-char infix))
    (destructuring-bind (thing object &rest args)
        (read stream t nil t)
      `(com:invoke-dispatch-get-property ,object (symbol-name ,thing) ,@args))))

(set-dispatch-macro-character #\# #\~ #'com-property-reader)


;; COM 'method call' read macro.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun com-method-reader (stream sub-char infix)
    (declare (ignore sub-char infix))
    (let ((method (read stream t nil t)))
      `(lambda (obj &rest rest)
         (apply #'com:invoke-dispatch-method obj (symbol-name ',method) rest)))))

(set-dispatch-macro-character #\# #\_ #'com-method-reader)


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
