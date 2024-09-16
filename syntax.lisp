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

(cclet* ((workbooks #p(workbooks global))             ; value of property 'workbooks' of object 'global'
         (workbook  #m(add workbooks +xl-worksheet+)) ; method 'add' on object 'workbooks' with arg '+xl-worksheet+'
         (sheets    #p(worksheets workbook))
         (worksheet #p(item sheets 1))                ; value of property 'item' of object 'sheets' with arg '1'
         (address   "A1")                             ; plain string
         (range     #p(range worksheet address)))
  (print #p(value2 range))
  (setf #p(name worksheet) "New name"))               ; modify value of property 'name' of object 'worksheet'
|#


;;; ----------------------------------------------------------------------
;;; COM property and method calls


;;; Create dispatch function for '#p' and '#m'.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun com-dispatch-reader (invoke-function)
    (lambda (stream sub-char infix)
      (declare (ignore sub-char infix))
      (destructuring-bind (thing object &rest args)
          (read stream t nil t)
        `(,invoke-function ,object (symbol-name ',thing) ,@args)))))


;;; COM 'get property' read macro.
(set-dispatch-macro-character
 #\# #\p
 (com-dispatch-reader 'com::invoke-dispatch-get-property))


;; COM 'method call' read macro.
(set-dispatch-macro-character
 #\# #\m
 (com-dispatch-reader 'com::invoke-dispatch-method))


;;; ----------------------------------------------------------------------
;;; Binding interface pointers


;;; Local binding context for both interface pointers and plain Lisp values.
(defmacro cclet* (bindings &body body)
  (when bindings
    (let ((vars (mapcar #'first bindings)))
      `(with-com-initialized
         (let* ,bindings
           (flet ((do-iptrs (fn &optional (order-fn #'identity))
                    (dolist (var (funcall order-fn ',vars))
                      (when (typep var 'com::com-interface)
                        (funcall fn var)))))
             (unwind-protect
                 (progn
                   (do-iptrs #'com:add-ref)
                   ,@body)
               (do-iptrs #'com:release #'reverse))))))))


#|;;; ----------------------------------------------------------------------
;;; Hidden app instances

(defparameter *running-instances* '())

(defun new-app-instance (progid)
  (cclet* ((ifp (com:create-object :progid progid)))
    (push (list progid ifp) *running-instances*)
    ifp))


(defun kill-running-instances ()
  (ignore-errors
    (dolist (rec *running-instances*)
      (com:invoke-dispatch-method (second rec) "Quit")))
  (setf *running-instances* '()))|#
