;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;; ----------------------------------------------------------------------
;; Basic COM wrapper


;; Is COM initialized?
(defparameter *com-initialized-p* nil)


;; Binding context for COM operations.
(defmacro with-com-initialized (&body body)
  `(unwind-protect
       (progn
         (unless *com-initialized-p*
           (co-initialize))
         (let ((*com-initialized-p* t))
           ,@body))
     (unless *com-initialized-p*
       (co-uninitialize))))


;; Generate binding forms for COMLET+.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun comhelper (params body)
    (if params
        (let* ((current (first params))
               (sym     (first current))
               (val     (second current)))
          `(let ((,sym ,val))
             (if (typep ,sym 'com-interface)
               (with-temp-interface (,sym) ,val
                 ,(comhelper (rest params) body))
               ,(comhelper (rest params) body))))
      `(progn ,@body))))


;; Locally bind both variables and interface pointers.
(defmacro comlet* (parameter-list &body body)
  (comhelper parameter-list body))


;; Set visibility for an application.
(defun set-app-visibility (var vis)
  (comlet* ((app #<(application var)))
    (setf #<(visible app) vis)))


(defun quit-app (app)
  #>(quit app))


;; Binding context to use given application's objects.
(defmacro with-app ((variable application &key (visible t) (observe-running nil)) &body body)
  (with-gensyms (running)
    `(with-com-initialized
       (comlet* ((,running  (get-active-object :progid ,application
                                               :riid   'i-dispatch
                                               :errorp nil))
                 (,variable (or ,running
                                (create-object :progid ,application))))
         (unwind-protect
             (progn
               (set-app-visibility ,variable ,visible)
               ,@body)
           (when ,observe-running
             (unless ,running
               (quit-app ,variable))))))))


;; ----------------------------------------------------------------------
;; Excel workbooks, worksheets


;; Describing Excel region of interest.
(defclass xlsx-handle ()
  ((fullname  :initarg :fullname  :accessor fullname)
   (name      :initarg :name      :accessor name)
   (sheetname :initarg :sheetname :accessor sheetname)))


(defconstant +xl-calculation-automatic+ -4105)
(defconstant +xl-calculation-manual+    -4135)


;; Set values for faster Excel operation.
(defun excel-speed-up (app)
  (setf #<(screenupdating app) nil
        #<(calculation app)    +xl-calculation-manual+
        #<(displayalerts app)  nil))


;; Set values for normal Excel behaviour.
(defun excel-slow-down (app)
  (setf #<(screenupdating app) t
        #<(calculation app)    +xl-calculation-automatic+
        #<(displayalerts app)  t))


;; Binding context for talking to Excel.
(defmacro excel ((xl &key (visible t) (observe-running nil)) &body body)
  `(with-app (,xl "Excel.Application" :visible ,visible :observe-running ,observe-running)
     (unwind-protect
         (progn 
           (excel-speed-up ,xl)
           ,@body)
       (excel-slow-down ,xl))))


;; Create a list of XLSX-HANDLE objects, each representing an open Excel workbook.
(defun open-workbooks ()
  (excel (xl)
    (comlet* ((workbooks #<(workbooks xl))
              (count     #<(count workbooks)))
      (loop for i from 1 upto count collecting
            (comlet* ((workbook    #<(item workbooks i))
                      (fullname    #<(fullname workbook))
                      (name        #<(name workbook))
                      (worksheets  #<(worksheets workbook))
                      (first-sheet #<(item worksheets 1))
                      (sheetname   #<(name first-sheet)))
              (make-instance 'xlsx-handle
                             :fullname  fullname
                             :name      name
                             :sheetname sheetname))))))


;; Return a list of the names of every worksheet in an open Excel file.
(defun worksheet-names (handle)
  (excel (xl)
    (comlet* ((workbooks #<(workbooks xl))
              (workbook  #<(item workbooks (name handle))))
      (when workbook
        (comlet* ((worksheets #<(worksheets workbook))
                  (count      #<(count worksheets)))
          (loop for i from 1 upto count collecting
                (comlet* ((worksheet #<(item worksheets i)))
                  #<(name worksheet))))))))


;; Create a copy of HANDLE.
(defun copy-xlsx-handle (handle)
  (make-instance 'xlsx-handle
                 :fullname  (fullname handle)
                 :name      (name handle)
                 :sheetname (sheetname handle)))


;; Return a list of every worksheet in every open Excel file.
(defun open-worksheets ()
  (let ((workbooks (open-workbooks))
        (results   '()))
    (dolist (handle workbooks)
      (dolist (sheetname (worksheet-names handle))
        (let ((new-handle (copy-xlsx-handle handle)))
          (setf (sheetname new-handle) sheetname)
          (push new-handle results))))
    (nreverse results)))


;; Get interface pointer for workbook described by XLSX.
(defun grab-workbook (handle)
  (excel (xl)
    (comlet* ((workbooks  #<(workbooks xl)))
      #<(item workbooks (slot-value handle 'name)))))


;; Get interface pointer for worksheet described by XLSX.
(defun grab-worksheet (handle)
  (excel (xl)
    (comlet* ((workbook   (grab-workbook handle))
              (worksheets #<(worksheets workbook)))
      #<(item worksheets (sheetname handle)))))


;; ----------------------------------------------------------------------
;; Excel ranges, values


;; Create range object within given 'coordinates'.
(defun range (worksheet &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (excel (xl)
    (let ((upper-left (cell-index x1 y1)))
      (if (and x2 y2)
          #<(range worksheet upper-left (cell-index x2 y2))
        #<(range worksheet upper-left)))))


;; Helper function to create binding clauses for LET-RANGE.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun let-range-helper (bindings)
    (when bindings
      (destructuring-bind (var (handle &optional x1 y1 x2 y2))
          (first bindings)
        (cons `(,var (range (grab-worksheet ,handle) ,x1 ,y1 ,x2 ,y2))
              (let-range-helper (rest bindings)))))))


;; Local range bindings.
(defmacro let-range (bindings &body body)
  `(comlet* ,(let-range-helper bindings)
     ,@body))

           
;; Get value(s) of given range.
(defun get-range (range)
  #<(value2 range))


;; Set value(s) of given range.
(defun set-range (range value)
  (setf #<(value2 range) value))


;; ----------------------------------------------------------------------
;; Excel cell formatting


;; Constants used to copy cell formatting.
(defconstant +xl-paste-formats+    -4122)
(defconstant +xl-paste-comments+   -4144)
(defconstant +xl-paste-validation+     6)


;; Copy cell formatting between ranges.
(defun copy-formatting (range-from range-into)
  (excel (xl)
    #>(copy range-from)
    #>(pastespecial range-into +xl-paste-formats+)
    #>(pastespecial range-into +xl-paste-comments+)
    #>(pastespecial range-into +xl-paste-validation+)))








(defun test (handle)
  (excel (xl)
    (let-range ((from (handle 1 1 1 10))
                (into (handle 3 1 3 10)))
      (set-range into
                 (get-range from))
      (copy-formatting from into))))
