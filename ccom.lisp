;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Basic COM wrapper


;;; Is COM initialized?
(defparameter *com-initialized-p* nil)


;;; Binding context for COM operations.
(defmacro with-com-initialized (&body body)
  `(unwind-protect
       (progn
         (unless *com-initialized-p*
           (co-initialize))
         (let ((*com-initialized-p* t))
           ,@body))
     (unless *com-initialized-p*
       (co-uninitialize))))


;;; ----------------------------------------------------------------------
;;; Excel workbooks, worksheets


(defconstant +xl-calculation-automatic+ -4105)
(defconstant +xl-calculation-manual+    -4135)


;;; Set values for faster Excel operation.
(defun excel-speed-up (app)
  (setf #<(screenupdating app) nil
        #<(calculation app)    +xl-calculation-manual+
        #<(displayalerts app)  nil))


;; Set values for normal Excel behaviour.
(defun excel-slow-down (app)
  (setf #<(screenupdating app) t
        #<(calculation app)    +xl-calculation-automatic+
        #<(displayalerts app)  t))


;;; Open given document with default application, hopefully Excel.
(defun get-document (fullname)
  (get-object fullname :riid 'i-dispatch))


;;; Create a list of XLSX-HANDLE objects, each representing an open Excel workbook.
(defun open-workbooks (xl)
  (cclet* ((workbooks #<(workbooks xl))
           (count     #<(count workbooks)))
    (loop for i from 1 upto count collecting
          #<(fullname #<(item workbooks i)))))


;;; Return a list of every worksheet in every open Excel file.
(defun open-worksheets (xl)
  (let ((workbooks (open-workbooks xl))
        (results   '()))
    (dolist (workbook workbooks)
      (cclet* ((worksheets #<(worksheets (get-document workbook)))
               (count      #<(count worksheets)))
        (push workbook results)
        (push (loop for i from 1 upto count collecting
                    #<(name #<(item worksheets i)))
              results)))
    (nreverse results)))


;;; ----------------------------------------------------------------------
;;; Excel ranges, values


;; Create range object within given 'coordinates'.
(defun range (worksheet &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (let ((upper-left (cell-index x1 y1)))
    (if (and x2 y2)
      #<(range worksheet upper-left (cell-index x2 y2))
      #<(range worksheet upper-left))))


;;; ----------------------------------------------------------------------
;;; Excel cell formatting


;;; Constants used to copy cell formatting.
(defconstant +xl-paste-formats+    -4122)
(defconstant +xl-paste-comments+   -4144)
(defconstant +xl-paste-validation+     6)


;;; Copy cell formatting between ranges.
(defun copy-formatting (source destination)
  #>(copy source)
  #>(pastespecial destination +xl-paste-formats+)
  #>(pastespecial destination +xl-paste-comments+)
  #>(pastespecial destination +xl-paste-validation+))


;;; ----------------------------------------------------------------------
;;; Sandbox


(defun test ()
  (with-com-initialized
    (cclet* ((file    "c:\\Users\\cselovszkid\\Desktop\\lisp-book-structure-comparison.xlsx")
             (wbook   (get-document file))
             (wsheets #<(worksheets wbook))
             (wsheet  #<(item wsheets 1))
             (source  (range wsheet 1 1 1 5))
             (dest    (range wsheet 3 1 3 5))
             (furth   (range wsheet 1 100 1 105)))
      (copy-formatting source dest)
      #<(value2 dest)
      (setf #<(value2 furth) "Grr"))))


