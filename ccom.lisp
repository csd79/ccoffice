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
           (com::co-initialize))
         (let ((*com-initialized-p* t))
           ,@body))
     (unless *com-initialized-p*
       (com::co-uninitialize))))


;;; ----------------------------------------------------------------------
;;; Excel workbooks, worksheets


(defconstant +xl-calculation-automatic+ -4105)
(defconstant +xl-calculation-manual+    -4135)


;;; Set values for faster Excel operation.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun excel-speed-up (app)
    (setf #p(screenupdating app) nil
          #p(calculation app)    +xl-calculation-manual+
          #p(displayalerts app)  nil)))


;; Set values for normal Excel behaviour.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun excel-slow-down (app)
    (setf #p(screenupdating app) t
          #p(calculation app)    +xl-calculation-automatic+
          #p(displayalerts app)  t)))


;; Get parrent application of OBJECT.
(eval-when (:load-toplevel :compile-toplevel :execute)
  (defun get-application (object)
    #p(application object)))


;;; Faster exchange with Excel.
(defmacro excellerate ((object) &body body)
  (let ((excel (gensym)))
    `(cclet* ((,excel (com-property 'application ,object)))
       (unwind-protect
           (progn
             (excel-speed-up ,excel)
             ,@body)
         (excel-slow-down ,excel)))))



;;; Open given document with default application, hopefully Excel.
(defun get-document (fullname)
  (com::get-object fullname :riid 'com::i-dispatch))


;;; Create a list of XLSX-HANDLE objects, each representing an open Excel workbook.
#|(defun open-workbooks (xl)
  (cclet* ((workbooks #p(workbooks xl))
           (count     #p(count workbooks)))
    (loop for i from 1 upto count collecting
          #p(fullname #p(item workbooks i)))))


;;; Return a list of every worksheet in every open Excel file.
(defun open-worksheets (xl)
  (let ((workbooks (open-workbooks xl))
        (results   '()))
    (dolist (workbook workbooks)
      (cclet* ((worksheets #p(worksheets (get-document workbook)))
               (count      #p(count worksheets)))
        (push workbook results)
        (push (loop for i from 1 upto count collecting
                    #p(name #p(item worksheets i)))
              results)))
    (nreverse results)))|#


;;; ----------------------------------------------------------------------
;;; Excel ranges, values


;; Create range object within given 'coordinates'.
(defun range (worksheet &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (let ((upper-left (cell-index x1 y1)))
    (if (and x2 y2)
      #p(range worksheet upper-left (cell-index x2 y2))
      #p(range worksheet upper-left))))


(defconstant +xl-to-left+          -4159)
(defconstant +xl-up+               -4162)
(defconstant +xl-values+           -4163)
(defconstant +xl-by-rows+              1)
(defconstant +xl-by-columns+           2)
(defconstant +xl-whole+                1)
(defconstant +xl-previous+             2) 


;;; Locate used range of WORKSHEET.
(defun used-range (worksheet)
  (cclet* ((whole  #p(cells worksheet))
           (top    #p(row #p(end whole +xl-up+)))
           (left   #p(column #p(end whole +xl-to-left+)))
           (from   (range worksheet left top))
           (bottom #p(row #m(find whole "*" from +xl-values+
                                  +xl-whole+ +xl-by-rows+ +xl-previous+)))
           (right  #p(column #m(find whole "*" from +xl-values+
                                     +xl-whole+ +xl-by-columns+ +xl-previous+))))
    (range worksheet left top right bottom)))


;;; ----------------------------------------------------------------------
;;; Excel cell formatting


;;; Constants used to copy cell formatting.
(defconstant +xl-paste-formats+    -4122)
(defconstant +xl-paste-comments+   -4144)
(defconstant +xl-paste-validation+     6)


;;; Copy cell formatting between ranges.
(defun copy-formatting (source destination)
  #m(copy source)
  #m(pastespecial destination +xl-paste-formats+)
  #m(pastespecial destination +xl-paste-comments+)
  #m(pastespecial destination +xl-paste-validation+))


;;; ----------------------------------------------------------------------
;;; Sandbox


(defun test ()
  (with-com-initialized
    (cclet* ((file    "c:\\Users\\cselovszkid\\Desktop\\lisp-book-structure-comparison.xlsx")
             (wbook   (get-document file))
             (wsheets #p(worksheets wbook))
             (wsheet  #p(item wsheets 1))
             (source  (range wsheet 1 1 1 5))
             (dest    (range wsheet 3 1 3 5))
             (furth   (range wsheet 1 100 1 105)))
      (copy-formatting source dest)
      #p(value2 dest)
      (setf #p(value2 furth) "Grr"))))


(defun test2 ()
  (with-com-initialized
    (cclet* ((file    "c:\\Users\\cselovszkid\\Desktop\\lisp-book-structure-comparison.xlsx")
             (wbook   (get-document file))
             (wsheets #p(worksheets wbook))
             (wsheet  #p(item wsheets 1))
             (used    (used-range wsheet)))
      #p(value2 used))))
