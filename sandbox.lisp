;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Sandbox


(defun k ()
  (with-workbook (:open-file "c:\\Users\\cselovszkid\\Desktop\\Ample Controls.xlsx"
                  :wsvars (wsheet) :save t :close t)
    (cclet* ((write  (range wsheet 1 100 1 105))
             (source (range wsheet 1 1 1 5))
             (dest   (range wsheet 9 1 9 5)))
      (with-used-range (wsheet left top right bottom)
        (format t "left: ~a, top: ~a, right: ~a, bottom: ~a~%" left top right bottom))
      (copy-formatting source dest)
      (setf #~('value2 write) "Grr")
      (apply-style (range wsheet 2 10 2 14) '(:border))
      (apply-style (font (range wsheet 1 11 1 14)) '(:bold))
      ;; Print formula & value.
      (format t "~a   =   ~a~%"
              (xcell wsheet 8 2 :prop 'formula)
              (xcell wsheet 8 2 :prop 'value))
      ;; Set formulas
      (setf (xrange wsheet 2 2 4 4 :prop 'formula) "=\"\"")
      (setf (xcell wsheet 8 8 :prop 'formula) "=8*8")
      )))


(defun l ()
  (with-workbook (:open-file "c:\\Users\\cselovszkid\\Downloads\\EXPORT_kinevezés.XLSX"
                  :wsvars (in) :close t :read-only t)
    (with-used-range (in l u r b)
      (xrange in l u r b :prop 'value))))
