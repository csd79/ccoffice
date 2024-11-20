;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Sandbox


(defun h ()
  (with-workbook (:open "c:\\Users\\cselovszkid\\Desktop\\Ample Controls.xlsx"
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


(defun j ()
  (with-workbook (:open "c:\\Users\\cselovszkid\\Downloads\\EXPORT_kinevezés.XLSX"
                  :wsvars (in) :close t :read-only t)
    (with-used-range (in l u r b)
      (xrange in l u r b :prop 'value))))


(defun k ()
  (with-workbook (:open "c:\\Users\\cselovszkid\\Downloads\\EXPORT_kinevezés.XLSX"
                  :wsvars (in) :close t :read-only t)
    (cclet* ((r (range in 1 1 6 6)))
      (with-range (r left top right bottom)
        (format t "left: ~a~%top: ~a~%right: ~a~%bottom: ~a~%" left top right bottom)
        (let ((obj (read-xarray r)))
          (loop for c from 2 upto 4 doing
                (setf (xaref obj 0 c) c))
          (loop for c from 1 upto 5 doing
                (print (xaref obj 0 c)))
          (print "-----")
          (do-xarows (row r (xarows obj #(0 1 2 3 4)))
            (print (xcref row 1)))
          (print "-----")
          (xaselect obj #'(lambda (row)
                            (string= (xcref row "Vállalat hosszú megnevezése")
                                     "Sárospataki Tankerületi Központ")))
          (xauniques obj "Vállalat hosszú megnevezése")
          (xadouniques (v obj "Vállalat hosszú megnevezése")
            (print v)))))))


(defun l ()
  (with-workbook (:open "c:\\Users\\cselovszkid\\Downloads\\EXPORT_kinevezés.XLSX"
                  :wsvars (in) :close t :read-only t)
    (let* ((xarray (read-xarray (used-range in)))
           (sorted (xasort xarray (xapred ("Vállalat hosszú megnevezése" #'astring< #'astring=)
                                          ("SZK" #'astring< #'astring=)
                                          ("SZTSZ" #'astring<=)))))
      (do-xarows (row r sorted)
        (format t "~a    ~a    ~a~%"
                (xcref row "Vállalat hosszú megnevezése")
                (xcref row "SZK")
                (xcref row "SZTSZ"))))))














(defun m (wb)
  (with-workbook (:use wb :wsvars (in))
    (let* ((xarray (read-xarray (used-range in)))
           (sorted (xasort xarray (xapred ("Vállalat hosszú megnevezése" #'astring< #'astring=)
                                          ("SZK" #'astring< #'astring=)
                                          ("SZTSZ" #'astring<=)))))
      (do-xarows (row r sorted)
        (format t "~a    ~a    ~a~%"
                (xcref row "Vállalat hosszú megnevezése")
                (xcref row "SZK")
                (xcref row "SZTSZ"))))))

(defun n ()
  (with-workbook (:wbook w :open "c:\\Users\\cselovszkid\\Downloads\\EXPORT_kinevezés.XLSX"
                  :close t :read-only t)
    (m w)))





(defun p ()
  (let ((xa nil))
    (with-workbook (:open "C:\\Users\\cselovszkid\\Downloads\\2024.10.09. 3 feladat\\2 2024.10.01 statisztikák\\B2 2024.10.01. ének- és testnevelõtanár.xlsx"
                    :wsvars (in) :close t :read-only t)
      (setf xa (read-xarray (used-range in))))
    (let ((enek '("ének"))
          (tesi '("testnevel" "torna" "gyógype" "gyógyte" "gyógyto"))
          (enek-found '())
          (tesi-found '()))
      (flet ((init (string list)
               (when (stringp string)
                 (position-if #'(lambda (sample)
                                  (search sample string))
                              list))))
        (do-xarows (row r xa)
          (when (init (xcref row :f)
                      enek)
            (push r enek-found))
          (when (init (xcref row :g)
                      enek)
            (push r enek-found))
          (when (init (xcref row :f)
                      tesi)
            (push r tesi-found))
          (when (init (xcref row :g)
                      tesi)
            (push r tesi-found))))
      (values
       (remove-duplicates (nreverse enek-found))
       (remove-duplicates (nreverse tesi-found))))))

