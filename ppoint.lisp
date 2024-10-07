;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; ...


;;; Pedagógusok létszáma vármegyénként 2024.05.31.
#|(defparameter *ppt-template* "c:\\Users\\cselovszkid\\Downloads\\2024.07.09\\sablon.pptx")|#
#|(defparameter *ppt-control*  "c:\\Users\\cselovszkid\\Downloads\\2024.07.09\\Létszámadatok_köznevelési intézmények_2024.05.31._WIP.xlsx")|#

#|(defparameter *input-directory* "c:\\Users\\cselovszkid\\Downloads\\2024.08.01. Pedagógus túlórák 2023-24 tanév\\éles\\")|#

#|(defparameter *bek*
  '("2704" "2703" "2785" "2701" "3116" "3115" "2702" "2709" "2708"))|#


#|(defun test9 ()
  (let* ((input-files (directory (concatenate 'string *input-directory* "*.xls*")))
         (input-files-sorted (sort input-files #'string< :key #'namestring)))
    (cclet* ((excel    (com:create-object :progid "Excel.Application"))
             (wbooks   #p(workbooks excel))
             (wbook    #m(add wbooks))
             (wsheets  #p(worksheets wbook))
             (wsheet   #p(item wsheets 1))
             (dest-row 2))
      (unwind-protect
          (progn
            ;; Fejléc másolása
            (with-workbook (wbook-in :open-file (namestring (first input-files-sorted))
                                     :wsvars (wsheet-in) :close t)
              (with-used-edges (wsheet-in left top right bottom)
                (setf (xrange wsheet 1 1 right 1)
                      (xrange wsheet-in 1 1 right 1))))
            ;; Bemeneti fájlok
            (dolist (input-file input-files-sorted)
              (with-workbook (wbook-in :open-file (namestring input-file) :wsvars (wsheet-in) :close t)
                (format t "~a~%"  (namestring input-file))
                (with-used-edges (wsheet-in left top right bottom)
                  ;; IDX = bérelemek oszlopa
                  (let ((idx (column->row (xrange wsheet-in 3 1 3 bottom))))
                    (loop for i from 0 below (length idx)
                          for row = (1+ i) doing
                          ;; Ha a bérelem benne van *BEK*-ben és a háttere nem színes:
                          (when
                              (and (member (svref idx i) *bek* :test #'string=)
                                   (= #p(colorindex #p(interior (range wsheet-in 3 row))) -4142))
                            ;; Sor másolása
                            (setf (xrange wsheet 1 dest-row right dest-row)
                                  (xrange wsheet-in 1 row right row))
                            (format t ".")
                            (incf dest-row))))
                  (format t "~%~a~%" dest-row))))
            (autofit-cols wsheet)
            (freeze-panes wsheet 5 1)
            #m(saveas wbook (concatenate 'string (namestring *input-directory*) "hihi.xlsx")))
        (progn
          #m(close wbook)
          #m(quit excel))))))|#
