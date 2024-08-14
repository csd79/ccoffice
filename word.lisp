;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Documents


;;; Create bindings for an Excel workbook with optional method calls
;;; (open, save at the end, close).
(defmacro with-document ((doc &key (open-file nil)
                                   (app       'word)
                                   (read-only nil)
                                   (close     t)
                                   (save      nil))
                         &body body)
    `(ccom:cclet* ((,app (com:create-object :progid "Word.Application"))
                   (docs #p(documents ,app))
                   (,doc (if ,open-file
                           #m(open docs ,open-file nil ,read-only)
                           #m(add docs))))
       (unwind-protect
           (progn
             (setf #p(screenupdating ,app) nil)
             ,@body)
         (progn
           (setf #p(screenupdating ,app) t)
           (when (and ,save (not ,read-only))
             #m(save ,doc))
           (when ,close
             (setf #p(saved ,doc) t)
             #m(close ,doc)
             #m(quit ,app))))))


(defun begining-of-doc (document)
  #p(first #p(characters document)))


(defun end-of-doc (document)
  #p(last #p(characters document)))


;;; ----------------------------------------------------------------------
;;; Text mangling


(defconstant +wd-find-continue+ 1)

(defun trim-text (text)
  (if (> (length text) 250)
    (let ((shorter  (subseq text 0 249)))
      (subseq shorter 0 (- 249 (count #\return shorter))))
    text))
  

(defun range-find-text (range text)
  (cclet* ((find #p(find range)))
    #m(execute find (trim-text text) nil nil nil nil nil t
               +wd-find-continue+ nil)
    (when #p(found find)
      #p(start range))))
  

(defun selection-overwrite (range start end text)
  #m(select range)
  (cclet* ((document  #p(document range))
           (selection #p(selection #p(activewindow document))))
    #m(setrange selection start end)
    #m(typetext selection text)))


(defun footer (document section type)
  #p(range #m(item #p(footers #m(item #p(sections document) section))
                           type)))


(defun header (document section type)
  #p(range #m(item #p(headers #m(item #p(sections document) section))
                           type)))


#|(defun range-replace-text (range orig-text new-text)


;;; Replace text in Word doc (including headers & footers).
(defun word-replace-text (document orig-text new-text)
  (cclet* ((body     #p(content document))
           (sections #p(sections document)))
    ;; Body
    (range-replace-text body orig-text new-text)
    ;; Headers & footers
    (loop for i from 1 upto #p(count sections)
          for section = #m(item sections i)
          for headers = #p(headers section)
          for footers = #p(footers section) doing
          ;; Headers
          (loop for j from 1 upto #p(count headers)
                for header = #m(item headers j) doing
                (range-replace-text #p(range header) orig-text new-text)
;                (print #p(text #p(range header)))
                )
          ;; Footers
          (loop for k from 1 upto #p(count footers)
                for footer = #m(item footers k) doing
                (range-replace-text #p(range footer) orig-text new-text)
;                (print #p(text #p(range footer)))
                ))))|#
  


;(defparameter *cr* (format nil "~C" #\Return))

(defun carriage-return (string)
  (let ((position (search "^M" string :test #'string=)))
    (if position
      (carriage-return (concatenate 'string
                                    (subseq string 0 position)
                                    (string #\return)
                                    (subseq string (+ position 2))))
      string)))


#|(defun word-replace1st (document old-text new-text
                                 &key (range #'(lambda (doc)
                                                 #p(content doc))))
  (cclet* ((old    (carriage-return old-text))
           (new    (carriage-return new-text))
           (select #p(selection #p(activewindow document)))
           (start  (range-find-text (funcall range document) old))
           (end    (when start
                     (+ start (length old)))))
;    (print start)))
    (when start
      #m(setrange select start end)
      #m(typetext select new))))|#
