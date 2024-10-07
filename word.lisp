;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Globals


(defconstant  +wd-section-break-next-page+  2)
(defconstant  +wd-format-document-default+ 16)
(defconstant  +wd-header-footer-first-page+ 2)
(defconstant  +wd-header-footer-primary+    1)
(defconstant  +wd-align-page-number-center+ 1)


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
  (let ((app2 (gensym)))
    `(cclet* ((,app2 (if (boundp ',app)
                       ,app
                       (com:create-object :progid "Word.Application")))
              (docs  #~('documents ,app2))
              (,doc  (if ,open-file
                       (#_open docs ,open-file nil ,read-only)
                       (#_add docs))))
       (unwind-protect 
           (progn
             (setf #~('screenupdating ,app2) nil)
             ,@body)
         (progn 
           (setf #~('screenupdating ,app2) t)
           (when (and ,save (not ,read-only))
             (#_save ,doc))
           (when ,close
             (setf #~('saved ,doc) t)
             (#_close ,doc)
             (unless (boundp ',app)
               (#_quit ,app2))))))))


(defun begining-of-doc (document)
  #~('first #~('characters document)))


(defun end-of-doc (document)
  #~('last #~('characters document)))


;;; ----------------------------------------------------------------------
;;; Text mangling


(defconstant +wd-find-continue+ 1)

(defun trim-text (text)
  (if (> (length text) 250)
    (let ((shorter  (subseq text 0 249)))
      (subseq shorter 0 (- 249 (count #\return shorter))))
    text))
  

(defun range-find-text (range text)
  (cclet* ((find #~('find range)))
    (#_execute find (trim-text text) nil nil nil nil nil t
               +wd-find-continue+ nil)
    (when #~('found find)
      #~('start range))))
  

(defun carriage-return (string)
  (let ((position (search "^M" string :test #'string=)))
    (if position
      (carriage-return (concatenate 'string
                                    (subseq string 0 position)
                                    (string #\return)
                                    (subseq string (+ position 2))))
      string)))


(defun selection-overwrite (range start end text)
  (#_select range)
  (cclet* ((document  #~('document range))
           (selection #~('selection #~('activewindow document)))
           (text2     (if (string= text "")
                        " "
                        text)))
    (#_setrange selection start end)
    (#_typetext selection (carriage-return text2))
    (when (string/= text text2)
      (#_typebackspace selection))))


(defun footer (document section type)
  #~('range (#_item #~('footers (#_item #~('sections document) section))
                           type)))


(defun header (document section type)
  #~('range (#_item #~('headers (#_item #~('sections document) section))
                           type)))
