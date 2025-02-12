;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Globals


(defconstant +wd-section-break-next-page+  2)
(defconstant +wd-section-break-odd-page+   5)
(defconstant +wd-page-break+               7)
(defconstant +wd-format-document-default+ 16)
(defconstant +wd-header-footer-first-page+ 2)
(defconstant +wd-header-footer-primary+    1)
(defconstant +wd-align-page-number-center+ 1)
(defconstant +wd-format-pdf+              17)
(defconstant +wd-header-footer-even-pages+ 3)
(defconstant +wd-align-paragraph-center+   1)


;;; ----------------------------------------------------------------------
;;; Documents


;;; Create bindings for an Excel workbook with optional method calls
;;; (open, save at the end, close).
(defmacro with-document ((&key (app       nil)
                               (doc       'doc)
                               (use       nil)
                               (open      nil)
                               (read-only nil)
                               (save      nil)
                               (close     t))
                         &body body)
  (let ((app-obj (gensym))
        (docs    (gensym))
        (doc-obj (gensym))
        (update  (gensym)))
    `(cclet* ((,app-obj (or ,app 
                            (cclet* ((global (com:create-object :progid "Word.Application"))
                                     (app    (?application global)))
                              (setf (?visible app) nil)
                              app)))
              (,docs    (?documents ,app-obj))
              (,doc-obj (or (when (typep ,use 'com::com-interface) ,use)
                            (when ,open (!open ,docs ,open nil ,read-only))
                            (!add ,docs)))
              (,doc     ,doc-obj)
              (,update  (?screenupdating ,app-obj)))
       (unwind-protect
           (progn
             (setf (?screenupdating ,app-obj) nil)
             ,@body)
         (progn 
           (setf (?screenupdating ,app-obj) ,update)
           (when (and ,save (not ,read-only))
             (!save ,doc))
           (when ,close
             (setf (?saved ,doc) t)
             (setf (?displayalerts ,app-obj) nil)
             (!close ,doc)
             (unless ,app
               (!quit ,app-obj))))))))


(defun begining-of-doc (document)
  (?first (?characters document)))


(defun end-of-doc (document)
  (?last (?characters document)))


;;; ----------------------------------------------------------------------
;;; Text mangling


(defconstant +wd-find-continue+ 1)

(defun trim-text (text)
  (if (> (length text) 250)
    (let ((shorter  (subseq text 0 249)))
      (subseq shorter 0 (- 249 (count #\return shorter))))
    text))
  

(defun range-find-text (range text)
  (cclet* ((find (?find range)))
    (!execute find (trim-text text) nil nil nil nil nil t
               +wd-find-continue+ nil)
    (when (?found find)
      (?start range))))
  

(defun carriage-return (string)
  (let ((position (search "^M" string :test #'string=)))
    (if position
      (carriage-return (concatenate 'string
                                    (subseq string 0 position)
                                    (string #\return)
                                    (subseq string (+ position 2))))
      string)))


(defun selection-overwrite (range start end text)
  (!select range)
  (cclet* ((document  (?document range))
           (selection (?selection (?activewindow document)))
           (text2     (if (string= text "")
                        " "
                        text)))
    (!setrange selection start end)
    (!typetext selection (carriage-return text2))
    (when (string/= text text2)
      (!typebackspace selection))))


(defun footer (document section type)
  (?range (!item (?footers (!item (?sections document) section))
                           type)))


(defun header (document section type)
  (?range (!item (?headers (!item (?sections document) section))
                           type)))


(defun copy-via-fragment (from to tempfile)
  (!exportfragment (?formattedtext from)
                   tempfile
                   +wd-format-document-default+)
;  (!importfragment to fragment)
;  (delete-file fragment))
  (!importfragment to tempfile)
  (delete-file tempfile))
