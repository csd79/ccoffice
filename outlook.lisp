;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccoffice)
#.(enable-ccom-syntax)


;;; ----------------------------------------------------------------------
;;; Workbooks, worksheets


(defconstant +ol-mail-item+ 0)
(defconstant +ol-by-value+  1)


(defun outlook-running-p ()
  (com:get-active-object :progid "Outlook.Application" :riid 'com:i-dispatch :errorp nil))


(defun new-mail (to subject &key on-behalf body (cc "") attch (command :view))
  (cclet* ((outlook (outlook-running-p))
           (mail    (!'createitem outlook +ol-mail-item+))
           (atts    (?'attachments mail)))
    (setf (?'to mail) to
          (?'cc mail) cc
          (?'subject mail) subject)
    (when on-behalf
      (setf (?'sentonbehalfofname mail) on-behalf))
    (when body
      (setf (?'htmlbody mail) body))
    (when attch
      (!'add atts attch +ol-by-value+))
    (cond ((eq command :view) (!'display mail))
          ((eq command :send) (!'send mail))
          (t mail))))


#.(disable-ccom-syntax)
