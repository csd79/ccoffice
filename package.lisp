;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:cl-user)


(defpackage #:ccom
  (:use #:cl)
  (:export #:#p
           #:#m
           #:cclet*
           #:column
           #:cell-index
           #:letters-column
           #:with-com-initialized
           #:file-in-dir
           #:flitmp
           #:new-app-instance
           #:kill-running-instances
           #:get-excel
           #:excellerate
           #:get-document
           #:open-worksheets
           #:unfreeze-panes
           #:freeze-panes
           #:range
           #:with-used-edges
           #:used-range
           #:last-row
           #:copy-formatting
           #:autofit-cols
           #:font
           #:apply-style
           #:title-column
           #:with-workbook
           #:xselect>
           #:close-workbook
           #:with-xselection
           #:column->row
           #:column-designator-p
           #:column-designator
           #:resolve-column-designator
           #:locate-row
           #:row-designator-p
           #:row-designator
           #:resolve-row-designator
           #:xcell
           #:excel-value-as-number
           #:*xrange-default-value*
           #:*xrange-target-too-small*
           #:*xrange-target-too-large*
           #:xrange
           #:excel-date
           #:excel-date-string
           #:+wd-section-break-next-page+
           #:+wd-format-document-default+
           #:+wd-header-footer-first-page+
           #:+wd-header-footer-primary+
           #:+wd-align-page-number-center+
           #:+wd-find-continue+
           #:with-document
           #:begining-of-doc
           #:end-of-doc
           #:range-find-text
           #:selection-overwrite
           #:footer
           #:header
           #:carriage-return
))
