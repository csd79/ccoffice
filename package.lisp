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
;           #:wsselect
           #:xselect>
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
           #:with-document
           #:begining-of-doc
           #:end-of-doc
           #:range-find-text
           #:selection-overwrite
           #:footer
           #:header
;           #:word-replace-text
           #:carriage-return
;           #:word-replace1st
))
