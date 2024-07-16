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
           #:get-excel
           #:excellerate
           #:get-document
           #:open-worksheets
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
           #:wsselect
           #:column->row
           #:column-designator-p
           #:column-designator
           #:resolve-column-designator
           #:locate-row
           #:row-designator-p
           #:row-designator
           #:resolve-row-designator
           #:xcell
           #:*xrange-default-value*
           #:*xrange-target-too-small*
           #:*xrange-target-too-large*
           #:xrange
           #:word-replace-text
))
