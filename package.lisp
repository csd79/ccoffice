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
           #:with-com-initialized
           #:get-excel
           #:excellerate
           #:get-document
           #:open-worksheets
           #:range
           #:with-used-edges
           #:used-range
           #:copy-formatting
           #:autofit-cols
           #:font
           #:apply-style
           #:title-column
           #:wsselect
           #:wsref
           #:column->row
#|           #:index
           #:title-column
           #:occurances
           #:filter-rows
           #:search-file
           #:search-files|#
           #:word-replace-text
))
