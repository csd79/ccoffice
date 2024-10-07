;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:cl-user)


(defpackage #:ccom
  (:use #:cl)
  (:export #:column
           #:cell-index
           #:letters-column
           #:with-com-initialized
           #:file-in-dir
           #:print-used-fli-templates
           #:#~
           #:#_
           #:cclet*
           #:excellerate
           #:column->row
           #:excel-value-as-number
           #:range
           #:with-used-range
           #:used-range
           #:last-row
           #:copy-formatting
           #:autofit-cols
           #:font
           #:apply-style
           #:close-workbook
           #:with-workbook
           #:title-column
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
           #:xselect>
           #:with-xselection
           #:excel-date
           #:excel-date-string
           #:title-xacolumn
           #:resolve-xacolumn-designator
           #:xarow
           #:xacell
           #:do-xarows
           #:xaselect
           #:xauniques
           #:xadouniques
           #:xapred
           #:xasort
           #:+wd-section-break-next-page+
           #:+wd-format-document-default+
           #:+wd-header-footer-first-page+
           #:+wd-header-footer-primary+
           #:+wd-align-page-number-center+
           #:with-document
           #:begining-of-doc
           #:end-of-doc
           #:range-find-text
           #:carriage-return
           #:selection-overwrite
           #:footer
           #:header))
