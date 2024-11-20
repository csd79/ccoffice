;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:cl-user)


(defpackage #:ccom
  (:use #:cl #:achar)
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
           #:freeze-panes
           #:range
           #:with-used-range
           #:used-range
           #:last-row
           #:with-range
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
           #:index
           #:read-xarray
           #:make-xarray
           #:rearrange
           #:write-xarray
           #:xaref
           #:xcref
           #:xarows
           #:do-xarows
           #:xaselect
           #:xauniques
           #:xadouniques
           #:xapred
           #:xasort
           #:+wd-section-break-next-page+
           #:+wd-section-break-odd-page+
           #:+wd-page-break+
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
