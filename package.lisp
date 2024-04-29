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
           #:used-range
           #:copy-formatting))
