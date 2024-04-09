;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:cl-user)


(defpackage #:csd79-lispworks-comwrapper
  (:nicknames "COMWRAPPER")
  (:use #:cl #:com #:alexandria)
  (:export #:+rgb-black+
           #:+rgb-light-grey+
	   #:+rgb-yellow+
           #:+rgb-light-yellow+
           #:+xl-worksheet+
           #:+xl-to-left+
           #:+xl-up+
           #:+xl-formulas+
           #:+xl-values+
           #:+xl-by-rows+
           #:+xl-by-columns+
           #:+xl-whole+
           #:+xl-previous+
           #:+xl-shift-down+
           #:+xl-align-center+
           #:+xl-workbook-default+
           #:+ol-mail-item+
           #:+ol-by-value+
;           #:*com*
;           #:define-fn-nickname
;           #:column
           #:cell-index
           #:with-com-initialized
           #:com-property
           #:com-method
;           #:comhelper
           #:comlet*
           #:set-app-visibility
           #:*observe-app-running*
           #:with-app
           #:excel
           #:outlook
           #:xlsx
           #:close-workbook
           #:save-and-close-workbook
           #:with-xlsx
           #:with-range
           #:getcell
           #:setcell
           #:locate-string
           #:worksheet-name
           #:list-sheet-names
           #:get-sheet
           #:select-sheet
           #:rename-sheet
           #:create-sheet
;           #:used-range
           #:with-used-range
           #:*style-elements*
           #:char-range
           #:apply-style
           #:insert-rows
           #:autofit-cols
           #:body-from-file
           #:new-mail))
