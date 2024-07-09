(require "automation")

(defsystem "ccom"
  :description "Wrapper layer for LispWorks 7.1 COM package v2"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.12"
  :serial      t
  :components  ((:file "package")
                (:file "utilities")
                (:file "syntax")
                (:file "ccom")))
