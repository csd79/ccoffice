(require "automation")

(defsystem "ccom"
  :description "Wrapper layer for LispWorks 7.1 COM package v2"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.05"
  :depends-on  ("alexandria")
  :serial      t
  :components  ((:file "package")
                (:file "syntax")
                (:file "utilities")
                (:file "ccom")))
