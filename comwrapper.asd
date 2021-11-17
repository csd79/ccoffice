(require "automation")

(defsystem "comwrapper"
  :description "Wrapper layer for LispWorks 7.1 COM package"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.47"
  :depends-on  ("alexandria")
  :serial      t
  :components  ((:file "package")
                (:file "comwrapper")))
