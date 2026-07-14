(require "automation")

(defsystem "ccoffice"
  :description "Office automation layer"
  :author      "Denes Cselovszki <denes.cselovszki@gmail.com>"
  :version     "0.25"
  :depends-on  ("ccom4" "achar" "local-time")
  :serial      t
  :components  ((:file "package")
                (:file "utilities")
                (:file "excel")
                (:file "word")
                (:file "outlook")
                (:file "ppoint")
                (:file "sandbox")))
