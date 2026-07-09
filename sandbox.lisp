;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccoffice)
#.(enable-ccom-syntax)


;;; ----------------------------------------------------------------------
;;; Sandbox


#|(defun h ()
  (with-workbook (:open "c:\\Users\\cselovszkid\\Desktop\\Ample Controls.xlsx"
                  :wsvars (wsheet) :save t :close t)
    (cclet* ((write  (range wsheet 1 100 1 105))
             (source (range wsheet 1 1 1 5))
             (dest   (range wsheet 9 1 9 5)))
      (with-used-range (wsheet left top right bottom)
        (format t "left: ~a, top: ~a, right: ~a, bottom: ~a~%" left top right bottom))
      (copy-formatting source dest)
      (setf ?'('value2 write) "Grr")
      (apply-style (range wsheet 2 10 2 14) '(:border))
      (apply-style (font (range wsheet 1 11 1 14)) '(:bold))
      ;; Print formula & value.
      (format t "~a   =   ~a~%"
              (xcell wsheet 8 2 :prop #'?'formula)
              (xcell wsheet 8 2 :prop #'?'value))
      ;; Set formulas
      (setf (xrange wsheet 2 2 4 4 :prop #'?'formula) "=\"\"")
      (setf (xcell wsheet 8 8 :prop #'?'formula) "=8*8")
      )))


(defun j ()
  (with-workbook (:open "c:\\Users\\cselovszkid\\Downloads\\EXPORT_kinevezés.XLSX"
                  :wsvars (in) :close t :read-only t)
    (with-used-range (in l u r b)
      (xrange in l u r b :prop #'?'value))))


(defun k ()
  (with-workbook (:open "c:\\Users\\cselovszkid\\Downloads\\EXPORT_kinevezés.XLSX"
                  :wsvars (in) :close t :read-only t)
    (cclet* ((r (range in 1 1 6 6)))
      (with-range (r left top right bottom)
        (format t "left: ~a~%top: ~a~%right: ~a~%bottom: ~a~%" left top right bottom)
        (let ((obj (read-xarray r)))
          (loop for c from 2 upto 4 doing
                (setf (xaref obj 0 c) c))
          (loop for c from 1 upto 5 doing
                (print (xaref obj 0 c)))
          (print "-----")
          (do-xarows (row r (xarows obj #(0 1 2 3 4)))
            (print (xcref row 1)))
          (print "-----")
          (xaselect obj #'(lambda (row)
                            (string= (xcref row "Vállalat hosszú megnevezése")
                                     "Sárospataki Tankerületi Központ")))
          (xauniques obj "Vállalat hosszú megnevezése")
          (xadouniques (v obj "Vállalat hosszú megnevezése")
            (print v)))))))


(defun l ()
  (with-workbook (:open "c:\\Users\\cselovszkid\\Downloads\\EXPORT_kinevezés.XLSX"
                  :wsvars (in) :close t :read-only t)
    (let* ((xarray (read-xarray (used-range in)))
           (sorted (xasort xarray (xapred ("Vállalat hosszú megnevezése" #'astring< #'astring=)
                                          ("SZK" #'astring< #'astring=)
                                          ("SZTSZ" #'astring<=)))))
      (do-xarows (row r sorted)
        (format t "~a    ~a    ~a~%"
                (xcref row "Vállalat hosszú megnevezése")
                (xcref row "SZK")
                (xcref row "SZTSZ"))))))


(defun m (wb)
  (with-workbook (:use wb :wsvars (in))
    (let* ((xarray (read-xarray (used-range in)))
           (sorted (xasort xarray (xapred ("Vállalat hosszú megnevezése" #'astring< #'astring=)
                                          ("SZK" #'astring< #'astring=)
                                          ("SZTSZ" #'astring<=)))))
      (do-xarows (row r sorted)
        (format t "~a    ~a    ~a~%"
                (xcref row "Vállalat hosszú megnevezése")
                (xcref row "SZK")
                (xcref row "SZTSZ"))))))

(defun n ()
  (with-workbook (:wbook w :open "c:\\Users\\cselovszkid\\Downloads\\EXPORT_kinevezés.XLSX"
                  :close t :read-only t)
    (m w)))


(defun p ()
  (let ((xa nil))
    (with-workbook (:open "C:\\Users\\cselovszkid\\Downloads\\2024.10.09. 3 feladat\\2 2024.10.01 statisztikák\\B2 2024.10.01. ének- és testnevelõtanár.xlsx"
                    :wsvars (in) :close t :read-only t)
      (setf xa (read-xarray (used-range in))))
    (let ((enek '("ének"))
          (tesi '("testnevel" "torna" "gyógype" "gyógyte" "gyógyto"))
          (enek-found '())
          (tesi-found '()))
      (flet ((init (string list)
               (when (stringp string)
                 (position-if #'(lambda (sample)
                                  (search sample string))
                              list))))
        (do-xarows (row r xa)
          (when (init (xcref row :f)
                      enek)
            (push r enek-found))
          (when (init (xcref row :g)
                      enek)
            (push r enek-found))
          (when (init (xcref row :f)
                      tesi)
            (push r tesi-found))
          (when (init (xcref row :g)
                      tesi)
            (push r tesi-found))))
      (values
       (remove-duplicates (nreverse enek-found))
       (remove-duplicates (nreverse tesi-found))))))|#


#|(defun q ()
  (with-property-accessors 
    (unwind-protect
        (progn
          (cclet* ((global  (com:create-object :progid "Excel.Application"))
                   (app     (?'application global))
                   (wbooks  (?'workbooks app))
                   (wbook   (!'open wbooks "c:\\Users\\cselovszkid\\Desktop\\Ample Controls.xlsx"))
                   (wsheets (?'worksheets wbook))
                   (wsheet  (?'item wsheets 1))
                   (range   (?'range wsheet "B1")))
            (print (?'value2 range))
            (setf (?'visible app) nil)
            (setf (?'value2 range) "P")
            (print (?'value2 range))   
            (!'close wbook)
            (!'quit app)))
      (print "LIBA!'"))))|#


;;; ----------------------------------------------------------------------
;;; Splt strikes back


(defparameter *in* "C:\\Users\\cselovszkid\\Downloads\\2024.11.29. Splt strikes back\\_TK_Adatok szerzõdéshez.xlsx")
(defparameter *out* "C:\\Users\\cselovszkid\\Downloads\\2024.11.29. Splt strikes back\\out\\")



(defun route-number (number current-range)
  (if (and current-range (= (1+ (second current-range)) number))
    (values (list (first current-range) number) nil)
    (values current-range (list number number))))

(defun find-ranges (list)
  (let* ((car (car list))
         (first-range (list car car)))
    (if (= (length list) 1)
      (list first-range)
      (let ((result '())
            (current-range first-range))
        (dolist (i (rest list))
          (multiple-value-bind (old new)
              (route-number i current-range)
            (if new
              (progn
                (push old result)
                (setf current-range new))
              (setf current-range old))))
        (push current-range result)))))
;        (nreverse result)))))

(defun delete-rows-unless (wsheet selector &key (start 2) (end nil))
  (with-used-range (wsheet left top right bottom)
    (let* ((y1 (max top (1- start)))
           (y2 (if end (min end bottom) bottom))
           (xarray (read-xarray (range wsheet left y1 right y2)))
           (negatives '()))
      (do-xarows (row r xarray)
        (unless (funcall selector row)
          (push (+ r start) negatives)))
      (when negatives
        (let ((ranges (sort (find-ranges (nreverse negatives)) #'> :key #'first)))
          (dolist (range ranges)
            (destructuring-bind (start end) range
              (delete-rows wsheet start end))))))))



(defun ssb ()
  (with-property-accessors
    (with-workbook (:open *in* :wsvars (form) :read-only nil :save t :close t)
      (!'saveas form (concatenate 'string *out* "Szegedi TK" " - Adatok szerzõdéshez.xlsx"))
      (delete-rows-unless form #'(lambda (row)
                                   (string= (xcref row 0) "Szegedi Tankerületi Központ"))
                          :start 3))))
#|

Most ezt egy ciklusba, uniques @ tks
  
  |#


;;; ----------------------------------------------------------------------
;;; Sandbox #2


(defparameter *kl* "C:\\Users\\cselovszkid\\Downloads\\laptop helyett.xlsx")

#|(defun w ()
  (let ((ccom4::*property-accessors-on* t))
    (with-property-accessors
      (with-workbook (:open *kl* :wsvars (data) :save t :close t)
;        (setf (xcell data 4 15) (xcell data 2 2))
        (setf (xrange data 4 15 6 17) (xrange data 2 2 4 4))
        ))))|#


;;; ----------------------------------------------------------------------
;;; Sandbox #2

;(defparameter *fx* "c:\\Users\\cselovszkid\\Downloads\\___\\…._TK_Iktatószámok iratgeneráláshoz.xlsx")
;(defparameter *fx* "c:\\Users\\cselovszkid\\Downloads\\___\\Adatszolgáltatás_legmagasabb illetménnyel rendelkezõ pedagógusok.xlsx")

#|(defun splt2 (file title)
  (with-property-accessors
    (setf (property-accessors-on) t)
    (with-workbook (:wbook wbook :open file :wsvars (ws1) :save t :close t)
      (cclet* ((last-row   (last-row ws1))
               (column     (title-column ws1 title))
               (raw        (?'value2 (range ws1 column 2 column last-row)))
               (list       (loop for i from 0 below (array-dimension raw 0) collecting (aref raw i 0)))
               (categories (remove-duplicates list :test #'string=))
               (sheets     (?'worksheets wbook))
               (index      0))
        (dolist (categorie categories)
;          (cclet* ((new-ws (wsselect ws1 `((,title ,categorie)) :paste-all t :paste-column-widths t)))
          (cclet* ((new-ws (xselect> ws1 `((,title ,categorie)) :paste-all t :paste-column-widths t)))
            (!'select new-ws)
            (!'copy new-ws (?'item sheets (incf index)))
            (setf (?'name (?'item sheets index)) (concatenate 'string (first (str:words categorie)) " TK"))
            (print categorie))))))
  (print "OK"))|#


;;; ----------------------------------------------------------------------
;;; Sandbox #3: not my type


(defparameter *a* "c:\\Users\\cselovszkid\\Desktop\\teszt.xlsx")


#|(defun xl-error-p (value2)
  "Is VALUE2 an errorcode, and if it is, what error does it stands for?'"
  (and (integerp value2)
       (cdr (assoc value2 '((-2146826259 . :name) (-2146826281 . :div/0) (-2146826265 . :ref)
                            (-2146826252 . :num)  (-2146826246 . :n/a)   (-2146826273 . :value)
                            (-2146826288 . :null))))))|#


(defun xl-spill-range (range)
  (cclet* ((spillparent (?'spillparent range)))
    (unless (eq spillparent :empty)
      (?'spillingtorange spillparent))))



; felesleges opcionális paraméterek :NOT-SPECIFIED


#|
  A XARRAY-t ki lehetne egészíteni egy TEXT réteggel, ami szintén egyben olvasható, de nem írható.
  Ez mindig az Excel beállított számformátum szerinti formában adja vissza az értéket.

  Érdekes módon ha a range 1 cella, a value2 nem tesz különbséget a hamis értéket tartalmazó és
  az üres cella értéke között (NIL).
  Viszont ha a range több cella (és az eredmény tömb) akkor az üres cella :EMPTY.

  Tehát érdemes lenne MINDIG tömböt lekérni.

  Ennek általánosítására átalakíthatnánk a cella-accessor függvényeket hogy közvetlenül megadható
  legyen a ws és sarkok.
  Ha csak egy sarok van, akkor azt küldje el a köv. paraméternek is.
  És legyen még két KEY: :single-cell-as-value és :column-as-vector

  Egyébként a XA és sima táblakezelõ függvényeket egységesíteni kéne, akár metódusként!'!'!'!'!'!'
  
  |#




(defun b (&optional (r 5))
  (with-property-accessors
    (setf (property-accessors-on) t)
    (with-workbook (:open *a* :wsvars (ws1 ws2) :save nil :close t)
#|      ;; Hibakódok
      (loop for row from 1 upto 8 doing
            (print (xl-error-p (?'value2 (range ws1 1 row)))))|#
      ;; Üres cellák
;      (loop for row from 10 upto 13 doing
;            (multiple-value-bind (&rest all)
;                (?'value2 (range ws1 1 row))
;              (print all)))
            (print (!'value2 (range ws1 1 10 1 13))))
#|      ;; Értékek
      (loop for col from 1 upto 13 doing
            (print (?'text (range ws2 col r))))|#
#|      ;; Dinamikus tömb
      (loop for col from 1 upto 13 doing
            (print (?'formulaarray (range ws2 col 14))))
      (loop for col from 1 upto 13 doing
            (print (?'value (range ws2 col 14))))|#
;      (print (?'numberformat (range ws2 1 4 13 4))) ; Ha egyformat, formátum, különben :NULL
;      (print (?'numberformat (range ws2 3 3 3 18)))

;      (print (?'spillingtorange (range ws2 2 14)))
;      (print (?'spillparent (range ws2 4 4)))
;      (print (?'text (xl-spill-range (range ws2 5 17)))) ; Spillrange-re a TEXT valamiért nem mûködik.
;      (print (?'value2 (xl-spill-range (range ws2 2 16))))
      ))




;;; ----------------------------------------------------------------------
;;; Sandbox #4: further



(defparameter *f* "c:\\Users\\cselovszkid\\common-lisp\\msoffice\\ff.xlsx")



(defun d (getter setter ws)
  (funcall !setter "Juhhéééé!'" (range ws 4 4))
  (funcall !getter (range ws 1 1)))

(defun s ()
  (with-property-accessors
    (setf (property-accessors-on) t)
    (with-workbook (:open *f* :wsvars (ws1) :read-only nil :save t)
      (d 'value2 'value2 ws1))))



(defun e (accessor-sym ws)
  (flet ((acc (row column)
           (funcall (fdefinition ?accessor-sym) ws row column))
         ((setf acc) (row column value)
           (funcall (fdefinition `(setf ?,accessor-sym)) value ws row column)))
    (setf (acc 6 6) "Nahát!'?'")
    (acc 6 6)))

(defun u ()
  (with-property-accessors
    (setf (property-accessors-on) t)
    (with-workbook (:open *f* :wsvars (ws1) :read-only nil :save t)
      (e 'value2 ws1))))






;;; ----------------------------------------------------------------------
;;; setf


(defparameter *english-numbers* (make-hash-table :test #'equal))

(defun eninit ()
  (setf (gethash "one" *english-numbers*) 1
        (gethash "two" *english-numbers*) 2
        (gethash "fifteen" *english-numbers*) 15))

(defun word-number (word)
  (values (gethash word *english-numbers*)))

(define-setf-expander word-number (word)
  (let ((new-value-var (gensym)))
    (values nil
            nil
            `(,new-value-var)
            `(setf (gethash ,word *english-numbers*) ,new-value-var)
            `(word-number ,word))))

(define-setf-expander word-number (word)
  (let ((word-var (gensym))
        (new-value-var (gensym)))
    (values `(,word-var)                        ; list of names for temp vars
            `(,word)                            ; value forms (will get bound the the vars above)
            `(,new-value-var)                   ; will be used by the compiler to deliver the new value to us
            `(setf (gethash ,word-var           ; the form that updates the place; also must return the new value
                            *english-numbers*)
                   ,new-value-var)
            `(word-number ,word-var))))         ; the form to access the place



(defparameter *alist*
  (list (cons 'donald-duck 'duckberg)
        (cons 'superman 'metropolis)
        (cons 'batman 'gotham-city)))

(defun cdr-assoc (item alist)
  (cdr (assoc item alist)))

(define-setf-expander cdr-assoc (item alist)
  (let ((item-var (gensym))
        (cons-found (gensym))
        (alist-var (gensym))
        (new-value-var (gensym)))
    (values `(,item-var ,alist-var ,cons-found)
            `(,item ,alist (assoc ,item-var ,alist-var))
            `(,new-value-var)
            `(cond (,cons-found
                    (setf (cdr ,cons-found) ,new-value-var))
                   (t
                    (setf ,alist (acons ,item-var ,new-value-var ,alist-var))
                    ,new-value-var))
            `(cdr ,cons-found))))
;; This solution ignores the presence of another implicit setf for NTH.


(define-setf-expander cdr-assoc (item alist &environment env)
  (multiple-value-bind (temp-vars temp-forms store-vars
                                  setter-form getter-form)
      (get-setf-expansion alist env)
    (let ((item-var (gensym))
          (cons-found (gensym))
          (new-value-var (gensym)))
      (values `(,@temp-vars ,item-var ,cons-found)
              `(,@temp-forms ,item (assoc ,item-var ,getter-form))
              `(,new-value-var)
              `(cond (,cons-found
                      (setf (cdr ,cons-found) ,new-value-var))
                     (t
                      (let ((,(first store-vars)
                             (acons ,item-var ,new-value-var
                                    ,getter-form)))
                        ,setter-form
                        ,new-value-var)))
              `(cdr ,cons-found)))))



;;; ----------------------------------------------------------------------
;;; compiled closure


(defun aa ()
  #'(lambda (v)
      (* v 3)))









;;; ----------------------------------------------------------------------
;;; ccom4 tests



(defparameter *foil* "C:\\Users\\cselovszkid\\common-lisp\\wax\\Munka\\SAP lekérdezések\\gayzax.xlsx")


(defun n ()
  (with-ccom-context
    (with-workbook (:open *foil* :read-only t :wsvars (ws1) :close t)
      (print (xcell ws1 2 2))
      (read-xarray (used-range ws1))
;      nil
      )))




(defun m ()
  (enable-ccom-syntax)
  (let ((result (read-from-string
   "(defmacro with-workbook ((&key (app       nil)
                               (!'wbook     'wbook)
                               (use       nil)
                               (open      nil)
                               (wsvars    '())
                               (read-only nil)
                               (save      nil)
                               (close     t))
                         &body body)
  (let ((app-obj (gensym))
        (wbooks  (gensym))
        (doc-obj (gensym))
        (wsheets (gensym)))
    `(cclet* ((,app-obj (or ,app 
                            (cclet* ((global (com:create-object :progid \"Excel.Application\")))
                              (?'application global))))
              (,wbooks  (?'workbooks ,app-obj))
              (,doc-obj (or (when (typep ,use 'com::com-interface) ,use)
                            (when ,open (!'open ,wbooks ,open nil ,read-only))
                            (!'add ,wbooks)))
              (,wbook   ,doc-obj)
              (,wsheets (?'worksheets ,doc-obj))
              ,@(loop for n from 1
                      for var in wsvars
                      when var collect
                      (list var `(?'item ,wsheets ,n))))
       (unwind-protect
           (excellerate (,app-obj)
             (when (eq ,use ,doc-obj)
               (com:add-ref ,doc-obj))
             (window-visible ,doc-obj 1 t)
             ,@body)
         (progn 
           (when (and ,save (not ,read-only))
             (!'save ,doc-obj))
           (when ,close
             (setf (?'saved ,doc-obj) t)
             (setf (?'displayalerts ,app-obj) nil)
             (!'close ,doc-obj)
             (unless ,app
               (!'quit ,app-obj))))))))")))
    (disable-ccom-syntax)
    result))

(defun o ()
  (enable-ccom-syntax)
  (print (read-from-string "(?'prop obj arg1 arg2)"))
  (disable-ccom-syntax))
  

  




#.(disable-ccom-syntax)
