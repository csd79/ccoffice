;;;; -*- Mode: Common-Lisp; Author: denes.cselovszky@gmail.com -*- 
                                                                              ;

(in-package #:ccom)


;;; ----------------------------------------------------------------------
;;; Excel workbooks, worksheets


;;; Get Excell app.
(defun get-excel (&optional object)
  (if object
    ;; Parrent instance of OBJECT.
    #p(application object)
    ;; Any running Excel instance.
    (com::get-active-object :progid "Excel.Application"
                            :riid   'i-dispatch
                            :errorp nil)))


(defconstant +xl-calculation-automatic+ -4105)
(defconstant +xl-calculation-manual+    -4135)


;;; Execute BODY using faster Excel interaction.
(defmacro excellerate ((excel) &body body)
  `(unwind-protect 
       (progn 
         (setf #p(screenupdating ,excel) nil
               #p(displayalerts ,excel)  nil)
         (when #p(visible ,excel)
           (setf #p(calculation ,excel) +xl-calculation-manual+))
         ,@body)
     (setf #p(screenupdating ,excel) t
           #p(displayalerts ,excel)  t)
     (when #p(visible ,excel)
       (setf #p(calculation ,excel) +xl-calculation-automatic+))))


;;; Open given document with default application, hopefully Excel.
(defun get-document (fullname)
  (com::get-object fullname :riid 'com::i-dispatch))


;;; List pathname of every open Excel workbook.
(defun open-workbooks (excel)
  (cclet* ((workbooks #p(workbooks excel))
           (count     #p(count workbooks)))
    (loop for i from 1 upto count collecting
          #p(fullname #p(item workbooks i)))))


;;; List every open worksheet.
(defun open-worksheets (excel)
  (let ((workbooks (open-workbooks excel))
        (results   '()))
    (dolist (workbook workbooks)
      (cclet* ((worksheets #p(worksheets (get-document workbook)))
               (count      #p(count worksheets)))
        (push workbook results)
        (push (loop for i from 1 upto count collecting
                    #p(name #p(item worksheets i)))
              results)))
    (nreverse results)))


;;; ----------------------------------------------------------------------
;;; Excel ranges, values


;; Create range object within given 'coordinates'.
(defun range (worksheet &optional (x1 nil) (y1 nil) (x2 nil) (y2 nil))
  (let ((upper-left (cell-index x1 y1)))
    (if (and x2 y2)
      #p(range worksheet upper-left (cell-index x2 y2))
      #p(range worksheet upper-left))))


(defconstant +xl-to-left+          -4159)
(defconstant +xl-up+               -4162)
(defconstant +xl-values+           -4163)
(defconstant +xl-by-rows+              1)
(defconstant +xl-by-columns+           2)
(defconstant +xl-whole+                1)
(defconstant +xl-previous+             2) 

(defmacro with-used-edges ((worksheet left top right bottom) &body body)
  (let ((whole (gensym))
        (from  (gensym)))
    `(cclet* ((,whole  #p(cells ,worksheet))
              (,top    #p(row #p(end ,whole +xl-up+)))
              (,left   #p(column #p(end ,whole +xl-to-left+)))
              (,from   (range ,worksheet ,left ,top))
              (,bottom #p(row #m(find ,whole "*" ,from +xl-values+
                                     +xl-whole+ +xl-by-rows+ +xl-previous+)))
              (,right  #p(column #m(find ,whole "*" ,from +xl-values+
                                        +xl-whole+ +xl-by-columns+
                                        +xl-previous+))))
         ,@body)))


;;; Locate used range of WORKSHEET.
(defun used-range (worksheet)
  (with-used-edges (worksheet left top right bottom)
    (range worksheet left top right bottom)))


;;; ----------------------------------------------------------------------
;;; Excel cell formatting


;;; Constants used to copy cell formatting.
(defconstant +xl-paste-formats+    -4122)
(defconstant +xl-paste-comments+   -4144)
(defconstant +xl-paste-validation+     6)


;;; Copy cell formatting between ranges.
(defun copy-formatting (source destination)
  #m(copy source)
  #m(pastespecial destination +xl-paste-formats+)
  #m(pastespecial destination +xl-paste-comments+)
  #m(pastespecial destination +xl-paste-validation+))


;;; Set auto width for all columns.
(defun autofit-cols (worksheet)
  (cclet* ((range #p(columns worksheet)))
    #m(auto-fit range)))


;;; Target range for font formatting.
(defun font (range &optional first last)
  #p(font (if (and first last)
            #p(characters range first last)
            range)))


(defconstant +xl-align-center+     -4108)
(defconstant +rgb-black+               0)

(defparameter *style-elements*
  `(:wrap-text     (:prop   "WrapText"            ,t)
    :halign-center (:prop   "HorizontalAlignment" ,+xl-align-center+)
    :valign-center (:prop   "VerticalAlignment"   ,+xl-align-center+)
    :border        (:method "BorderAround"        ,:null ,:null ,:null ,+rgb-black+)
    :bold          (:prop   "Bold"                ,t)))


;;; Apply any style defined in *STYLE-ELEMENTS*.
(defun apply-style (range style)
  (dolist (style-element style)
    (destructuring-bind (type name &rest values)
        (getf *style-elements* style-element)
      (case type
        (:prop   (setf (com::invoke-dispatch-get-property range name)
                       (first values)))
        (:method (apply #'com::invoke-dispatch-method range name values))))))


;;; ----------------------------------------------------------------------
;;; Excel indexing & searching basics


(defun column->row (array &key (element-type t) (getter #'identity))
  (let* ((height (array-dimension array 0))
         (result (make-array height :element-type element-type)))
    (loop for i from 0 below height doing
          (setf (svref result i)
                (funcall getter
                         (aref array i 0))))
    result))


;;; Create index vector from values of given COLUMN in WORKSHEET.
;;; :ELEMENT-TYPE and :GETTER must cooperate!
(defun index (worksheet column &key (element-type t) (getter #'identity))
  (cclet* ((app (get-excel worksheet)))
    (excellerate (app)
      (with-used-edges (worksheet left top right bottom)
        (declare (ignore left top right))
        (cclet* ((range  (range worksheet column 2 column bottom))
                 (raw    #p(value2 range)))
          (column->row raw
                       :element-type element-type
                       :getter getter))))))


;;; Find the column that has TITLE in its header.
(defun title-column (worksheet title &key (key #'identity) (header-row 1) (test #'string-equal)
                         (from-end nil) (start nil) (end nil))
  (with-used-edges (worksheet left top right bottom)
;    (declare (ignore top bottom))
    (cclet* ((header #p(value2 (range worksheet left header-row right header-row))))
      (if from-end
        (loop for i from (1- (or end right)) downto (1- (or start left))
              for v = (funcall key (aref header 0 i))
              thereis (and (funcall test title v) (1+ i)))
        (loop for i from (1- (or start left)) upto (1- (or end right))
              for v = (funcall key (aref header 0 i))
              thereis (and (funcall test title v) (1+ i)))))))


;;; Search INDEX for value, collect positions of occurences found.
(defun occurances (index value &optional (test #'equalp))
  (loop for i from 0 below (array-dimension index 0)
        when (funcall test value (svref index i))
        collect i))


;;; List rows of WORKSHEET where INDEX is VALUE.
(defun filter-rows (worksheet index-col values &key (idx-element-type t) (idx-getter #'identity) (test #'equalp))
  (cclet* ((app   (get-excel worksheet))
           (index (index worksheet index-col :element-type idx-element-type :getter idx-getter))
           (rows  (apply #'append (mapcar #'(lambda (value)
                                              (occurances index value test))
                                          values))))
    (with-used-edges (worksheet left top right bottom)
;      (declare (ignore bottom))
      (excellerate (app)
        (loop for r in rows
              for rr = (+ r 2)
              collecting #p(value2 (range worksheet left rr right rr)))))))


;(defun filter-rows-if (worksheet index-col &optional (predicate #'equalp))


;;; List rows from given FILE/SHEET whose index equals VALUE.
(defun search-file (values file sheet title &key (element-type t) (getter #'identity) (header-row 1))
  (cclet* ((wbook   (get-document (namestring file)))
           (wsheets #p(worksheets wbook))
           (wsheet  #p(item wsheets sheet))
           (column  (title-column wsheet title :header-row header-row)))
    (filter-rows wsheet column values :idx-element-type element-type :idx-getter getter)))


;;; List rows corresponding to VALUE from multiple files.
(defun search-files (values files &key (sheets '()) (titles '())
                           (element-type t) (getter #'identity) (header-row 1))
  (apply #'append
         (mapcar #'(lambda (file sheet title)
                     (search-file values file sheet title
                                  :element-type element-type
                                  :getter getter
                                  :header-row header-row))
                 files sheets titles)))


;;; ----------------------------------------------------------------------
;;; Word


(defconstant +wd-find-continue+ 1)

;;; Replace text in Word doc.
(defun word-replace-text (document orig-text new-text)
  (cclet* ((find #p(find #p(content document))))
    #m(execute find orig-text nil nil nil nil nil t
               +wd-find-continue+ nil new-text)))


;;; ----------------------------------------------------------------------
;;; Sandbox


(defun test1 ()
  (cclet* ((file    "c:\\Users\\cselovszkid\\Desktop\\lisp-book-structure-comparison.xlsx")
           (wbook   (get-document file))
           (wsheets #p(worksheets wbook))
           (wsheet  #p(item wsheets 1))
           (source  (range wsheet 1 1 1 5))
           (dest    (range wsheet 3 1 3 5))
           (write   (range wsheet 1 100 1 105))
           (used    (used-range wsheet))
           (app     (get-excel wbook)))
    (excellerate (app)
      ;; Copy formatting.
      (copy-formatting source dest)
      ;; Set values.
      (setf #p(value2 write) "Grr")
      ;; Read values.
      #p(value2 used)
      ;; Add borders.
      (apply-style (range wsheet 1 9 1 14) '(:border))
      ;; Bold text.
      (apply-style (font (range wsheet 1 11 1 14)) '(:bold)))))


(defparameter *root* "c:\\Users\\cselovszkid\\Downloads\\2024.06.27. Újabb kinevezések elõkészület\\Lekérdezés\\")
(defparameter *alap* "B8_0008IT_20240701.XLSX")
(defparameter *rendszeres* "B8_0014IT_20240701.XLSX")

(defparameter *rendszeres-map*
  '("SZTSZ"                          :sztsz
    "Név"                            :teljesnev
    "Belépés dátuma"                 :belepdatum
    "Vállalat hosszú megnevezése"    :vallalat
    "szervezeti egys hosszú megnev." :szerv-egys
    "Személyi kör megnevezése"       :szemelyikor
    "Hely.dolg.neve."                :helyettesitett-dolg
    "Munkakör"                       :munkakor
    "Születési dátum"                :szul-datum
    "Születési hely"                 :szul-hely
    "Születési vezetéknév"           :szul-veznev
    "Születési utónév"               :szul-utonev1
    "2.születési utónév"             :szul-utonev2
    "Anya"                           :anya-veznev
    "Anyja keresztneve"              :anya-utonev1
    "Anyja 2.keresztneve"            :anya-utonev2
    "MI-%"                           :mi-hanyad
    "Kinevezés/szerzõdés jellege"    :kinevszerz-jelleg
    "Szerz.vége"                     :szerz-vege
    "Próbaidõ  vége"                 :probaido-vege
    "FEOR-szám standard"             :feor
    "Vége"                           :berelem-vege
    "Bérelem"                        :berelem
    "Összeg"                         :osszeg
    "Pénznem"                        :penznem))
(defparameter *alap-map*
  (append *rendszeres-map*
          '("Bérr.cs." :berr-csop)))

(defparameter *header-row* 1)


(defun workfile (file)
  (probe-file
   (pathname
    (concatenate 'string (namestring *root*)
                 (namestring file)))))


(defun test2 (&optional (sztsz 10004006))
   (cclet* ((wbook   (get-document (namestring (workfile *alap*))))
            (wsheets #p(worksheets wbook))
            (wsheet  #p(item wsheets 1)))
     (let ((index (index wsheet 1 :element-type 'integer
                         :getter #'read-from-string)))
       index)))


(defun test3 (values)
  (search-files values
                (list (namestring (workfile *alap*))
                      (namestring (workfile *rendszeres*)))
                :sheets '(1 1)
                :titles '("sztsz" "sztsz")
                :element-type 'integer
                :getter #'read-from-string))


(defun test4 ()
  (cclet* ((wbook   (get-document (namestring (workfile *alap*))))
           (wsheets #p(worksheets wbook))
           (wsheet  #p(item wsheets 1)))
    (idx-column wsheet "SZtSz" :from-end t)))


(defun test5 ()
  (cclet* ((wbook   (get-document (namestring (workfile *alap*))))
           (wsheets #p(worksheets wbook))
           (wsheet  #p(item wsheets 1))
           (col-no  (title-column wsheet "sztsz")))
    (filter-rows wsheet col-no '(10049639) :idx-element-type 'integer :idx-getter #'read-from-string)))


#|
10006313	2
10015516	2
10016245	2
10045937	2
10049639	4
10055012	2
10066570	2
10066638	2
10066785	2
10066837	2
10078582	2
10084503	2
10090861	2
10097191	2
10145704	2
10145934	2
10160161	2
10160960	2
10162576	2
10164999	2
10191618	2
10203001	2
10204824	2
|#


#| Vigyázat, egymás alá rakja a 8 és 14 találatait, de a jobb oldali értékek
nem biztos hogy ugyanabban az oszlopban vannak! |#
