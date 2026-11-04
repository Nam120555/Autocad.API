;;;========================================================================
;;; XUAT TOA DO TU EXCEL LEN CAD (TEXT + BLOCK) - LISP NHE FILE
;;;========================================================================
;;; Tac gia  : AI Assistant cho Huynh Ngoc Nam
;;; Ngay     : 28/07/2026
;;; Phien ban: 1.1
;;;
;;; CHUC NANG:
;;;   Copy du lieu toa do tu Excel (STT, X, Y, Z, Ten diem)
;;;   Phun len AutoCAD duoi dang Block dau cong (+) va Text nhe file.
;;;   Thay the COGO Points nang de giam dung luong ban ve.
;;;
;;; CACH DUNG:
;;;   1. Mo Excel, chon vung du lieu toa do (5 cot: STT, X/N, Y/E, Z, Ten diem)
;;;   2. Nhan Ctrl+C de copy
;;;   3. Sang AutoCAD, go lenh: XTD  (hoac XUATTOADO)
;;;   4. Chon cac tuy chon (cao chu, thu tu XY, layer...)
;;;   5. Xong! Toa do duoc phun len ban ve.
;;;
;;; CAC LENH:
;;;   XTD        - Lenh chinh: Doc clipboard Excel va phun toa do
;;;   XTDF       - Doc tu file TXT/CSV thay vi clipboard
;;;   TDBLOCK    - Tao lai block dinh nghia (neu can)
;;;========================================================================

;;; ====== BIEN TOAN CUC (CAI DAT MAC DINH) ======
(setq *xtd-textheight*   2.5)    ; Cao chu mac dinh
(setq *xtd-layer-point*  "DIEM_TOADO")  ; Layer cho block diem
(setq *xtd-layer-stt*    "STT_DIEM")    ; Layer cho so thu tu
(setq *xtd-layer-caodo*  "CAO_DO")      ; Layer cho cao do Z
(setq *xtd-layer-desc*   "TEN_DIEM")    ; Layer cho ten diem / mo ta
(setq *xtd-block-name*   "TD_DIEM")     ; Ten block dau cong
(setq *xtd-block-size*   1.5)    ; Kich thuoc dau cong (ti le voi cao chu)
(setq *xtd-xy-order*     0)     ; 0 = X,Y (Mac dinh Excel); 1 = Y,X
(setq *xtd-show-stt*     T)     ; Hien thi so thu tu
(setq *xtd-show-z*       T)     ; Hien thi cao do
(setq *xtd-show-desc*    T)     ; Hien thi ten diem / mo ta
(setq *xtd-color-point*  1)     ; Mau block diem (1=Do)
(setq *xtd-color-stt*    7)     ; Mau so thu tu (7=Trang)
(setq *xtd-color-z*      3)     ; Mau cao do (3=Xanh la)
(setq *xtd-color-desc*   5)     ; Mau ten diem (5=Xanh duong)
(setq *xtd-offset-x*     0.0)   ; Offset X cho text (ti le voi cao chu)
(setq *xtd-offset-stt*   1.2)   ; Offset Y cho STT (ti le voi cao chu)
(setq *xtd-offset-z*    -1.5)   ; Offset Y cho Z (ti le voi cao chu)
(setq *xtd-offset-desc* -3.0)   ; Offset Y cho ten diem (ti le voi cao chu)

;;; ====== HAM TIEN ICH: DOC CLIPBOARD ======
(defun XTD:ReadClipboard (/ htmlobj result)
  (vl-load-com)
  (setq result nil)
  (vl-catch-all-apply
    '(lambda ()
       (setq htmlobj (vlax-create-object "htmlfile"))
       (setq result
         (vlax-invoke
           (vlax-get
             (vlax-get htmlobj 'parentwindow)
             'clipboarddata)
           'getdata "Text"))
       (vlax-release-object htmlobj)
     )
  )
  result
)

;;; ====== HAM TIEN ICH: TACH CHUOI THEO KY TU PHAN CACH ======
(defun XTD:SplitString (str delim / pos result token)
  (setq result '())
  (while (setq pos (vl-string-search delim str))
    (setq token (substr str 1 pos))
    (setq result (append result (list token)))
    (setq str (substr str (+ pos (strlen delim) 1)))
  )
  (setq result (append result (list str)))
  result
)

;;; ====== HAM TIEN ICH: LOAI BO KHOANG TRANG ======
(defun XTD:Trim (str / )
  (if (= (type str) 'STR)
    (progn
      (while (and (> (strlen str) 0) (= (substr str 1 1) " "))
        (setq str (substr str 2))
      )
      (while (and (> (strlen str) 0) (= (substr str (strlen str) 1) " "))
        (setq str (substr str 1 (1- (strlen str))))
      )
      (while (vl-string-search "\r" str)
        (setq str (vl-string-subst "" "\r" str))
      )
      str
    )
    ""
  )
)

;;; ====== HAM TIEN ICH: CHUYEN CHUOI SANG SO (XU LY DAU PHAY) ======
(defun XTD:Str2Num (str / )
  (setq str (XTD:Trim str))
  (while (vl-string-search "," str)
    (setq str (vl-string-subst "." "," str))
  )
  (if (and str (/= str "") (distof str 2))
    (distof str 2)
    nil
  )
)

;;; ====== HAM: PARSE DU LIEU TU CHUOI CLIPBOARD ======
;;; Tra ve danh sach: ((stt x y z desc) (stt x y z desc) ...)
(defun XTD:ParseData (cliptext xy-order / lines cols stt xval yval zval desc
                                         result line tokens skip-first)
  (setq result '())
  (setq lines (XTD:SplitString cliptext "\n"))
  
  ;; Kiem tra dong dau tien co phai header khong
  (setq skip-first nil)
  (if (> (length lines) 0)
    (progn
      (setq tokens (XTD:SplitString (car lines) "\t"))
      (if (> (length tokens) 1)
        (progn
          (if (null (XTD:Str2Num (nth 1 tokens)))
            (setq skip-first T)
          )
        )
      )
    )
  )
  
  ;; Duyet tung dong
  (foreach line lines
    (if skip-first
      (setq skip-first nil)
      (progn
        (setq line (XTD:Trim line))
        (if (and line (/= line ""))
          (progn
            (setq tokens (XTD:SplitString line "\t"))
            (setq desc "") ;; Mac dinh khong co mo ta
            
            (cond
              ;; 5 cot: STT, X/N, Y/E, Z, Ten diem
              ((>= (length tokens) 5)
               (setq stt  (XTD:Trim (nth 0 tokens)))
               (if (= xy-order 0)
                 (progn
                   (setq xval (XTD:Str2Num (nth 1 tokens)))
                   (setq yval (XTD:Str2Num (nth 2 tokens)))
                 )
                 (progn
                   (setq yval (XTD:Str2Num (nth 1 tokens)))
                   (setq xval (XTD:Str2Num (nth 2 tokens)))
                 )
               )
               (setq zval (XTD:Str2Num (nth 3 tokens)))
               (setq desc (XTD:Trim (nth 4 tokens)))
               (if (and xval yval)
                 (setq result (append result
                   (list (list stt xval yval (if zval zval 0.0) desc))))
               )
              )
              
              ;; 4 cot: STT, X/N, Y/E, Z
              ((= (length tokens) 4)
               (setq stt  (XTD:Trim (nth 0 tokens)))
               (if (= xy-order 0)
                 (progn
                   (setq xval (XTD:Str2Num (nth 1 tokens)))
                   (setq yval (XTD:Str2Num (nth 2 tokens)))
                 )
                 (progn
                   (setq yval (XTD:Str2Num (nth 1 tokens)))
                   (setq xval (XTD:Str2Num (nth 2 tokens)))
                 )
               )
               (setq zval (XTD:Str2Num (nth 3 tokens)))
               (if (and xval yval)
                 (setq result (append result
                   (list (list stt xval yval (if zval zval 0.0) ""))))
               )
              )
              
              ;; 3 cot: STT, X, Y (khong co Z va Ten diem)
              ((= (length tokens) 3)
               (setq stt  (XTD:Trim (nth 0 tokens)))
               (if (= xy-order 0)
                 (progn
                   (setq xval (XTD:Str2Num (nth 1 tokens)))
                   (setq yval (XTD:Str2Num (nth 2 tokens)))
                 )
                 (progn
                   (setq yval (XTD:Str2Num (nth 1 tokens)))
                   (setq xval (XTD:Str2Num (nth 2 tokens)))
                 )
               )
               (if (and xval yval)
                 (setq result (append result
                   (list (list stt xval yval 0.0 ""))))
               )
              )
              
              ;; 2 cot: X, Y (khong co gi khac)
              ((= (length tokens) 2)
               (if (= xy-order 0)
                 (progn
                   (setq xval (XTD:Str2Num (nth 0 tokens)))
                   (setq yval (XTD:Str2Num (nth 1 tokens)))
                 )
                 (progn
                   (setq yval (XTD:Str2Num (nth 0 tokens)))
                   (setq xval (XTD:Str2Num (nth 1 tokens)))
                 )
               )
               (if (and xval yval)
                 (setq result (append result
                   (list (list "" xval yval 0.0 ""))))
               )
              )
            ) ;; end cond
          ) ;; end progn
        ) ;; end if line not empty
      ) ;; end progn (not skip)
    ) ;; end if skip-first
  ) ;; end foreach
  
  result
)

;;; ====== HAM: TAO LAYER NEU CHUA CO ======
(defun XTD:MakeLayer (layname color / )
  (if (not (tblsearch "LAYER" layname))
    (entmake
      (list
        '(0 . "LAYER")
        '(100 . "AcDbSymbolTableRecord")
        '(100 . "AcDbLayerTableRecord")
        (cons 2 layname)
        '(70 . 0)
        (cons 62 color)
        '(6 . "Continuous")
      )
    )
  )
)

;;; ====== HAM: TAO BLOCK DAU CONG (+) ======
(defun XTD:CreateBlock (blkname blksize / halfsize)
  (setq halfsize (* blksize 0.5))
  
  (if (not (tblsearch "BLOCK" blkname))
    (progn
      (entmake
        (list
          '(0 . "BLOCK")
          (cons 2 blkname)
          '(70 . 2)
          '(10 0.0 0.0 0.0)
        )
      )
      
      ;; Duong ngang (-)
      (entmake
        (list
          '(0 . "LINE")
          '(8 . "0")
          (list 10 (- halfsize) 0.0 0.0)
          (list 11 halfsize 0.0 0.0)
        )
      )
      
      ;; Duong doc (|)
      (entmake
        (list
          '(0 . "LINE")
          '(8 . "0")
          (list 10 0.0 (- halfsize) 0.0)
          (list 11 0.0 halfsize 0.0)
        )
      )
      
      ;; Vong tron nho o tam
      (entmake
        (list
          '(0 . "CIRCLE")
          '(8 . "0")
          '(10 0.0 0.0 0.0)
          (cons 40 (* halfsize 0.3))
        )
      )
      
      ;; Dong block
      (entmake
        (list
          '(0 . "ENDBLK")
          (cons 2 blkname)
        )
      )
      
      (princ (strcat "\n[XTD] Da tao block dinh nghia: " blkname))
    )
  )
)

;;; ====== HAM: CHEN 1 DIEM TOA DO LEN BAN VE ======
(defun XTD:InsertPoint (stt px py pz desc texth blkname
                        show-stt show-z show-desc
                        lay-pt lay-stt lay-z lay-desc
                        col-pt col-stt col-z col-desc
                        off-stt-x off-stt-y off-z-x off-z-y
                        off-desc-x off-desc-y
                        / pt pt-stt pt-z pt-desc ztext)
  
  (setq pt (list px py 0.0))
  
  ;; 1. Chen block dau cong
  (entmake
    (list
      '(0 . "INSERT")
      (cons 8 lay-pt)
      (cons 62 col-pt)
      (cons 2 blkname)
      (cons 10 pt)
      '(41 . 1.0)
      '(42 . 1.0)
      '(43 . 1.0)
      '(50 . 0.0)
    )
  )
  
  ;; 2. Text so thu tu (STT) - phia tren phai
  (if (and show-stt stt (/= stt ""))
    (progn
      (setq pt-stt (list
        (+ px (* off-stt-x texth))
        (+ py (* off-stt-y texth))
        0.0))
      (entmake
        (list
          '(0 . "TEXT")
          (cons 8 lay-stt)
          (cons 62 col-stt)
          (cons 10 pt-stt)
          (cons 11 pt-stt)
          (cons 40 texth)
          (cons 1 stt)
          '(72 . 0)
          '(73 . 0)
          '(50 . 0.0)
        )
      )
    )
  )
  
  ;; 3. Text cao do Z - phia duoi phai
  (if (and show-z (/= pz 0.0))
    (progn
      (setq ztext (rtos pz 2 3))
      (setq pt-z (list
        (+ px (* off-z-x texth))
        (+ py (* off-z-y texth))
        0.0))
      (entmake
        (list
          '(0 . "TEXT")
          (cons 8 lay-z)
          (cons 62 col-z)
          (cons 10 pt-z)
          (cons 11 pt-z)
          (cons 40 (* texth 0.8))
          (cons 1 ztext)
          '(72 . 0)
          '(73 . 0)
          '(50 . 0.0)
        )
      )
    )
  )
  
  ;; 4. Text ten diem / mo ta - phia duoi cao do
  (if (and show-desc desc (/= desc ""))
    (progn
      (setq pt-desc (list
        (+ px (* off-desc-x texth))
        (+ py (* off-desc-y texth))
        0.0))
      (entmake
        (list
          '(0 . "TEXT")
          (cons 8 lay-desc)
          (cons 62 col-desc)
          (cons 10 pt-desc)
          (cons 11 pt-desc)
          (cons 40 (* texth 0.7))
          (cons 1 desc)
          '(72 . 0)
          '(73 . 0)
          '(50 . 0.0)
        )
      )
    )
  )
)

;;; ====== HAM: HOI NGUOI DUNG CAI DAT ======
(defun XTD:AskSettings (/ inp)
  (princ "\n╔═══════════════════════════════════════════════╗")
  (princ "\n║     XUAT TOA DO TU EXCEL LEN CAD (v1.1)      ║")
  (princ "\n║     Text + Block - Nhe File                   ║")
  (princ "\n║     Ho tro 5 cot: STT, X, Y, Z, Ten diem     ║")
  (princ "\n╚═══════════════════════════════════════════════╝")
  
  ;; Cao chu
  (setq inp (getreal
    (strcat "\nCao chu text <" (rtos *xtd-textheight* 2 1) ">: ")))
  (if inp (setq *xtd-textheight* inp))
  
  ;; Thu tu toa do
  (initget "XY YX")
  (setq inp (getkword
    (strcat "\nThu tu cot toa do [XY/YX] <"
            (if (= *xtd-xy-order* 0) "XY" "YX") ">: ")))
  (cond
    ((= inp "XY") (setq *xtd-xy-order* 0))
    ((= inp "YX") (setq *xtd-xy-order* 1))
  )
  
  ;; Hien thi STT
  (initget "Co Khong")
  (setq inp (getkword
    (strcat "\nHien thi so thu tu? [Co/Khong] <"
            (if *xtd-show-stt* "Co" "Khong") ">: ")))
  (cond
    ((= inp "Co") (setq *xtd-show-stt* T))
    ((= inp "Khong") (setq *xtd-show-stt* nil))
  )
  
  ;; Hien thi Z
  (initget "Co Khong")
  (setq inp (getkword
    (strcat "\nHien thi cao do Z? [Co/Khong] <"
            (if *xtd-show-z* "Co" "Khong") ">: ")))
  (cond
    ((= inp "Co") (setq *xtd-show-z* T))
    ((= inp "Khong") (setq *xtd-show-z* nil))
  )
  
  ;; Hien thi ten diem
  (initget "Co Khong")
  (setq inp (getkword
    (strcat "\nHien thi ten diem / mo ta? [Co/Khong] <"
            (if *xtd-show-desc* "Co" "Khong") ">: ")))
  (cond
    ((= inp "Co") (setq *xtd-show-desc* T))
    ((= inp "Khong") (setq *xtd-show-desc* nil))
  )
  
  ;; Kich thuoc block
  (setq inp (getreal
    (strcat "\nKich thuoc dau cong (+) <" (rtos *xtd-block-size* 2 1) ">: ")))
  (if inp (setq *xtd-block-size* inp))
)

;;; ====== LENH CHINH: XTD - XUAT TOA DO TU CLIPBOARD ======
(defun c:XTD (/ cliptext data count oldcmd oldosm oldecho)
  (vl-load-com)
  
  (setq oldcmd  (getvar "CMDECHO"))
  (setq oldosm  (getvar "OSMODE"))
  (setq oldecho (getvar "NOMUTT"))
  (setvar "CMDECHO" 0)
  (setvar "OSMODE" 0)
  
  ;; Hoi cai dat
  (XTD:AskSettings)
  
  ;; Doc clipboard
  (princ "\n[XTD] Dang doc du lieu tu clipboard (Ctrl+C tu Excel)...")
  (setq cliptext (XTD:ReadClipboard))
  
  (if (or (null cliptext) (= cliptext ""))
    (progn
      (princ "\n[LOI] Clipboard trong! Hay copy du lieu tu Excel truoc (Ctrl+C).")
      (princ "\n      Dinh dang: STT | X(N) | Y(E) | Z | Ten diem")
      (setvar "CMDECHO" oldcmd)
      (setvar "OSMODE" oldosm)
      (princ)
    )
    (progn
      (setq data (XTD:ParseData cliptext *xtd-xy-order*))
      
      (if (or (null data) (= (length data) 0))
        (progn
          (princ "\n[LOI] Khong phan tich duoc du lieu! Kiem tra dinh dang Excel:")
          (princ "\n      Cot 1: STT | Cot 2: X/N | Cot 3: Y/E | Cot 4: Z | Cot 5: Ten diem")
          (setvar "CMDECHO" oldcmd)
          (setvar "OSMODE" oldosm)
          (princ)
        )
        (progn
          (princ (strcat "\n[XTD] Tim thay " (itoa (length data)) " diem toa do."))
          
          ;; Tao cac layer
          (XTD:MakeLayer *xtd-layer-point* *xtd-color-point*)
          (XTD:MakeLayer *xtd-layer-stt*   *xtd-color-stt*)
          (XTD:MakeLayer *xtd-layer-caodo* *xtd-color-z*)
          (XTD:MakeLayer *xtd-layer-desc*  *xtd-color-desc*)
          
          ;; Tao block dinh nghia
          (XTD:CreateBlock *xtd-block-name* *xtd-block-size*)
          
          ;; Bat dau undo group
          (command "_.UNDO" "_Begin")
          
          ;; Chen tung diem
          (setq count 0)
          (foreach pt data
            (XTD:InsertPoint
              (nth 0 pt)              ; STT
              (nth 1 pt)              ; X
              (nth 2 pt)              ; Y
              (nth 3 pt)              ; Z
              (nth 4 pt)              ; Ten diem
              *xtd-textheight*        ; Cao chu
              *xtd-block-name*        ; Ten block
              *xtd-show-stt*          ; Hien STT?
              *xtd-show-z*            ; Hien Z?
              *xtd-show-desc*         ; Hien ten diem?
              *xtd-layer-point*       ; Layer diem
              *xtd-layer-stt*         ; Layer STT
              *xtd-layer-caodo*       ; Layer Z
              *xtd-layer-desc*        ; Layer ten diem
              *xtd-color-point*       ; Mau diem
              *xtd-color-stt*         ; Mau STT
              *xtd-color-z*           ; Mau Z
              *xtd-color-desc*        ; Mau ten diem
              *xtd-offset-x*          ; Offset X STT
              *xtd-offset-stt*        ; Offset Y STT
              *xtd-offset-x*          ; Offset X Z
              *xtd-offset-z*          ; Offset Y Z
              *xtd-offset-x*          ; Offset X ten diem
              *xtd-offset-desc*       ; Offset Y ten diem
            )
            (setq count (1+ count))
          )
          
          ;; Ket thuc undo group
          (command "_.UNDO" "_End")
          
          ;; Zoom toi tat ca diem
          (command "_.ZOOM" "_Extents")
          
          ;; Thong bao ket qua
          (princ "\n╔═══════════════════════════════════════════════╗")
          (princ (strcat
            "\n║  HOAN TAT! Da phun "
            (itoa count) " diem toa do len ban ve"))
          (princ "\n║                                               ║")
          (princ (strcat "\n║  Layer diem    : " *xtd-layer-point*))
          (princ (strcat "\n║  Layer STT     : " *xtd-layer-stt*))
          (princ (strcat "\n║  Layer Z       : " *xtd-layer-caodo*))
          (princ (strcat "\n║  Layer ten diem: " *xtd-layer-desc*))
          (princ (strcat "\n║  Cao chu       : " (rtos *xtd-textheight* 2 1)))
          (princ "\n╚═══════════════════════════════════════════════╝")
          
          (setvar "CMDECHO" oldcmd)
          (setvar "OSMODE" oldosm)
          (princ)
        )
      )
    )
  )
)

;;; ====== LENH PHU: XTDF - DOC TU FILE TXT/CSV ======
(defun c:XTDF (/ filepath fhandle fline cliptext data count
                 oldcmd oldosm)
  (vl-load-com)
  
  (setq oldcmd  (getvar "CMDECHO"))
  (setq oldosm  (getvar "OSMODE"))
  (setvar "CMDECHO" 0)
  (setvar "OSMODE" 0)
  
  ;; Hoi cai dat
  (XTD:AskSettings)
  
  ;; Chon file
  (setq filepath
    (getfiled "Chon file toa do (TXT hoac CSV)"
              "" "txt;csv" 4))
  
  (if (null filepath)
    (progn
      (princ "\n[XTD] Da huy. Khong chon file nao.")
      (setvar "CMDECHO" oldcmd)
      (setvar "OSMODE" oldosm)
      (princ)
    )
    (progn
      (setq cliptext "")
      (setq fhandle (open filepath "r"))
      (if fhandle
        (progn
          (while (setq fline (read-line fhandle))
            (if (and (not (vl-string-search "\t" fline))
                     (vl-string-search "," fline))
              (progn
                (while (vl-string-search ";" fline)
                  (setq fline (vl-string-subst "\t" ";" fline)))
                (while (vl-string-search "," fline)
                  (setq fline (vl-string-subst "\t" "," fline)))
              )
            )
            (setq cliptext (strcat cliptext fline "\n"))
          )
          (close fhandle)
          
          (setq data (XTD:ParseData cliptext *xtd-xy-order*))
          
          (if (or (null data) (= (length data) 0))
            (progn
              (princ "\n[LOI] Khong doc duoc du lieu tu file!")
              (princ "\n      Dinh dang: STT,X,Y,Z,Ten diem")
            )
            (progn
              (princ (strcat "\n[XTD] Tim thay " (itoa (length data))
                             " diem tu file: " filepath))
              
              (XTD:MakeLayer *xtd-layer-point* *xtd-color-point*)
              (XTD:MakeLayer *xtd-layer-stt*   *xtd-color-stt*)
              (XTD:MakeLayer *xtd-layer-caodo* *xtd-color-z*)
              (XTD:MakeLayer *xtd-layer-desc*  *xtd-color-desc*)
              (XTD:CreateBlock *xtd-block-name* *xtd-block-size*)
              
              (command "_.UNDO" "_Begin")
              (setq count 0)
              (foreach pt data
                (XTD:InsertPoint
                  (nth 0 pt) (nth 1 pt) (nth 2 pt) (nth 3 pt) (nth 4 pt)
                  *xtd-textheight* *xtd-block-name*
                  *xtd-show-stt* *xtd-show-z* *xtd-show-desc*
                  *xtd-layer-point* *xtd-layer-stt* *xtd-layer-caodo* *xtd-layer-desc*
                  *xtd-color-point* *xtd-color-stt* *xtd-color-z* *xtd-color-desc*
                  *xtd-offset-x* *xtd-offset-stt*
                  *xtd-offset-x* *xtd-offset-z*
                  *xtd-offset-x* *xtd-offset-desc*
                )
                (setq count (1+ count))
              )
              (command "_.UNDO" "_End")
              (command "_.ZOOM" "_Extents")
              
              (princ (strcat "\n[XTD] HOAN TAT! Da phun "
                             (itoa count) " diem toa do."))
            )
          )
        )
        (princ (strcat "\n[LOI] Khong mo duoc file: " filepath))
      )
      
      (setvar "CMDECHO" oldcmd)
      (setvar "OSMODE" oldosm)
      (princ)
    )
  )
)

;;; ====== LENH PHU: TDBLOCK - TAO LAI BLOCK ======
(defun c:TDBLOCK (/ inp)
  (princ "\n[TDBLOCK] Tao lai block dau diem toa do.")
  
  (if (tblsearch "BLOCK" *xtd-block-name*)
    (progn
      (princ (strcat "\n  Block '" *xtd-block-name* "' da ton tai."))
      (initget "Co Khong")
      (setq inp (getkword "\n  Xoa va tao lai? [Co/Khong] <Khong>: "))
      (if (= inp "Co")
        (progn
          (command "_.PURGE" "_Blocks" *xtd-block-name* "_No")
          (setq inp (getreal
            (strcat "\nKich thuoc moi <" (rtos *xtd-block-size* 2 1) ">: ")))
          (if inp (setq *xtd-block-size* inp))
          (XTD:CreateBlock *xtd-block-name* *xtd-block-size*)
        )
        (princ "\n  Da huy.")
      )
    )
    (progn
      (setq inp (getreal
        (strcat "\nKich thuoc block <" (rtos *xtd-block-size* 2 1) ">: ")))
      (if inp (setq *xtd-block-size* inp))
      (XTD:CreateBlock *xtd-block-name* *xtd-block-size*)
    )
  )
  (princ)
)

;;; ====== ALIAS LENH ======
(defun c:XUATTOADO () (c:XTD))

;;; ====== THONG BAO KHI LOAD ======
(princ "\n╔═══════════════════════════════════════════════╗")
(princ "\n║  XUAT TOA DO EXCEL -> CAD (Text + Block)     ║")
(princ "\n║  Phien ban 1.1 - Ho tro 5 cot                ║")
(princ "\n╠═══════════════════════════════════════════════╣")
(princ "\n║  Lenh:                                       ║")
(princ "\n║    XTD    - Phun toa do tu clipboard Excel    ║")
(princ "\n║    XTDF   - Phun toa do tu file TXT/CSV      ║")
(princ "\n║    TDBLOCK - Tao lai block dau diem           ║")
(princ "\n╠═══════════════════════════════════════════════╣")
(princ "\n║  Dinh dang Excel:                            ║")
(princ "\n║    STT | X(N) | Y(E) | Z | Ten diem          ║")
(princ "\n╚═══════════════════════════════════════════════╝")
(princ)
