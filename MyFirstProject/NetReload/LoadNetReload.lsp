;;; LoadNetReload.lsp
;;; Tự động load NetReload.dll để sử dụng lệnh NRL và RELOAD
;;; Đặt file này vào thư mục Support của AutoCAD hoặc load thủ công bằng APPLOAD

(defun C:LOADNRL ()
  (princ "\nLoading NetReload.dll...")
  (command "._NETLOAD" "c:\\Google\\Autocad 2026_API\\Autocad 2026_API\\MyFirstProject\\NetReload\\bin\\Debug\\NetReload.dll")
  (princ "\nNetReload.dll loaded successfully!")
  (princ "\nAvailable commands: NRL, RELOAD")
  (princ)
)

;;; Tự động load khi file LISP được load
(defun s::startup ()
  (C:LOADNRL)
)

;;; Load ngay khi file LISP được load lần đầu
(C:LOADNRL)

(princ "\n*** LoadNetReload.lsp loaded ***")
(princ "\n*** Type LOADNRL to reload NetReload.dll manually ***")
(princ)
