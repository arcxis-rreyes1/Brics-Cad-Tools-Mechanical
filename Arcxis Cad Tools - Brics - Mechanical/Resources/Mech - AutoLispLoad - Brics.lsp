; ============================================================
; ARCXIS MEP MECHANICAL LOADER (AUTOLOAD)
; - Detects the mapped Egnyte drive letter by finding:
;     X:\Shared\Arcxis\Engineering\
; - Loader folder:
;     X:\Shared\Arcxis\Engineering\Drafting Standards\MEP\LISP\MainLispLoader\
; - DLL:
;     Arcxis Cad Tools - Brics - Mechanical.dll
;
; Command:
;   WM = Load Arcxis Mechanical tools
; ============================================================

(vl-load-com)

; Capture loader folder at load time (avoids A:..Z: drive scan on drawing open)
(setq *rr-loader-dir*
  (cond
    ((and (boundp '*load-truename*) *load-truename*)
     (vl-filename-directory *load-truename*))
    ((setq *rr-loader-file* (findfile "Mech - AutoLispLoad - Brics.lsp"))
     (vl-filename-directory *rr-loader-file*))
    (T nil)
  )
)

(setq *rr-egnyte-drive-cache* nil)

(defun rr-dir-exists-p (p) (and p (vl-file-directory-p p)))
(defun rr-drive-root (letter) (strcat letter ":\\"))
(defun rr-path-join (a b)
  (cond
    ((or (null a) (= a "")) b)
    ((or (null b) (= b "")) a)
    ((= (substr a (strlen a) 1) "\\")
     (if (= (substr b 1 1) "\\")
       (strcat (substr a 1 (- (strlen a) 1)) b)
       (strcat a b)))
    (T
     (if (= (substr b 1 1) "\\")
       (strcat a b)
       (strcat a "\\" b)))
  )
)

(defun rr-str-find-ci (needle haystack / n h)
  (setq n (strcase needle) h (strcase haystack))
  (vl-string-search n h)
)

; Registry fallback: HKCU\Network\<Letter>\RemotePath
(defun rr-regread (key / sh val)
  (setq sh (vlax-create-object "WScript.Shell"))
  (setq val (vl-catch-all-apply '(lambda () (vlax-invoke sh 'RegRead key))))
  (vlax-release-object sh)
  (if (vl-catch-all-error-p val) nil val)
)

(defun rr-egnyte-drive-from-registry ( / letters found key unc)
  (setq letters (vl-string->list "ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
  (setq found nil)
  (foreach code letters
    (if (null found)
      (progn
        (setq key (strcat "HKEY_CURRENT_USER\\Network\\" (chr code) "\\RemotePath"))
        (setq unc (rr-regread key))
        (if (and unc (rr-str-find-ci "egnyte" unc))
          (setq found (chr code))
        )
      )
    )
  )
  found
)

; Find Egnyte drive: registry first, then A:..Z: probe as last resort
(defun rr-egnyte-drive-letter ( / letters code letter root test regLetter)
  (if *rr-egnyte-drive-cache*
    *rr-egnyte-drive-cache*
    (progn
      (setq letter nil)
      (setq regLetter (rr-egnyte-drive-from-registry))
      (if regLetter
        (progn
          (setq test (rr-path-join (rr-drive-root regLetter) "Shared\\Arcxis\\Engineering\\"))
          (if (rr-dir-exists-p test)
            (setq letter regLetter)
          )
        )
      )
      (if (null letter)
        (progn
          (setq letters (vl-string->list "ABCDEFGHIJKLMNOPQRSTUVWXYZ"))
          (foreach code letters
            (if (null letter)
              (progn
                (setq root (rr-drive-root (chr code)))
                (if (rr-dir-exists-p root)
                  (progn
                    (setq test (rr-path-join root "Shared\\Arcxis\\Engineering\\"))
                    (if (rr-dir-exists-p test)
                      (setq letter (chr code))
                    )
                  )
                )
              )
            )
          )
        )
      )
      (setq *rr-egnyte-drive-cache* letter)
      *rr-egnyte-drive-cache*
    )
  )
)

(defun rr-mech-loader-root-from-loader ( / dir)
  (setq dir *rr-loader-dir*)
  (if (and dir (rr-dir-exists-p dir))
    (if (= (substr dir (strlen dir) 1) "\\")
      dir
      (strcat dir "\\")
    )
    nil
  )
)

(defun rr-mech-loader-root ( / eng l loaderRoot)
  (setq loaderRoot (rr-mech-loader-root-from-loader))
  (if loaderRoot
    loaderRoot
    (progn
      (setq l (rr-egnyte-drive-letter))
      (if l
        (progn
          (setq eng (rr-path-join (rr-drive-root l) "Shared\\Arcxis\\Engineering\\"))
          (rr-path-join eng "Drafting Standards\\MEP\\LISP\\MainLispLoader\\")
        )
        nil
      )
    )
  )
)

(defun rr-safe-netload (dllFullPath)
  (if (and dllFullPath (findfile dllFullPath))
    (progn
      (setvar "FILEDIA" 0)
      (setvar "CMDECHO" 0)
      (command "_.NETLOAD" dllFullPath)
      (setvar "CMDECHO" 1)
      (setvar "FILEDIA" 1)
      T
    )
    (progn
      (prompt (strcat "\nDLL not found: " (if dllFullPath dllFullPath "<nil>")))
      nil
    )
  )
)

; ===========================
; COMMAND: WM (Mechanical)
; ===========================
(defun C:WM (/ olddyn root dllPath)

  (vl-load-com)
  (setq olddyn (getvar "DYNMODE"))
  (setvar "DYNMODE" 3)

  (setq root (rr-mech-loader-root))
  (if (null root)
    (progn
      (prompt "\nCould not resolve MainLispLoader path (Shared\\Arcxis\\Engineering\\Drafting Standards\\MEP\\LISP\\MainLispLoader).")
      (setvar "DYNMODE" olddyn)
      (princ)
      (exit)
    )
  )

  (setq dllPath (rr-path-join root "Arcxis Cad Tools - Brics - Mechanical.dll"))
  (rr-safe-netload dllPath)
  (prompt "\nArcxis Mechanical tools are loaded..")

  (setvar "DYNMODE" olddyn)
  (princ)
)

(princ "\nArcxis Mechanical Loader loaded. Command: WM")
(princ)
