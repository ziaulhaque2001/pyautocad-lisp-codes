(defun c:SeqName ( / ss baseName startNum count blkObj atts i)
  (princ "\nSelect blocks to rename sequentially: ")
  (setq ss (ssget))  ; select multiple blocks
  (if ss
    (progn
      ;; Ask for base name and starting number
      (setq baseName (getstring T "\nEnter base name (e.g., PL): "))
      (setq startNum (getint "\nEnter starting number: "))
      ;; Loop through selected blocks
      (setq i 0)
      (while (< i (sslength ss))
        (setq blkObj (vlax-ename->vla-object (ssname ss i)))
        (setq atts (vlax-invoke blkObj 'GetAttributes))
        ;; Update NAME attribute
        (foreach a atts
          (if (= (strcase (vla-get-TagString a)) "NAME")
            (vla-put-TextString a (strcat baseName (itoa startNum)))
          )
        )
        (setq startNum (1+ startNum))
        (setq i (1+ i))
      )
      (princ "\nSequential naming complete.")
    )
    (princ "\nNo blocks selected.")
  )
  (princ)
)
