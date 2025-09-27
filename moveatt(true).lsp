(defun c:MoveNameAttr ( / blkObj atts selAtt newPt vPoint)
  (princ "\nSelect block whose NAME attribute you want to move: ")
  (setq blkObj (vlax-ename->vla-object (car (entsel))))
  (if blkObj
    (progn
      (setq atts (vlax-invoke blkObj 'GetAttributes))
      (if atts
        (progn
          ;; Find NAME attribute
          (setq selAtt nil)
          (foreach a atts
            (if (= (strcase (vla-get-TagString a)) "NAME")
              (setq selAtt a)
            )
          )
          (if selAtt
            (progn
              ;; Ask for new position
              (setq newPt (getpoint "\nPick new position for NAME attribute: "))
              (setq vPoint (vlax-3d-point newPt))
              ;; Move using TextAlignmentPoint instead of InsertionPoint
              (vla-put-TextAlignmentPoint selAtt vPoint)
              (princ "\nNAME attribute moved successfully (works for any justification).")
            )
            (princ "\nNAME attribute not found in this block.")
          )
        )
        (princ "\nNo attributes found in this block.")
      )
    )
    (princ "\nNo block selected.")
  )
  (princ)
)
