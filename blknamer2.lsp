(defun c:BlkNamerSmart (/ selArea blkPick blkEnt blkName txtsize prefix ssAll blocks tol grouped pt added row rowSorted sorted count e ed i j)

  (setvar "CMDECHO" 0)

  ;; 1. Select area
  (princ "\nSelect the AREA containing the blocks to rename: ")
  (setq selArea (ssget))
  (if (not selArea)
    (progn (princ "\nNo area selected.") (setvar "CMDECHO" 1) (exit))
  )

  ;; 2. Select one block type inside the area
  (princ "\nSelect ONE block type inside the area to rename: ")
  (setq blkPick (ssget ":S" '((0 . "INSERT"))))
  (if (not blkPick)
    (progn (princ "\nNo block selected.") (setvar "CMDECHO" 1) (exit))
  )

  ;; Get block name from picked block
  (setq blkEnt (ssname blkPick 0))
  (setq blkName (cdr (assoc 2 (entget blkEnt))))

  ;; 3. User input
  (initget 7)
  (setq txtsize (getreal "\nEnter text size: "))
  (setq prefix (getstring T "\nEnter name prefix: "))

  ;; 4. Collect all blocks of same type inside area
  (setq ssAll (ssadd))
  (setq i 0)
  (repeat (sslength selArea)
    (setq e (ssname selArea i))
    (setq ed (entget e))
    (if (and (= (cdr (assoc 0 ed)) "INSERT")
             (= (cdr (assoc 2 ed)) blkName))
      (ssadd e ssAll)
    )
    (setq i (1+ i))
  )

  (if (= (sslength ssAll) 0)
    (progn (princ "\nNo matching blocks found.") (setvar "CMDECHO" 1) (exit))
  )

  ;; 5. Extract insertion points
  (setq blocks '())
  (setq j 0)
  (repeat (sslength ssAll)
    (setq e (ssname ssAll j))
    (setq pt (cdr (assoc 10 (entget e))))
    (setq blocks (cons (list e pt) blocks))
    (setq j (1+ j))
  )

  ;; 6. Grouping tolerance
  (setq tol (* txtsize 2.0))

  ;; 7. Group into rows by Y value
  (setq grouped '())
  (foreach b (vl-sort blocks (function (lambda (a b) (< (cadr (cadr a)) (cadr (cadr b)))))))
    (setq pt (cadr b))
    (setq added nil)
    (foreach row grouped
      (if (< (abs (- (cadr (cadar row)) (cadr pt))) tol)
        (progn
          (setq grouped (subst (append row (list b)) row grouped))
          (setq added T)
        )
      )
    )
    (if (not added)
      (setq grouped (append grouped (list (list b))))
    )
  )

  ;; 8. Sort rows bottom→top, inside row left→right
  (setq grouped
    (vl-sort grouped
      (function (lambda (r1 r2)
                  (< (cadr (cadar r1)) (cadr (cadar r2))))))
  )

  (setq sorted '())
  (foreach row grouped
    (setq rowSorted
      (vl-sort row (function (lambda (a b) (< (car (cadr a)) (car (cadr b)))))))
    (setq sorted (append sorted rowSorted))
  )

  ;; 9. Place labels
  (setq count 1)
  (foreach b sorted
    (setq pt (cadr b))
    (entmake
      (list
        (cons 0 "TEXT")
        (cons 10 pt)
        (cons 11 pt)
        (cons 40 txtsize)
        (cons 1 (strcat prefix (itoa count)))
        (cons 7 "Standard")
        (cons 72 1) ; center
        (cons 73 2) ; middle
        (cons 50 0.0) ; horizontal
      )
    )
    (setq count (1+ count))
  )

  (princ (strcat "\n✅ Placed " (itoa (1- count)) " labels."))
  (setvar "CMDECHO" 1)
  (princ)
)
