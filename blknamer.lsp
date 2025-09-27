(defun c:BLKNAMER ( / prefix txtsize startnum pos choice offset count txtpt )

  ;; 1. Name
  (setq prefix (getstring "\nEnter prefix (e.g. PL): "))

  ;; 2. Text size & starting count
  (setq txtsize (getreal "\nEnter text height (e.g. 2.5): "))
  (setq startnum (getint "\nCount starts from <1>: "))
  (if (null startnum) (setq startnum 1))

  ;; 3. Placement details
  (setq pos (getstring "\nChoose label position [Right/Left/Above/Below/Manual] <Right>: "))
  (if (or (null pos) (= pos "")) (setq pos "Right"))
  (setq pos (strcase pos))

  (cond
    ((or (= pos "R") (= pos "RIGHT"))   (setq choice "Right"))
    ((or (= pos "L") (= pos "LEFT"))    (setq choice "Left"))
    ((or (= pos "A") (= pos "ABOVE"))   (setq choice "Above"))
    ((or (= pos "B") (= pos "BELOW"))   (setq choice "Below"))
    ((or (= pos "M") (= pos "MANUAL"))  (setq choice "Manual"))
    (T (progn (princ "\nInvalid option, defaulting to Right.") (setq choice "Right")))
  )

  ;; 4. Manual mode → direct placement, no block selection
  (if (= choice "Manual")
    (progn
      (setq count startnum)
      (princ "\n👉 Click on drawing to place labels, press ESC to stop.")
      (while
        (setq txtpt (getpoint "\nPick label position or ESC to stop: "))
        (entmake
          (list
            (cons 0 "TEXT")
            (cons 10 txtpt)
            (cons 11 txtpt)
            (cons 40 txtsize)
            (cons 1 (strcat prefix (itoa count)))
            (cons 7 "Standard")
            (cons 72 1) ; middle align
            (cons 73 2)
          )
        )
        (setq count (1+ count))
      )
    )
    ;; 5. Auto placement mode (needs offset + object selection)
    (progn
      (setq offset (getreal "\nEnter offset distance from block: "))
      (setq sel (ssget '((0 . "INSERT"))))
      (if sel
        (progn
          (setq count startnum idx 0)
          (repeat (sslength sel)
            (setq ent (ssname sel idx))  
            (setq idx (1+ idx))

            ;; bounding box
            (vla-getboundingbox (vlax-ename->vla-object ent) 'minpt 'maxpt)
            (setq minpt (vlax-safearray->list minpt))
            (setq maxpt (vlax-safearray->list maxpt))

            ;; compute text point
            (cond
              ((= choice "Right")
                (setq txtpt (list (+ (car maxpt) offset)
                                  (/ (+ (cadr minpt) (cadr maxpt)) 2.0))))
              ((= choice "Left")
                (setq txtpt (list (- (car minpt) offset)
                                  (/ (+ (cadr minpt) (cadr maxpt)) 2.0))))
              ((= choice "Above")
                (setq txtpt (list (/ (+ (car minpt) (car maxpt)) 2.0)
                                  (+ (cadr maxpt) offset))))
              ((= choice "Below")
                (setq txtpt (list (/ (+ (car minpt) (car maxpt)) 2.0)
                                  (- (cadr minpt) offset))))
            )

            ;; create text entity
            (entmake
              (list
                (cons 0 "TEXT")
                (cons 10 txtpt)
                (cons 11 txtpt)
                (cons 40 txtsize)
                (cons 1 (strcat prefix (itoa count)))
                (cons 7 "Standard")
                (cons 72 1)
                (cons 73 2)
              )
            )
            (setq count (1+ count))
          )
          (princ "\n✅ Blocks labeled successfully.")
        )
        (princ "\n⚠️ No blocks selected.")
      )
    )
  )
  (princ)
)
