(defun c:WIREARC_CHAIN (/ startBlock startPt startName sel nextBlock nextPt nextName midPt wireEnt dx dy dist height maxHeight perpVec perpLen)
  ;; Pick first block
  (setq startBlock (car (entsel "\nSelect first block to start wiring: ")))
  ;; Base point from block object
  (setq startPt (vlax-get (vlax-ename->vla-object startBlock) 'InsertionPoint))
  ;; Get NAME
  (setq startName nil)
  (foreach att (vlax-invoke (vlax-ename->vla-object startBlock) 'GetAttributes)
    (if (= (strcase (vla-get-TagString att)) "NAME")
      (setq startName (vla-get-TextString att))
    )
  )

  (princ "\nContinuous chaining mode: pick next block. Press ENTER or ESC to exit.")

  ;; Loop
  (while T
    (setq sel (entsel "\nSelect next block: "))
    (if (and sel (= sel "")) (exit))
    (if sel
      (progn
        (setq nextBlock (car sel))
        (if (= nextBlock startBlock)
          (princ "\nCannot connect to the same block. Pick another block.")
          (progn
            ;; Base point
            (setq nextPt (vlax-get (vlax-ename->vla-object nextBlock) 'InsertionPoint))
            ;; Get NAME
            (setq nextName nil)
            (foreach att (vlax-invoke (vlax-ename->vla-object nextBlock) 'GetAttributes)
              (if (= (strcase (vla-get-TagString att)) "NAME")
                (setq nextName (vla-get-TextString att))
              )
            )

            ;; Vector and distance
            (setq dx (- (car nextPt) (car startPt)))
            (setq dy (- (cadr nextPt) (cadr startPt)))
            (setq dist (sqrt (+ (* dx dx) (* dy dy))))

            ;; Arc height proportional
            (setq height (* dist 0.12))
            (if (> height 15.0) (setq height 15.0))

            ;; Perpendicular vector for bend
            (setq perpVec (list (- dy) dx))
            (setq perpLen (sqrt (+ (* (car perpVec) (car perpVec)) (* (cadr perpVec) (cadr perpVec)))))
            (setq perpVec (list (* (/ (car perpVec) perpLen) height)
                                (* (/ (cadr perpVec) perpLen) height)
                                0))

            ;; Midpoint
            (setq midPt (list (+ (/ (+ (car startPt) (car nextPt)) 2) (car perpVec))
                              (+ (/ (+ (cadr startPt) (cadr nextPt)) 2) (cadr perpVec))
                              0))

            ;; Draw arc using ARC command, exact start/mid/end points
            (command "_ARC" (list (car startPt) (cadr startPt) 0)
                              (list (car midPt) (cadr midPt) 0)
                              (list (car nextPt) (cadr nextPt) 0))
            (setq wireEnt (entlast))

            ;; Register XData
            (regapp "WIREAPP")
            ;; Attach XData
            (entmod
              (append (entget wireEnt)
                (list
                  (list -3
                    (list "WIREAPP"
                      (cons 1000 (strcat "FROM=" startName))
                      (cons 1000 (strcat "TO=" nextName))
                    )
                  )
                )
              )
            )

            (princ (strcat "\nStored connection: " startName " -> " nextName))

            ;; Update start
            (setq startBlock nextBlock)
            (setq startPt nextPt)
            (setq startName nextName)
          )
        )
      )
    )
    (princ)
  )
)