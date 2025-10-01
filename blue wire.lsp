(defun c:BWire (/ selArea sourceBlock sourceEnt sourceBox ssTargetBlocks 
                  destBlkName destBlock destEnt destInsertPt sourceEdgePt 
                  usedPoints tolerance midPoint)
  (setvar "OSMODE" 0)
  (setvar "CMDECHO" 0)
  
  ;; 1. Get selection area for fixtures
  (princ "\nSelect the area containing the fixtures to connect: ")
  (setq selArea (ssget))
  (if (not selArea)
    (progn (princ "\nNo area selected. Exiting.") (exit))
  )

  ;; 2. Get the SOURCE block (SDB)
  (princ "\nSelect the SOURCE Sub Distribution Board (SDB) block: ")
  (while (not sourceBlock)
    (setq sourceBlock (ssget ":S" '((0 . "INSERT"))))
    (if (not sourceBlock)
      (princ "\n*** Please select a BLOCK object for the SDB: ")
    )
  )
  (setq sourceEnt (ssname sourceBlock 0))
  (setq sourceBox (getblockboundingbox sourceEnt))

  ;; 3. SIMPLIFIED: Select destination blocks
  (princ "\nSelect DESTINATION FIXTURE blocks to connect to. Press Enter when done: ")
  (setq ssTargetBlocks (ssadd))
  
  ;; Keep selecting until user presses Enter
  (while (setq destBlock (ssget ":S" '((0 . "INSERT"))))
    (setq destEnt (ssname destBlock 0))
    (setq destBlkName (cdr (assoc 2 (entget destEnt))))
    
    ;; Add all blocks of this type from the selection area
    (setq i 0)
    (repeat (sslength selArea)
      (setq e (ssname selArea i))
      (setq ent (entget e))
      (if (and (= (cdr (assoc 0 ent)) "INSERT")
               (= (cdr (assoc 2 ent)) destBlkName))
        (ssadd e ssTargetBlocks)
      )
      (setq i (1+ i))
    )
    (princ (strcat "\nFound " (itoa (sslength ssTargetBlocks)) " '" destBlkName "' blocks. Select another type or press Enter to finish: "))
  )

  (if (= (sslength ssTargetBlocks) 0)
    (progn (princ "\nNo destination blocks found. Exiting.") (exit))
  )

  (princ (strcat "\nConnecting to " (itoa (sslength ssTargetBlocks)) " fixtures with subtle ARC wires..."))

  ;; 4. Track used connection points
  (setq usedPoints '())
  (setq tolerance 10.0)

  ;; 5. Process each target block
  (setq i 0)
  (repeat (sslength ssTargetBlocks)
    (setq e (ssname ssTargetBlocks i))
    
    ;; Get the INSERTION POINT (base point) of the destination block
    (setq destInsertPt (cdr (assoc 10 (entget e))))
    
    ;; Find connection point on SOURCE SDB
    (setq sourceEdgePt (findsourcepointonperimeter sourceBox destInsertPt usedPoints tolerance))
    (setq usedPoints (cons sourceEdgePt usedPoints))

    ;; Calculate midpoint for subtle ARC
    (setq midPoint (calculatesubtlearcmidpoint sourceEdgePt destInsertPt))
    
    ;; DRAW SUBTLE ARC WIRE
    (command "_.ARC" sourceEdgePt midPoint destInsertPt "")
    
    (setq i (1+ i))
  )

  (princ "\nWiring complete! Subtle ARC wires connected to block insertion points.")
  (setvar "OSMODE")
  (setvar "CMDECHO" 1)
  (princ)
)

;;; ==========================================================
;;; SUBTLE ARC Wiring Function
;;; ==========================================================
(defun calculatesubtlearcmidpoint (pt1 pt2 / mid ang dist offset)
  (setq mid (mapcar '(lambda (a b) (/ (+ a b) 2.0)) pt1 pt2))
  (setq ang (angle pt1 pt2))
  (setq dist (distance pt1 pt2))
  (setq offset (* dist 0.04))
  
  (mapcar '+ mid (list (* offset (cos (+ ang (* pi 0.5)))) (* offset (sin (+ ang (* pi 0.5)))) 0.0))
)

;;; ==========================================================
;;; Helper Functions (unchanged)
;;; ==========================================================
(defun getblockboundingbox (ent / ll ur)
  (vla-getboundingbox (vlax-ename->vla-object ent) 'll 'ur)
  (list (vlax-safearray->list ll) (vlax-safearray->list ur))
)

(defun findsourcepointonperimeter (sourceBox destPt usedPoints tolerance / minpt maxpt xmin ymin xmax ymax
                                    candidatePt bestPt minDist tooClose dist testX testY testPt)
  (setq minpt (car sourceBox))
  (setq maxpt (cadr sourceBox))
  (setq xmin (car minpt))
  (setq ymin (cadr minpt))
  (setq xmax (car maxpt))
  (setq ymax (cadr maxpt))
  
  (setq candidatePt (closestpointonbbox sourceBox destPt))
  (setq bestPt candidatePt)
  (setq minDist (distance candidatePt destPt))
  (setq tooClose nil)
  
  (foreach usedPt usedPoints
    (if (< (distance candidatePt usedPt) tolerance)
      (setq tooClose T)
    )
  )
  
  (if tooClose
    (progn
      (foreach offset '(0.1 0.2 0.3 0.4 0.5 0.6 0.7 0.8 0.9)
        (setq testX (+ xmin (* offset (- xmax xmin))))
        (setq testPt (list testX ymin 0.0))
        (if (and (not (pointtooclose testPt usedPoints tolerance))
                 (< (setq dist (distance testPt destPt)) minDist))
          (setq bestPt testPt minDist dist)
        )
        
        (setq testPt (list testX ymax 0.0))
        (if (and (not (pointtooclose testPt usedPoints tolerance))
                 (< (setq dist (distance testPt destPt)) minDist))
          (setq bestPt testPt minDist dist)
        )
        
        (setq testY (+ ymin (* offset (- ymax ymin))))
        (setq testPt (list xmin testY 0.0))
        (if (and (not (pointtooclose testPt usedPoints tolerance))
                 (< (setq dist (distance testPt destPt)) minDist))
          (setq bestPt testPt minDist dist)
        )
        
        (setq testPt (list xmax testY 0.0))
        (if (and (not (pointtooclose testPt usedPoints tolerance))
                 (< (setq dist (distance testPt destPt)) minDist))
          (setq bestPt testPt minDist dist)
        )
      )
    )
  )
  
  bestPt
)

(defun pointtooclose (pt ptList tolerance / tooClose)
  (setq tooClose nil)
  (foreach usedPt ptList
    (if (< (distance pt usedPt) tolerance)
      (setq tooClose T)
    )
  )
  tooClose
)

(defun centroid (bbox / minpt maxpt)
  (setq minpt (car bbox))
  (setq maxpt (cadr bbox))
  (list
    (/ (+ (car minpt) (car maxpt)) 2.0)
    (/ (+ (cadr minpt) (cadr maxpt)) 2.0)
    0.0
  )
)

(defun closestpointonbbox (bbox targetpt / minpt maxpt xmin ymin xmax ymax x y)
  (setq minpt (car bbox))
  (setq maxpt (cadr bbox))
  (setq xmin (car minpt))
  (setq ymin (cadr minpt))
  (setq xmax (car maxpt))
  (setq ymax (cadr maxpt))
  (setq x (car targetpt))
  (setq y (cadr targetpt))

  (cond
    ((<= x xmin)
     (list xmin (max ymin (min y ymax)) 0.0)
    )
    ((>= x xmax)
     (list xmax (max ymin (min y ymax)) 0.0)
    )
    ((<= y ymin)
     (list (max xmin (min x xmax)) ymin 0.0)
    )
    ((>= y ymax)
     (list (max xmin (min x xmax)) ymax 0.0)
    )
    (t
     (list
       (cond
         ((< (- x xmin) (- xmax x)) xmin)
         (t xmax)
       )
       (cond
         ((< (- y ymin) (- ymax y)) ymin)
         (t ymax)
       )
       0.0
     )
    )
  )
)