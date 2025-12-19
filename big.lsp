;;; ========================================================================
;;; INTEGRATED ELECTRICAL WIRING SYSTEM FOR AUTOCAD
;;; ========================================================================
;;; 
;;; PROJECT OVERVIEW:
;;; This system automates electrical wiring design for architectural projects,
;;; specifically handling the connection between distribution boards (SB, SDB, MDB)
;;; and electrical fixtures (lights, fans, sockets, AC units, etc.).
;;;
;;; WORKFLOW:
;;; 1. Manual fixture placement on architectural drawing
;;; 2. Wire SB→SDB connections using WIREARC_CHAIN (manual click-through)
;;; 3. Wire SDB→fixtures using BWIRE (batch auto-wiring)
;;; 4. Auto-assign circuits using POPULATE_CKT (traces wire chains)
;;; 5. Export load calculations, schedules, and reports
;;;
;;; BLOCK ATTRIBUTE STRUCTURE:
;;; All blocks must have these 6 attributes:
;;;   NAME   = Unique identifier (e.g., "PL-01", "SB-01", "SDB-1")
;;;            Emergency fixtures use "E" prefix (e.g., "EPL-01")
;;;   TYPE   = Block category: "SB", "SDB", "MDB", "LIGHT", "FAN", "SOCKET", "AC", etc.
;;;   LOAD   = Wattage (fixtures only, empty for distribution boards)
;;;   CKT    = Circuit assignment (populated by POPULATE_CKT, e.g., "C1", "EC2")
;;;            Only SB blocks get CKT values, fixtures inherit from parent SB
;;;   SPARE1 = Auto-populated: "NORMAL" or "EMERGENCY" based on NAME prefix
;;;   SPARE2 = Reserved for future use
;;;
;;; HIERARCHY STRUCTURE:
;;;   Main Switchboard (SB)
;;;     ├─ Sub Distribution Board 1 (SDB)
;;;     │   └─ Switchboard 1 (SB) [CKT="C1"]
;;;     │       ├─ Normal fixtures (lights, fans, sockets)
;;;     │       └─ Emergency fixtures (E-prefix, separate circuit tracking)
;;;     └─ Sub Distribution Board 2 (SDB)
;;;         └─ Switchboard 2 (SB) [CKT="C2"]
;;;
;;; CIRCUIT ASSIGNMENT LOGIC:
;;;   - Circuits numbered per SDB (SDB-1: C1, C2... | SDB-2: C1, C2...)
;;;   - Emergency circuits: EC1, EC2... (separate from normal)
;;;   - Circuit assignment by wire chain tracing:
;;;     * SDB→SB1→SB2→SB3 (daisy chain) = All get same circuit (e.g., C1)
;;;     * SDB→SB4 (separate wire) = New circuit (e.g., C2)
;;;   - Circuit capacity: 800W max per circuit (configurable)
;;;
;;; WIRE METADATA (XData):
;;;   Each wire stores:
;;;     - FROM/TO block names
;;;     - FROM/TO block handles (for tracking)
;;;     - FROM/TO block types
;;;     - CHAIN_ID (unique batch/chain identifier)
;;;     - TIMESTAMP (when created)
;;;     - WIRE_TYPE ("BATCH" for BWIRE, "CHAIN" for WIREARC_CHAIN)
;;;     - STATUS (1=active, 0=disabled)
;;;
;;; LOAD CALCULATION:
;;;   Follows standard electrical engineering practice:
;;;   - Connected Load: Sum of all fixture wattages
;;;   - Demand Factors (D.F.): Account for diversity (not all loads run simultaneously)
;;;     * Light/Fan: 0.6 (60% usage)
;;;     * Socket: 0.2 (20% usage)
;;;     * AC: 0.7 (70% usage)
;;;     * Emergency: 0.6 (60% usage)
;;;   - Used Load = Connected Load × Demand Factor
;;;
;;; MULTI-FLOOR PROJECTS:
;;;   - Same fixture types repeat per floor
;;;   - Use floor suffix for unique naming (e.g., "PL-01-F1", "PL-01-F2")
;;;   - Reports combine all floors in single file
;;;
;;; COMMANDS AVAILABLE:
;;;   Wiring:
;;;     BWIRE          - Batch SDB→fixtures wiring (select area, auto-wire all)
;;;     WIREARC_CHAIN  - Manual SB→SDB chaining (click-through connections)
;;;   
;;;   Circuit Assignment:
;;;     POPULATE_CKT   - Auto-assign circuits by tracing wire chains
;;;   
;;;   Validation:
;;;     VALIDATE_WIRES - Check wire integrity and metadata
;;;     CHECK_CIRCUITS - Validate circuit loads (warn if >800W)
;;;   
;;;   Export:
;;;     EXPORT_WIRES     - Connection list (TXT + CSV)
;;;     EXPORT_TREE      - Hierarchy tree structure (TXT)
;;;     EXPORT_LOAD_CALC - Full load calculation matching Excel format (TXT + CSV)
;;;     EXPORT_SUMMARY   - Quick summary (fixture counts, totals)
;;;     EXPORT_ALL       - Run all exports at once
;;;
;;; ========================================================================

(vl-load-com)

;;; ========================================================================
;;; GLOBAL CONFIGURATION
;;; ========================================================================

(setq *WIRE-ARC-BULGE* 0.12)        ; Arc bulge for manual connections (12%)
(setq *BATCH-ARC-BULGE* 0.04)       ; Arc bulge for batch connections (4%)
(setq *CONNECTION-TOLERANCE* 10.0)   ; Min spacing between connection points
(setq *CIRCUIT-CAPACITY* 800)        ; Max wattage per circuit (configurable)

;; Demand Factors (Diversity Factors)
(setq *DF-LIGHT* 0.6)                ; 60% light usage
(setq *DF-SOCKET* 0.2)               ; 20% socket usage
(setq *DF-AC* 0.7)                   ; 70% AC usage
(setq *DF-EMERGENCY* 0.6)            ; 60% emergency usage

;; Parent/Child Classification
(setq *PARENT-TYPES* '("SB" "SDB" "MDB"))  ; Distribution boards (no load)
(setq *FIXTURE-TYPES* '("LIGHT" "FAN" "SOCKET" "AC" "GEYSER" "SWITCH"))

;;; ========================================================================
;;; BWIRE - BATCH SDB WIRING
;;; Auto-connects multiple fixtures to one SDB
;;; ========================================================================

(defun c:BWIRE (/ *doc* oldOsnap oldCmdecho selArea sourceBlock sourceEnt sourceName 
                  sourceHandle sourceBox ssTargetBlocks destBlkName destBlock destEnt 
                  destInsertPt sourceEdgePt usedPoints tolerance midPoint wireEnt 
                  chainID timestamp i e total layerName)
  
  (vl-load-com)
  (setq *doc* (vla-get-activedocument (vlax-get-acad-object)))
  
  ;; Save settings
  (setq oldOsnap (getvar "OSMODE"))
  (setq oldCmdecho (getvar "CMDECHO"))
  (setq layerName (getvar "CLAYER"))  ; Use current layer
  (setvar "OSMODE" 0)
  (setvar "CMDECHO" 0)
  
  ;; Start undo group
  (vla-startundomark *doc*)
  
  ;; Generate unique chain ID
  (setq chainID (strcat "BWIRE_" (rtos (getvar "CDATE") 2 0)))
  (setq timestamp (rtos (getvar "CDATE") 2 0))
  
  (princ "\n========================================")
  (princ "\n      BATCH SDB WIRING (BWire)")
  (princ "\n========================================")
  
  ;; 1. Get selection area for fixtures
  (princ "\nSelect the area containing fixtures to connect: ")
  (setq selArea (ssget))
  (if (not selArea)
    (progn 
      (princ "\nNo area selected. Exiting.")
      (restore-settings oldOsnap oldCmdecho)
      (vla-endundomark *doc*)
      (exit)
    )
  )
  
  ;; 2. Get SOURCE SDB block
  (princ "\nSelect the SOURCE Sub Distribution Board (SDB): ")
  (while (not sourceBlock)
    (setq sourceBlock (ssget ":S" '((0 . "INSERT"))))
    (if (not sourceBlock)
      (princ "\n*** Please select a BLOCK for the SDB: ")
    )
  )
  
  (setq sourceEnt (ssname sourceBlock 0))
  (setq sourceName (get-block-name sourceEnt))
  (setq sourceHandle (vla-get-handle (vlax-ename->vla-object sourceEnt)))
  (setq sourceBox (getblockboundingbox sourceEnt))
  
  (if (not sourceName)
    (setq sourceName (vla-get-effectivename (vlax-ename->vla-object sourceEnt)))
  )
  
  (princ (strcat "\nSource SDB: " sourceName))
  
  ;; 3. Select destination fixture types
  (princ "\nSelect DESTINATION FIXTURE blocks (one of each type).")
  (princ "\nPress Enter when done selecting types: ")
  (setq ssTargetBlocks (ssadd))
  
  (while (setq destBlock (ssget ":S" '((0 . "INSERT"))))
    (setq destEnt (ssname destBlock 0))
    (setq destBlkName (cdr (assoc 2 (entget destEnt))))
    
    ;; Find all blocks of this type in selection area
    (setq i 0)
    (repeat (sslength selArea)
      (setq e (ssname selArea i))
      (if (and (= (cdr (assoc 0 (entget e))) "INSERT")
               (= (cdr (assoc 2 (entget e))) destBlkName))
        (ssadd e ssTargetBlocks)
      )
      (setq i (1+ i))
    )
    
    (princ (strcat "\nAdded " (itoa (sslength ssTargetBlocks)) 
                   " blocks. Select another type or Enter to finish: "))
  )
  
  (if (= (sslength ssTargetBlocks) 0)
    (progn 
      (princ "\nNo destination blocks found. Exiting.")
      (restore-settings oldOsnap oldCmdecho)
      (vla-endundomark *doc*)
      (exit)
    )
  )
  
  (setq total (sslength ssTargetBlocks))
  (princ (strcat "\n\nConnecting " (itoa total) " fixtures to " sourceName "..."))
  
  ;; 4. Initialize connection point tracking
  (setq usedPoints '())
  (setq tolerance *CONNECTION-TOLERANCE*)
  
  ;; 5. Process each target fixture
  (setq i 0)
  (repeat total
    (setq e (ssname ssTargetBlocks i))
    (setq destInsertPt (cdr (assoc 10 (entget e))))
    
    ;; Find available connection point on SDB perimeter
    (setq sourceEdgePt (findsourcepointonperimeter sourceBox destInsertPt usedPoints tolerance))
    (setq usedPoints (cons sourceEdgePt usedPoints))
    
    ;; Calculate subtle arc midpoint
    (setq midPoint (calculatearcmidpoint sourceEdgePt destInsertPt *BATCH-ARC-BULGE*))
    
    ;; Draw arc wire on current layer
    (command "_.ARC" sourceEdgePt midPoint destInsertPt)
    (setq wireEnt (entlast))
    
    ;; Ensure wire stays on current layer
    (entmod (subst (cons 8 layerName) (assoc 8 (entget wireEnt)) (entget wireEnt)))
    
    ;; Attach metadata
    (attach-wire-metadata wireEnt sourceEnt e chainID timestamp "BATCH")
    
    ;; Mark emergency status
    (mark-emergency-status e)
    
    ;; Progress indicator
    (princ (strcat "\r  Wired " (itoa (1+ i)) "/" (itoa total) " fixtures..."))
    
    (setq i (1+ i))
  )
  
  (princ (strcat "\n\n✓ Batch wiring complete! " (itoa total) " fixtures connected."))
  (princ "\n========================================")
  
  ;; Restore settings
  (restore-settings oldOsnap oldCmdecho)
  (vla-endundomark *doc*)
  (princ)
)

;;; ========================================================================
;;; WIREARC_CHAIN - MANUAL SB CHAINING
;;; Click-through connections for Switch Boards
;;; ========================================================================

(defun c:WIREARC_CHAIN (/ *doc* oldOsnap oldCmdecho oldOrtho startBlock startPt startName 
                          startHandle startType sel nextBlock nextPt nextName 
                          nextHandle nextType midPt wireEnt dx dy dist height 
                          perpVec perpLen chainID timestamp layerName)
  
  (vl-load-com)
  (setq *doc* (vla-get-activedocument (vlax-get-acad-object)))
  
  ;; Save settings
  (setq oldOsnap (getvar "OSMODE"))
  (setq oldCmdecho (getvar "CMDECHO"))
  (setq oldOrtho (getvar "ORTHOMODE"))
  (setq layerName (getvar "CLAYER"))  ; Use current layer
  (setvar "OSMODE" 0)
  (setvar "CMDECHO" 0)
  (setvar "ORTHOMODE" 0)
  
  ;; Start undo group
  (vla-startundomark *doc*)
  
  ;; Generate unique chain ID
  (setq chainID (strcat "CHAIN_" (rtos (getvar "CDATE") 2 0)))
  (setq timestamp (rtos (getvar "CDATE") 2 0))
  
  (princ "\n========================================")
  (princ "\n    MANUAL SB CHAINING (WIREARC)")
  (princ "\n========================================")
  
  ;; Pick first block (SB or SDB)
  (setq startBlock (car (entsel "\nSelect first block (SB/SDB) to start: ")))
  
  (if (not startBlock)
    (progn
      (princ "\nNo block selected. Exiting.")
      (restore-settings-ortho oldOsnap oldCmdecho oldOrtho)
      (vla-endundomark *doc*)
      (exit)
    )
  )
  
  ;; Validate it's a block
  (if (/= (cdr (assoc 0 (entget startBlock))) "INSERT")
    (progn
      (princ "\nSelected object is not a block. Exiting.")
      (restore-settings-ortho oldOsnap oldCmdecho oldOrtho)
      (vla-endundomark *doc*)
      (exit)
    )
  )
  
  ;; Get start block data
  (setq startPt (cdr (assoc 10 (entget startBlock))))
  (setq startHandle (vla-get-handle (vlax-ename->vla-object startBlock)))
  (setq startType (vla-get-effectivename (vlax-ename->vla-object startBlock)))
  (setq startName (get-block-name startBlock))
  
  (if (not startName)
    (setq startName startType)
  )
  
  (princ (strcat "\nStarting chain: " chainID))
  (princ (strcat "\nSource: " startName))
  (princ "\n\nContinuous chaining mode active.")
  (princ "\nPick next block to connect, or press ENTER/ESC to finish.")
  
  ;; Main chaining loop
  (while T
    (setq sel (entsel "\nSelect next block (ENTER to finish): "))
    
    (cond
      ;; User pressed ENTER/ESC - exit gracefully
      ((null sel)
       (princ "\n\n✓ Chaining completed.")
       (princ "\n========================================")
       (restore-settings-ortho oldOsnap oldCmdecho oldOrtho)
       (vla-endundomark *doc*)
       (exit)
      )
      
      ;; User selected something
      (sel
       (progn
         (setq nextBlock (car sel))
         
         (cond
           ;; Validate it's a block
           ((not (= (cdr (assoc 0 (entget nextBlock))) "INSERT"))
            (princ "\n✗ Not a block. Select a valid block."))
           
           ;; Check not same block
           ((= nextBlock startBlock)
            (princ "\n✗ Cannot connect to same block."))
           
           ;; Valid selection - create wire
           (T
            (progn
              ;; Get next block data
              (setq nextPt (cdr (assoc 10 (entget nextBlock))))
              (setq nextHandle (vla-get-handle (vlax-ename->vla-object nextBlock)))
              (setq nextType (vla-get-effectivename (vlax-ename->vla-object nextBlock)))
              (setq nextName (get-block-name nextBlock))
              
              (if (not nextName)
                (setq nextName nextType)
              )
              
              ;; Calculate dramatic arc geometry
              (setq dx (- (car nextPt) (car startPt)))
              (setq dy (- (cadr nextPt) (cadr startPt)))
              (setq dist (sqrt (+ (* dx dx) (* dy dy))))
              
              ;; Arc height proportional to distance
              (setq height (* dist *WIRE-ARC-BULGE*))
              (if (> height 15.0) (setq height 15.0))
              
              ;; Perpendicular vector for arc bend
              (setq perpVec (list (- dy) dx))
              (setq perpLen (sqrt (+ (* (car perpVec) (car perpVec)) 
                                     (* (cadr perpVec) (cadr perpVec)))))
              (setq perpVec (list (* (/ (car perpVec) perpLen) height)
                                  (* (/ (cadr perpVec) perpLen) height)
                                  0))
              
              ;; Calculate arc midpoint
              (setq midPt (list (+ (/ (+ (car startPt) (car nextPt)) 2) (car perpVec))
                                (+ (/ (+ (cadr startPt) (cadr nextPt)) 2) (cadr perpVec))
                                0))
              
              ;; Draw arc wire on current layer
              (command "_.ARC" startPt midPt nextPt)
              (setq wireEnt (entlast))
              
              ;; Ensure wire stays on current layer
              (entmod (subst (cons 8 layerName) (assoc 8 (entget wireEnt)) (entget wireEnt)))
              
              ;; Attach metadata
              (attach-wire-metadata wireEnt startBlock nextBlock chainID timestamp "CHAIN")
              
              (princ (strcat "\n✓ Connected: " startName " → " nextName 
                            " [" (rtos dist 2 1) " units]"))
              
              ;; Update start point for next connection
              (setq startBlock nextBlock)
              (setq startPt nextPt)
              (setq startName nextName)
              (setq startHandle nextHandle)
              (setq startType nextType)
            )
           )
         )
       ))
      
      ;; Clicked empty space
      (T
       (princ "\n✗ Missed. Try again."))
    )
  )
)

;;; ========================================================================
;;; POPULATE_CKT - AUTO-ASSIGN CIRCUITS
;;; Traces wire chains to assign circuit numbers to SB blocks
;;; ========================================================================

(defun c:POPULATE_CKT (/ ss sdbList sdb sdbName sdbHandle circuitNum visited 
                         allSBs sb connections normalCount emergCount)
  
  (princ "\n========================================")
  (princ "\n    AUTO-ASSIGN CIRCUITS (POPULATE_CKT)")
  (princ "\n========================================")
  
  ;; Find all SDB blocks
  (setq sdbList (get-blocks-by-type "SDB"))
  
  (if (not sdbList)
    (progn
      (princ "\n✗ No SDB blocks found in drawing.")
      (princ)
      (exit)
    )
  )
  
  (princ (strcat "\nFound " (itoa (length sdbList)) " SDB block(s)"))
  
  ;; Process each SDB
  (foreach sdb sdbList
    (setq sdbName (get-block-name sdb))
    (setq sdbHandle (vla-get-handle (vlax-ename->vla-object sdb)))
    
    (princ (strcat "\n\nProcessing " sdbName "..."))
    
    ;; Find all SB blocks connected to this SDB
    (setq allSBs (find-connected-sbs sdb))
    
    (if (not allSBs)
      (princ (strcat "\n  No SB blocks connected to " sdbName))
      (progn
        (princ (strcat "\n  Found " (itoa (length allSBs)) " SB block(s)"))
        
        ;; Build connection graph for chain detection
        (setq connections (build-sb-connection-graph allSBs))
        
        ;; Assign circuits by tracing chains
        (setq circuitNum 1)
        (setq visited '())
        (setq normalCount 0)
        (setq emergCount 0)
        
        (foreach sb allSBs
          (if (not (member sb visited))
            (progn
              ;; Trace chain from this SB
              (setq chain (trace-sb-chain sb connections visited))
              (setq visited (append visited chain))
              
              ;; Calculate total load for this circuit
              (setq totalLoad (calculate-chain-load chain))
              
              ;; Determine if emergency or normal circuit
              (setq isEmergency (has-emergency-fixtures chain))
              
              ;; Assign circuit number
              (if isEmergency
                (progn
                  (setq emergCount (1+ emergCount))
                  (setq cktName (strcat "EC" (itoa emergCount)))
                )
                (progn
                  (setq normalCount (1+ normalCount))
                  (setq cktName (strcat "C" (itoa normalCount)))
                )
              )
              
              ;; Populate CKT attribute for all SBs in chain
              (foreach chainSB chain
                (set-block-attribute chainSB "CKT" cktName)
              )
              
              ;; Report
              (princ (strcat "\n  ✓ " cktName ": " (itoa (length chain)) " SB(s), "
                            (rtos totalLoad 2 0) "W"))
              
              ;; Warn if overloaded
              (if (> totalLoad *CIRCUIT-CAPACITY*)
                (princ (strcat " ⚠ OVERLOAD! (>" (itoa *CIRCUIT-CAPACITY*) "W)"))
              )
            )
          )
        )
        
        (princ (strcat "\n  Summary: " (itoa normalCount) " normal circuits, "
                      (itoa emergCount) " emergency circuits"))
      )
    )
  )
  
  (princ "\n========================================")
  (princ "\n✓ Circuit assignment complete!")
  (princ "\n========================================")
  (princ)
)

;;; ========================================================================
;;; VALIDATION COMMANDS
;;; ========================================================================

(defun c:VALIDATE_WIRES (/ ss count validCount invalidCount batchCount chainCount 
                           ent xdata fromName toName wireType)
  (princ "\n========================================")
  (princ "\n      WIRE VALIDATION REPORT")
  (princ "\n========================================")
  
  (setq ss (ssget "_X" '((0 . "ARC")))
        count 0
        validCount 0
        invalidCount 0
        batchCount 0
        chainCount 0)
  
  (if ss
    (progn
      (repeat (sslength ss)
        (setq ent (ssname ss count))
        (setq xdata (assoc -3 (entget ent '("WIREAPP" "WIRE_METADATA"))))
        
        (if xdata
          (progn
            (setq validCount (1+ validCount))
            (setq fromName "Unknown")
            (setq toName "Unknown")
            (setq wireType "UNKNOWN")
            
            ;; Extract data
            (foreach item (cdr xdata)
              (cond
                ((= (car item) "WIREAPP")
                 (foreach subitem (cdr item)
                   (if (= (car subitem) 1000)
                     (cond
                       ((wcmatch (cdr subitem) "FROM=*")
                        (setq fromName (substr (cdr subitem) 6)))
                       ((wcmatch (cdr subitem) "TO=*")
                        (setq toName (substr (cdr subitem) 4)))
                     )
                   )
                 ))
                ((= (car item) "WIRE_METADATA")
                 (foreach subitem (cdr item)
                   (if (= (car subitem) 1000)
                     (if (wcmatch (cdr subitem) "WIRE_TYPE=*")
                       (setq wireType (substr (cdr subitem) 11))
                     )
                   )
                 ))
              )
            )
            
            ;; Count wire types
            (cond
              ((= wireType "BATCH") (setq batchCount (1+ batchCount)))
              ((= wireType "CHAIN") (setq chainCount (1+ chainCount)))
            )
            
            (princ (strcat "\n✓ [" wireType "] " fromName " → " toName))
          )
          (progn
            (setq invalidCount (1+ invalidCount))
            (princ (strcat "\n✗ Wire " (itoa (1+ count)) ": Missing metadata"))
          )
        )
        (setq count (1+ count))
      )
      
      (princ "\n========================================")
      (princ (strcat "\nTotal wires: " (itoa count)))
      (princ (strcat "\n  Valid: " (itoa validCount)))
      (princ (strcat "\n  Invalid: " (itoa invalidCount)))
      (princ (strcat "\n\nBy type:"))
      (princ (strcat "\n  Batch (SDB): " (itoa batchCount)))
      (princ (strcat "\n  Chain (SB): " (itoa chainCount)))
      (princ "\n========================================")
    )
    (princ "\n✗ No wires found in drawing.")
  )
  (princ)
)

(defun c:CHECK_CIRCUITS (/ sdbList sdb sdbName allSBs circuits ckt cktSBs totalLoad overloadCount)
  (princ "\n========================================")
  (princ "\n      CIRCUIT LOAD VALIDATION")
  (princ "\n========================================")
  
  (setq sdbList (get-blocks-by-type "SDB"))
  (setq overloadCount 0)
  
  (if (not sdbList)
    (princ "\n✗ No SDB blocks found.")
    (progn
      (foreach sdb sdbList
        (setq sdbName (get-block-name sdb))
        (princ (strcat "\n\n" sdbName ":"))
        
        (setq allSBs (find-connected-sbs sdb))
        
        (if allSBs
          (progn
            ;; Group SBs by circuit
            (setq circuits (group-by-circuit allSBs))
            
            (foreach ckt circuits
              (setq cktName (car ckt))
              (setq cktSBs (cdr ckt))
              (setq totalLoad (calculate-chain-load cktSBs))
              
              (princ (strcat "\n  " cktName ": " (rtos totalLoad 2 0) "W"))
              
              (if (> totalLoad *CIRCUIT-CAPACITY*)
                (progn
                  (princ (strcat " ⚠ OVERLOAD! (Max: " (itoa *CIRCUIT-CAPACITY*) "W)"))
                  (setq overloadCount (1+ overloadCount))
                )
                (princ " ✓")
              )
            )
          )
          (princ "\n  No circuits assigned")
        )
      )
      
      (princ "\n========================================")
      (if (> overloadCount 0)
        (princ (strcat "\n⚠ WARNING: " (itoa overloadCount) " overloaded circuit(s) found!"))
        (princ "\n✓ All circuits within capacity limits")
      )
      (princ "\n========================================")
    )
  )
  (princ)
)

;;; ========================================================================
;;; EXPORT COMMANDS
;;; ========================================================================

(defun c:EXPORT_WIRES (/ filePath ss count ent xdata fromName toName fromHandle 
                         toHandle chainID timestamp wireType file csvFile)
  (setq filePath (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") "-WIRES.txt"))
  (setq csvFile (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") "-WIRES.csv"))
  
  ;; TXT Export
  (setq file (open filePath "w"))
  
  (if file
    (progn
      (write-line "========================================" file)
      (write-line "       WIRE CONNECTION REPORT" file)
      (write-line "========================================" file)
      (write-line (strcat "Drawing: " (getvar "DWGNAME")) file)
      (write-line (strcat "Date: " (menucmd "M=$(edtime,$(getvar,date),DD-MON-YYYY)")) file)
      (write-line "========================================" file)
      (write-line "" file)
      
      (setq ss (ssget "_X" '((0 . "ARC")))
            count 0)
      
      (if ss
        (progn
          (repeat (sslength ss)
            (setq ent (ssname ss count))
            (setq xdata (assoc -3 (entget ent '("WIREAPP" "WIRE_METADATA"))))
            
            (if xdata
              (progn
                (setq fromName "Unknown"
                      toName "Unknown"
                      fromHandle "N/A"
                      toHandle "N/A"
                      chainID "N/A"
                      timestamp "N/A"
                      wireType "N/A")
                
                (foreach item (cdr xdata)
                  (cond
                    ((= (car item) "WIREAPP")
                     (foreach subitem (cdr item)
                       (if (= (car subitem) 1000)
                         (cond
                           ((wcmatch (cdr subitem) "FROM=*")
                            (setq fromName (substr (cdr subitem) 6)))
                           ((wcmatch (cdr subitem) "TO=*")
                            (setq toName (substr (cdr subitem) 4)))
                         )
                       )
                     ))
                    ((= (car item) "WIRE_METADATA")
                     (foreach subitem (cdr item)
                       (if (= (car subitem) 1000)
                         (cond
                           ((wcmatch (cdr subitem) "FROM_HANDLE=*")
                            (setq fromHandle (substr (cdr subitem) 13)))
                           ((wcmatch (cdr subitem) "TO_HANDLE=*")
                            (setq toHandle (substr (cdr subitem) 11)))
                           ((wcmatch (cdr subitem) "CHAIN_ID=*")
                            (setq chainID (substr (cdr subitem) 10)))
                           ((wcmatch (cdr subitem) "TIMESTAMP=*")
                            (setq timestamp (substr (cdr subitem) 11)))
                           ((wcmatch (cdr subitem) "WIRE_TYPE=*")
                            (setq wireType (substr (cdr subitem) 11)))
                         )
                       )
                     ))
                  )
                )
                
                (write-line (strcat "Connection " (itoa (1+ count)) ": [" wireType "]") file)
                (write-line (strcat "  From: " fromName " [" fromHandle "]") file)
                (write-line (strcat "  To: " toName " [" toHandle "]") file)
                (write-line (strcat "  Chain ID: " chainID) file)
                (write-line "" file)
              )
            )
            (setq count (1+ count))
          )
          
          (write-line "========================================" file)
          (write-line (strcat "Total connections: " (itoa count)) file)
          (write-line "========================================" file)
        )
        (write-line "No wires found in drawing." file)
      )
      
      (close file)
      (princ (strcat "\n✓ Wire data exported to: " filePath))
    )
    (princ "\n✗ Error: Could not create TXT file.")
  )
  
  ;; CSV Export
  (setq file (open csvFile "w"))
  
  (if file
    (progn
      (write-line "Type,From,To,From Handle,To Handle,Chain ID,Timestamp" file)
      
      (setq ss (ssget "_X" '((0 . "ARC")))
            count 0)
      
      (if ss
        (repeat (sslength ss)
          (setq ent (ssname ss count))
          (setq xdata (assoc -3 (entget ent '("WIREAPP" "WIRE_METADATA"))))
          
          (if xdata
            (progn
              (setq fromName "Unknown"
                    toName "Unknown"
                    fromHandle "N/A"
                    toHandle "N/A"
                    chainID "N/A"
                    timestamp "N/A"
                    wireType "N/A")
              
              (foreach item (cdr xdata)
                (cond
                  ((= (car item) "WIREAPP")
                   (foreach subitem (cdr item)
                     (if (= (car subitem) 1000)
                       (cond
                         ((wcmatch (cdr subitem) "FROM=*")
                          (setq fromName (substr (cdr subitem) 6)))
                         ((wcmatch (cdr subitem) "TO=*")
                          (setq toName (substr (cdr subitem) 4)))
                       )
                     )
                   ))
                  ((= (car item) "WIRE_METADATA")
                   (foreach subitem (cdr item)
                     (if (= (car subitem) 1000)
                       (cond
                         ((wcmatch (cdr subitem) "FROM_HANDLE=*")
                          (setq fromHandle (substr (cdr subitem) 13)))
                         ((wcmatch (cdr subitem) "TO_HANDLE=*")
                          (setq toHandle (substr (cdr subitem) 11)))
                         ((wcmatch (cdr subitem) "CHAIN_ID=*")
                          (setq chainID (substr (cdr subitem) 10)))
                         ((wcmatch (cdr subitem) "TIMESTAMP=*")
                          (setq timestamp (substr (cdr subitem) 11)))
                         ((wcmatch (cdr subitem) "WIRE_TYPE=*")
                          (setq wireType (substr (cdr subitem) 11)))
                       )
                     )
                   ))
                )
              )
              
              (write-line (strcat wireType "," fromName "," toName "," 
                                 fromHandle "," toHandle "," chainID "," timestamp) file)
            )
          )
          (setq count (1+ count))
        )
      )
      
      (close file)
      (princ (strcat "\n✓ CSV data exported to: " csvFile))
    )
    (princ "\n✗ Error: Could not create CSV file.")
  )
  
  (princ)
)

(defun c:EXPORT_TREE (/ ss count ent xdata fromName toName connections 
                        graph chains chain sb-parent tree-data orphaned-chains
                        file filePath)
  (setq filePath (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") "-TREE.txt"))
  (setq file (open filePath "w"))
  
  (if file
    (progn
      (write-line "========================================" file)
      (write-line "       WIRE TREE STRUCTURE" file)
      (write-line "========================================" file)
      (write-line (strcat "Drawing: " (getvar "DWGNAME")) file)
      (write-line (strcat "Date: " (menucmd "M=$(edtime,$(getvar,date),DD-MON-YYYY)")) file)
      (write-line "========================================" file)
      (write-line "" file)
      
      (setq ss (ssget "_X" '((0 . "ARC")))
            count 0
            connections '())
      
      (if ss
        (progn
          (repeat (sslength ss)
            (setq ent (ssname ss count))
            (setq xdata (assoc -3 (entget ent '("WIREAPP"))))
            
            (if xdata
              (progn
                (setq fromName nil
                      toName nil)
                
                (foreach item (cdr xdata)
                  (if (= (car item) "WIREAPP")
                    (foreach subitem (cdr item)
                      (if (= (car subitem) 1000)
                        (cond
                          ((wcmatch (cdr subitem) "FROM=*")
                           (setq fromName (substr (cdr subitem) 6)))
                          ((wcmatch (cdr subitem) "TO=*")
                           (setq toName (substr (cdr subitem) 4)))
                        )
                      )
                    )
                  )
                )
                
                (if (and fromName toName)
                  (progn
                    (setq connections (cons (list fromName toName) connections))
                    (setq connections (cons (list toName fromName) connections))
                  )
                )
              )
            )
            (setq count (1+ count))
          )
          
          (write-line (strcat "Found " (itoa (/ (length connections) 2)) " wire connections.") file)
          (write-line "" file)
          
          (setq graph (build-graph connections))
          (setq chains (find-all-chains graph))
          
          (write-line (strcat "Found " (itoa (length chains)) " separate chain(s).") file)
          (write-line "" file)
          
          (setq tree-data '()
                orphaned-chains '())
          
          (foreach chain chains
            (setq sb-parent (find-sb-in-chain chain))
            
            (if sb-parent
              (setq tree-data (cons (list sb-parent chain) tree-data))
              (setq orphaned-chains (cons chain orphaned-chains))
            )
          )
          
          (write-line "TREE STRUCTURE" file)
          (write-line "========================================" file)
          (write-line "" file)
          
          (foreach tree-item tree-data
            (setq sb-parent (car tree-item))
            (setq chain (cadr tree-item))
            
            (write-line (strcat sb-parent " (Parent - SB)") file)
            
            (foreach block chain
              (if (/= block sb-parent)
                (write-line (strcat "  ├─ " block) file)
              )
            )
            (write-line "" file)
          )
          
          (if orphaned-chains
            (progn
              (write-line "========================================" file)
              (write-line "⚠ WARNING: ORPHANED CHAINS FOUND" file)
              (write-line "========================================" file)
              (foreach orphan orphaned-chains
                (write-line "" file)
                (write-line "Chain with NO SB parent:" file)
                (foreach block orphan
                  (write-line (strcat "  • " block) file)
                )
              )
              (write-line "" file)
            )
          )
          
          (write-line "========================================" file)
          (write-line (strcat "Total SB Parents: " (itoa (length tree-data))) file)
          (write-line (strcat "Orphaned Chains: " (itoa (length orphaned-chains))) file)
          (write-line (strcat "Total Connections: " (itoa (/ (length connections) 2))) file)
          (write-line "========================================" file)
        )
        (write-line "No wires found in drawing." file)
      )
      
      (close file)
      (princ (strcat "\n✓ Tree structure exported to: " filePath))
    )
    (princ "\n✗ Error: Could not create export file.")
  )
  (princ)
)

(defun c:EXPORT_LOAD_CALC (/ sdbList allFixtures fixtureData sdbData floorData 
                             file txtFile csvFile)
  (princ "\n========================================")
  (princ "\n    GENERATING LOAD CALCULATION")
  (princ "\n========================================")
  
  (setq txtFile (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") "-LOAD-CALC.txt"))
  (setq csvFile (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") "-LOAD-CALC.csv"))
  
  ;; Collect all fixture data
  (setq allFixtures (collect-all-fixtures))
  
  (if (not allFixtures)
    (progn
      (princ "\n✗ No fixtures found in drawing.")
      (princ)
      (exit)
    )
  )
  
  (princ (strcat "\nCollected " (itoa (length allFixtures)) " fixture(s)"))
  
  ;; Group by SDB
  (setq sdbData (group-fixtures-by-sdb allFixtures))
  
  ;; Generate TXT report
  (generate-load-calc-txt txtFile sdbData allFixtures)
  
  ;; Generate CSV report
  (generate-load-calc-csv csvFile sdbData allFixtures)
  
  (princ "\n========================================")
  (princ (strcat "\n✓ TXT report: " txtFile))
  (princ (strcat "\n✓ CSV report: " csvFile))
  (princ "\n========================================")
  (princ)
)

(defun c:EXPORT_SUMMARY (/ file filePath fixtureCount wireCount totalLoad)
  (setq filePath (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") "-SUMMARY.txt"))
  (setq file (open filePath "w"))
  
  (if file
    (progn
      (write-line "========================================" file)
      (write-line "       PROJECT SUMMARY" file)
      (write-line "========================================" file)
      (write-line (strcat "Drawing: " (getvar "DWGNAME")) file)
      (write-line (strcat "Date: " (menucmd "M=$(edtime,$(getvar,date),DD-MON-YYYY)")) file)
      (write-line "========================================" file)
      (write-line "" file)
      
      ;; Fixture counts
      (setq fixtureCount (count-fixtures-by-type))
      
      (write-line "FIXTURE SUMMARY" file)
      (write-line "----------------" file)
      (foreach item fixtureCount
        (write-line (strcat (car item) ": " (itoa (cdr item)) " nos") file)
      )
      (write-line "" file)
      
      ;; Wire summary
      (setq wireCount (count-wires))
      (write-line "WIRING SUMMARY" file)
      (write-line "---------------" file)
      (write-line (strcat "Total wires: " (itoa wireCount)) file)
      (write-line "" file)
      
      ;; Load summary
      (setq totalLoad (calculate-total-load))
      (write-line "LOAD SUMMARY" file)
      (write-line "-------------" file)
      (write-line (strcat "Total Connected Load: " (rtos totalLoad 2 0) "W") file)
      (write-line (strcat "Demand Factor Applied: " (rtos (* totalLoad *DF-LIGHT*) 2 0) "W (estimate)") file)
      (write-line "" file)
      
      (write-line "========================================" file)
      
      (close file)
      (princ (strcat "\n✓ Summary exported to: " filePath))
    )
    (princ "\n✗ Error: Could not create summary file.")
  )
  (princ)
)

(defun c:EXPORT_ALL (/ )
  (princ "\n========================================")
  (princ "\n    EXPORTING ALL REPORTS")
  (princ "\n========================================")
  
  (princ "\n\n1. Exporting wires...")
  (c:EXPORT_WIRES)
  
  (princ "\n\n2. Exporting tree...")
  (c:EXPORT_TREE)
  
  (princ "\n\n3. Exporting load calculation...")
  (c:EXPORT_LOAD_CALC)
  
  (princ "\n\n4. Exporting summary...")
  (c:EXPORT_SUMMARY)
  
  (princ "\n\n========================================")
  (princ "\n✓ All exports complete!")
  (princ "\n========================================")
  (princ)
)

;;; ========================================================================
;;; HELPER FUNCTIONS - Attribute Management
;;; ========================================================================

(defun get-block-name (blkEnt / blkObj attList name)
  (setq blkObj (vlax-ename->vla-object blkEnt))
  (setq name nil)
  
  (if (= (vla-get-hasattributes blkObj) :vlax-true)
    (progn
      (setq attList (vlax-invoke blkObj 'GetAttributes))
      (foreach att attList
        (if (= (strcase (vla-get-TagString att)) "NAME")
          (setq name (vla-get-TextString att))
        )
      )
    )
  )
  name
)

(defun get-block-attribute (blkEnt attName / blkObj attList value)
  (setq blkObj (vlax-ename->vla-object blkEnt))
  (setq value nil)
  
  (if (= (vla-get-hasattributes blkObj) :vlax-true)
    (progn
      (setq attList (vlax-invoke blkObj 'GetAttributes))
      (foreach att attList
        (if (= (strcase (vla-get-TagString att)) (strcase attName))
          (setq value (vla-get-TextString att))
        )
      )
    )
  )
  value
)

(defun set-block-attribute (blkEnt attName attValue / blkObj attList)
  (setq blkObj (vlax-ename->vla-object blkEnt))
  
  (if (= (vla-get-hasattributes blkObj) :vlax-true)
    (progn
      (setq attList (vlax-invoke blkObj 'GetAttributes))
      (foreach att attList
        (if (= (strcase (vla-get-TagString att)) (strcase attName))
          (vla-put-TextString att attValue)
        )
      )
    )
  )
)

(defun mark-emergency-status (blkEnt / name)
  (setq name (get-block-name blkEnt))
  
  (if name
    (if (wcmatch (strcase name) "E*")
      (set-block-attribute blkEnt "SPARE1" "EMERGENCY")
      (set-block-attribute blkEnt "SPARE1" "NORMAL")
    )
  )
)

;;; ========================================================================
;;; HELPER FUNCTIONS - Wire Metadata
;;; ========================================================================

(defun attach-wire-metadata (wireEnt sourceEnt destEnt chainID timestamp wireType / 
                              sourceName destName sourceHandle destHandle sourceType destType)
  
  (setq sourceName (get-block-name sourceEnt))
  (setq destName (get-block-name destEnt))
  (setq sourceHandle (vla-get-handle (vlax-ename->vla-object sourceEnt)))
  (setq destHandle (vla-get-handle (vlax-ename->vla-object destEnt)))
  (setq sourceType (vla-get-effectivename (vlax-ename->vla-object sourceEnt)))
  (setq destType (vla-get-effectivename (vlax-ename->vla-object destEnt)))
  
  (if (not sourceName) (setq sourceName sourceType))
  (if (not destName) (setq destName destType))
  
  (regapp "WIREAPP")
  (regapp "WIRE_METADATA")
  
  (entmod
    (append (entget wireEnt)
      (list
        (list -3
          (list "WIREAPP"
            (cons 1000 (strcat "FROM=" sourceName))
            (cons 1000 (strcat "TO=" destName))
          )
          (list "WIRE_METADATA"
            (cons 1000 (strcat "FROM_HANDLE=" sourceHandle))
            (cons 1000 (strcat "TO_HANDLE=" destHandle))
            (cons 1000 (strcat "FROM_TYPE=" sourceType))
            (cons 1000 (strcat "TO_TYPE=" destType))
            (cons 1000 (strcat "CHAIN_ID=" chainID))
            (cons 1000 (strcat "TIMESTAMP=" timestamp))
            (cons 1000 (strcat "WIRE_TYPE=" wireType))
            (cons 1070 1)
          )
        )
      )
    )
  )
)

;;; ========================================================================
;;; HELPER FUNCTIONS - System Settings
;;; ========================================================================

(defun restore-settings (osnap cmdecho / )
  (setvar "OSMODE" osnap)
  (setvar "CMDECHO" cmdecho)
)

(defun restore-settings-ortho (osnap cmdecho ortho / )
  (setvar "OSMODE" osnap)
  (setvar "CMDECHO" cmdecho)
  (setvar "ORTHOMODE" ortho)
)

;;; ========================================================================
;;; HELPER FUNCTIONS - Geometry
;;; ========================================================================

(defun getblockboundingbox (ent / ll ur)
  (vla-getboundingbox (vlax-ename->vla-object ent) 'll 'ur)
  (list (vlax-safearray->list ll) (vlax-safearray->list ur))
)

(defun calculatearcmidpoint (pt1 pt2 bulge / mid ang dist offset)
  (setq mid (mapcar '(lambda (a b) (/ (+ a b) 2.0)) pt1 pt2))
  (setq ang (angle pt1 pt2))
  (setq dist (distance pt1 pt2))
  (setq offset (* dist bulge))
  
  (mapcar '+ mid 
    (list (* offset (cos (+ ang (* pi 0.5)))) 
          (* offset (sin (+ ang (* pi 0.5)))) 
          0.0))
)

(defun findsourcepointonperimeter (sourceBox destPt usedPoints tolerance / 
                                    minpt maxpt xmin ymin xmax ymax candidatePt 
                                    bestPt minDist tooClose dist testX testY testPt)
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

;;; ========================================================================
;;; HELPER FUNCTIONS - Graph Building
;;; ========================================================================

(defun build-graph (connections / graph node neighbor)
  (setq graph '())
  
  (foreach conn connections
    (setq node (car conn)
          neighbor (cadr conn))
    
    (if (assoc node graph)
      (setq graph (subst 
                    (cons node (cons neighbor (cdr (assoc node graph))))
                    (assoc node graph)
                    graph))
      (setq graph (cons (list node neighbor) graph))
    )
  )
  graph
)

(defun find-all-chains (graph / visited chains all-nodes node chain)
  (setq visited '()
        chains '()
        all-nodes (mapcar 'car graph))
  
  (foreach node all-nodes
    (if (not (member node visited))
      (progn
        (setq chain (dfs-chain graph node visited))
        (setq visited (append visited chain))
        (setq chains (cons chain chains))
      )
    )
  )
  chains
)

(defun dfs-chain (graph node visited / chain neighbors neighbor)
  (setq chain (list node))
  (setq neighbors (cdr (assoc node graph)))
  
  (if neighbors
    (foreach neighbor neighbors
      (if (not (member neighbor visited))
        (progn
          (setq visited (cons neighbor visited))
          (setq chain (append chain (dfs-chain graph neighbor visited)))
        )
      )
    )
  )
  
  (setq chain (remove-duplicates chain))
  chain
)

(defun remove-duplicates (lst / result item)
  (setq result '())
  (foreach item lst
    (if (not (member item result))
      (setq result (cons item result))
    )
  )
  (reverse result)
)

(defun find-sb-in-chain (chain / sb-block)
  (setq sb-block nil)
  
  (foreach block chain
    (if (or (wcmatch (strcase block) "*SB*")
            (wcmatch (strcase block) "SB-*")
            (wcmatch (strcase block) "SB_*"))
      (setq sb-block block)
    )
  )
  sb-block
)

;;; ========================================================================
;;; HELPER FUNCTIONS - Block Query
;;; ========================================================================

(defun get-blocks-by-type (blockType / ss result i ent type)
  (setq ss (ssget "_X" '((0 . "INSERT")))
        result '())
  
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq type (get-block-attribute ent "TYPE"))
        
        (if (and type (= (strcase type) (strcase blockType)))
          (setq result (cons ent result))
        )
        
        (setq i (1+ i))
      )
    )
  )
  result
)

(defun find-connected-sbs (sdbEnt / sdbHandle wires result)
  (setq sdbHandle (vla-get-handle (vlax-ename->vla-object sdbEnt)))
  (setq wires (get-all-wires))
  (setq result '())
  
  (foreach wire wires
    (setq fromHandle (get-wire-data wire "FROM_HANDLE"))
    (setq toHandle (get-wire-data wire "TO_HANDLE"))
    (setq toEnt nil)
    (setq fromEnt nil)
    
    (if (= fromHandle sdbHandle)
      (progn
        (setq toEnt (handent toHandle))
        (if toEnt
          (if (= (get-block-attribute toEnt "TYPE") "SB")
            (setq result (cons toEnt result))
          )
        )
      )
    )
    
    (if (= toHandle sdbHandle)
      (progn
        (setq fromEnt (handent fromHandle))
        (if fromEnt
          (if (= (get-block-attribute fromEnt "TYPE") "SB")
            (setq result (cons fromEnt result))
          )
        )
      )
    )
  )
  
  (remove-duplicates result)
)

(defun get-all-wires (/ ss result i)
  (setq ss (ssget "_X" '((0 . "ARC")))
        result '())
  
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq result (cons (ssname ss i) result))
        (setq i (1+ i))
      )
    )
  )
  result
)

(defun get-wire-data (wireEnt dataKey / xdata value)
  (setq xdata (assoc -3 (entget wireEnt '("WIRE_METADATA"))))
  (setq value nil)
  
  (if xdata
    (foreach item (cdr xdata)
      (if (= (car item) "WIRE_METADATA")
        (foreach subitem (cdr item)
          (if (= (car subitem) 1000)
            (if (wcmatch (cdr subitem) (strcat dataKey "=*"))
              (setq value (substr (cdr subitem) (+ (strlen dataKey) 2)))
            )
          )
        )
      )
    )
  )
  value
)

;;; ========================================================================
;;; HELPER FUNCTIONS - Circuit Management
;;; ========================================================================

(defun build-sb-connection-graph (sbList / graph wires wire fromHandle toHandle fromEnt toEnt)
  (setq graph '())
  (setq wires (get-all-wires))
  
  (foreach wire wires
    (setq fromHandle (get-wire-data wire "FROM_HANDLE"))
    (setq toHandle (get-wire-data wire "TO_HANDLE"))
    
    (if (and fromHandle toHandle)
      (progn
        (setq fromEnt (handent fromHandle))
        (setq toEnt (handent toHandle))
        
        (if (and fromEnt toEnt
                 (member fromEnt sbList)
                 (member toEnt sbList))
          (progn
            ;; Add bidirectional connections
            (if (assoc fromEnt graph)
              (setq graph (subst 
                            (cons fromEnt (cons toEnt (cdr (assoc fromEnt graph))))
                            (assoc fromEnt graph)
                            graph))
              (setq graph (cons (list fromEnt toEnt) graph))
            )
            
            (if (assoc toEnt graph)
              (setq graph (subst 
                            (cons toEnt (cons fromEnt (cdr (assoc toEnt graph))))
                            (assoc toEnt graph)
                            graph))
              (setq graph (cons (list toEnt fromEnt) graph))
            )
          )
        )
      )
    )
  )
  graph
)

(defun trace-sb-chain (startSB connections visited / chain neighbors neighbor)
  (setq chain (list startSB))
  (setq neighbors (cdr (assoc startSB connections)))
  
  (if neighbors
    (foreach neighbor neighbors
      (if (and (not (member neighbor visited))
               (not (member neighbor chain)))
        (setq chain (append chain (trace-sb-chain neighbor connections (append visited chain))))
      )
    )
  )
  chain
)

(defun calculate-chain-load (sbChain / totalLoad sb fixtures fixture load)
  (setq totalLoad 0)
  
  (foreach sb sbChain
    (setq fixtures (find-connected-fixtures sb))
    
    (foreach fixture fixtures
      (setq load (get-block-attribute fixture "LOAD"))
      
      (if (and load (numberp (read load)))
        (setq totalLoad (+ totalLoad (atof load)))
      )
    )
  )
  totalLoad
)

(defun find-connected-fixtures (sbEnt / sbHandle wires result wire fromHandle toHandle toEnt)
  (setq sbHandle (vla-get-handle (vlax-ename->vla-object sbEnt)))
  (setq wires (get-all-wires))
  (setq result '())
  
  (foreach wire wires
    (setq fromHandle (get-wire-data wire "FROM_HANDLE"))
    (setq toHandle (get-wire-data wire "TO_HANDLE"))
    
    (if (= fromHandle sbHandle)
      (progn
        (setq toEnt (handent toHandle))
        (if (and toEnt (member (get-block-attribute toEnt "TYPE") *FIXTURE-TYPES*))
          (setq result (cons toEnt result))
        )
      )
    )
  )
  
  (remove-duplicates result)
)

(defun has-emergency-fixtures (sbChain / hasEmerg sb fixtures fixture name)
  (setq hasEmerg nil)
  
  (foreach sb sbChain
    (setq fixtures (find-connected-fixtures sb))
    
    (foreach fixture fixtures
      (setq name (get-block-name fixture))
      (if (and name (wcmatch (strcase name) "E*"))
        (setq hasEmerg T)
      )
    )
  )
  hasEmerg
)

(defun group-by-circuit (sbList / circuits sb cktName cktGroup)
  (setq circuits '())
  
  (foreach sb sbList
    (setq cktName (get-block-attribute sb "CKT"))
    
    (if cktName
      (progn
        (setq cktGroup (assoc cktName circuits))
        
        (if cktGroup
          (setq circuits (subst 
                           (cons cktName (cons sb (cdr cktGroup)))
                           cktGroup
                           circuits))
          (setq circuits (cons (list cktName sb) circuits))
        )
      )
    )
  )
  circuits
)

;;; ========================================================================
;;; HELPER FUNCTIONS - Export Support
;;; ========================================================================

(defun collect-all-fixtures (/ ss result i ent type)
  (setq ss (ssget "_X" '((0 . "INSERT")))
        result '())
  
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq type (get-block-attribute ent "TYPE"))
        
        (if (and type (member (strcase type) (mapcar 'strcase *FIXTURE-TYPES*)))
          (setq result (cons ent result))
        )
        
        (setq i (1+ i))
      )
    )
  )
  result
)

(defun group-fixtures-by-sdb (fixtures / sdbGroups fixture sb sdb sdbName group)
  (setq sdbGroups '())
  
  (foreach fixture fixtures
    (setq sb (find-parent-sb fixture))
    
    (if sb
      (progn
        (setq sdb (find-parent-sdb sb))
        
        (if sdb
          (progn
            (setq sdbName (get-block-name sdb))
            (setq group (assoc sdbName sdbGroups))
            
            (if group
              (setq sdbGroups (subst 
                                (cons sdbName (cons fixture (cdr group)))
                                group
                                sdbGroups))
              (setq sdbGroups (cons (list sdbName fixture) sdbGroups))
            )
          )
        )
      )
    )
  )
  sdbGroups
)

(defun find-parent-sb (fixtureEnt / fixtureHandle wires wire fromHandle toHandle fromEnt)
  (setq fixtureHandle (vla-get-handle (vlax-ename->vla-object fixtureEnt)))
  (setq wires (get-all-wires))
  
  (foreach wire wires
    (setq toHandle (get-wire-data wire "TO_HANDLE"))
    
    (if (= toHandle fixtureHandle)
      (progn
        (setq fromHandle (get-wire-data wire "FROM_HANDLE"))
        (setq fromEnt (handent fromHandle))
        
        (if (and fromEnt (= (get-block-attribute fromEnt "TYPE") "SB"))
          (return fromEnt)
        )
      )
    )
  )
  nil
)

(defun find-parent-sdb (sbEnt / sbHandle wires wire fromHandle toHandle fromEnt)
  (setq sbHandle (vla-get-handle (vlax-ename->vla-object sbEnt)))
  (setq wires (get-all-wires))
  
  (foreach wire wires
    (setq toHandle (get-wire-data wire "TO_HANDLE"))
    
    (if (= toHandle sbHandle)
      (progn
        (setq fromHandle (get-wire-data wire "FROM_HANDLE"))
        (setq fromEnt (handent fromHandle))
        
        (if (and fromEnt (= (get-block-attribute fromEnt "TYPE") "SDB"))
          (return fromEnt)
        )
      )
    )
  )
  nil
)

(defun generate-load-calc-txt (filePath sdbData allFixtures / file)
  (setq file (open filePath "w"))
  
  (if file
    (progn
      (write-line "========================================" file)
      (write-line "     ELECTRICAL LOAD CALCULATION" file)
      (write-line "========================================" file)
      (write-line (strcat "Project: " (getvar "DWGNAME")) file)
      (write-line (strcat "Date: " (menucmd "M=$(edtime,$(getvar,date),DD-MON-YYYY)")) file)
      (write-line "========================================" file)
      (write-line "" file)
      
      (foreach sdbGroup sdbData
        (generate-sdb-load-section file (car sdbGroup) (cdr sdbGroup))
      )
      
      (generate-total-load-section file allFixtures)
      
      (close file)
    )
  )
)

(defun generate-sdb-load-section (file sdbName fixtures / typeGroups connLoad usedLoad)
  (write-line (strcat "SDB: " sdbName) file)
  (write-line "----------------------------------------" file)
  
  (setq typeGroups (group-fixtures-by-type fixtures))
  
  (foreach typeGroup typeGroups
    (setq connLoad (calculate-type-connected-load (cdr typeGroup)))
    (setq usedLoad (calculate-type-used-load (car typeGroup) connLoad))
    
    (write-line (strcat "  " (car typeGroup) ": " 
                       (itoa (length (cdr typeGroup))) " nos, "
                       (rtos connLoad 2 0) "W connected, "
                       (rtos usedLoad 2 0) "W used") file)
  )
  
  (write-line "" file)
)

(defun generate-total-load-section (file fixtures / totalConn totalUsed)
  (write-line "========================================" file)
  (write-line "TOTAL PROJECT LOAD" file)
  (write-line "========================================" file)
  
  (setq totalConn (calculate-total-connected-load fixtures))
  (setq totalUsed (calculate-total-used-load fixtures))
  
  (write-line (strcat "Total Connected Load: " (rtos totalConn 2 0) "W") file)
  (write-line (strcat "Total Used Load: " (rtos totalUsed 2 0) "W") file)
  (write-line "========================================" file)
)

(defun generate-load-calc-csv (filePath sdbData allFixtures / file)
  (setq file (open filePath "w"))
  
  (if file
    (progn
      (write-line "SDB,Type,Quantity,Connected Load (W),Used Load (W)" file)
      
      (foreach sdbGroup sdbData
        (setq sdbName (car sdbGroup))
        (setq fixtures (cdr sdbGroup))
        (setq typeGroups (group-fixtures-by-type fixtures))
        
        (foreach typeGroup typeGroups
          (setq typeName (car typeGroup))
          (setq typeFixtures (cdr typeGroup))
          (setq connLoad (calculate-type-connected-load typeFixtures))
          (setq usedLoad (calculate-type-used-load typeName connLoad))
          
          (write-line (strcat sdbName "," typeName "," 
                             (itoa (length typeFixtures)) ","
                             (rtos connLoad 2 0) ","
                             (rtos usedLoad 2 0)) file)
        )
      )
      
      (close file)
    )
  )
)

(defun group-fixtures-by-type (fixtures / groups fixture type group)
  (setq groups '())
  
  (foreach fixture fixtures
    (setq type (get-block-attribute fixture "TYPE"))
    
    (if type
      (progn
        (setq group (assoc type groups))
        
        (if group
          (setq groups (subst 
                         (cons type (cons fixture (cdr group)))
                         group
                         groups))
          (setq groups (cons (list type fixture) groups))
        )
      )
    )
  )
  groups
)

(defun calculate-type-connected-load (fixtures / total fixture load)
  (setq total 0)
  
  (foreach fixture fixtures
    (setq load (get-block-attribute fixture "LOAD"))
    
    (if (and load (numberp (read load)))
      (setq total (+ total (atof load)))
    )
  )
  total
)

(defun calculate-type-used-load (typeName connLoad / df)
  (setq df 1.0)
  
  (cond
    ((= (strcase typeName) "LIGHT") (setq df *DF-LIGHT*))
    ((= (strcase typeName) "FAN") (setq df *DF-LIGHT*))
    ((= (strcase typeName) "SOCKET") (setq df *DF-SOCKET*))
    ((= (strcase typeName) "AC") (setq df *DF-AC*))
  )
  
  (* connLoad df)
)

(defun calculate-total-connected-load (fixtures / total)
  (setq total 0)
  
  (foreach fixture fixtures
    (setq load (get-block-attribute fixture "LOAD"))
    
    (if (and load (numberp (read load)))
      (setq total (+ total (atof load)))
    )
  )
  total
)

(defun calculate-total-used-load (fixtures / total fixture type load df)
  (setq total 0)
  
  (foreach fixture fixtures
    (setq type (get-block-attribute fixture "TYPE"))
    (setq load (get-block-attribute fixture "LOAD"))
    
    (if (and load (numberp (read load)))
      (progn
        (setq df 1.0)
        
        (cond
          ((= (strcase type) "LIGHT") (setq df *DF-LIGHT*))
          ((= (strcase type) "FAN") (setq df *DF-LIGHT*))
          ((= (strcase type) "SOCKET") (setq df *DF-SOCKET*))
          ((= (strcase type) "AC") (setq df *DF-AC*))
        )
        
        (setq total (+ total (* (atof load) df)))
      )
    )
  )
  total
)

(defun count-fixtures-by-type (/ fixtures groups)
  (setq fixtures (collect-all-fixtures))
  (setq groups (group-fixtures-by-type fixtures))
  
  (mapcar '(lambda (g) (cons (car g) (length (cdr g)))) groups)
)

(defun count-wires (/ ss)
  (setq ss (ssget "_X" '((0 . "ARC"))))
  (if ss (sslength ss) 0)
)

(defun calculate-total-load (/ )
  (calculate-total-connected-load (collect-all-fixtures))
)

;;; ========================================================================
;;; END OF CODE
;;; ========================================================================

(princ "\n========================================")
(princ "\n  ELECTRICAL WIRING SYSTEM LOADED")
(princ "\n========================================")
(princ "\nAvailable commands:")
(princ "\n  BWIRE          - Batch wire SDB to fixtures")
(princ "\n  WIREARC_CHAIN  - Manual SB chaining")
(princ "\n  POPULATE_CKT   - Auto-assign circuits")
(princ "\n  VALIDATE_WIRES - Check wire integrity")
(princ "\n  CHECK_CIRCUITS - Validate circuit loads")
(princ "\n  EXPORT_WIRES   - Export wire connections")
(princ "\n  EXPORT_TREE    - Export hierarchy tree")
(princ "\n  EXPORT_LOAD_CALC - Export load calculation")
(princ "\n  EXPORT_SUMMARY - Export project summary")
(princ "\n  EXPORT_ALL     - Export all reports")
(princ "\n========================================")
(princ)