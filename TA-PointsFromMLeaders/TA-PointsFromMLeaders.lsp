;;;-----------------------------------------------------------------------------
;;; TA-PointsFromMLeaders.lsp
;;; 
;;; Description: Creates points at MLeader locations with elevation values extracted
;;;              from MLeader content (block attributes or MText)
;;;
;;; Developer:   Nikita Prokhor
;;; Version:     2.0
;;; Date:        2024-03-24
;;;
;;; Features:
;;; - Extracts elevation values from MLeader content (block attributes or MText)
;;; - Creates points at MLeader vertex locations with extracted elevations
;;; - Automatically creates layers for correct/incorrect points
;;; - Handles both block-based and MText-based MLeaders
;;; - Supports decimal numbers with both comma and dot separators
;;;
;;; Usage:
;;; 1. Load the LISP file
;;; 2. Type TA-PointsFromMLeaders at the command prompt
;;; 3. Select MLeaders containing elevation values
;;; 4. Points will be created at MLeader vertices with corresponding elevations
;;;
;;; Notes:
;;; - Points with valid elevations are placed on layer "-points-from-mleaders-correct"
;;; - Points with invalid elevations are placed on layer "-points-from-mleaders-incorrect"
;;; - Invalid elevation values return -100 as default
;;;-----------------------------------------------------------------------------

(defun parse-text-for-elevation (txt / len i char start-flag point-flag num-str)
  (setq len (strlen txt)
        i len
        start-flag nil
        point-flag nil
        num-str ""
  )

  (while (> i 0)
    (setq i (- i 1))
    (setq char (substr txt (+ i 1) 1))

    ;; Start parsing when a digit is found
    (if (and (not start-flag) (wcmatch char "[0-9]"))
        (setq start-flag T))

    ;; If parsing started, process characters
    (if start-flag
        (cond
          ;; Convert ',' to '.' for decimal handling
          ((or (= char ",") (= char "."))
           (if (not point-flag)
               (progn (setq point-flag T) (setq num-str (strcat "." num-str))))
          )
          ;; If it's a digit, add to result string
          ((wcmatch char "[0-9]")
           (setq num-str (strcat char num-str))
          )
          ;; Stop parsing on first non-numeric character
          (T (setq i 0))
        )
    )
  )

  ;; Convert extracted string to float
  (if (> (strlen num-str) 0)
      (distof num-str)
      -100 ;; Return -100 if no valid number found
  )
)

(defun AddLeaderToMLeader (vlaMLeader)
  (vl-load-com) ;; Load ActiveX support

  ;; Add a new leader
  (setq newLeaderIndex (vl-catch-all-apply 'vla-AddLeader (list vlaMLeader)))

  ;; Check if the operation was successful
  (if (vl-catch-all-error-p newLeaderIndex)
      (progn 
        (princ "\nError adding leader.")
        nil) ;; Return nil on error
      (progn
        (setq newLeaderIndex (fix newLeaderIndex)) ;; Convert to integer

        ;; Create an array of points for the new line
        (setq points (vlax-make-safearray vlax-vbDouble '(0 . 5)))
        (vlax-safearray-fill points '(0.0 0.0 0.0   ;; Start (0,0,0)
                                      4.0 4.0 0.0)) ;; Second point for the line

        ;; Add a line to the new leader
        (setq result (vl-catch-all-apply 'vla-AddLeaderLine (list vlaMLeader newLeaderIndex points)))

        ;; Check if the line addition was successful
        (if (vl-catch-all-error-p result)
            (progn
              (princ "\nError adding leader line.")
              nil) ;; Return nil on error
            (progn
              newLeaderIndex)) ;; Return the leader index on success
      )
  )
)
(defun create-mleader-layers ()
  (vl-load-com)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq layers (vla-get-Layers doc))
  
  ;; Create layer for incorrect points if it doesn't exist
  (vl-catch-all-apply
    (function
      (lambda ()
        (setq newLayer (vla-Add layers "-points-from-mleaders-incorrect"))
        (vla-put-Color newLayer 1) ;; Red color
      )
    )
  )

  ;; Create layer for correct points if it doesn't exist  
  (vl-catch-all-apply
    (function
      (lambda ()
        (setq newLayer (vla-Add layers "-points-from-mleaders-correct"))
        (vla-put-Color newLayer 3) ;; Green color
      )
    )
  )
)

(defun c:TA-PointsFromMLeaders ()
  (vl-load-com) ;; Load ActiveX support
  ;; Create required layers
  (create-mleader-layers)

  ;; Prompt user to select multiple MLeaders
  (setq ss (ssget '((0 . "MULTILEADER"))))

  (if ss
    (progn
      ;; Get the Active Document and Model Space
      (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
      (setq ms (vla-get-ModelSpace doc))

      ;; Loop through each selected MLeader
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq vlaMLeader (vlax-ename->vla-object ent))

        ;; Get content type (1 = Block Content, 2 = MText Content)
        (setq contentType (vlax-get vlaMLeader 'ContentType))

        (if (= contentType 1)  ;; Ensure the MLeader has a block
          (progn
            ;; Get the block name from the MultiLeader
            (setq sBlock (vlax-get vlaMLeader 'ContentBlockName))
            (princ (strcat "\nBlock Name: " sBlock))

            ;; Get the Blocks collection from the document
            (setq blkDef (vla-Item (vla-get-Blocks doc) sBlock))

            ;; Retrieve attribute value
            (setq attValue "")
            (vlax-for obj blkDef
              (if (= (vla-get-ObjectName obj) "AcDbAttributeDefinition")
                (setq attValue (vla-GetBlockAttributeValue vlaMLeader (vla-get-ObjectID obj)))
              )
            )
            (princ (strcat "\nAttribute Value: " attValue))
          )
          (if (= contentType 2)  ;; Ensure the MLeader has MText
            (progn
              ;; Get MText value
              (setq attValue (vlax-get vlaMLeader 'TextString))
              (princ (strcat "\nMText Value: " attValue))
            )
          )
        )

        ;; Convert attribute or MText value to elevation
        (setq elevation (parse-text-for-elevation attValue))
        ;; rtos converts real number to string
        ;; 2 = decimal format (vs scientific/engineering)
        ;; 4 = precision (number of decimal places)
        (princ (strcat "\nParsed Elevation: " (rtos elevation 2 4)))

        ;; Get the number of leader lines
        (setq leaderIndex 0)

        ;; Add a new leader to MLeader and save the index
        (setq newLeaderIndex (AddLeaderToMLeader vlaMLeader))

        ;; Loop through leader lines
        (while (not (vl-catch-all-error-p (setq vertices (vl-catch-all-apply 'vlax-invoke (list vlaMLeader 'GetLeaderLineVertices leaderIndex)))))
          (if vertices
              (progn
                ;; Extract the first vertex (starting point of the leader)
                (setq firstVert (list (nth 0 vertices)  ;; X
                                      (nth 1 vertices)  ;; Y
                                      elevation)) ;; Use parsed elevation

                ;; Only create point if X and Y are not both 0
                (if (not (and (= (nth 0 vertices) 0.0) (= (nth 1 vertices) 0.0)))
                    (progn
                      ;; Create a point at the first vertex location
                      (setq ptObj (vla-AddPoint ms (vlax-3d-point firstVert)))
                      
                      ;; Set layer based on elevation
                      (vl-catch-all-apply 
                        (function 
                          (lambda ()
                            (if (= elevation -100)
                                (progn
                                  (vla-put-Layer ptObj "-points-from-mleaders-incorrect")
                                )
                                (vla-put-Layer ptObj "-points-from-mleaders-correct")
                            )
                          )
                        )
                      )
                      
                      (princ (strcat "\nPoint created at: " 
                                     (rtos (car firstVert) 2 4) ", " 
                                     (rtos (cadr firstVert) 2 4) ", " 
                                     (rtos (caddr firstVert) 2 4)))
                    )
                )
              )
          )

          ;; Increment leader index
          (setq leaderIndex (1+ leaderIndex))
        )
        ;; Delete the newly added leader
        (if (vl-catch-all-error-p 
              (setq removeResult (vl-catch-all-apply 'vla-RemoveLeader (list vlaMLeader newLeaderIndex)))
            )
            (princ "\nError removing addititonal leader")
        )

        ;; Move to the next MLeader in selection set
        (setq i (1+ i))
      )

      (princ "\nFinished processing all selected MLeaders.")
    )
    (princ "\nNo MLeaders selected.")
  )
  (princ)
) 