;;;-----------------------------------------------------------------------------
;;; Developer:   Nikita Prokhor
;;; Version:     1.8
;;; Date:        2025-10-06
;;;
;;; Commands:
;;;   TA-SET-ELEV-FOR-PADS     - Sets elevation for polylines based on text inside them
;;;   TA-POINTS-FROM-MLEADERS  - Creates points from MLeader vertices with elevation values
;;;   TA-CONVER-BROKEN-LEADER  - Creates points from broken leaders with elevation values from nearest text
;;;   TA-SCALE-LIST-RESET      - Resets and configures the scale list with either Metric or Imperial scales
;;;   TA-MULTY-POLYLINE-OFFSET - Creates offset polylines with automatic selection of correct offset direction
;;;   TA-ADD-PREFIX-SUFFIX-TO-TEXT - Adds a prefix or suffix to selected text objects
;;;   TA-POINTS-AT-POLY-ANGLES - Creates points at polyline vertices by angle range, with optional start/end points
;;;   TA-3dPOLY-BY-POINTS-BLOCKS - Creates 3D polylines from selected points and blocks
;;;   TA-EXP-SLOPE             - Exports slope lines to GeoJSON format for external processing
;;; 
;;;
;;; Usage:
;;;   TA-SET-ELEV-FOR-PADS
;;;     1. Run command and select polylines and text objects in one selection set
;;;     2. Script finds text inside each polyline. If multiple texts exist inside a polyline,
;;;        only the first one is processed. Avoid having multiple text objects inside a polyline.
;;;     3. Extracts elevation value from text content (e.g. "23.5", "FG=23.5", "23,5", "23.5\PFG")
;;;     4. Sets elevation for polyline if valid value found
;;;     5. Moves objects to appropriate layers:
;;;        - "-TA-pline-to-elev-processed" (Green) for successfully processed polylines
;;;        - "-TA-pline-to-elev-failed" (Red) for polylines with errors
;;;
;;;   TA-POINTS-FROM-MLEADERS
;;;     1. Run command and select MLeaders containing elevation values
;;;     2. Script extracts elevation value from text content (e.g. "23.5", "FG=23.5", "23,5", "23.5\PFG")
;;;     3. Creates points at MLeader vertex locations
;;;     4. Points placed on layers:
;;;        - "-TA-points-from-mleaders-processed" (Green) for points with valid elevations
;;;        - "-TA-points-from-mleaders-invalid" (Red) for points with invalid elevations
;;;
;;;   TA-CONVER-BROKEN-LEADER
;;;     1. Run command and select leaders, ellipses, and text objects in one selection set
;;;     2. Script finds the nearest text to each leader's end point
;;;     3. Extracts elevation value from the nearest text content (e.g. "23.5", "FG=23.5", "23,5", "23.5\PFG")
;;;     4. Creates points at leader's start point with the found elevation
;;;     5. Points are created as AutoCAD point entities
;;;
;;;   TA-SCALE-LIST-RESET
;;;     1. Run command and choose unit system (M for Metric, I for Imperial)
;;;     2. Command deletes all existing scales in the scale list
;;;     3. Adds new scales based on selected unit system:
;;;        - Imperial: 1"=20' to 1"=300' (1:20 to 1:300)
;;;        - Metric: 1:100 to 1:2000
;;;     4. Updates the scale list in the current drawing
;;;
;;;   TA-MULTY-POLYLINE-OFFSET
;;;     1. Run command and select closed polylines to offset (non closed polylines will be filtered)
;;;     2. Choose DELETE or KEEP original polylines (DELETE by default)
;;;     3. Enter offset distance (positive for outside, negative for inside)
;;;     4. Command automatically determines correct offset direction by comparing areas
;;;     5. Moves objects to appropriate layers:
;;;        - "-TA-offset-poly-shifted" (Light blue) for successfully offset polylines
;;;        - "-TA-offset-poly-processed" (Green) for original polylines if kept
;;;        - "-TA-offset-poly-filtered" (Red) for polylines that failed to offset
;;;
;;;   TA-ADD-PREFIX-SUFFIX-TO-TEXT
;;;     1. Run command and select text objects (TEXT or MTEXT)
;;;     2. Choose whether to add the string as a prefix or suffix (Prefix by default)
;;;     3. Enter the string to add
;;;     4. Command modifies all selected text objects by adding the string either at the beginning or end
;;;     5. Works with both regular text (TEXT) and multiline text (MTEXT) objects
;;;
;;;   TA-POINTS-AT-POLY-ANGLES
;;;     1. Run command and input angle range in degrees (defaults to 5–135)
;;;     2. Optionally choose to create points at start and end nodes (Yes/No)
;;;     3. Select one or more polylines (supports LWPOLYLINE, POLYLINE, 3DPOLYLINE)
;;;     4. Command computes interior angles per vertex and creates points where angle ∈ [min,max]
;;;     5. If enabled, adds points at start and end nodes as well
;;;     6. Outputs total number of created points
;;;     Notes:
;;;       - Angle association is applied to the pivot vertex (middle of each triplet)
;;;       - Start/end node points are created only when explicitly confirmed
;;;       - Results list omits duplicate coordinates after the start node
;;;
;;;   TA-3dPOLY-BY-POINTS-BLOCKS
;;;     1. Run command and choose to Keep or Delete original objects (Keep by default)
;;;     2. Select points and/or block references (INSERT entities)
;;;     3. Command extracts coordinates from selected objects:
;;;        - For points: uses point coordinates
;;;        - For blocks: uses block insertion points
;;;     4. Sorts points based on bounding box dimensions (wider than tall = sort by X, else by Y)
;;;     5. Creates a 3D polyline connecting all sorted points
;;;     6. Optionally deletes original objects if Delete option was chosen
;;;
;;;   TA-EXP-SLOPE
;;;     1. Run command and select lines, polylines, or 3D polylines representing slopes
;;;     2. Command extracts coordinates from all selected slope lines
;;;     3. Creates individual line segments from polyline vertices
;;;     4. Exports data to GeoJSON format in the current drawing folder
;;;     5. Output file: slopes-input_local_CRS.geojson
;;;     6. Each line segment gets a unique slopeId for external processing
;;;-----------------------------------------------------------------------------

;;;-----------------------------------------------------------------------------
;;; Constants and Configuration
;;;-----------------------------------------------------------------------------

;; Layer definitions
(setq TA-LAYERS
  '(
    ("-TA-pline-to-elev-processed" 3)  ; Green
    ("-TA-pline-to-elev-failed" 1)     ; Red
    ("-TA-points-from-mleaders-processed" 3)  ; Green
    ("-TA-points-from-mleaders-invalid" 1)    ; Red
    ("-TA-offset-poly-processed" 3) ; Green
    ("-TA-offset-poly-filtered" 1) ; Red
    ("-TA-offset-poly-shifted" 4) ; Light blue
    
  )
)

;; Constants
(setq TA-INVALID-ELEVATION -100)
(setq TA-COORD-PRECISION 6)
(setq TA-POINT-PRECISION 4)

;;;-----------------------------------------------------------------------------
;;; Window command manager to be implemented
;;;-----------------------------------------------------------------------------


;;;-----------------------------------------------------------------------------
;;; Layer Management Functions
;;;-----------------------------------------------------------------------------

(defun create-layers-with-colors (layer-list / doc layers created-layers i)
  (vl-load-com)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq layers (vla-get-Layers doc))
  (setq created-layers '())

  ;; Process each layer tuple
  (foreach layer-tuple layer-list
    (setq layer-name (car layer-tuple))
    (setq layer-color (cadr layer-tuple))
    
    ;; Try to create layer if it doesn't exist
    (vl-catch-all-apply
      (function
        (lambda ()
          (setq newLayer (vla-Add layers layer-name))
          (vla-put-Color newLayer layer-color)
          (setq created-layers (cons layer-name created-layers))
        )
      )
    )
  )

  ;; Print created layers
  (if created-layers
    (princ (strcat "\nCreated layers: " (vl-princ-to-string created-layers)))
    (princ "\nNo new layers were created")
  )
  (princ)
)

;;;-----------------------------------------------------------------------------
;;; Object Filtering Functions
;;;-----------------------------------------------------------------------------

(defun filter-objects-by-type (obj-list type-list / filtered-list obj obj-type i)
  (setq filtered-list
    (vl-remove-if-not
      (function
        (lambda (obj)
          (setq obj-type (vla-get-ObjectName obj))
          (vl-some 
            (function 
              (lambda (type) (= obj-type type))
            )
            type-list
          )
        )
      )
      obj-list
    )
  )
  filtered-list
)


(defun IsPolyline (ent / vla objname dxf0 i)
  (cond
    ((null ent) nil)
    ((= (type ent) 'VLA-OBJECT)
     (setq vla ent))
    ((= (type ent) 'ENAME)
     (setq vla (vlax-ename->vla-object ent)))
    (T (setq vla nil))
  )
  (if vla
    (progn
      (setq objname (vla-get-ObjectName vla))
      (or (= objname "AcDbPolyline")
          (= objname "AcDb2dPolyline")
          (= objname "AcDb3dPolyline"))
    )
    (if (= (type ent) 'ENAME)
      (progn
        ;; Fallback to DXF 0 if VLA conversion failed
        (setq dxf0 (cdr (assoc 0 (entget ent))))
        (if (null dxf0)
          nil
          (member dxf0 '("LWPOLYLINE" "POLYLINE"))
        )
      )
      nil
    )
  )
)
  

;;;-----------------------------------------------------------------------------
;;; Text Processing Functions
;;;-----------------------------------------------------------------------------

(defun get-text-insertion-point (txt-obj / result i)
  (setq result (vl-catch-all-apply 
    (function
      (lambda ()
        (cond
          ;; Get Text insertion point
          ((= (vla-get-ObjectName txt-obj) "AcDbText")
           (setq insertPoint (vlax-safearray->list (vlax-variant-value (vla-get-InsertionPoint txt-obj)))))
          ;; Get MText insertion point  
          ((= (vla-get-ObjectName txt-obj) "AcDbMText")
           (setq insertPoint (vlax-safearray->list (vlax-variant-value (vla-get-InsertionPoint txt-obj)))))
          ;; Return nil for non-text objects
          (T (setq insertPoint nil))
        )
        (if (and insertPoint 
                 (= (length insertPoint) 3)
                 (not (and (= (nth 0 insertPoint) 0.0)
                          (= (nth 1 insertPoint) 0.0))))
            ;; Format each coordinate to use decimal format
            (list (rtos (nth 0 insertPoint) 2 6)
                  (rtos (nth 1 insertPoint) 2 6)
                  (rtos (nth 2 insertPoint) 2 6))
            nil
        )
      )
    )
  ))
  (if (vl-catch-all-error-p result)
      nil
      result
  )
)

(defun get-text-contents (txt-obj / result i)
  (setq result (vl-catch-all-apply
    (function 
      (lambda ()
        (cond
          ;; Get Text string
          ((= (vla-get-ObjectName txt-obj) "AcDbText")
           (vla-get-TextString txt-obj))
          ;; Get MText contents
          ((= (vla-get-ObjectName txt-obj) "AcDbMText")
           (vla-get-TextString txt-obj))
          ;; Return nil for non-text objects
          (T nil)
        )
      )
    )
  ))
  (if (vl-catch-all-error-p result)
      nil
      result
  )
)

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

;;;-----------------------------------------------------------------------------
;;; Geometry Functions
;;;-----------------------------------------------------------------------------

(defun point-inside-polyline (polyline point / coords i j xi yi xj yj intersect)
  ;; Get polyline coordinates with error handling
  (setq coords (vl-catch-all-apply
    (function
      (lambda ()
        (vlax-safearray->list (vlax-variant-value (vla-get-Coordinates polyline)))
      )
    )
  ))
  
  (if (vl-catch-all-error-p coords)
    (progn
      (princ "\nError getting polyline coordinates")
      nil
    )
    (progn
      ;; Initialize intersection counter
      (setq intersect 0)

      
      ;; Check each edge of the polyline
      (setq i 0)
      (while (< i (length coords))
        (setq j (if (= (+ i 2) (length coords)) 0 (+ i 2)))
        
        ;; Get current edge endpoints
        (setq xi (nth i coords))
        (setq yi (nth (1+ i) coords))
        (setq xj (nth j coords))
        (setq yj (nth (1+ j) coords))
        
        ;; Check if edge intersects with ray from point
        (if (and (or (and (> yi (nth 1 point)) (<= yj (nth 1 point)))
                     (and (> yj (nth 1 point)) (<= yi (nth 1 point))))
                 (<= (nth 0 point) (+ xi (* (/ (- xj xi) (- yj yi)) (- (nth 1 point) yi)))))
          (setq intersect (1+ intersect)))
        
        (setq i (+ i 2)))
      
      ;; Point is inside if number of intersections is odd
      (setq result (= (rem intersect 2) 1))
      result
    )
  )
)

(defun close-polyline (polyline / closed-state i)
  ;; Get the closed state of the polyline
  (setq closed-state (vl-catch-all-apply 
    (function 
      (lambda () 
        (vla-get-Closed polyline)
      )
    )
  ))

  ;; Check if there was an error getting the closed state
  (if (vl-catch-all-error-p closed-state)
    (progn
      (princ "\nError checking if polyline is closed")
      nil) ; Return nil on error
    (progn
      ;; If polyline is not closed, close it
      (if (/= closed-state :vlax-true)
        (vla-put-Closed polyline :vlax-true)
      )
      ;; Return T if closed (either originally or after closing)
      T
    )
  )
)

(defun filter-invalid-polylines (polyline-list error-layer / filtered-list i)
  ;; Initialize filtered list
  (setq filtered-list '())

  ;; Process each polyline
  (foreach polyline polyline-list
    (if (not (has-self-intersections polyline))
      (setq filtered-list (cons polyline filtered-list))
      (vla-put-Layer polyline error-layer)
    )
  )

  (foreach polyline filtered-list
    (close-polyline polyline)
  )

  ;; Return filtered list
  filtered-list
)

(defun has-self-intersections (polyline / coords i j x1 y1 x2 y2 x3 y3 x4 y4 denom ua ub found-intersection)
  ;; Get polyline coordinates
  (setq coords (vlax-safearray->list (vlax-variant-value (vla-get-Coordinates polyline))))
  
  ;; Initialize found-intersection flag
  (setq found-intersection nil)
  
  ;; Check each pair of edges for intersections
  (setq i 0)
  (while (and (< i (- (length coords) 2)) (not found-intersection))
    (setq j (+ i 2))
    (while (and (< j (length coords)) (not found-intersection))
      ;; Get edge endpoints
      (setq x1 (nth i coords))
      (setq y1 (nth (1+ i) coords))
      (setq x2 (nth (if (= (+ i 2) (length coords)) 0 (+ i 2)) coords))
      (setq y2 (nth (if (= (+ i 2) (length coords)) 1 (+ i 3)) coords))
      
      (setq x3 (nth j coords))
      (setq y3 (nth (1+ j) coords))
      (setq x4 (nth (if (= (+ j 2) (length coords)) 0 (+ j 2)) coords))
      (setq y4 (nth (if (= (+ j 2) (length coords)) 1 (+ j 3)) coords))
      
      ;; Check if edges intersect
      (if (and (/= (+ i 2) j) ; Skip adjacent edges
               (/= i (if (= (+ j 2) (length coords)) 0 (+ j 2)))) ; Skip first/last edge connection
        (progn
          (setq denom (- (* (- x4 x3) (- y2 y1))
                        (* (- y4 y3) (- x2 x1))))
          (if (/= denom 0) ; Check if lines are not parallel
            (progn
              (setq ua (/ (- (* (- y3 y1) (- x2 x1))
                            (* (- x3 x1) (- y2 y1)))
                         denom))
              (setq ub (/ (- (* (- y3 y1) (- x4 x3))
                            (* (- x3 x1) (- y4 y3)))
                         denom))
              (if (and (< 0 ua 1) (< 0 ub 1))
                (setq found-intersection T)
              )
            )
          )
        )
      )
      (setq j (+ j 2))
    )
    (setq i (+ i 2))
  )
  found-intersection
)

(defun set-polyline-elevation (polyline elevation / i)
  (vl-catch-all-apply
    (function
      (lambda ()
        ;; Set elevation for the polyline
        (vla-put-Elevation polyline elevation)
        ;; Update the polyline
        (vla-Evaluate polyline)
        ;; Move polyline to processed layer
        (vla-put-Layer polyline "-TA-pline-to-elev-processed")
        ;; Return T for success
        T
      )
    )
  )
)

(defun AddLeaderToMLeader (vlaMLeader / i)
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

(defun create-text-points-list (ss3 / main-list i txt-obj)
  (setq main-list nil
        i 0)
  (repeat (sslength ss3)
    (setq txt-obj (vlax-ename->vla-object (ssname ss3 i)))
    (setq main-list (append main-list 
                           (list (list (get-text-insertion-point txt-obj)
                                     (get-text-contents txt-obj)))))
    (setq i (1+ i))
  )
  main-list
)

(defun find-nearest-text-elevation (end-point main-list / min-dist nearest-text i)
  (setq min-dist 1e99
        nearest-text nil)
  
  (foreach text-info main-list
    (setq text-point (car text-info)
          text-content (cadr text-info))
    (if text-point
        (progn
          (setq text-point-num (mapcar 'distof text-point))
          (setq dist (distance end-point text-point-num))
          (if (< dist min-dist)
              (progn
                (setq min-dist dist)
                (setq nearest-text text-content)
              )
          )
        )
    )
  )
  
  (if nearest-text
      (parse-text-for-elevation nearest-text)
      -100
  )
)

;;;-----------------------------------------------------------------------------
;;; Processing of polyline list for verticies and coordinates
;;;-----------------------------------------------------------------------------

(defun calculate-flat-angle-between-points (pt1 pt2 pt3 / ang ang1 ang2 i)
  ;; (princ "\nStart calc flat angle: \n")
  
  ;; Calculate scalar vector production: BA · BC = (x1 - x2)(x3 - x2) + (y1 - y2)(y3 - y2)
  (setq scalarVecProd 
        (+
          (* 
            (- (car pt1) (car pt2))
            (- (car pt3) (car pt2))
          )
          (*
            (- (cadr pt1) (cadr pt2))
            (- (cadr pt3) (cadr pt2))
          )
        )  
  )
  
  ;; Calculate module vector: |BA| = sqrt((x1 - x2)² + (y1 - y2)²); |BC| = sqrt((x3 - x2)² + (y3 - y2)²)
  (defun calc-moduleVec (lcl-pt1 lcl-pt2 / moduleVec i)
      (setq moduleVec
            (sqrt 
              (+
                (expt (- (car lcl-pt1) (car lcl-pt2)) 2) ;(x1 - x2)²
                (expt (- (cadr lcl-pt1) (cadr lcl-pt2)) 2) ;(y1 - y2)²
              )
            )
      )

  )
  
  (setq moduleVec1 (calc-moduleVec pt1 pt2))
  (setq moduleVec2 (calc-moduleVec pt2 pt3))
  (setq den (* moduleVec1 moduleVec2))
  
  ;; Safely compute cosAng, guarding against zero-length vectors and rounding
  (setq cosAng
          (if (equal den 0.0 1e-12)
            1.0
            (/ scalarVecProd den)
          )
        
  )
  

  ;; (print (rtos cosAng 2 8))
  ;; acos using two-argument atan for correct quadrant handling: acos(x) = atan2(√(1-x²), x)
  (defun acos (lcl-cosAng / i)
    (atan
      (sqrt (- 1 (* lcl-cosAng lcl-cosAng)))
      lcl-cosAng
    )
  )
  ;; check if cosAng = 0; angRad = pi/2
  ;; check if cosAng 
  (setq angRad (acos cosAng))
  (setq angDeg (* angRad (/ 180 pi)))    
  
)

(defun get-list-of-polyline-vert-coords (polyline / coords i lcl-vert-coords x y)
  (vl-load-com)

  (if (or (eq (vla-get-ObjectName polyline) "AcDb2dPolyline")
          (eq (vla-get-ObjectName polyline) "AcDbPolyline")
      )
    ;; Handle 2D polylines and lightweight polylines
    (progn
      ;; (princ "poly or 2d poly processing \n")
      ;; Get coordinates as a flat list: (x1 y1 x2 y2 ...)
      (setq coords (vlax-safearray->list (vlax-variant-value (vla-get-Coordinates polyline))))
      (setq i 0)
      (setq lcl-vert-coords '())
      (setq z (vla-get-Elevation polyline))
      ;; Iterate through coordinate pairs and build (x y z) lists
      (while (< i (length coords))
        (setq x (nth i coords))
        (setq y (nth (1+ i) coords))
        (setq lcl-vert-coords (cons (list x y z) lcl-vert-coords))
        (setq i (+ i 2))
      )
      (reverse lcl-vert-coords)
    )
    ;; Handle 3D polylines
    (progn
      (if (eq (vla-get-ObjectName polyline) "AcDb3dPolyline")
        (progn
          ;; Get coordinates as a flat list: (x1 y1 z1 x2 y2 z2 ...)
          (setq coords (vlax-safearray->list (vlax-variant-value (vla-get-Coordinates polyline))))
          ;; (princ (strcat "\n3d poly coords: " (vl-princ-to-string coords)))
          (setq i 0)
          (setq lcl-vert-coords '())
          (while (< i (length coords))
            (setq x (nth i coords))
            (setq y (nth (1+ i) coords))
            (setq z (nth (+ i 2) coords))
            (setq lcl-vert-coords (cons (list x y z) lcl-vert-coords))
            (setq i (+ i 3))
          )
          (reverse lcl-vert-coords)
        )
        ;; Not a supported polyline type
        (progn 
          (if 
            (eq (vla-get-ObjectName polyline) "AcDbLine")
            (progn
            ;; For AcDbLine, return a list of its start and end point coordinates as ((x1 y1 z1) (x2 y2 z2))
            (progn
              (setq startPt (vlax-get-property polyline 'StartPoint))
              (setq endPt   (vlax-get-property polyline 'EndPoint))
              (list
                (vlax-safearray->list (vlax-variant-value startPt))
                (vlax-safearray->list (vlax-variant-value endPt))
              )
            )
            )
            (princ "Unexpected error: not poly and not 3d poly")
          )
          
        )
      )
    )
  )
)

;;;-----------------------------------------------------------------------------
;;; Bounding box calculation for a list of point lists
;;; Input: points-list = ((x y [z]) (x y [z]) ...)
;;; Output: (minX minY maxX maxY) or nil if input invalid/empty
;;;-----------------------------------------------------------------------------
(defun get-bbox (points-list / min-x min-y max-x max-y pt x y i)
  (cond
    ((or (null points-list) (not (listp points-list)))
     nil)
    (T
     (setq min-x 1e99
           min-y 1e99
           max-x -1e99
           max-y -1e99)
     (foreach pt points-list
       (if (and (listp pt) (>= (length pt) 2))
         (progn
           (setq x (car pt))
           (setq y (cadr pt))
           (if (< x min-x) (setq min-x x))
           (if (< y min-y) (setq min-y y))
           (if (> x max-x) (setq max-x x))
           (if (> y max-y) (setq max-y y))
         )
       )
     )
     (if (or (= min-x 1e99) (= min-y 1e99) (= max-x -1e99) (= max-y -1e99))
       nil
       (list min-x min-y max-x max-y)
     )
    )
  )
)

(defun get-polyline-verticies-angles ( insert-poly / polyline i)
  
  ;; Gets polyline  
  
  (cond
    ;; If caller provided an entity and it is a polyline, use it
    ((and insert-poly (IsPolyline insert-poly))
     (setq polyline insert-poly))
    ;; If caller provided a non-polyline (e.g., LINE), skip gracefully
    ((and insert-poly (not (IsPolyline insert-poly)))
     (setq polyline nil))
    ;; If nothing provided, prompt the user
    (T
     (princ "\nSelect polyline: ")
     (setq ss (ssget '((0 . "POLYLINE,LWPOLYLINE"))))
     (if (not ss) (progn (princ "\nNo polyline selected.") (exit)))
     (setq polyline (ssname ss 0))
    )
  )
  ;; If nothing valid to process, return nil silently
  (if (null polyline)
    nil
    (progn
	  (setq polyline (vlax-ename->vla-object polyline))
      (setq vert-coords (get-list-of-polyline-vert-coords polyline))
      (setq i 0)

      ;; (print "Start angle calculation")
      
      (setq vert-coords-ang '())
      (setq start-point (list (car vert-coords) "start_node"))
      (setq end-point (list (nth (- (length vert-coords) 1) vert-coords) "end_node"))
      (setq vect-corods-ang (list start-point))

      
      (if (> (length vert-coords) 2)
        (progn
          (setq i 1)
          (while (< i (- (length vert-coords) 1))
            (setq pt1 (nth (- i 1) vert-coords))
            (setq pt2 (nth i vert-coords))
            (setq pt3 (nth (+ i 1) vert-coords))
            (setq ang (calculate-flat-angle-between-points pt1 pt2 pt3))
            ;; Associate the angle with the current pivot vertex (pt2), not the previous (pt1)
            (setq tmp (list pt2 ang))
            (print tmp)
            (setq vect-corods-ang (cons tmp vect-corods-ang))
            (setq i (+ i 1))
          )
        )
        (print "Two vertecies")
      )
      
      (setq vect-corods-ang (cons end-point vect-corods-ang))
      (setq vect-corods-ang (reverse vect-corods-ang))
      vect-corods-ang
    )
  )
)

;;;-----------------------------------------------------------------------------
;;; Main Commands
;;;-----------------------------------------------------------------------------

(defun c:TA-SET-ELEV-FOR-PADS ( / ss obj-list i)
  (vl-load-com)
  ;; Create required layers
  (create-layers-with-colors TA-LAYERS)
  ;; Get selection set from user
  (princ "\nSelect objects: ")
  (setq ss (ssget))
  
  ;; Convert selection set to list of objects
  (if ss
    (progn
      (setq obj-list '())
      (setq i 0)
      (while (< i (sslength ss))
        (setq obj-list (cons (vlax-ename->vla-object (ssname ss i)) obj-list))
        (setq i (1+ i))
      )
      obj-list
    )
    (progn
      (princ "\nNo objects selected.")
      (exit)
    )
  )

  ;; Filter for polylines only
  (setq polyline-list (filter-objects-by-type obj-list '("AcDbPolyline" "AcDb2dPolyline")))
  ;; Set color ByLayer for all polylines
  (foreach pline polyline-list
    (vla-put-Color pline acByLayer)
  )
  (princ)

  ;; Filter out invalid (self-intersecting) polylines
  (setq polyline-list (filter-invalid-polylines polyline-list "-TA-pline-to-elev-failed"))

  ;; Filter for text and mtext objects
  (setq txt-list (filter-objects-by-type obj-list '("AcDbText" "AcDbMText")))
  ;; Set color ByLayer for all text objects
  (foreach txt txt-list
    (vla-put-Color txt acByLayer)
  )

  ;; Process each polyline
  (foreach polyline polyline-list
    (princ (strcat "\n\nProcessing polyline: " (vla-get-Handle polyline)))
    
    ;; Process each text object until we find one inside
    (setq found-text nil)
    (foreach txt-obj txt-list
      (if (not found-text) ; Only continue if we haven't found a text yet
        (progn
          (setq txt-point (get-text-insertion-point txt-obj))
          
          (if txt-point
            (progn
              ;; Convert string coordinates to numbers for point-inside-polyline
              (setq point-x (distof (nth 0 txt-point)))
              (setq point-y (distof (nth 1 txt-point)))
              (setq point-z (distof (nth 2 txt-point)))
              
              (if (and point-x point-y point-z)
                (if (point-inside-polyline polyline (list point-x point-y point-z))
                  (progn
                    (setq found-text t)
                    (setq text-contents (get-text-contents txt-obj))
                    (if text-contents
                      (progn
                        (princ (strcat "\n  Found text inside polyline: " text-contents))
                        (setq elevation (parse-text-for-elevation text-contents))
                        (princ (strcat "\n  Parsed elevation: " (rtos elevation 2 3)))
                        (if (= elevation -100)
                          (progn
                            (vla-put-Layer polyline "-TA-pline-to-elev-failed")
                            (vla-put-Layer txt-obj "-TA-pline-to-elev-failed")
                          )
                          (progn
                            (set-polyline-elevation polyline elevation)
                            (vla-put-Layer polyline "-TA-pline-to-elev-processed")
                          )
                        )
                      )
                      (progn
                        (princ "\n  Error getting text contents")
                        (vla-put-Layer txt-obj "-TA-pline-to-elev-failed")
                        (vla-put-Layer polyline "-TA-pline-to-elev-failed")
                      )
                    )
                  )
                )
                (progn
                  (princ "\n  Could not convert text point coordinates to numbers")
                  (vla-put-Layer txt-obj "-TA-pline-to-elev-failed")
                  (vla-put-Layer polyline "-TA-pline-to-elev-failed")
                )
              )
            )
            (progn
              (princ "\n  Could not get text insertion point")
              (vla-put-Layer txt-obj "-TA-pline-to-elev-failed")
              (vla-put-Layer polyline "-TA-pline-to-elev-failed")
            )
          )
        )
      )
    )
    (if (not found-text)
      (progn
        (princ "\n  No text found inside polyline")
        (vla-put-Layer polyline "-TA-pline-to-elev-failed")
      )
    )
  )
  (princ)
)

(defun c:TA-POINTS-FROM-MLEADERS (/ i)
  (vl-load-com) ;; Load ActiveX support
  ;; Create required layers
  (create-layers-with-colors TA-LAYERS)

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
                                  (vla-put-Layer ptObj "-TA-points-from-mleaders-invalid")
                                )
                                (vla-put-Layer ptObj "-TA-points-from-mleaders-processed")
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

(defun c:TA-CONVER-BROKEN-LEADER
  (/ doc ss1 ss2 ss3 i)
  (vl-load-com)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  
  ;; Get selection from user
  (setq ss (ssget '((0 . "LEADER,ELLIPSE,TEXT,MTEXT"))))
  (if (not ss)
    (progn
      (princ "\nNo objects selected.")
      (exit)
    )
  )
  
  ;; Filter selection into separate groups
  (setq ss1 (ssadd))
  (setq ss2 (ssadd))
  (setq ss3 (ssadd))
  
  (setq i 0)
  (repeat (sslength ss)
    (setq ent (ssname ss i))
    (setq type (cdr (assoc 0 (entget ent))))
    (cond
      ((= type "LEADER") (ssadd ent ss1))
      ((= type "ELLIPSE") (ssadd ent ss2))
      ((or (= type "TEXT") (= type "MTEXT")) (ssadd ent ss3))
    )
    (setq i (1+ i))
  )

  (princ (strcat "\nFound "
                 (itoa (sslength ss1)) " leaders, "
                 (itoa (sslength ss2)) " ellipses, and "
                 (itoa (sslength ss3)) " text objects."))
  (princ)

  ;; Process each leader
  (setq i 0)
  (repeat (sslength ss1)
    (setq leader (ssname ss1 i))
    (setq ent (entget leader))
    (setq vertices (mapcar 'cdr (vl-remove-if-not '(lambda (x) (= 10 (car x))) ent)))
    (princ (strcat "\nLeader " (itoa (1+ i)) ":"))
    (setq start-point (car vertices)
          end-point (last vertices))
    (princ (strcat "\n  Start point: (" 
                   (rtos (car start-point) 2 8) " " 
                   (rtos (cadr start-point) 2 8) " " 
                   (rtos (caddr start-point) 2 8) ")"))
    (princ (strcat "\n  End point: (" 
                   (rtos (car end-point) 2 8) " "
                   (rtos (cadr end-point) 2 8) " "
                   (rtos (caddr end-point) 2 8) ")"))
    
    ;; Find nearest text and get elevation
    (setq text-list (create-text-points-list ss3))
    (setq elevation (find-nearest-text-elevation end-point text-list))
    
    ;; Create point at start-point with found elevation
    (if (and elevation (/= elevation -100))
        (progn
          (setq new-point (list (car start-point) (cadr start-point) elevation))
          (entmake (list
                    '(0 . "POINT")
                    '(100 . "AcDbEntity")
                    '(100 . "AcDbPoint")
                    (cons 10 new-point)
                  )
          )
          (princ (strcat "\n  Created point with elevation: " (rtos elevation 2 6)))
        )
        (princ "\n  No valid elevation found for this leader")
    )
    
    (setq i (1+ i))
  )
  (princ)
)

(defun c:TA-SCALE-LIST-RESET (/ unit_choice i)
  ;; Initialize keywords for dropdown selection
  (initget "Metric Imperial")
  (setq unit_choice (getkword "\nChoose unit system [Metric/Imperial] <Metric>: "))
  
  ;; If no choice made, default to Metric
  (if (not unit_choice)
    (setq unit_choice "Metric")
  )
  
  ;; Delete all existing scales
  (command "_.-scalelistedit" "d" "*" "e")
  
  (cond
    ;; Imperial scales
    ((= unit_choice "Imperial")
     (command "_.-scalelistedit" "a" "1\"=20'" "1:20" "e")
     (command "_.-scalelistedit" "a" "1\"=30'" "1:30" "e")
     (command "_.-scalelistedit" "a" "1\"=40'" "1:40" "e")
     (command "_.-scalelistedit" "a" "1\"=50'" "1:50" "e")
     (command "_.-scalelistedit" "a" "1\"=60'" "1:60" "e")
     (command "_.-scalelistedit" "a" "1\"=70'" "1:70" "e")
     (command "_.-scalelistedit" "a" "1\"=80'" "1:80" "e")
     (command "_.-scalelistedit" "a" "1\"=90'" "1:90" "e")
     (command "_.-scalelistedit" "a" "1\"=100'" "1:100" "e")
     (command "_.-scalelistedit" "a" "1\"=150'" "1:150" "e")
     (command "_.-scalelistedit" "a" "1\"=200'" "1:200" "e")
     (command "_.-scalelistedit" "a" "1\"=300'" "1:300" "e")
    )
    
    ;; Metric scales
    ((= unit_choice "Metric")
     (command "_.-scalelistedit" "a" "1:100" "1:100" "e")
     (command "_.-scalelistedit" "a" "1:200" "1:200" "e")
     (command "_.-scalelistedit" "a" "1:250" "1:250" "e")
     (command "_.-scalelistedit" "a" "1:400" "1:400" "e")
     (command "_.-scalelistedit" "a" "1:500" "1:500" "e")
     (command "_.-scalelistedit" "a" "1:750" "1:750" "e")
     (command "_.-scalelistedit" "a" "1:1000" "1:1000" "e")
     (command "_.-scalelistedit" "a" "1:1500" "1:1500" "e")
     (command "_.-scalelistedit" "a" "1:2000" "1:2000" "e")
    )
  )
  (princ)
)

(defun c:TA-MULTY-POLYLINE-OFFSET (/ obj-list-base obj-list-offset i)
  (vl-load-com)
  (create-layers-with-colors TA-LAYERS)

  (princ "\nSelect base polylines: ")
  (setq ss-obj-list-base (ssget '((0 . "POLYLINE,LWPOLYLINE"))))
  
  (initget "Delete Keep")
  (setq keep-or-delete (getkword "\nDelete or keep the original polylines? [Delete/Keep] <Delete>: "))
  (if (not keep-or-delete)
    (setq keep-or-delete "Delete")
  )
  
  (setq obj-list-base '())
  (setq i 0)
  (while (< i (sslength ss-obj-list-base))
    (setq obj-list-base (cons (vlax-ename->vla-object (ssname ss-obj-list-base i)) obj-list-base))
    (setq i (1+ i))
  )
  
  (princ (strcat "\nSelected objects: " (itoa (length obj-list-base))))
  
  ;; get offset distance
  (setq offset-distance (getreal "\nEnter offset distance (minus for inside, plus for outside): "))
  
  ;; Check if offset distance is zero
  (if (= offset-distance 0.0)
    (progn
      (princ "\nOffset distance cannot be zero. Command cancelled.")
      (quit)
    )
  )
  
  (setq i 0)
  (while (< i (length obj-list-base))
    ;; (princ (strcat "\nObject " (itoa i) ":"))
    (setq tmp-polyline (nth i obj-list-base))
    ;; Set color and lineweight to ByLayer
    (vla-put-Color tmp-polyline acByLayer)
    (vla-put-Lineweight tmp-polyline acLnWtByLayer)
    
    (setq is-closed (vlax-get-property tmp-polyline 'Closed))
    
    (if (= is-closed :vlax-true)
      (progn
        ;; offset the polyline
        (setq tmp-polyline-offset-plus (vl-catch-all-apply 'vla-Offset (list tmp-polyline offset-distance)))
        (setq tmp-polyline-offset-minus (vl-catch-all-apply 'vla-Offset (list tmp-polyline (- offset-distance))))
        
        ;; Check if offset operations were successful
        (if (or (vl-catch-all-error-p tmp-polyline-offset-plus)
                (vl-catch-all-error-p tmp-polyline-offset-minus))
            (progn
              (princ "\nFailed to create offset for polyline")
              (vla-put-Layer tmp-polyline "-TA-offset-poly-filtered")
            )
            (progn
              ;; Get the first offset polyline from the array
              (setq tmp-polyline-offset-plus (vlax-safearray-get-element (vlax-variant-value tmp-polyline-offset-plus) 0))
              (setq tmp-polyline-offset-minus (vlax-safearray-get-element (vlax-variant-value tmp-polyline-offset-minus) 0))
              
              ;; check area of the offset polylines
              (setq area-plus (vlax-get-property tmp-polyline-offset-plus 'Area))
              (setq area-minus (vlax-get-property tmp-polyline-offset-minus 'Area))
              
              (if (> offset-distance 0)
                (progn
                  (if (> area-plus area-minus)
                    (progn
                      (setq tmp-offset-polyline tmp-polyline-offset-plus)
                      (vla-put-Layer tmp-offset-polyline "-TA-offset-poly-shifted")
                      ;; delete the minus polyline
                      (vla-Delete tmp-polyline-offset-minus)
                      ;; delete the original polyline if needed
                      (if (= keep-or-delete "Delete")
                        (vla-Delete tmp-polyline)
                        (vla-put-Layer tmp-polyline "-TA-offset-poly-processed")
                      )
                    )
                    (progn
                      (setq tmp-offset-polyline tmp-polyline-offset-minus)
                      (vla-put-Layer tmp-offset-polyline "-TA-offset-poly-shifted")
                      ;; delete the plus polyline
                      (vla-Delete tmp-polyline-offset-plus)
                      ;; delete the original polyline if needed
                      (if (= keep-or-delete "Delete")
                        (vla-Delete tmp-polyline)
                        (vla-put-Layer tmp-polyline "-TA-offset-poly-processed")
                      )
                    )
                  )
                )
                (progn
                  (if (< area-plus area-minus)
                    (progn
                      (setq tmp-offset-polyline tmp-polyline-offset-plus)
                      (vla-put-Layer tmp-offset-polyline "-TA-offset-poly-shifted")
                      ;; delete the minus polyline
                      (vla-Delete tmp-polyline-offset-minus)
                      ;; delete the original polyline if needed
                      (if (= keep-or-delete "Delete")
                        (vla-Delete tmp-polyline)
                        (vla-put-Layer tmp-polyline "-TA-offset-poly-processed")
                      )
                    )
                    (progn
                      (setq tmp-offset-polyline tmp-polyline-offset-minus)
                      (vla-put-Layer tmp-offset-polyline "-TA-offset-poly-shifted")
                      ;; delete the plus polyline
                      (vla-Delete tmp-polyline-offset-plus)
                      ;; delete the original polyline if needed
                      (if (= keep-or-delete "Delete")
                        (vla-Delete tmp-polyline)
                        (vla-put-Layer tmp-polyline "-TA-offset-poly-processed")
                      )
                    )
                  )
                )
              )
            )
        )
      )
      (progn
        (princ "\nOpen polyline")
        (vla-put-Layer tmp-polyline "-TA-offset-poly-filtered")
      )
    )
    (setq i (1+ i))
  )
  (princ)
)

(defun c:TA-ADD-PREFIX-SUFFIX-TO-TEXT (/ obj-list str-to-add suff-pref i txt-obj)


  (princ "\nSelect texts and/or MText for adding suff/pref: ")
  (setq obj-list (ssget '((0 . "TEXT,MTEXT"))))
  
  (if obj-list
    (progn
      (initget "Suffix Prefix")
      (setq suff-pref (getkword "\nWhere to add? [Prefix/Suffix] <Prefix>: "))
      (if (not suff-pref)
        (setq suff-pref "Prefix")
      )
      
      (setq str-to-add (getstring (strcat "\nString to add as " suff-pref ": ")))
      
      (setq i 0)
      (repeat (sslength obj-list)
        (setq txt-obj (vlax-ename->vla-object (ssname obj-list i)))
        (if (= (vla-get-ObjectName txt-obj) "AcDbText")
          (progn
            (if (= suff-pref "Suffix")
              (vla-put-TextString txt-obj (strcat (vla-get-TextString txt-obj) str-to-add))
              (vla-put-TextString txt-obj (strcat str-to-add (vla-get-TextString txt-obj)))
            )
          )
          (if (= (vla-get-ObjectName txt-obj) "AcDbMText")
            (progn
              (if (= suff-pref "Suffix")
                (vla-put-TextString txt-obj (strcat (vla-get-TextString txt-obj) str-to-add))
                (vla-put-TextString txt-obj (strcat str-to-add (vla-get-TextString txt-obj)))
              )
            )
          )
        )
        (setq i (1+ i))
      )
      (princ "\nText modification completed.")
    )
    (princ "\nNo text objects selected.")
  )
  (princ)
)

(defun c:TA-POINTS-AT-POLY-ANGLES (/ polyline results-list i)
  ;; Prompt user to enter an angle range (from 1 to 179 degrees), default is 5 to 135
  (defun get-angle-range (/ min-angle max-angle user-input i)
    (princ "\nEnter minimum angle in degrees [1-179] <5>: ")
    (setq user-input (getint))
    (if (or (null user-input) (< user-input 1) (> user-input 179))
      (setq min-angle 5)
      (setq min-angle user-input)
    )
    (princ (strcat "\nEnter maximum angle in degrees [" (itoa (1+ min-angle)) "-179] <135>: "))
    (setq user-input (getint))
    (if (or (null user-input) (<= user-input min-angle) (> user-input 179))
      (setq max-angle 135)
      (setq max-angle user-input)
    )
    (list min-angle max-angle)
  )
  (setq angle-range (get-angle-range))
  
  ;; Ask user if they want to create points at the start and end of the polyline
  (initget "Yes No")
  (setq addStartEnd (getkword "\nCreate points at start and end of polyline? [Yes/No] <No>: "))
  (if (null addStartEnd) (setq addStartEnd "No"))
  
  (princ "\nSelect polylines, lines, or 3D polylines: ")
  (setq polyline (ssget '((0 . "LWPOLYLINE,POLYLINE,LINE,3DPOLYLINE"))))
  (setq results-list '())

  (if polyline
    (progn
      (setq i 0)
      (print "Strat repeat")
      (repeat (sslength polyline)
        (print i)
        (setq ent (ssname polyline i))
        (if (IsPolyline ent)
          (progn
            (print "Is polyline")
            (setq res (get-polyline-verticies-angles ent))
            (if res (setq results-list (cons res results-list)))
          )
          (print "Not a poly")
        )
        (setq i (1+ i))
      )
    )
  )
  
  ;|
  (
   (
      ((2419.91 7405.36 0.0) "start_node") 
      ((2419.91 7405.36 0.0) 88.542) 
      ((2530.47 7470.61 0.0) 166.84) 
      ((2590.48 7362.74 0.0) 78.1015) 
      ((2688.4 7407.99 0.0) "end_node")
    ) 
   (
      ((2603.12 7503.23 0.0) "start_node") 
      ((2603.12 7503.23 0.0) 73.108) 
      ((2664.71 7555.85 0.0) 63.4341) 
      ((2750.0 7562.69 0.0) "end_node")
    )
  )
  |;
  
  
  ;; Iterate over results-list and add points at vertices for angle entries
  (if results-list
    (progn
      (setq created-count 0)
      (foreach poly-results results-list
        (foreach item poly-results
          ;; Expected item formats:
          ;;  - ((x y z) "start_node") or ((x y z) "end_node")
          ;;  - ((x y z) angle)
          
          ;; If user chose to add points at start and end nodes, create them
          (if (and 
                (or (equal addStartEnd "Yes") (equal addStartEnd "yes"))
                (or (equal (cadr item) "start_node") (equal (cadr item) "end_node"))
              )
            (progn
              (setq vtx (car item))
              (if (= (length vtx) 2)
                (setq vtx (append vtx (list 0.0)))
              )
              (if (= (length vtx) 3)
                (progn
                  (entmake (list
                            '(0 . "POINT")
                            '(100 . "AcDbEntity")
                            '(100 . "AcDbPoint")
                            (cons 10 vtx)
                          )
                  )
                  (setq created-count (1+ created-count))
                  (print (strcat "Start/End node point created at: " (vl-princ-to-string vtx)))
                )
                (print "Start/End node point not created (invalid vertex)")
              )
            )
          )
          
          
          ;; processing intermediate points
          (if (numberp (cadr item))
            (progn              
              (print "cheked as number, processing")
              (setq ang(cadr item))
              (print ang)
              
              (if 
                  (and
                    (>= ang (car angle-range))
                    (<= ang (cadr angle-range))
                  )
                  (progn
                    (print "angle is in range")
                    (setq vtx (car item))
                    (if (= (length vtx) 2)
                      (setq vtx (append vtx (list 0.0)))
                    )
                    (print vtx)
                    (if (= (length vtx) 3)
                      (progn
                        (entmake (list
                                  '(0 . "POINT")
                                  '(100 . "AcDbEntity")
                                  '(100 . "AcDbPoint")
                                  (cons 10 vtx)
                                )
                        )
                        (setq created-count (1+ created-count))
                      )
                    )
                  )
               )
              
            )

          )

          
        )
      )
      (princ (strcat "\nCreated points at corners: " (itoa created-count)))
    )
  )
  (princ)
)

(defun c:TA-3dPOLY-BY-POINTS-BLOCKS (/ polyline i deleteOriginalObjects ent vlaObj pt pointList)
  (initget "Keep Delete")
  (setq deleteOriginalObjects (getkword "\nKeep or delete original points and blocks: [Keep/Delete] <Keep>: "))
  (if (null deleteOriginalObjects) (setq deleteOriginalObjects "Keep"))
  
  (princ "\nSelect blocks and/or points: ")
  ;; Select only POINT and INSERT (block reference) entities
  (setq initObjects (ssget '((0 . "POINT,INSERT"))))
  
  ;; Get a list of insertion points (as (x y z) lists) from selected POINT and INSERT entities
  (setq pointList '())
  (if initObjects
    (progn
      (setq i 0)
      (while (< i (sslength initObjects))
        (setq ent (ssname initObjects i))
        (setq vlaObj (vlax-ename->vla-object ent))
        (cond
          ;; For POINT entities
          ((= (vla-get-ObjectName vlaObj) "AcDbPoint")
           (setq pt (vlax-safearray->list (vlax-variant-value (vla-get-Coordinates vlaObj))))
           (if (>= (length pt) 3)
             (setq pointList (cons (list (nth 0 pt) (nth 1 pt) (nth 2 pt)) pointList))
           )
          )
          ;; For INSERT (block reference) entities
          ((= (vla-get-ObjectName vlaObj) "AcDbBlockReference")
           (setq insPt (vlax-safearray->list (vlax-variant-value (vla-get-InsertionPoint vlaObj))))
           (if (>= (length insPt) 3)
             (setq pointList (cons (list (nth 0 insPt) (nth 1 insPt) (nth 2 insPt)) pointList))
           )
          )
        )
        (setq i (1+ i))
      )
      (setq pointList (reverse pointList))
    )
  )
  
  (print pointList)
  
  (setq bbox(get-bbox pointList))

  ;; Sort points depending on bbox dimensions (mimics the Python snippet above)
  (if bbox
    (if (> (- (nth 2 bbox) (nth 0 bbox)) (- (nth 3 bbox) (nth 1 bbox)))
      ;; Wider than tall: sort by X
      (setq pointList
            (vl-sort pointList
              (function (lambda (a b) (< (car a) (car b))))))
      ;; Taller (or equal): sort by Y
      (setq pointList
            (vl-sort pointList
              (function (lambda (a b) (< (cadr a) (cadr b))))))
    )
  )
    
  ;; Create a 3D polyline from the sorted pointList
  (if (and pointList (> (length pointList) 1))
    (progn
      (setq plData (list '(0 . "POLYLINE")
                         '(100 . "AcDbEntity")
                         '(100 . "AcDb3dPolyline")
                         '(66 . 1)
                         '(70 . 8))) ; 8 = 3D polyline
      (entmake plData)
      (setq plEnt (entlast))
      (foreach pt pointList
        (entmake (list
                   '(0 . "VERTEX")
                   '(100 . "AcDbEntity")
                   '(100 . "AcDbVertex")
                   '(100 . "AcDb3dPolylineVertex")
                   (cons 10 pt)
                   '(70 . 32) ; 32 = 3D polyline vertex
                 ))
      )
      (entmake '((0 . "SEQEND")))
      (princ "\n3D Polyline created from selected points.")
    )
    (princ "\nNot enough points to create a 3D polyline.")
  )
  (if (equal deleteOriginalObjects "Delete")
    (command "_.ERASE" initObjects "")
  )


)

(defun c:TA-EXP-SLOPE(/ slopeLines coordList pathToCurrCadFolder pathToResultJSON ent i finalLineList slopeId finalSorted)
  
  (princ "\nSelect lines of slopes to be exported: ")
  (setq slopeLines (ssget '((0 . "LWPOLYLINE,POLYLINE,LINE,3DPOLYLINE"))))
  
  (setq pathToCurrCadFolder (getvar "DWGPREFIX"))
  (setq pathToResultJSON (strcat pathToCurrCadFolder "slopes-input_local_CRS.geojson"))
  
  (setq coordList '())
  
  (if slopeLines
    (progn
      (setq i 0)
      (repeat (sslength slopeLines)
        (setq ent (ssname slopeLines i))
        (setq ent (vlax-ename->vla-object ent))
        (setq coords (get-list-of-polyline-vert-coords ent))
        (if coords
          (setq coordList (cons coords coordList))
        )
        (setq i (1+ i))
      )
      (setq coordList (reverse coordList))
    )
  )
  
  (print coordList)

  
  (setq finalLineList '())
  ;; coordList template (((2642.08 7638.89 0.0) (2710.87 7587.62 150.0)) ((2456.23 7580.58 0.0) (2494.67 7618.5 0.0) (2533.1 7580.58 0.0)))
  (setq i 0)
  (setq slopeId 0)
    
  (foreach polyLine coordList
    (setq j 0)
    (while (< j (- (length polyLine) 1))
      (setq curLine '())
      (setq curLine (cons (nth j polyLine) curLine))
      (setq curLine (cons (nth (+ j 1) polyLine) curLine))
      (setq curLine (cons slopeId curLine))
      (setq finalLineList (cons curLine finalLineList))
      (setq j (+ j 1))
      (setq slopeId (+ slopeId 1))
      
    )
 
  )
  (princ "\nFinalList: ")
  
  ;|
    (
      (5 (2533.1 7580.58 0.0) (2494.67 7618.5 0.0))
      (4 (2494.67 7618.5 0.0) (2456.23 7580.58 0.0))
      (3 (2710.87 7587.62 150.0) (2642.08 7638.89 0.0))
      (2 (2688.4 7407.99 300.0) (2626.81 7322.74 200.0))
      (1 (2626.81 7322.74 200.0) (2588.41 7334.66 200.0))
      (0 (2588.41 7334.66 200.0) (2582.33 7391.28 100.0))
    )|;
  
  ;; geojson template:
  ;; res_dict = {"type": "FeatureCollection", "name": "slopes-input", "crs": None, "features": []}
  ;; feature template:
  ;| { "type": "Feature", 
     "properties": { "elevationsTargetM": null, "padId": null, "slopeId": (iterable)), "slopeTarget": nul }, 
     "geometry": { "type": "LineString", "coordinates": [ [ 875249.77, 1113911.30 ], [ 875291.47, 1113804.12 ] ] } }
  |;
  
  
  
  (print finalLineList)
  (princ)
  
 
  ;; Build GeoJSON FeatureCollection and write to file
  ;; Sort features by slopeId to ensure predictable order
  (setq finalSorted
    (if finalLineList
      (vl-sort finalLineList (function (lambda (a b) (< (car a) (car b)))))
    )
  )

  (setq f (open pathToResultJSON "w"))
  (if f
    (progn
      (write-line "{ \"type\": \"FeatureCollection\", \"name\": \"slopes-input\", \"crs\": null, \"features\": [" f)
      (setq firstFeature T)
      (foreach item finalSorted
        (setq slopeId (car item))
        (setq p1 (cadr item))
        (setq p2 (caddr item))
        ;; Extract XY only; format numbers
        (setq x1 (rtos (car p1) 2 TA-POINT-PRECISION))
        (setq y1 (rtos (cadr p1) 2 TA-POINT-PRECISION))
        (setq x2 (rtos (car p2) 2 TA-POINT-PRECISION))
        (setq y2 (rtos (cadr p2) 2 TA-POINT-PRECISION))

        (if firstFeature
          (setq firstFeature nil)
          (write-line "," f)
        )

        (write-line
          (strcat
            "{ \"type\": \"Feature\", \"properties\": { \"elevationsTargetM\": null, \"padId\": null, \"slopeId\": " (itoa slopeId) ", \"slopeTarget\": null }, "
            "\"geometry\": { \"type\": \"LineString\", \"coordinates\": [[ " x1 ", " y1 " ], [ " x2 ", " y2 " ]] } }"
          )
          f
        )
      )
      (write-line "]}" f)
      (close f)
      (princ (strcat "\nGeoJSON saved to: " pathToResultJSON))
    )
    (princ "\nError opening result file for writing.")
  )
 
)

