;;  DimPoly.lsp [command names: DPI, DPO]
;;  To dimension the lengths of all segments of a Polyline on the Inboard or Outboard
;;    side; for self-intersecting or open Polyline without a clear "inside" and "outside,"
;;    will determine a side -- if not as desired, undo and run other command.
;;  Uses current Dimension and Units settings; dimension line location distance
;;    from Polyline segment = dimension text height.
;;  Accepts LW and 2D "heavy" Polylines, but not 3D Polylines or meshes.
;;  Kent Cooper, 7 March 2014

(vl-load-com)

(defun DP (side / *error* clay cmde styht plsel pl cw inc pt1 pt2 pt3 pt4)

  (defun *error* (errmsg)
    (if (not (wcmatch errmsg "Function cancelled,quit / exit abort,console break"))
      (princ (strcat "\nError: " errmsg))
    ); if
    (setvar 'clayer clay)
    (setvar 'osmode osm)
    (command "_.undo" "_end")
    (setvar 'cmdecho cmde)
    (princ)
  ); defun -- *error*

  (setq clay (getvar 'clayer) osm (getvar 'osmode) cmde (getvar 'cmdecho))
  (setvar 'cmdecho 0)
  (setvar 'osmode 0)
  (command
    "_.undo" "_begin"
    "_.layer" "_make" "YourDimensionLayer" "" ;; <---EDIT
  ); command
    ;; if Layer does not exist, will use default color etc.; add those to command if desired
  (setq styht (cdr (assoc 40 (tblsearch "style" (getvar 'dimtxsty))))); height of text style in current dimension style
  (if (= styht 0.0) (setq styht (* (getvar 'dimtxt) (getvar 'dimscale)))); if above is non-fixed-height
  (while
    (not
      (and
        (setq plsel (entsel "\nSelect Polyline: "))
        (wcmatch (cdr (assoc 0 (entget (car plsel)))) "*POLYLINE")
        (= (logand (cdr (assoc 70 (entget (car plsel)))) 88) 0)
          ;; not 3D or mesh [88 = 8 (3D) + 16 (polygon mesh) + 64 (polyface mesh)]
      ); and
    ); not
    (prompt "\nNothing selected, or not a LW or 2D Polyline.")
  ); while
  (setq pl (vlax-ename->vla-object (car plsel)))
  (vla-offset pl styht); temporary
  (setq cw (< (vla-get-area (vlax-ename->vla-object (entlast))) (vla-get-area pl)))
    ;; clockwise for closed or clearly inside/outside open; may not give
    ;; desired result for open without obvious inside/outside
  (entdel (entlast))
  (repeat (setq inc (fix (vlax-curve-getEndParam pl)))
    (setq
      pt1 (vlax-curve-getPointAtParam pl inc)
      pt2 (vlax-curve-getPointAtParam pl (- inc 0.5)); segment midpoint
      pt3 (vlax-curve-getPointAtParam pl (1- inc))
    ); setq
    (if (equal (angle pt1 pt2) (angle pt2 pt3) 1e-8); line segment
      (command "_.dimaligned" pt1 pt3); then [leaves at dimension line location prompt]
      (command ; else [arc segment]
        "_.dimangular" ""
        (inters ; arc center
          (setq pt4 (mapcar '/ (mapcar '+ pt1 pt2) '(2 2 2)))
          (polar pt4 (+ (angle pt1 pt2) (/ pi 2)) 1)
          (setq pt4 (mapcar '/ (mapcar '+ pt2 pt3) '(2 2 2)))
          (polar pt4 (+ (angle pt2 pt3) (/ pi 2)) 1)
          nil
        ); inters
        pt1 pt3
        "_text" (rtos (abs (- (vlax-curve-getDistAtParam pl inc) (vlax-curve-getDistAtParam pl (1- inc)))))
          ;; [include mode and precision if current dimension style's settings not desired]
      ); command [leaves at dimension line location prompt]
    ); if
    (command ; complete Dimension
      (polar
        pt2
        (apply
          (if (or (and cw (= side "in")) (and (not cw) (= side "out"))) '- '+)
          (list
            (angle '(0 0 0) (vlax-curve-getFirstDeriv pl (- inc 0.5)))
            (/ pi 2)
          ); list
        ); apply
        styht
          ;; [If you use stacked fractions, consider multiplying styht by e.g. 1.5]
      ); polar
    ); command
    (setq inc (1- inc))
  ); repeat
  (setvar 'clayer clay)
  (setvar 'osmode osm)
  (command "_.undo" "_end")
  (setvar 'cmdecho cmde)
  (princ) 
); defun -- C:DPI

(defun C:DPI () (DP "in")); = Dimension Polyline Inside
(defun C:DPO () (DP "out")); = Dimension Polyline Outside

(prompt "\nType DPI to Dimension a Polyline on the Inside, DPO to do so on the Outside.")