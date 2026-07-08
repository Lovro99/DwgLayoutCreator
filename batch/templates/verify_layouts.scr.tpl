(setq f (open "__OUT_PATH__" "w"))
(setq d (dictsearch (namedobjdict) "ACAD_LAYOUT"))
(foreach pair d (if (= 3 (car pair)) (write-line (cdr pair) f)))
(close f)
