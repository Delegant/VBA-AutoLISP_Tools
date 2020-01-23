(defun c:replace (/ fobj sobj str color bool)
(vl-load-com)
  (defun *error* (msg)
    (vla-Highlight fobj :vlax-false)
    (princ)
  )

	(while (if  (null (not (progn
		(setq fobj (vlax-ename->vla-object(ssname (ssget "_:S") 0)))
		(vla-Highlight fobj :vlax-true)
		(setq str(vla-get-TextString fobj))
   		(setq color (vla-get-TrueColor fobj)) 
		(setq sobj (vlax-ename->vla-object(ssname (ssget "_:S") 0)))
	      ))) T)
 
  	(vla-put-TextString sobj str)
	(vla-put-TrueColor sobj color)
	(vla-Highlight fobj :vlax-false)
  	(vla-update sobj))
  
(princ)
)
	      