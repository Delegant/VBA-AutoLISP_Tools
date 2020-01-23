(vl-load-com)

(defun c:apsnum	(/	fil    i      all    sel    lst	   new
		 idx	key    name   nabor  fun    ord	   minp
		 maxp	delta  newpt1 newpt2  keepline  
		)
  (setq	fil '((0 . "INSERT"))
	i   -1
  )
  (cond
    ((null (setq all (ssget "_X" fil)))
     (count:popup
       "No Blocks Found"
       64
       (princ "No blocks were found in the active drawing.")
     )
    )
    ((=	4
	(logand	4
		(cdr (assoc 70 (tblsearch "layer" (getvar 'clayer))))
	)
     )
     (count:popup
       "Current Layer Locked"
       64
       (princ
	 "Please unlock the current layer before using this program."
       )
     )
    )
    ((progn
       (setvar 'nomutt 1)
       (princ "\nSelect blocks to count <all>: ")
       (setq sel
	      (cond
		((null (setq sel (vl-catch-all-apply 'ssget (list fil))))
		 all
		)
		((null (vl-catch-all-error-p sel))
		 sel
		)
	      )
       )
       (setvar 'nomutt 0)
     (initget "Yes No Да Нет _Yes No Да Нет")
       (setq keepline
	      (not
		(= "Нет"
		   (getkword
		     "\nРаспределить автоматически? [Да/Нет] <Да>:"
		   )
		)
	      )
       )
       (initget 32)
       (setq newpt1 (getpoint)

	     newpt2 (getpoint newpt1)
	     zn	    (< (car newpt1) (car newpt2))
	     lngth (abs (- (car newpt1) (car newpt2)))
	)

       
	;(setq oldlngth lngth)
       ;(if(not lngth)(setq lngth 0))
       ;(if (not keepline)
       ;(setq lngth (getint (strcat "\nВведите расстояние между блоками <"
	;(itoa lngth)">: "))))
  	;(if(null lngth)(setq lngth oldlngth))
       (repeat (setq idx (sslength sel))
	 (setq
	   lst (cons++
		 (get-att (setq new (ssname sel (setq idx (1- idx))))
		 )
		 new
		 lst
	       )
	 )
       )
       (setq fun (eval (list 'lambda
			     '(a b)
			     (list (if (= "az" ord)
				     '>
				     '<
				   )
				   '(atof (car a))
				   '(atof (car b))
			     )
		       )
		 )
       )
       (setq lst (vl-sort lst 'fun))
       (repeat (length lst)
	 (setq new
		(vlax-ename->vla-object (cdr (nth (setq i (1+ i)) lst)))
	 )
	 (if (or keepline (null lngth))
	   (progn
	     (vla-GetBoundingBox new 'minp 'maxp)
	     (setq delta (- (car (vlax-safearray->list minp))
			    (car (vlax-safearray->list maxp))
			 )
	     )
	     (vlax-invoke
	       new
	       'Move
	       (vlax-safearray->list
		 (vlax-variant-value (vla-get-InsertionPoint new))
	       )
	       (inspt)
	     )
	   )

	   (progn
	     (vlax-invoke
	       new
	       'Move
	       (vlax-safearray->list
		 (vlax-variant-value (vla-get-InsertionPoint new))
	       )
	       (insptfix)
	     )
	   )
	 )
       )
     )
    )
  )
  (princ)
)
;;----------------------------------------------------------------------;;

(defun cons++ (key name nabor)
  (if (not (null key))
    (cons (cons	key
		name
	  )
	  nabor
    )
    nabor
  )
)


;;----------------------------------------------------------------------;;
(defun inspt ()
  (if (null zn)
    (setq newpt1 (cons (- (car newpt1) (+ 5 (abs delta))) (cdr newpt1)))
    (setq newpt1 (cons (+ (car newpt1) (+ 5 (abs delta))) (cdr newpt1)))
  )

)

;;----------------------------------------------------------------------;;

(defun insptfix	()
  (if (= i 0)
    newpt1
    (if	(null zn)
      (setq newpt1 (cons (- (car newpt1) lngth) (cdr newpt1)))
      (setq newpt1 (cons (+ (car newpt1) lngth) (cdr newpt1)))
    )
   )
  )
  ;;----------------------------------------------------------------------;;

  (defun assoc+	(key new lst / itm)
    (cons (cons key new) lst)

  )

  ;;----------------------------------------------------------------------;;

  (defun get-att (obj / idx ida ina key)
    (if	(= (type obj) 'ENAME)
      (setq obj (vlax-ename->vla-object obj))
    )
    (if	(and obj
	     (vlax-property-available-p obj 'Hasattributes)
	     (eq :vlax-true (vla-get-HasAttributes obj))
	)
      (progn
	(setq ida (vlax-invoke obj 'Getattributes)
	      idx (length ida)
	)
	(while
	  (and
	    (/=	"НОМЕР"
		(vla-get-TagString
		  (setq	ina (nth (setq idx (1- idx))
				 ida
			    )
		  )
		)
	    )
	    (< 0 idx)
	  )
	)
	(if
	  (/= 0 (atoi (setq key (vla-get-TextString ina))))
	   key
	   (if (= "лс" (substr key 1 2))
	     (rtos (+ 1000 (atoi (substr key 4))))
	   )
	)
      )
    )
  )

  ;;----------------------------------------------------------------------;;

  (defun count:popup (ttl flg msg / err)
    (setq err (vl-catch-all-apply
		'vlax-invoke-method
		(list (count:wsh) 'popup msg 0 ttl flg)
	      )
    )
    (if	(null (vl-catch-all-error-p err))
      err
    )
  )

  ;;----------------------------------------------------------------------;;

  (defun count:wsh nil
    (cond (count:wshobject)
	  ((setq count:wshobject (vlax-create-object "wscript.shell"))
	  )
    )
  )

  ;;----------------------------------------------------------------------;;

