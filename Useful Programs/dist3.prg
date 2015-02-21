PROCEDURE dist3
	PARAMETERS flist, cond, outname

	SET FULLPATH ON
	IF LEN(DBF())=0 THEN
		USE ?
		IF LEN(DBF())=0 THEN
			RETURN
		ENDIF
	ENDIF
	dbfname = DBF()

	IF LEN(ALLTRIM(cond))=0 THEN
		cond = '.T.'
	ENDIF

	&& create field list with just names or aliases (names after "as ")
	flist_all = STRTRAN(STRTRAN(flist, '[', ''), ']', '')
	flist_sort_temp = flist_all
	IF('{'$flist_sort_temp) THEN
		flist_sort_temp = LEFT(flist_sort_temp, AT('{', flist_sort_temp) -1) ;
			+SUBSTR(flist_sort_temp, AT('}', flist_sort_temp) +1)
	ENDIF
	flist_sort = names_only(flist_sort_temp)
	flist_all = STRTRAN(STRTRAN(flist_all, '{', ''), '}', '')
	SELECT &flist_all, cnt(*) as quantity FROM '&dbfname' ORDER BY &flist_sort GROUP BY &flist_sort ;
		WHERE &cond INTO CURSOR zyx1

	IF ']'$flist THEN
		flist_rest = SUBSTR(flist, AT(']', flist) +1)
		&& create field list with just names or aliases (names after "as ")
		flist_sort_temp = flist_rest
		IF('{'$flist_sort_temp) THEN
			flist_sort_temp = LEFT(flist_sort_temp, AT('{', flist_sort_temp) -1) ;
				+SUBSTR(flist_sort_temp, AT('}', flist_sort_temp) +1)
		ENDIF
		flist_sort = names_only(flist_sort_temp)
		flist_rest = STRTRAN(STRTRAN(flist_rest, '{', ''), '}', '')
		SELECT &flist_rest, cnt(*) as quantity FROM zyx1 ORDER BY &flist_sort GROUP BY &flist_sort ;
			INTO CURSOR zyx2
	ENDIF

	SELECT IIF(']'$flist, 'zyx2', 'zyx1')
	SET SAFETY OFF
	COPY TO &outname
	COPY TO &outname TYPE XL5
	SET SAFETY ON

	CLOSE TABLES ALL
	USE '&dbfname'
ENDPROC

FUNCTION names_only(sel_list as String)
	PRIVATE c, i, name_list, sel_item, sel_list_noargs

	sel_list_noargs = ''
	depth= 0
	FOR i=1 TO LEN(sel_list)
		c = SUBSTR(sel_list, i, 1)
		DO CASE
		CASE c='('
	   		depth = depth +1
	   	CASE c=')'
	   		depth = depth -1
	   	OTHERWISE
	   		IF depth = 0 THEN
	   			sel_list_noargs = sel_list_noargs +c
	   		ENDIF
	   	ENDCASE
	ENDFOR
		   		
	name_list = ''
	DO WHILE AT(',', sel_list_noargs)>0
		sel_item = ALLTRIM(LEFT(sel_list_noargs, AT(',', sel_list_noargs) -1))
		sel_item = SUBSTR(sel_item, RAT(' ', sel_item) +1)
		name_list = name_list +sel_item +', '
		sel_list_noargs = SUBSTR(sel_list_noargs, AT(',', sel_list_noargs) +1)
	ENDDO
	sel_item = ALLTRIM(sel_list_noargs)
	sel_item = SUBSTR(sel_item, RAT(' ', sel_item) +1)
	name_list = name_list +sel_item
RETURN name_list