**********************Request Parameter from Function
PARAMETERS cInputText, cWhere, cInputTable
**********************On Error
*SET SAFETY OFF
ON ERROR DO errhand WITH ;
   ERROR( ), MESSAGE( ), MESSAGE(1), PROGRAM( ), LINENO( )
**********************Variable to determine what kind of table is being processed (fullpath pass(1), table name pass(2), currently open(3))
cTableType = "0"
cProgramName = PROGRAM()
IF PARAMETERS() = 3 THEN
	cInputTable= STRTRAN(LOWER(cInputTable),'.dbf','')+".dbf"
ENDIF
cFieldName = "date"
**********************Check Number of parameters for 0
IF PARAMETERS() = 0 THEN
	cMessageText = 'You must enter a field paprameter.'
	nMessageType = 16
	cMessageTitle = cProgramName + '("Table Name") - No Parameter Entered Error.'
	MESSAGEBOX(cMessageText, nMessageType, cMessageTitle)
	RETURN
ENDIF
IF PARAMETERS() =1 OR PARAMETERS() = 2 THEN
	SET FULLPATH ON
	cDBFFullPath = DBF()
	cInputTable= cDBFFullPath 		&&Assign Table Variable from current open table
	cTableType = "3"					&&currently open
ELSE 
	cTableType = "1"					&&fullpath pass
ENDIF
**********************Check for populated table variable
IF EMPTY(ALLTRIM(cInputTable)) THEN
	cMessageText = 'There is no valid table in use or specified.' + CHR(13) + 'You must have a table open, or specify a table inside the function.'
	nMessageType = 16
	cMessageTitle = cProgramName + '("Table Name") - No Table Specified Error.'
	MESSAGEBOX(cMessageText, nMessageType, cMessageTitle)
	RETURN								&&If still no table name, exit function and give error message.
ENDIF 
**********************Set table variable with fullpath
cInputTablePath = SUBSTR(ALLTRIM(cInputTable),1,RAT("\",ALLTRIM(cInputTable),1))
IF EMPTY(ALLTRIM(cInputTablePath)) THEN
	cCurrDir = ALLTRIM(SYS(5))+ALLTRIM(CURDIR())
	cTableExists = cCurrDir+cInputTable
	cTableType = "2"					&&table name pass
ELSE 
	cTableExists = cInputTable
ENDIF
**********************Check for single quote in table name
IF "'"$"&cTableExists" THEN
	IF !FILE("&cTableExists")
		cMessageText = "Please do not use tables with single quotes in the name."
		nMessageType = 16
		cMessageTitle = cProgramName + "('Table Name') - Invalid Table Name Error"
		MESSAGEBOX(cMessageText, nMessageType, cMessageTitle)
		RETURN
	ENDIF
	**********************Use table to create array of field names
	USE "&cInputTable"
	**********************Create array of field names
	iArrayLoop=1
	bArrayLoop=.T.
	ifields=AFIELDS(aFieldNameArray)
	**********************Initial String with table name
	cAlterTableString='ALTER TABLE "' + cInputTable+ '" '
	**********************Build alter table string
	DO WHILE bArrayLoop
	 IF iArrayLoop<=ifields THEN
	  cArrayFieldName=aFieldNameArray(iArrayLoop,1)
	  cAlterTableString=cAlterTableString+"ALTER COLUMN "+cArrayFieldName+" NOT NULL "
	  IF iArrayLoop=ifields
	   bArrayLoop=.F.
	  ENDIF
	  iArrayLoop=iArrayLoop+1
	 ENDIF  
	ENDDO
	**********************Run alter table string
	&cAlterTableString
	CLOSE TABLES ALL
	**********************If table was processed from open table (using AlltrimAll()) the table is reopened.
	IF cTableType = "3" THEN
		USE "&cInputTable"
	ENDIF
ELSE
	**********************Check for table exists
	IF !FILE('&cTableExists')
		cMessageText = 'The FILE  "' + cTableExists + '"  does not exist.' + CHR(13) + 'Please check the table name and try again.'
		nMessageType = 16
		cMessageTitle = cProgramName + '("Table Name") - Table Does Not Exist Error.'
		MESSAGEBOX(cMessageText, nMessageType, cMessageTitle)
		RETURN
	ENDIF
	**********************Use table to create array of field names
	USE '&cInputTable'
	**************************************************************************************************************
	**********************             Insert Function Here           ********************************************
	**************************************************************************************************************
	IF PARAMETERS() > 1 THEN
		cSelectStatement = 'select ' + cInputText + ', count(*) as quantity from "' + cInputTable + '" order by ' + cInputText + ' group by ' + cInputText + ' where ' + cWhere + ' into cursor temp12345678'
	ELSE
		cSelectStatement = 'select ' + cInputText + ', count(*) as quantity from "' + cInputTable + '" order by ' + cInputText + ' group by ' + cInputText + ' into cursor temp12345678'
	ENDIF
	strDateTime=ALLTRIM(STR(YEAR(date())))+PADL(ALLTRIM(STR(month(date()))),2,"0")+PADL(ALLTRIM(STR(day(date()))),2,"0")+"_"+PADL(ALLTRIM(STR(hour(datetime()))),2,"0")+PADL(ALLTRIM(STR(minute(datetime()))),2,"0")+PADL(ALLTRIM(STR(sec(datetime()))),2,"0")
	strFileName = STRTRAN(UPPER(ALLTRIM(cInputTable)),'.DBF','') + '_Dist_' + strDateTime
	&cSelectStatement 
	COPY TO '&strFileName' TYPE xl5
	COPY TO '&strFileName'
	**************************************************************************************************************
	**************************************************************************************************************
	**************************************************************************************************************
	CLOSE TABLES ALL
	**********************If table was processed from open table (using AlltrimAll()) the table is reopened.
	IF cTableType = "3" THEN
		USE '&cInputTable'
	ENDIF
ENDIF

SET SAFETY ON
**********************Error Handler
ON ERROR  && restore system error handler
PROCEDURE errhand
	PARAMETER merror, mess, mess1, mprog, mlineno
	DO CASE
		CASE merror = 15
			cMessageText = 'The table  "' + cTableExists + '"  is not a valid table.' + CHR(13) + 'You must have a table open, or specify a valid table inside the function.'
			nMessageType = 16
			cMessageTitle = cProgramName + '("Table Name") - File is not a table.'
			MESSAGEBOX(cMessageText, nMessageType, cMessageTitle)
			SET SAFETY on
			CANCEL
		CASE merror = 1806
			cMessageText = mess
			nMessageType = 16
			cMessageTitle = cProgramName + '("Fields") Error- Field name does not exists.'
			MESSAGEBOX(cMessageText, nMessageType, cMessageTitle)
			SET SAFETY on
			CANCEL		
		OTHERWISE
			cProgrammerEmail = 'reaton@carolina.sourcelink.com'
			erroremail(merror, mess, mess1, mprog, mlineno, cProgrammerEmail)
			cMessageText = 'An unknown error has occured.' + CHR(13) + 'The programmer has been notified.'
			nMessageType = 16
			cMessageTitle = cProgramName + '("Table Name") - Unknown Error'
			MESSAGEBOX(cMessageText, nMessageType, cMessageTitle)
			SET SAFETY on
			CANCEL
	ENDCASE
ENDFUNC