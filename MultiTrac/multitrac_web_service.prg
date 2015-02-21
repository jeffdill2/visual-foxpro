****************************************************************************************************
** New IMB Request system 
****************************************************************************************************

DEFINE CLASS pCodeGetterFS as Custom 

	PROTECTED strXMLFileFS as String
	PROTECTED strXMLHeaderFS as String
	PROTECTED strSoapHeaderFS as String
	PROTECTED strCreatePCodeElementFS as String
	PROTECTED strUserIDFS as String 
	PROTECTED strPasswordFS as String 
	PROTECTED strJobNoFS as String 
	PROTECTED strClientIDFS as String 
	PROTECTED strExternalJobNoFS as String 
	PROTECTED strJobNameFS as String 
	PROTECTED strDivIDFS as String 
	PROTECTED strMailClassFS as String
	PROTECTED strIMBFS as String 
	PROTECTED strByStoreFS as String
	PROTECTED strEntryPointFS as String
	PROTECTED strStartDateFS as String
	PROTECTED strEndDateFS as String 
	PROTECTED strDropQtyFS as String
	PROTECTED strVersionFS as String   
	PROTECTED strInhStartDateFS as String
	PROTECTED strInhEndDateFS as String 
	PROTECTED strDropNameFS as String
	
	PROTECTED strDivIDTableFS as String
	PROTECTED strClientIDTableFS as String
	PROTECTED strMailClassTableFS as String
	
	PROTECTED strRetMTJobNoFS as String
	PROTECTED strRetPlanetCodeFS as String
	PROTECTED strVendorIDFS as String 
	PROTECTED strOriginServiceTypeFS as String
	PROTECTED strOriginZipFS as String 
	PROTECTED strServiceOptionFS as String
	PROTECTED strACSFS as String
	PROTECTED strOELFS as String 
	
	PROTECTED bolTestFS as Boolean
		
	&& These Variables will keep track of how many of each Job Details elements I have 
	PROTECTED intStartDateCountFS as Integer 
	PROTECTED intEndDateCountFS as Integer 
	PROTECTED intDropQtyCountFS as Integer 
	PROTECTED intVersionCountFS as Integer 
	PROTECTED intInhStartDateCountFS as Integer 
	PROTECTED intInhEndDateCountFS as Integer
	PROTECTED intDropNameCountFS as Integer 
	PROTECTED intVendorIDCountFS as Integer 
	PROTECTED intMailClassCountFS as Integer 
	
	&& Variables for use with new runGenerator method
	PROTECTED strRunXmlFile as String
    PROTECTED strLoginID as String 
    PROTECTED strLoginPassword as String 
    PROTECTED strMTJobNo as String 
    PROTECTED strFileType as String 
    PROTECTED strInputFile as String 
    PROTECTED strOutputFile as String 
    PROTECTED strClientIDField as String 
    PROTECTED strClientIDStart as String 
    PROTECTED strClientIDLength as String 
    PROTECTED strNameField as String 
    PROTECTED strNameStart as String 
    PROTECTED strNameLength as String 
    PROTECTED strAddress1Field as String 
    PROTECTED strAddress1Start as String 
    PROTECTED strAddress1Length as String 
    PROTECTED strAddress2Field as String 
    PROTECTED strAddress2Start as String 
    PROTECTED strAddress2Length as String 
    PROTECTED strCityField as String 
    PROTECTED strCityStart as String 
    PROTECTED strCityLength as String 
    PROTECTED strStateField as String
    PROTECTED strStateStart as String 
    PROTECTED strStateLength as String 
    PROTECTED strZipField as String 
    PROTECTED strZipStart as String 
    PROTECTED strZipLength as String 
    PROTECTED strZip4Field as String 
    PROTECTED strZip4Start as String 
    PROTECTED strZip4Length as String 
    PROTECTED strDpbcField as String 
    PROTECTED strDpbcStart as String 
    PROTECTED strDpbcLength as String 
    PROTECTED strBarcodeField as String 
    PROTECTED strBarcodeStart as String 
    PROTECTED strBarcodeLength as String 
    PROTECTED strVersionField as String 
    PROTECTED strVersionStart as String 
    PROTECTED strVersionLength as String 
    PROTECTED strPhoneField as String 
    PROTECTED strPhoneStart as String 
    PROTECTED strPhoneLength as String
    PROTECTED strEmailField as String 
    PROTECTED strEmailStart as String 
    PROTECTED strEmailLength as String 
    PROTECTED strStoreField as String 
    PROTECTED strStoreStart as String 
    PROTECTED strStoreLength as String 
    PROTECTED strOptEndrsField as String 
    PROTECTED strOptEndrsStart as String 
    PROTECTED strOptEndrsLength as String 
    PROTECTED strSeedFlagField as String 
    PROTECTED strSeedFlagStart as String 
    PROTECTED strSeedFlagLength as String 
    PROTECTED strEntryField as String 
    PROTECTED strEntryStart as String
    PROTECTED strEntryLength as String 
    PROTECTED strPRateField as String 
    PROTECTED strPRateStart as String 
    PROTECTED strPRateLength as String 
    PROTECTED strPCStartChar as String 
    PROTECTED strPCEndChar as String 
	
	DIMENSION aryJobDetailsFS[1,9]
	**********************************************
	** Array is defined as follows:				**
	** Column 1 = Start Date					**
	** Column 2 = End Date						**
	** Column 3 = Drop Quantity					**
	** Column 4 = Version						**
	** Column 5 = InhStartDate					**
	** Column 6 = InhEndDate					**
	** Column 7 = Drop Name						**
	** Column 8 = Vendor ID						**
	**********************************************
	**********************************************************************
	**	Set variable default values										**
	**********************************************************************
	&& Constants
	strXMLHeaderFS = '<?xml version="1.0" encoding="utf-8"?>'
	strSoapHeaderFS = '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
	strCreatePCodeElementFS = '<CreatePCodeFS xmlns="http://www.multitrac.com/">'
	strUserIDFS = "" 
	strPasswordFS = "" 
	strRetMTJobNoFS = ""
	strRetPlanetCodeFS = ""
	strRetErrorCodeFS = ""
	strRetErrorDescFS = ""
	
	strJobNoFS = "0"
	strClientIDFS = ""
	strExternalJobNoFS = ""
	strJobNameFS = ""
	strDivIDFS = ""
	strMailClassFS = ""
	strIMBFS = ""
	strByStoreFS = ""
	strEntryPointFS = ""
	strStartDateFS = ""
	strEndDateFS = ""
	strDropQtyFS = ""
	strVersionFS = ""
	strInhStartDateFS = ""
	strInhEndDateFS = ""
	strDropNameFS = ""
	strVendorIDFS = ""
	strOriginServiceTypeFS = ""
	strOriginZipFS = ""	
	strServiceOptionFS = ""
	strACSFS = ""
	strOELFS = ""
	
	&& Set Default Values for new variables for the runGenerator Method
	strRunXmlFile = ""
    strLoginID = ""
    strLoginPassword = ""
    strMTJobNo = ""
    strFileType = ""
    strInputFile = ""
    strOutputFile = ""
    strClientIDField = ""
    strClientIDStart = ""
    strClientIDLength = ""
    strNameField = ""
    strNameStart = ""
    strNameLength = ""
    strAddress1Field = ""
    strAddress1Start = ""
    strAddress1Length = ""
    strAddress2Field = ""
    strAddress2Start = ""
    strAddress2Length = "" 
    strCityField = ""
    strCityStart = ""
    strCityLength = ""
    strStateField = ""
    strStateStart = ""
    strStateLength = ""
    strZipField = ""
    strZipStart = ""
    strZipLength = ""
    strZip4Field = ""
    strZip4Start = ""
    strZip4Length = "" 
    strDpbcField = ""
    strDpbcStart = "" 
    strDpbcLength = ""
    strBarcodeField = ""
    strBarcodeStart = ""
    strBarcodeLength = ""
    strVersionField = ""
    strVersionStart = ""
    strVersionLength = ""
    strPhoneField = ""
    strPhoneStart = ""
    strPhoneLength = ""
    strEmailField = ""
    strEmailStart = ""
    strEmailLength = ""
    strStoreField = ""
    strStoreStart = ""
    strStoreLength = ""
    strOptEndrsField = ""
    strOptEndrsStart = ""
    strOptEndrsLength = ""
    strSeedFlagField = ""
    strSeedFlagStart = ""
    strSeedFlagLength = ""
    strEntryField = ""
    strEntryStart = ""
    strEntryLength = ""
    strPRateField = ""
    strPRateStart = ""
    strPRateLength = ""
    strPCStartChar = ""
    strPCEndChar = ""
	
	&& Default is to go to test mode
	bolTestFS = .t.
	
	intStartDateCountFS = 0
	intEndDateCountFS = 0
	intDropQtyCountFS = 0
	intVersionCountFS = 0
	intInhStartDateCountFS = 0
	intInhEndDateCountFS = 0
	intDropNameCountFS = 0
	intVendorIDCountFS = 0
	intMailClassCountFS = 0

	*******************************************************************************
	*******************************************************************************
	&& Begin Procedure Declarations
	PROCEDURE setStartDate(strInputStartDate as Date)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputStartDate)))!="d"
			RETURN "bad entry - date entry must be a date variable type"
		ENDIF 
		&& Iterate the Count 
		this.intStartDateCountFS = this.intStartDateCountFS + 1
		
		&& Declare Variable and set value
		PRIVATE strInputDate as String
		strInputDate = TTOC(strInputStartDate,1)
		
		this.strStartDateFS = SUBSTR(ALLTRIM(strInputDate),1,4) + "-" + SUBSTR(ALLTRIM(strInputDate),5,2) + "-" + SUBSTR(ALLTRIM(strInputDate),7,2)
		
		this.checkArray(this.intStartDateCountFS)	
		
		this.aryJobDetailsFS(this.intStartDateCountFS,1) = this.strStartDateFS
				
		RELEASE strInputDate
		
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setEndDate(strInputEndDate as Date)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputEndDate)))!="d"
			RETURN "bad entry - date entry must be a date variable type"
		ENDIF 
		&& Iterate the Count 
		this.intEndDateCountFS = this.intEndDateCountFS + 1
		
		&& Declare Variable and set value
		PRIVATE strInputDate as String
		strInputDate = TTOC(strInputEndDate,1)
		
		this.strEndDateFS = SUBSTR(ALLTRIM(strInputDate),1,4) + "-" + SUBSTR(ALLTRIM(strInputDate),5,2) + "-" + SUBSTR(ALLTRIM(strInputDate),7,2)
		
		this.checkArray(this.intEndDateCountFS)	
		
		this.aryJobDetailsFS(this.intEndDateCountFS,2) = this.strEndDateFS
				
		RELEASE strInputDate
		
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setInhStartDate(strInputInhStartDate as Date)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputInhStartDate)))!="d"
			RETURN "bad entry - date entry must be a date variable type"
		ENDIF 
		&& Iterate the Count 
		this.intInhStartDateCountFS = this.intInhStartDateCountFS + 1
		
		&& Declare Variable and set value
		PRIVATE strInputDate as String
		strInputDate = TTOC(strInputInhStartDate,1)
		
		this.strInhStartDateFS = SUBSTR(ALLTRIM(strInputDate),1,4) + "-" + SUBSTR(ALLTRIM(strInputDate),5,2) + "-" + SUBSTR(ALLTRIM(strInputDate),7,2)
		
		this.checkArray(this.intInhStartDateCountFS)	
		
		this.aryJobDetailsFS(this.intInhStartDateCountFS,5) = this.strInhStartDateFS
				
		RELEASE strInputDate
		
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setInhEndDate(strInputInhEndDate as Date)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputInhEndDate)))!="d"
			RETURN "bad entry - date entry must be a date variable type"
		ENDIF 
		&& Iterate the Count 
		this.intInhEndDateCountFS = this.intInhEndDateCountFS + 1
		
		&& Declare Variable and set value
		PRIVATE strInputDate as String
		strInputDate = TTOC(strInputInhEndDate,1)
		
		this.strInhEndDateFS = SUBSTR(ALLTRIM(strInputDate),1,4) + "-" + SUBSTR(ALLTRIM(strInputDate),5,2) + "-" + SUBSTR(ALLTRIM(strInputDate),7,2)
		
		this.checkArray(this.intInhEndDateCountFS)	
		
		this.aryJobDetailsFS(this.intInhEndDateCountFS,6) = this.strInhEndDateFS
				
		RELEASE strInputDate
		
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setDropName(strInputDropName as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputDropName)))!="c"
			RETURN "bad entry - Drop Name entry must be a character variable type"
		ENDIF 
		&& Iterate the Count 
		this.intDropNameCountFS = this.intDropNameCountFS + 1
		
		this.checkArray(this.intDropNameCountFS)	
		
		this.strDropNameFS = strInputDropName
		
		this.aryJobDetailsFS(this.intDropNameCountFS,7) = this.strDropNameFS
				
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setDropQty(intInputDropQty as Integer)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(intInputDropQty)))!="n"
			RETURN "bad entry - Drop Qty entry must be a numeric variable type"
		ENDIF 
		&& Iterate the Count 
		this.intDropQtyCountFS = this.intDropQtyCountFS + 1
		
		this.checkArray(this.intDropQtyCountFS)	
		
		this.strDropQtyFS = ALLTRIM(STR(intInputDropQty,10,0))
		
		this.aryJobDetailsFS(this.intDropQtyCountFS,3) = this.strDropQtyFS
				
		RETURN "good"
	ENDPROC
	
	PROCEDURE setVersion(intInputVersion as Integer)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(intInputVersion)))!="n"
			RETURN "bad entry - Version entry must be a numeric variable type"
		ENDIF 
		&& Iterate the Count 
		this.intVersionCountFS = this.intVersionCountFS + 1
		
		this.checkArray(this.intVersionCountFS)	
		
		this.strVersionFS = ALLTRIM(STR(intInputVersion,10,0))
		
		this.aryJobDetailsFS(this.intDropQtyCountFS,4) = this.strVersionFS
				
		RETURN "good"
	ENDPROC
	
	PROCEDURE setVendorID(strInputVendor as String)
		&& Check Input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputVendor)))!="c"
			RETURN "bad entry - Vendor ID must be a character variable type"
		ENDIF 
		&& Iterate the Count 
		this.intVendorIDCountFS = this.intVendorIDCountFS + 1
		
		this.checkArray(this.intVendorIDCountFS)
		
		this.strVendorIDFS = strInputVendor
		
		this.aryJobDetailsFS(this.intVendorIDCountFS,8) = this.strVendorIDFS
		
		RETURN "good"
	ENDPROC
	
	PROCEDURE setJobNo(intInputJobNo as Integer)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(intInputJobNo)))!="n"
			RETURN "bad entry - JobNo entry must be a numeric variable type"
		ENDIF 
				
		this.strJobNoFS = ALLTRIM(STR(intInputJobNo,10,0))
		
		RETURN "good"
	ENDPROC
	
	PROCEDURE setClientID(intInputClientID as Integer)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(intInputClientID)))!="n"
			RETURN "bad entry - ClientID entry must be a numeric variable type"
		ENDIF 
				
		this.strClientIDFS = ALLTRIM(STR(intInputClientID,10,0))
		
		RETURN "good"
	ENDPROC
	
	PROCEDURE setDivID(intInputDivID as Integer)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(intInputDivID)))!="n"
			RETURN "bad entry - DivID entry must be a numeric variable type"
		ENDIF 
				
		this.strDivIDFS = ALLTRIM(STR(intInputDivID,10,0))
		
		RETURN "good"
	ENDPROC
	
	PROCEDURE setExternalJobNo(strInputExternalJobNo as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputExternalJobNo)))!="c"
			RETURN "bad entry - ExternalJobNo must be a character variable type"
		ENDIF 
						
		this.strExternalJobNoFS = strInputExternalJobNo
		
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setJobName(strInputJobName as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputJobName)))!="c"
			RETURN "bad entry - JobName must be a character variable type"
		ENDIF 
						
		this.strJobNameFS = strInputJobName
		
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setIMB(bolInputIMB as Boolean)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(bolInputIMB)))!="l"
			RETURN "bad entry - IMB must be a logical variable type"
		ENDIF 
		
		IF bolInputIMB				
			this.strIMBFS = "true"
		ELSE 
			this.strIMBFS = "false"
		ENDIF 
		
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setByStore(bolInputByStore as Boolean)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(bolInputByStore)))!="l"
			RETURN "bad entry - ByStore must be a logical variable type"
		ENDIF 
		
		IF bolInputByStore				
			this.strByStoreFS = "true"
		ELSE 
			this.strByStoreFS = "false"
		ENDIF 
		
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setMailClass(strInputMailClass as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputMailClass)))!="c"
			RETURN "bad entry - MailClass must be a character variable type"
		ENDIF 
		
		&& 310 = First;  311 = Standard
		&& They look like Service Type IDs, but they are not
		IF !INLIST(ALLTRIM(strInputMailClass),"310","311")
			RETURN "bad entry - Invalid MailClass Entry"
		ENDIF 
		
		&& Iterate the Count 
		this.intMailClassCountFS = this.intMailClassCountFS + 1
		
		this.checkArray(this.intMailClassCountFS)
		
		this.strMailClassFS = strInputMailClass
		
		this.aryJobDetailsFS(this.intMailClassCountFS,9) = this.strMailClassFS
				
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setEntryPoint(strInputEntryPoint as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputEntryPoint)))!="c"
			RETURN "bad entry - EntryPoint must be a character variable type (5 digit zip expected)"
		ENDIF 
		
		LOCAL intChar
		FOR intChar = 1 TO LEN(ALLTRIM(strInputEntryPoint))
			IF !ISDIGIT(SUBSTR(ALLTRIM(strInputEntryPoint),intChar,1))
				RETURN "bad entry - EntryPoint should be the 5-digit zip of entry. Non-Numeric characters found"
			ENDIF 
		ENDFOR 
		RELEASE intChar
		
		IF LEN(ALLTRIM(strInputEntryPoint))<>5
			RETURN "bad entry - EntryPoint should be the 5-digit zip of entry. Incorrect length found."
		ENDIF 
		
		this.strEntryPointFS = strInputEntryPoint
				
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setUserID(strInputUserID as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputUserID)))!="c"
			RETURN "bad entry - UserID must be a character variable type"
		ENDIF 
		
		this.strUserIDFS = strInputUserID
				
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setPassword(strInputPassword as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputPassword)))!="c"
			RETURN "bad entry - Password must be a character variable type"
		ENDIF 
		
		this.strPasswordFS = strInputPassword
				
		RETURN "good"
	ENDPROC
	
	PROCEDURE setTest(bolInputTest as Boolean)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(bolInputTest)))!="l"
			RETURN "bad entry - Test must be a boolean variable type"
		ENDIF 
 
		this.bolTestFS = bolInputTest
				
		RETURN "good"
	ENDPROC
	
	PROCEDURE setOriginServiceType(strOSType as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strOSType)))!="c"
			RETURN "bad entry - Origin Service Type must be a character variable type"
		ENDIF 
		IF !INLIST(ALLTRIM(strOSType),"050","052","051")
			RETURN "bad entry - Invalid Origin Service Type Entry"
		ENDIF 
		
		this.strOriginServiceTypeFS = strOSType
				
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setOriginZip(strOZip as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strOZip)))!="c"
			RETURN "bad entry - Origin Zip must be a character variable type (9 digit zip expected)"
		ENDIF 
		
		LOCAL intChar
		FOR intChar = 1 TO LEN(ALLTRIM(strOZip))
			IF !ISDIGIT(SUBSTR(ALLTRIM(strOZip),intChar,1))
				RETURN "bad entry - Origin Zip should be the 9-digit zip of registered for return mail tracking. Non-Numeric characters found"
			ENDIF 
		ENDFOR 
		RELEASE intChar
		
		IF LEN(ALLTRIM(strOZip))<>9
			RETURN "bad entry - EntryPoint should be the 9-digit zip of registered for return mail tracking. Incorrect length found."
		ENDIF 
		
		this.strOriginZipFS = strOZip
				
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setServiceOption(strInputService as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputService)))!="c"
			RETURN "bad entry - Service Option must be a character variable type"
		ENDIF
		
		IF !INLIST(UPPER(ALLTRIM(strInputService)),"F","B")
			RETURN "bad entry - Service Option must be F or S"
		ENDIF 
		
		this.strServiceOptionFS = UPPER(ALLTRIM(strInputService))
		
		RETURN "good"		
	ENDPROC 
	
	PROCEDURE setACS(strInputAcs as String)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(strInputAcs)))!="c"
			RETURN "bad entry - ACS must be a character variable type"
		ENDIF 
		
		IF !INLIST(UPPER(ALLTRIM(strInputAcs)),"N","A","C")
			RETURN "bad entry - ACS must be N, A, or C"
		ENDIF 
		
		this.strACSFS = UPPER(ALLTRIM(strInputAcs))
		
		RETURN "good"
	ENDPROC 

	PROCEDURE setOEL(bolInputOel as Boolean)
		&& Check input
		IF PARAMETERS()<>1
			RETURN "bad entry - you should have only one parameter"
		ENDIF 
		IF LOWER(ALLTRIM(VARTYPE(bolInputOel)))!="l"
			RETURN "bad entry - OEL must be a logical variable type"
		ENDIF 
		
		IF bolInputOel				
			this.strOELFS = "true"
		ELSE 
			this.strOELFS = "false"
		ENDIF 
		
		RETURN "good"
	ENDPROC 
	
	PROCEDURE setRunXMLFile(parRunXMLFile as string)
		IF PARAMETERS()<>1
			LOCAL strCurrDate 
			strCurrDate = STRTRAN(STRTRAN(STRTRAN(ALLTRIM(TTOC(DATETIME())),"/",""),":","")," ","")
			parRunXMLFile = "Q:\stepheng\xmltesting\" + strCurrDate + ".xml"
			RELEASE strCurrDate
		ENDIF 
		
		IF LOWER(RIGHT(ALLTRIM(parRunXMLFile),4)) != ".xml"
			parRunXMLFile = parRunXMLFile + ".xml"
		ENDIF 
		
		IF EMPTY(ALLTRIM(JUSTPATH(parRunXMLFile)))
			RETURN "ERROR - Path must be passed!"
		ENDIF 			
		
		this.strRunXmlFile = parRunXMLFile
		RETURN "good"
	ENDPROC 
	
	PROCEDURE getRunXMLFile()
		IF EMPTY(ALLTRIM(this.strRunXMLFile))
			this.setRunXMLFile()
		ENDIF 
		RETURN this.strRunXMLFile
	ENDPROC 
		
	PROCEDURE setMTJobNo(parMTJobNo as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - Only one parameter can be passed for the Multitrac Job Number!"
    	ENDIF 
    	
    	IF ALLTRIM(STR(VAL(ALLTRIM(parMTJobNo)),10,0))!=ALLTRIM(parMTJobNo)
    		RETURN "ERROR - MTJob number is a number passed as a string"
    	ENDIF 
    	
    	this.strMTJobNo = parMTJobNo
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setFileType(parFileType as String)
    	IF PARAMETERS()<>1
    		RETURN "ERROR - File type should be DBF or CSV"
    	ENDIF 
    	
    	IF !INLIST(LOWER(ALLTRIM(parFileType)),"dbf","csv")
    		RETURN "ERROR - File type should be DBF or CSV"
    	ENDIF 
    	
    	this.strFileType = parFileType
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setInputFile(parInputFile as string)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No File parameter passed!"
    	ENDIF 
    	
    	IF !FILE((parInputFile))
    		RETURN "ERROR - " + ALLTRIM(parInputFile) + " not found!"
    	ENDIF 
    	
    	this.strInputFile = parInputFile
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setOutputFile(parOutputFile as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No File parameter passed!"
    	ENDIF 
    	
    	this.strOutputFile = parOutputFile
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setClientIDField(parClientIDField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
	    this.strClientIDField = parClientIDField
	   	RETURN "good"
	ENDPROC 
	
	PROCEDURE setNameField(parNameField as String)
		IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strNameField = parNameField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setAddress1Field(parAddress1Field as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strAddress1Field = parAddress1Field
    	RETURN "good"
    ENDPROC
     
    PROCEDURE setAddress2Field(parAddress2Field as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strAddress2Field = parAddress2Field
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setCityField(parCityField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strCityField = parCityField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setStateField(parStateField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strStateField = parStateField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setZipField(parZipField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strZipField = parZipField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setZip4Field(parZip4Field as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strZip4Field = parZip4Field
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setDpbcField(parDpbcField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strDpbcField = parDpbcField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setBarcodeField(parBarcodeField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strBarcodeField = parBarcodeField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setVersionField(parVersionField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strVersionField = parVersionField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setPhoneField(parPhoneField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strPhoneField = parPhoneField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setEmailField(parEmailField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strEmailField = parEmailField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setStoreField(parStoreField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strStoreField = parStoreField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setOptEndrsField(parOptEndrsField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strOptEndrsField = parOptEndrsField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setSeedFlagField(parSeedFlagField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strSeedFlagField = parSeedFlagField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setEntryField(parEntryField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strEntryField = parEntryField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setPRateField(parPRateField as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	this.strPRateField = parPRateField
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setPCStartChar(parPCStartChar as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	IF LEN(ALLTRIM(parPCStartChar))<>1
    		RETURN "ERROR - Start Char must be one Character"
    	ENDIF 
    	
    	this.strPCStartChar = parPCStartChar
    	RETURN "good"
    ENDPROC 
    
    PROCEDURE setPCEndChar(parPCEndChar as String)
    	IF PARAMETERS() <> 1
    		RETURN "ERROR - No parameter passed!"
    	ENDIF 
    	
    	IF LEN(ALLTRIM(parPCEndChar))<>1
    		RETURN "ERROR - End Char must be one Character"
    	ENDIF 
    	
    	this.strPCEndChar = parPCEndChar
    	RETURN "good"
    ENDPROC
	
	PROCEDURE sendJob()
		this.strXMLFileFS = "Q:\stepheng\xmltesting\" + STRTRAN(STRTRAN(this.strJobNameFS," ",""),"/","") + ".xml"
		tempFileName = this.strXMLFileFS
		
		oXMLFile = FCREATE('&tempFilename')
		
		FPUTS(oXMLFile,this.strXMLHeaderFS)
		FPUTS(oXMLFile,this.strSoapHeaderFS)
		FPUTS(oXMLFile,'<soap:Body>')
		FPUTS(oXMLFile,this.strCreatePCodeElementFS)
		FPUTS(oXMLFile,'<userId>' + this.strUserIDFS + '</userId>')
		FPUTS(oXMLFile,'<password>' + this.strPasswordFS + '</password>')
		FPUTS(oXMLFile,'<jobSpecs>')
		FPUTS(oXMLFile,'<JobDetails>')
		
		FOR i = 1 TO ALEN(this.aryJobDetailsFS,1)
			FPUTS(oXMLFile,'<JobDet_FS>')
			FPUTS(oXMLFile,'<StartDate>' + this.aryJobDetailsFS(i,1) + '</StartDate>')
			FPUTS(oXMLFile,'<EndDate>' + this.aryJobDetailsFS(i,2) + '</EndDate>')
			FPUTS(oXMLFile,'<DropQty>' + this.aryJobDetailsFS(i,3) + '</DropQty>')
			FPUTS(oXMLFile,'<Version>' + this.aryJobDetailsFS(i,4) + '</Version>')
			FPUTS(oXMLFile,'<InhStartDate>' + this.aryJobDetailsFS(i,5) + '</InhStartDate>')
			FPUTS(oXMLFile,'<InhEndDate>' + this.aryJobDetailsFS(i,6) + '</InhEndDate>')
			FPUTS(oXMLFile,'<DropName>' + this.aryJobDetailsFS(i,7) + '</DropName>')
			FPUTS(oXMLFile,'<VendorID>' + this.aryJobDetailsFS(i,8) + '</VendorID>')
			FPUTS(oXMLFile,'<MailClass>' + this.aryJobDetailsFS(i,9) + '</MailClass>')
			FPUTS(oXMLFile,'</JobDet_FS>')
		ENDFOR 
		
		FPUTS(oXMLFile,'</JobDetails>')
		FPUTS(oXMLFile,'<JobNo>' + this.strJobNoFS + '</JobNo>')
		FPUTS(oXMLFile,'<ClientId>' + this.strClientIDFS + '</ClientId>')
		FPUTS(oXMLFile,'<ExternalJobNo>' + this.strExternalJobNoFS + '</ExternalJobNo>')
		FPUTS(oXMLFile,'<JobName>' + this.strJobNameFS + '</JobName>')
		FPUTS(oXMLFile,'<DivId>' + this.strDivIDFS + '</DivId>')
		FPUTS(oXMLFile,'<IMB>' + this.strIMBFS + '</IMB>')
		FPUTS(oXMLFile,'<ByStore>' + this.strByStoreFS + '</ByStore>')
		FPUTS(oXMLFile,'<EntryPoint>' + this.strEntryPointFS + '</EntryPoint>')
		
		&& Only add origin inforamtion if needed
		IF !EMPTY(ALLTRIM(this.strOriginServiceTypeFS)) AND !EMPTY(ALLTRIM(this.strOriginZipFS))
			FPUTS(oXMLFile,'<OriginZip>' + this.strOriginZipFS + '</OriginZip>')
			FPUTS(oXMLFile,'<OriginServiceType>' + this.strOriginServiceTypeFS + '</OriginServiceType>')
		ENDIF 
		
		FPUTS(oXMLFile,'<ACS>' + this.strACSFS + '</ACS>')
		FPUTS(oXMLFile,'<ServiceOption>' + this.strServiceOptionFS + '</ServiceOption>')
		FPUTS(oXMLFile,'<OEL>' + this.strOELFS + '</OEL>')
		
		FPUTS(oXMLFile,'</jobSpecs>')
		FPUTS(oXMLFile,'</CreatePCodeFS>')
		FPUTS(oXMLFile,'</soap:Body>')
		FPUTS(oXMLFile,'</soap:Envelope>')

		FCLOSE(oXMLFile)
	
		RELEASE oXMLFile
		
		&& Declare objects to be used
		PRIVATE oHTTP as Object 
		PRIVATE oXML as Object 
		PRIVATE retDoc as Object 
		PRIVATE jobNoList as Object 
		PRIVATE pCodeList as Object 
		
		oHTTP = CREATEOBJECT("Microsoft.XMLHTTP")
		oXML = CREATEOBJECT("msxml2.DOMDocument.6.0")
		retDoc = CREATEOBJECT("msxml2.DOMDocument.6.0")
		
		PRIVATE xx = []
		xx = FILETOSTR('&tempFilename')
		oXML.loadXML(xx)

		&&DELETE FILE (tempFilename)

		&& If this is a test, post to the test site, otherwise to the live site
		IF this.bolTestFS
			oHTTP.open("POST","http://172.30.26.240/pcws/pcode.asmx", .F.)
		ELSE 
			&&oHTTP.open("POST","http://172.30.26.240/pcws/pcode.asmx", .F.)
			oHTTP.open("POST","https://www.multitrac.com/pcws/pcode.asmx", .F.)
		ENDIF 
		
		oHTTP.send(oXML)
		
		retDoc = oHTTP.responseXML
		
		jobNoList = retDoc.getElementsByTagName("JobNo")
			
		IF jobNoList.length <> 1
			this.strRetMTJobNoFS = retDoc.text
			RELEASE jobNoList
		ELSE 
			this.strRetMTJobNoFS = JobNoList.Item(0).text
			RELEASE jobNoList
		ENDIF 
		
		errorCodeList = retDoc.getElementsByTagName("ErrorCode")
		errorDescList = retDoc.getElementsByTagName("ErrorDesc")
		
		IF errorCodeList.length <> 1
			this.strRetErrorCodeFS = retDoc.text
		ELSE 
			this.strRetErrorCodeFS = errorCodeList.Item(0).text
		ENDIF
		 
		IF errorDescList.length <> 1
			this.strRetErrorDescFS = retDoc.text
		ELSE 
			this.strRetErrorDescFS = errorDescList.Item(0).text
		ENDIF 			
		
	ENDPROC 
	
	PROCEDURE getMTJob()
		
		RETURN this.strRetMTJobNoFS
				
	ENDPROC 
	
	PROCEDURE getErrorCode()
	
		RETURN this.strRetErrorCodeFS
	
	ENDPROC 
	
	PROCEDURE getErrorDesc()
	
		RETURN this.strRetErrorDescFS
	
	ENDPROC 
	
	PROCEDURE runMTGenerator()
		&& Method will take defined fields and run the Mutitrac Generator
		&& An XML file must be passed to execute the C#.NET EXE
		
		&& Check to make sure all required fields have been populated
		IF EMPTY(ALLTRIM(this.strRunXmlFile))
			this.setRunXmlFile()
		ENDIF 
		
		IF EMPTY(ALLTRIM(this.strMTJobNo))
			RETURN "ERROR - Multitrac Job Number has not been set"
		ENDIF 
		
		IF EMPTY(ALLTRIM(this.strFileType))
			RETURN "ERROR - File Type has not been set. Valid types are DBF and CSV"
		ENDIF 
		
		IF EMPTY(ALLTRIM(this.strInputFile))
			RETURN "ERROR - Input File Not Defined"
		ENDIF 
		
		IF EMPTY(ALLTRIM(this.strOutputFile))
			RETURN "ERROR - Output File Not Defined"
		ENDIF 
		
		IF EMPTY(ALLTRIM(this.strClientIDField)) AND EMPTY(ALLTRIM(this.strNameField))
			RETURN "ERROR - Client ID or Name field must be defined"
		ENDIF 
		
		IF EMPTY(ALLTRIM(this.strAddress1Field))
			RETURN "ERROR - Address 1 Field must be defined"
		ENDIF 
		
		IF EMPTY(ALLTRIM(this.strBarcodeField)) AND EMPTY(ALLTRIM(this.strZipField))
			RETURN "ERROR - Barcode Field must be defined in the absence of Zip Code Field"
		ENDIF 
		
		IF !EMPTY(ALLTRIM(this.strBarcodeField)) AND !EMPTY(ALLTRIM(this.strZipField))
			RETURN "ERROR - Barcode and Zip fields should not both be defined for the same file"
		ENDIF 
		
		IF !EMPTY(ALLTRIM(this.strZipField))
			IF EMPTY(ALLTRIM(this.strZip4Field))
				RETURN "ERROR - You must define a Zip+4 Field when using the zip field rather than Barcode Field"
			ENDIF 
			IF EMPTY(ALLTRIM(this.strDpbcField))
				RETURN "ERROR - You must define a Dpbc Field when using the zip field rather than Barcode Field"
			ENDIF 
		ENDIF 
		
		IF EMPTY(ALLTRIM(this.strVersionField))
			RETURN "ERROR - Version Field must be defined"
		ENDIF 
		
		&& if we pass the checks, we can build the XML File
		oRunXml = FCREATE((this.strRunXmlFile))
		
		FPUTS(oRunXml,'<?xml version="1.0" encoding="utf-8"?>')
		FPUTS(oRunXml,'<parameters>')
		FPUTS(oRunXml,'<LoginID>' + ALLTRIM(this.strUserIDFS) + '</LoginID>')
  		FPUTS(oRunXml,'<LoginPSW>' + ALLTRIM(this.strPasswordFS) + '</LoginPSW>')
		FPUTS(oRunXml,'<!-- SourceLink Multitrac Job ID -->')
  		FPUTS(oRunXml,'<JobNo>' + ALLTRIM(this.strMTJobNo) + '</JobNo>')
		FPUTS(oRunXml,'<!-- File Type - Input file type -->') 
		FPUTS(oRunXml,'<FileType>' + ALLTRIM(this.strFileType) + '</FileType>')
		FPUTS(oRunXml,'<!-- Input file - full path to input dbf file -->')
		FPUTS(oRunXml,'<InputFile>' + ALLTRIM(this.strInputFile) + '</InputFile>')
		FPUTS(oRunXml,'<!-- Output file - full path to output dbf file -->')
		FPUTS(oRunXml,'<OutputFile>' + ALLTRIM(this.strOutputFile) + '</OutputFile>')
		FPUTS(oRunXml,'<ClientIDFieldName>' + ALLTRIM(this.strClientIDField) + '</ClientIDFieldName>')
		FPUTS(oRunXml,'<ClientIDStart />')
		FPUTS(oRunXml,'<ClientIDLength />')
		FPUTS(oRunXml,'<NameFieldName>' + ALLTRIM(this.strNameField) + '</NameFieldName>')
		FPUTS(oRunXml,'<NameStart />')
		FPUTS(oRunXml,'<NameLength />')
		FPUTS(oRunXml,'<Address1FieldName>' + ALLTRIM(this.strAddress1Field) + '</Address1FieldName>')
		FPUTS(oRunXml,'<Address1Start />')
		FPUTS(oRunXml,'<Address1Length />')
		FPUTS(oRunXml,'<Address2FieldName>' + ALLTRIM(this.strAddress2Field) + '</Address2FieldName>')
		FPUTS(oRunXml,'<Address2Start />')
		FPUTS(oRunXml,'<Address2Length />')
		FPUTS(oRunXml,'<CityFieldName>' + ALLTRIM(this.strCityField) + '</CityFieldName>')
		FPUTS(oRunXml,'<CityStart />')
		FPUTS(oRunXml,'<CityLength />')
		FPUTS(oRunXml,'<StateFieldName>' + ALLTRIM(this.strStateField) + '</StateFieldName>')
		FPUTS(oRunXml,'<StateStart />')
		FPUTS(oRunXml,'<StateLength />')
		FPUTS(oRunXml,'<ZipFieldName>' + ALLTRIM(this.strZipField) + '</ZipFieldName>')
		FPUTS(oRunXml,'<ZipStart />')
		FPUTS(oRunXml,'<ZipLength />')
		FPUTS(oRunXml,'<Zip4FieldName>' + ALLTRIM(this.strZip4Field) + '</Zip4FieldName>')
		FPUTS(oRunXml,'<Zip4Start />')
		FPUTS(oRunXml,'<Zip4Length />')
		FPUTS(oRunXml,'<DPBCFieldName>' + ALLTRIM(this.strDpbcField) + '</DPBCFieldName>')
		FPUTS(oRunXml,'<DPBCStart />')
		FPUTS(oRunXml,'<DPBCLength />')
		FPUTS(oRunXml,'<BarcodeFieldName>' + ALLTRIM(this.strBarcodeField) + '</BarcodeFieldName>')
		FPUTS(oRunXml,'<BarcodeStart />')
		FPUTS(oRunXml,'<BarcodeLength />')
		FPUTS(oRunXml,'<VersionFieldName>' + ALLTRIM(this.strVersionField) + '</VersionFieldName>')
		FPUTS(oRunXml,'<VersionStart />')
		FPUTS(oRunXml,'<VersionLength />')
		FPUTS(oRunXml,'<PhoneFieldName>' + ALLTRIM(this.strPhoneField) + '</PhoneFieldName>')
		FPUTS(oRunXml,'<PhoneStart />')
		FPUTS(oRunXml,'<PhoneLength />')
		FPUTS(oRunXml,'<EmailFieldName>' + ALLTRIM(this.strEmailField) + '</EmailFieldName>')
		FPUTS(oRunXml,'<EmailStart />')
		FPUTS(oRunXml,'<EmailLength />')
		FPUTS(oRunXml,'<StoreFieldName>' + ALLTRIM(this.strStoreField) + '</StoreFieldName>')
		FPUTS(oRunXml,'<StoreStart />')
		FPUTS(oRunXml,'<StoreLength />')
		FPUTS(oRunXml,'<OptEndrsFieldName>' + ALLTRIM(this.strOptEndrsField) + '</OptEndrsFieldName>')
		FPUTS(oRunXml,'<OptEndrsStart />')
		FPUTS(oRunXml,'<OptEndrsLength />')
		FPUTS(oRunXml,'<SeedFlagFieldName>' + ALLTRIM(this.strSeedFlagField) + '</SeedFlagFieldName>')
		FPUTS(oRunXml,'<SeedFlagStart />')
		FPUTS(oRunXml,'<SeedFlagLength />')
		FPUTS(oRunXml,'<EntryFieldName>' + ALLTRIM(this.strEntryField) + '</EntryFieldName>')
		FPUTS(oRunXml,'<EntryStart />')
		FPUTS(oRunXml,'<EntryLength />')
		FPUTS(oRunXml,'<PRateFieldName>' + ALLTRIM(this.strPRateField) + '</PRateFieldName>')
		FPUTS(oRunXml,'<PRateStart />')
		FPUTS(oRunXml,'<PRateLength />')
		FPUTS(oRunXml,'<!-- Planetcode starting and endging characters -->')
		FPUTS(oRunXml,'<PCStartChar>' + ALLTRIM(this.strPCStartChar) + '</PCStartChar>')
		FPUTS(oRunXml,'<PCEndChar>' + ALLTRIM(this.strPCEndChar) + '</PCEndChar>')
		FPUTS(oRunXml,'</parameters>')
		
		FCLOSE(oRunXML)
		
		LOCAL strMTError
		LOCAL strServer
		&& Need to get which server we are on, since there is a different path for different servers
		strServer = ALLTRIM(SUBSTR(ALLTRIM(SYS(0)),1,ATC("#",ALLTRIM(SYS(0)))-1))
		
		&& If you have multiple servers, check them here
		DO CASE 
			CASE LOWER(ALLTRIM(strServer))=="vader"
				strErrorFolder = "C:\program files (x86)\sourcelink\MultitracGenerator\log\"
			CASE LOWER(ALLTRIM(strServer))=="red"
				strErrorFolder = "C:\program files\sourcelink\MultitracGenerator\log\"
			CASE LOWER(ALLTRIM(strServer))=="blue"
				strErrorFolder = "C:\program files\sourcelink\MultitracGenerator\log\"
			CASE LOWER(ALLTRIM(strServer))=="chewie"
				strErrorFolder = "C:\Program Files (x86)\Sourcelink\MultitracGenerator\log\"
			CASE LOWER(ALLTRIM(strServer))=="fett"
				strErrorFolder = "C:\program files\sourcelink\MultitracGenerator\log\"
			CASE LOWER(ALLTRIM(strServer))=="archer"
				strErrorFolder = "C:\program files\sourcelink\MultitracGenerator\log\"
			OTHERWISE
				RETURN "ERROR - Invalid Server"
		ENDCASE 
		strMTError = ALLTRIM(strErrorFolder) + "MTErrors.log"
		
		IF FILE((strMTError))
			LOCAL strMTPastError 
			strMTPastError = ALLTRIM(strErrorFolder) + "MTErrors_" + ;
								PADL(ALLTRIM(STR(MONTH(DATE()),10,0)),2,"0") + ;
								PADL(ALLTRIM(STR(DAY(DATE()),10,0)),2,"0") + ;
								ALLTRIM(STR(YEAR(DATE()),10,0)) + "_" + ;
								SYS(2) + ".log"
								
			RENAME (strMTError) TO (strMTPastError)
		ENDIF 
		
		&& Now that the file is written, we can run the Multitrac Generator Executable
		LOCAL strMTExe
		
		strServer = ALLTRIM(SUBSTR(LOWER(ALLTRIM(SYS(0))),1,ATC("#",ALLTRIM(SYS(0)))-1))
		
		&& In case you have multiple servers with different executable locations
		DO CASE 
			CASE LOWER(ALLTRIM(strServer))=="red"
				strMTExe = "C:\program files\sourcelink\MultitracGenerator\MultitracGenerator.exe"
				
			CASE LOWER(ALLTRIM(strServer))=="vader"
				strMTExe = "C:\program files (x86)\sourcelink\MultitracGenerator\MultitracGenerator.exe"
				
			CASE LOWER(ALLTRIM(strServer))=="blue"
				strMTExe = "C:\program files\sourcelink\MultitracGenerator\MultitracGenerator.exe"
				
			CASE LOWER(ALLTRIM(strServer))=="chewie"
				strMTExe = "C:\program files (x86)\sourcelink\MultitracGenerator\MultitracGenerator.exe"
				
			CASE LOWER(ALLTRIM(strServer))=="fett"
				strMTExe = "C:\program files\sourcelink\MultitracGenerator\MultitracGenerator.exe"
				
			CASE LOWER(ALLTRIM(strServer))=="archer"
				strMTExe = "C:\program files\sourcelink\MultitracGenerator\MultitracGenerator.exe"
				
			OTHERWISE
				RETURN "ERROR - Invalid Server"
				
		ENDCASE 
		
		spawn('"' + ALLTRIM(strMTExe) + '" "' + this.strRunXmlFile + '"')
		
		IF FILE((strMTError)) OR !FILE((this.strOutputFile))
			RETURN "ERROR - Multitrac Generator Failed see C:\program files\sourcelink\multitracgenerator\log\mterrors.log for details"
		ELSE 
			RETURN "Multitrac Successful"
		ENDIF 		
		
	ENDPROC 
	
	PROCEDURE postMailDat(parFileName as String, parMtJob as String)

		IF !FILE((parFileName))
			RETURN "ERROR - File " + ALLTRIM(parFileName) + " not found"
		ENDIF 

		IF SUBSTR(ALLTRIM(STR(VERSION(5),10,0)),1,1)!="9"
			RETURN "ERROR - Mail.Dat Upload Procedure can only be used with FoxPro 9"
		ENDIF 
		
		&& Get the file into a character string
		xx = []
		xx = FILETOSTR(parFileName)
		
		&& as I understand it, passing a binary string will act like a byte array with .Net
		&& convert the string to binary
		fBinary = STRCONV(xx,13)
		
		&& write the xml file for posting to the web service
		tempFilename = ADDBS(JUSTPATH(ALLTRIM(parFileName))) + ALLTRIM(parMtJob) + "_post_maildat.xml"
		oXMLFile = FCREATE('&tempFilename')
		
		FPUTS(oXMLFile,this.strXMLHeaderFS)
		FPUTS(oXMLFile,this.strSoapHeaderFS)
		FPUTS(oXMLFile,'<soap:Body>')
		FPUTS(oXMLFile,'<LoadJobMailDatFile xmlns="http://www.multitrac.com/">')
		FPUTS(oXMLFile,'<userId>' + this.strUserIDFS + '</userId>')
		FPUTS(oXMLFile,'<password>' + this.strPasswordFS + '</password>')
		FPUTS(oXMLFile,'<jobid>' + ALLTRIM(parMtJob) + '</jobid>')
		FPUTS(oXMLFile,'<filename>' + ALLTRIM(JUSTFNAME(parFileName)) + '</filename>')
		FPUTS(oXMLFile,'<file>' + fbinary + '</file>')
		FPUTS(oXMLFile,"</LoadJobMailDatFile>")
		FPUTS(oXMLFile,"</soap:Body>")
		FPUTS(oXMLFile,"</soap:Envelope>")
		
		FCLOSE(oXMLFile)
		
		oHTTP = CREATEOBJECT("Microsoft.XMLHTTP")
		oXML = CREATEOBJECT("msxml2.DOMDocument.6.0")
		retDoc = CREATEOBJECT("msxml2.DOMDocument.6.0")
		
		xy = []
		xy = FILETOSTR((tempFilename))
		
		oXML.loadXML(xy)
		oHTTP.open("POST","https://www.multitrac.com/pcws/pcode.asmx", .F.)
		
		oHTTP.send(oXML)
		
		retDoc = oHTTP.responseXML
		
		errorCodeList = retDoc.getElementsByTagName("ErrorCode")
		errorDescList = retDoc.getElementsByTagName("ErrorDesc")
		
		strErrorCode =  errorCodeList.Item(0).text
		
		IF VAL(ALLTRIM(strErrorCode)) = 0
			RETURN "File Posted"
		ELSE 			
			RETURN errorDescList.Item(0).text
		ENDIF 
		
	ENDPROC 
	
	PROCEDURE resetJob()
		
		this.strJobNoFS = "0"
		this.strClientIDFS = ""
		this.strExternalJobNoFS = ""
		this.strJobNameFS = ""
		this.strDivIDFS = ""
		this.strMailClassFS = ""
		this.strIMBFS = ""
		this.strByStoreFS = ""
		this.strEntryPointFS = ""
		this.strStartDateFS = ""
		this.strEndDateFS = ""
		this.strDropQtyFS = ""
		this.strVersionFS = ""
		this.strInhStartDateFS = ""
		this.strInhEndDateFS = ""
		this.strDropNameFS = ""
		this.strVendorIDFS = ""
		this.strOriginServiceTypeFS = ""
		this.strOriginZipFS = ""	
		this.strServiceOptionFS = ""
		this.strACSFS = ""
		this.strOELFS = ""
		
		this.strRetErrorDescFS = ""
		this.strRetErrorCodeFS = ""
		
		this.strRunXmlFile = ""
	    this.strLoginID = ""
	    this.strLoginPassword = ""
	    this.strMTJobNo = ""
	    this.strFileType = ""
	    this.strInputFile = ""
	    this.strOutputFile = ""
	    this.strClientIDField = ""
	    this.strClientIDStart = ""
	    this.strClientIDLength = ""
	    this.strNameField = ""
	    this.strNameStart = ""
	    this.strNameLength = ""
	    this.strAddress1Field = ""
	    this.strAddress1Start = ""
	    this.strAddress1Length = ""
	    this.strAddress2Field = ""
	    this.strAddress2Start = ""
	    this.strAddress2Length = "" 
	    this.strCityField = ""
	    this.strCityStart = ""
	    this.strCityLength = ""
	    this.strStateField = ""
	    this.strStateStart = ""
	    this.strStateLength = ""
	    this.strZipField = ""
	    this.strZipStart = ""
	    this.strZipLength = ""
	    this.strZip4Field = ""
	    this.strZip4Start = ""
	    this.strZip4Length = "" 
	    this.strDpbcField = ""
	    this.strDpbcStart = "" 
	    this.strDpbcLength = ""
	    this.strBarcodeField = ""
	    this.strBarcodeStart = ""
	    this.strBarcodeLength = ""
	    this.strVersionField = ""
	    this.strVersionStart = ""
	    this.strVersionLength = ""
	    this.strPhoneField = ""
	    this.strPhoneStart = ""
	    this.strPhoneLength = ""
	    this.strEmailField = ""
	    this.strEmailStart = ""
	    this.strEmailLength = ""
	    this.strStoreField = ""
	    this.strStoreStart = ""
	    this.strStoreLength = ""
	    this.strOptEndrsField = ""
	    this.strOptEndrsStart = ""
	    this.strOptEndrsLength = ""
	    this.strSeedFlagField = ""
	    this.strSeedFlagStart = ""
	    this.strSeedFlagLength = ""
	    this.strEntryField = ""
	    this.strEntryStart = ""
	    this.strEntryLength = ""
	    this.strPRateField = ""
	    this.strPRateStart = ""
	    this.strPRateLength = ""
	    this.strPCStartChar = ""
	    this.strPCEndChar = ""
		
		this.bolTestFS=.f.
		
		this.intStartDateCountFS = 0
		this.intEndDateCountFS = 0
		this.intDropQtyCountFS = 0
		this.intVersionCountFS = 0
		this.intInhStartDateCountFS = 0
		this.intInhEndDateCountFS = 0
		this.intDropNameCountFS = 0
		this.intVendorIDCountFS = 0
		
		DIMENSION this.aryJobDetailsFS(1,8)
		
		LOCAL i
		FOR i=1 TO ALEN(this.aryJobDetailsFS)
			this.aryJobDetailsFS(i) = ""
		ENDFOR 
		RELEASE i
		
		RETURN "reset"
	ENDPROC 
	
	HIDDEN PROCEDURE checkArray(intRowsNeeded)
	**********************************************************
	**	This procedure will re-dimension the Job Details	**
	**	array if there are not enough rows for the next		**
	**	iteration of a given job detail variable			**
	**********************************************************
		IF intRowsNeeded > ALEN(this.aryJobDetailsFS,1)
			DIMENSION this.aryJobDetailsFS(intRowsNeeded,9)
		ENDIF 
		
		RETURN 
	ENDPROC 
	
ENDDEFINE 