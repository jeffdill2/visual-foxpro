SET CPDIALOG OFF

**********************************************************************
*****IF UNEXPECTED ERROR OCCURS, HANDLED BY procUnexpectedErrorHandler
**********************************************************************
ON ERROR DO procUnexpectedErrorHandler WITH ERROR(), MESSAGE(), MESSAGE(1), LINENO(), PROGRAM()

**********************************************************************
*****************************************************VARIABLES DEFINED
**********************************************************************
nProgramStartTime = SECONDS()
strProcessingServer = "RED"
strSpectrumSavedInfoDatabase = "\\sulu\dev\foxstuff\spectrum_bofa_imac\templates\spectrum_bofa_imac_saved_info.dbf"
strProcessDateTime = STRTRAN(ALLTRIM(DTOC(DATE())),"/","") + "_" + STRTRAN(ALLTRIM(TIME()),":","")
strJobNumberContainer = "\\sulu\public\Traci\Spectrum\BofA_IMAC\Current_Job_Number\spectrum_bofa_imac_job_number.txt"
strCurrentJobNumber = FILETOSTR(ALLTRIM(strJobNumberContainer))
strClientDirectory = "\\" + strProcessingServer + "\data\spectrum\bofa\imac_matchback\"
strDriasiFTPDirectory = "\\ftp2\ftp\driasi\"
strIncomingFileDirectory = ALLTRIM(strClientDirectory) + "From_FTP\"
strArchiveDirectory = ALLTRIM(strClientDirectory) + "Archive\"
strRootJobDirectory = ALLTRIM(strClientDirectory) + ALLTRIM(strCurrentJobNumber) + "\"
strCurrentJobDirectory = ALLTRIM(strRootJobDirectory) + ALLTRIM(strProcessDateTime) + "\"
strCurrentInDirectory = ALLTRIM(strCurrentJobDirectory) + "In\"
strCurrentOutDirectory = ALLTRIM(strCurrentJobDirectory) + "Out\"

**********************************************************************
*********************************RUNNING funcInitialProcesses FUNCTION
**********************************************************************
strFileToProcessPGP = funcInitialProcesses(strCurrentJobNumber, strClientDirectory, strIncomingFileDirectory, strArchiveDirectory, strRootJobDirectory, strCurrentJobDirectory, strCurrentInDirectory, strCurrentOutDirectory, strDriasiFTPDirectory, strSpectrumSavedInfoDatabase)

**********************************************************************
*********************************RUNNING funcPrimaryProcesses FUNCTION
**********************************************************************
strOutputFileGPG = funcPrimaryProcesses(strCurrentJobNumber, strFileToProcessPGP, strCurrentInDirectory, strCurrentJobDirectory)

**********************************************************************
********************************RUNNING procBackendProcesses PROCEDURE
**********************************************************************
procBackendProcesses(strCurrentJobNumber, strOutputFileGPG, strDriasiFTPDirectory, nProgramStartTime, strSpectrumSavedInfoDatabase, strFileToProcessPGP)

**********************************************************************
***************CLOSING DOWN THE PROGRAM NOW THAT IT'S FINISHED RUNNING
**********************************************************************
RELEASE ALL
CLOSE ALL
CANCEL

**********************************************************************
****FUNCTION - DOES ALL INITIAL PROCESSING AND RETURNS FILE TO PROCESS
**********************************************************************
FUNCTION funcInitialProcesses(strCurrentJobNumber as String, ;
							  strClientDirectory as String, ;
							  strIncomingFileDirectory as String, ;
							  strArchiveDirectory as String, ;
							  strRootJobDirectory as String, ;
							  strCurrentJobDirectory as String, ;
							  strCurrentInDirectory as String, ;
							  strCurrentOutDirectory as String, ;
							  strDriasiFTPDirectory as String, ;
							  strSpectrumSavedInfoDatabase as String)

	NOTE: Checking to make sure all parameters successfully passed to funcInitialProcesses
	IF PARAMETERS()<>10 THEN
		strErrorDescription = "Not all parameters passed to 'funcInitialProcesses'. Something went wrong - double check all work!"
		procPlannedErrorHandler(strErrorDescription)
	ENDIF
	
	NOTE: Checking to make sure the job number seems to be properly formatted
	IF LEN(ALLTRIM(strCurrentJobNumber))<>5 THEN
		strErrorDescription = "The job number provided in '\\sulu\public\Traci\Spectrum\BofA_IMAC\Current_Job_Number\spectrum_bofa_imac_job_number.txt' is not properly formatted."
		procPlannedErrorHandler(strErrorDescription)
	ELSE
		FOR i = 1 TO 5
			IF !ISDIGIT(SUBSTR(ALLTRIM(strCurrentJobNumber),i,1)) THEN
				strErrorDescription = "The job number provided in '\\sulu\public\Traci\Spectrum\BofA_IMAC\Current_Job_Number\spectrum_bofa_imac_job_number.txt' is not properly formatted."
				procPlannedErrorHandler(strErrorDescription)
			ENDIF
		ENDFOR
		RELEASE i
	ENDIF
	
	NOTE: Building an array of any of the incoming files
	CD (strIncomingFileDirectory)
	ADIR(aryIncomingFiles,"*.txt.pgp")
	
	NOTE: Checking to make sure files really are there
	IF ADIR(aryIncomingFiles,"*.txt.pgp")=0 THEN
		strErrorDescription = "No files could be found in '\\" + strProcessingServer + "\data\spectrum\bofa\imac_matchback\from_ftp\'. Something went wrong in 'from FTP' transmission"
		procPlannedErrorHandler(strErrorDescription)
	ELSE
		NOTE: Saving the name of the file that we will process for this current go-round to a variable for later use.
		STORE aryIncomingFiles(1,1) TO strFileToProcessPGP
		
		NOTE: If more than one file is present, all but the first file will be moved back to the FTP so that the program can properly process the one file, and then PlanetWatch will pick up the rest of the files when it runs again.
		IF ALEN(aryIncomingFiles,1)>1 THEN
			CD c:\
			
			FOR i = 2 TO ALEN(aryIncomingFiles,1)
				spawn("cmd.exe /c move " + ALLTRIM(strIncomingFileDirectory) + ALLTRIM(aryIncomingFiles(i,1)) + " " + ALLTRIM(strDriasiFTPDirectory) + ALLTRIM(aryIncomingFiles(i,1)))
			ENDFOR
		ENDIF
	ENDIF
	
	NOTE: If the root job directory doesn't already exist, then we are going to make it - also, if there are any old root job directories sitting out there, we will move those to \Archive.
	IF !DIRECTORY(ALLTRIM(strRootJobDirectory)) THEN
		NOTE: If the root job directory doesn't exist, that's a good indicator that this is a new job ticket, therefore, we'll check to see if any old job directories are sitting out there. If they are we'll move them into the Archive to keep things as clean as possible.
		USE (strSpectrumSavedInfoDatabase) IN 0 ALIAS saved_info
		SELECT saved_info
		CD c:\
		
		1
		SCAN
			IF ALLTRIM(job_number)<>ALLTRIM(strCurrentJobNumber) AND DIRECTORY(STRTRAN(ALLTRIM(strRootJobDirectory),ALLTRIM(strCurrentJobNumber),ALLTRIM(job_number))) THEN
				NOTE: Using the Spawn program to execute this Command Line function to move a folder into Archive
				spawn("cmd.exe /c move " + ALLTRIM(strClientDirectory) + ALLTRIM(job_number) + " " + ALLTRIM(strArchiveDirectory) + ALLTRIM(job_number))
			ENDIF
		ENDSCAN

		CLOSE TABLES ALL

		NOTE: Creating the root job directory we need for the current job.
		MKDIR ALLTRIM(strRootJobDirectory)
	ENDIF

	NOTE: Making the current directory which has a date_time naming convention to keep them 100% unique, and a subsequent "In" and "Out" folder
	MKDIR ALLTRIM(strCurrentJobDirectory)
	MKDIR ALLTRIM(strCurrentInDirectory)
	MKDIR ALLTRIM(strCurrentOutDirectory)

	NOTE: Moving the file that will be processed to it's 'In' folder
	spawn("cmd.exe /c move " + ALLTRIM(strIncomingFileDirectory) + ALLTRIM(strFileToProcessPGP) + " " + ALLTRIM(strCurrentInDirectory) + ALLTRIM(strFileToProcessPGP))

	NOTE: Cleaning up memory
	RELEASE strClientDirectory
	RELEASE strIncomingFileDirectory
	RELEASE strArchiveDirectory
	RELEASE aryIncomingFiles

	NOTE: Returning the filename that will be processed throughout the rest of this program
	RETURN ALLTRIM(strFileToProcessPGP)
ENDFUNC

**********************************************************************
***********************FUNCTION - DOES ALL OF THE MAIN FILE PROCESSING
**********************************************************************
FUNCTION funcPrimaryProcesses(strCurrentJobNumber as String, ;
							   strFileToProcessPGP as String, ;
							   strCurrentInDirectory as String, ;
							   strCurrentJobDirectory as String)

	NOTE: Checking to make sure all parameters successfully passed to funcPrimaryProcesses
	IF PARAMETERS()<>4 THEN
		strErrorDescription = "Not all parameters passed to 'funcPrimaryProcesses'. Something went wrong - double check all work!"
		procPlannedErrorHandler(strErrorDescription)
	ENDIF

	NOTE: Defining some variables
	strFileToProcessTXT = ALLTRIM(strCurrentInDirectory) + STRTRAN(UPPER(ALLTRIM(strFileToProcessPGP)),".TXT.PGP",".TXT")
	strFileToProcessDBF = ALLTRIM(strCurrentJobDirectory) + STRTRAN(UPPER(ALLTRIM(STRTRAN(UPPER(ALLTRIM(JUSTFNAME(strFileToProcessTXT))),".TXT","_Data.DBF"))),"BA_LINK",ALLTRIM(strCurrentJobNumber) + "_")
	strMatchedDataDBF = STRTRAN(UPPER(ALLTRIM(strFileToProcessDBF)),".DBF","_Matched.dbf")
	strOutputFileTXT = ALLTRIM(strCurrentOutDirectory) + STRTRAN(UPPER(ALLTRIM(STRTRAN(UPPER(ALLTRIM(JUSTFNAME(strFileToProcessTXT))),"LINKIN","LINKOUT"))),"LINK","Link")
	strOutputFileGPG = STRTRAN(ALLTRIM(strOutputFileTXT),".TXT",".TXT.GPG")
	strSpectrumIMACUniverse = "\\" + strProcessingServer + "\data\Spectrum\bofa\imac_matchback\Universe\Spectrum_BofA_IMAC_UNIVERSE.dbf"

	NOTE: Decrypting the strFileToProcess file
	strDecryptOnCmdLine = 'cmd.exe /c echo collinsdirect|\\' + strProcessingServer + '\c$\gnupg\gpg.exe --home \\' + strProcessingServer + '\c$\gnupg\ --passphrase-fd 0 --decrypt-files "' + ALLTRIM(strCurrentInDirectory) + ALLTRIM(strFileToProcessPGP) + '"'
	spawn(strDecryptOnCmdLine)

	NOTE: Creating a table that we can append the current input file into, so that we can perform the match, etc.
	CREATE TABLE ALLTRIM(strFileToProcessDBF) (finder_num c(13), certif_num c(10))
	CLOSE TABLES ALL

	NOTE: Bringing in the data from the IN file
	USE (strFileToProcessDBF) IN 0 ALIAS process_dbf
	SELECT process_dbf

	NOTE: Making sure the PGP file has time to be decrypted before we continue, because we need the subsequent TXT file to append into the DBF
	DO WHILE !FILE(ALLTRIM(strFileToProcessTXT))
		WAIT "" NOWAIT
	ENDDO

	NOTE: Appending in the TXT file that was decrypted from the PGP file
	APPEND FROM (strFileToProcessTXT) TYPE SDF

	NOTE: Performing the match to the Spectrum Universe
	SELECT a.finder_num, ;
		   a.certif_num, ;
		   b.first, ;
		   b.last, ;
		   b.add1, ;
		   b.add2, ;
		   b.city, ;
		   b.st, ;
		   b.zip, ;
		   b.loan_num, ;
		   b.res ;
		   FROM process_dbf as a LEFT OUTER JOIN ALLTRIM(strSpectrumIMACUniverse) as b ON ALLTRIM(a.finder_num)==ALLTRIM(b.driasi_cd) INTO TABLE ALLTRIM(strMatchedDataDBF)

	CLOSE TABLES ALL

	NOTE: Creating the output TXT file that will be GPG'd and then placed on the FTP
	USE (strMatchedDataDBF) IN 0 ALIAS matched
	SELECT matched
	
	oOutputTXTFile = FCREATE(ALLTRIM(strOutputFileTXT))

	1
	SCAN
		FPUTS(oOutputTXTFile, PADR(ALLTRIM(finder_num),13," ") + ;
							  PADR(ALLTRIM(certif_num),10," ") + ;
							  PADR(ALLTRIM(first),20," ") + ;
							  PADR(ALLTRIM(last),20," ") + ;
							  PADR(ALLTRIM(add1),40," ") + ;
							  PADR(ALLTRIM(add2),40," ") + ;
							  PADR(ALLTRIM(city),20," ") + ;
							  PADR(ALLTRIM(st),2," ") + ;
							  PADR(STRTRAN(ALLTRIM(zip),"-",""),9," ") + ;
							  PADR(ALLTRIM(loan_num),9," ") + ;
							  PADR(ALLTRIM(res),14," ") + ;
							  SPACE(19) + ;
							  SPACE(11) ;
							  )
	ENDSCAN
	FCLOSE(oOutputTXTFile)
	RELEASE oOutputTXTFile

	CLOSE TABLES ALL

	NOTE: Encrypting the output file with Driasipgp key
	strEncryptOnCmdLine = '\\' + strProcessingServer + '\c$\gnupg\gpg.exe --home \\' + strProcessingServer + '\c$\gnupg\ --recipient 7CB3BD09 --output "' + ALLTRIM(strOutputFileGPG) + '" --encrypt "' + ALLTRIM(strOutputFileTXT) + '"'
	spawn(strEncryptOnCmdLine)

	NOTE: Cleaning up memory
	RELEASE strMatchedDataDBF
	RELEASE strFileToProcessTXT
	RELEASE strFileToProcessDBF
	RELEASE strEncryptOnCmdLine
	RELEASE strDecryptOnCmdLine
	RELEASE strOutputFileTXT
	RELEASE strCurrentInDirectory
	RELEASE strSpectrumUniverse

	NOTE: Returning the output PGP file to be placed onto the FTP
	RETURN ALLTRIM(strOutputFileGPG)
ENDFUNC

**********************************************************************
***********PROCEDURE - DOES ALL BACKEND PROCESSING, FILE POSTING, ETC.
**********************************************************************
PROCEDURE procBackendProcesses(strCurrentJobNumber as String, ;
							   strOutputFileGPG as String, ;
							   strDriasiFTPDirectory as String, ;
							   nProgramStartTime as Number, ;
							   strSpectrumSavedInfoDatabase as String, ;
							   strFileToProcessPGP as String)

	NOTE: Checking to make sure all parameters successfully passed to procBackendProcesses
	IF PARAMETERS()<>6 THEN
		strErrorDescription = "Not all parameters passed to 'procBackendProcesses'. Something went wrong - double check all work!"
		procPlannedErrorHandler(strErrorDescription)
	ENDIF

	NOTE: Placing the output GPG file on the Driasi FTP for retrieval
	spawn("cmd.exe /c copy " + ALLTRIM(strOutputFileGPG) + " " + ALLTRIM(strDriasiFTPDirectory) + ALLTRIM(JUSTFNAME(strOutputFileGPG)))

	NOTE: Sending email alert to let everyone know that a file was just processed/posted
	strEmailSubject = "Spectrum - BofA IMAC - matchback processed"
	strEmailMessage = "Spectrum - BofA IMAC - matchback has been processed." + CHR(13) + CHR(13) + ;
					  "The following file has been uploaded to the Driasi FTP:" + CHR(13) + ;
					  ALLTRIM(JUSTFNAME(strOutputFileGPG))
	strEmailFrom = "automation@sourcelink.com"
	strEmailTo = "ftpnotification@driasi.com,serazzis@driasi.com"
	strEmailCC = "thunter@sourcelink.com,mteague@sourcelink.com,jdill@sourcelink.com"

	procSendEmail(strEmailSubject, strEmailMessage, strEmailFrom, strEmailTo, strEmailCC)

	NOTE: Figuring out how long it took the program to run all the way through
	nProgramEndTime = SECONDS()
	strTotalProcessTime = ALLTRIM(STR((nProgramEndTime - nProgramStartTime)/60))

	NOTE: Updating the saved info .DBF
	USE (strSpectrumSavedInfoDatabase) IN 0 ALIAS saved_info
	SELECT saved_info
	
	APPEND BLANK
	GOTO BOTTOM
	replace job_number WITH ALLTRIM(strCurrentJobNumber), ;
			date_time WITH DATETIME(), ;
			in_file WITH UPPER(ALLTRIM(strFileToProcessPGP)), ;
			out_file WITH UPPER(ALLTRIM(JUSTFNAME(strOutputFileGPG)))
			
			CLOSE TABLES ALL
	
	NOTE: Sending email alert just to Jeff Dill to give some good information about what has just run.
	strEmailSubject = "Spectrum - BofA IMAC - matchback processed"
	strEmailMessage = "Spectrum - BofA IMAC - matchback has been processed." + CHR(13) + CHR(13) + ;
					  "Here is the job info :" + CHR(13) + ;
					  "Job Number :  " + ALLTRIM(strCurrentJobNumber) + CHR(13) + ;
					  "Input File :  " + ALLTRIM(strFileToProcessPGP) + CHR(13) + ;
					  "Output File :  " + ALLTRIM(JUSTFNAME(strOutputFileGPG))
	strEmailFrom = "automation@sourcelink.com"
	strEmailTo = "jdill@sourcelink.com"
	strEmailCC = "jdill@sourcelink.com"
	
	procSendEmail(strEmailSubject, strEmailMessage, strEmailFrom, strEmailTo, strEmailCC)
	
	NOTE: Cleaning up memory
	RELEASE strSpectrumSavedInfoDatabase
	RELEASE strCurrentJobNumber
	RELEASE strTotalProcessTime
	RELEASE strFileToProcessPGP
	RELEASE strOutputFileGPG
	RELEASE nProgramStartTime
	RELEASE nProgramEndTime
	RELEASE strDriasiFTPDirectory
ENDPROC

**********************************************************************
***************************PROCEDURE - DOES ANY EMAILS THAT ARE NEEDED
**********************************************************************
PROCEDURE procSendEmail(strEmailSubject as String, ;
						strEmailMessage as String, ;
						strEmailFrom as String, ;
						strEmailTo as String, ;
						strEmailCC as String)

	NOTE: Checking to make sure all parameters successfully passed to procSendEmail
	IF PARAMETERS()<>5 THEN
		strErrorDescription = "Not all parameters passed to 'procSendEmail'. Something went wrong - double check all work!"
		procPlannedErrorHandler(strErrorDescription)
	ENDIF

	NOTE: Sending the email
	oEmail = CREATEOBJECT("SMTPControl.SMTP")
	oEmail.Server = "mail.carolina.sourcelink.com"
	oEmail.MailFrom = ALLTRIM(strEmailFrom)
	oEmail.SendTo = ALLTRIM(strEmailTo)
	oEmail.CC = ALLTRIM(strEmailCC)
	oEmail.MessageSubject = ALLTRIM(strEmailSubject)
	oEmail.MessageText = ALLTRIM(strEmailMessage)
	oEmail.connect()
	
	NOTE: Cleaning up memory
	RELEASE oEmail
	RELEASE strEmailSubject
	RELEASE strEmailMessage
	RELEASE strEmailFrom
	RELEASE strEmailTo
	RELEASE strEmailCC
ENDPROC

**********************************************************************
***********PROCEDURE - HANDLES ANY ERRORS THAT ARE SET UP TO BE CAUGHT
**********************************************************************
PROCEDURE procPlannedErrorHandler(strErrorDescription as String)
	strErrorLogFile = "\\sulu\dev\foxstuff\spectrum_bofa_imac\templates\spectrum_bofa_imac_error_log_file.log"
	strErrorLogLineItem = "--- " + ALLTRIM(TTOC(DATETIME())) + " - " + ALLTRIM(strErrorDescription)
	
	NOTE: Adding the error line item to the error log file.
	STRTOFILE(CHR(13) + CHR(10) + ALLTRIM(strErrorLogLineItem),ALLTRIM(strErrorLogFile),.T.)
	
	NOTE: Sending an email alerting me that an error has occured in the program - the reason that I am setting up the email object individually for this procedure as opposed to calling the Send Email function is that if the error occurs in the Send Email function, that's an issue!
	oEmail = CREATEOBJECT("SMTPControl.SMTP")
	oEmail.Server = "mail.carolina.sourcelink.com"
	oEmail.MailFrom = "automation@sourcelink.com"
	oEmail.SendTo = "jdill@sourcelink.com"
	oEmail.MessageSubject = "Error - Spectrum BofA IMAC"
	oEmail.MessageText = "Log file has been updated with error info"
	oEmail.connect()
	
	NOTE: Making sure all memory is reclaimed before cancelling the program.
	RELEASE ALL
	CLOSE ALL
	CANCEL
ENDPROC
	
**********************************************************************
*******PROCEDURE - HANDLES ANY ERRORS THAT ARE NOT SET UP TO BE CAUGHT
**********************************************************************
PROCEDURE procUnexpectedErrorHandler(nErrorNum as Number, ;
									 strErrorMsg as String, ;
									 strCodeInError as String, ;
									 nLineNumOfError as Number, ;
									 strProgramWithError as String)

	strErrorLogFile = "\\sulu\dev\foxstuff\spectrum_bofa_imac\templates\spectrum_bofa_imac_error_log_file.log"
	
	NOTE: Creating the line item of the error to place in the log file.
	strErrorLogLineItem = "--- " + ALLTRIM(TTOC(DATETIME())) + ;
						  " - Error number:" + ALLTRIM(STR(nErrorNum)) + ;
						  " - Error message:" + ALLTRIM(strErrorMsg) + ;
						  " - Code in error:" + ALLTRIM(strCodeInError) + ;
						  " - Line number of error:" + ALLTRIM(STR(nLineNumOfError)) + ;
						  " - Program with error:" + ALLTRIM(strProgramWithError)
	
	NOTE: Adding the error line item to the error log file.
	STRTOFILE(CHR(13) + CHR(10) + ALLTRIM(strErrorLogLineItem),ALLTRIM(strErrorLogFile),.T.)
	
	NOTE: Sending an email alerting me that an error has occured in the program - the reason that I am setting up the email object individually for this procedure as opposed to calling the Send Email function is that if the error occurs in the Send Email function, that's an issue!
	oEmail = CREATEOBJECT("SMTPControl.SMTP")
	oEmail.Server = "mail.carolina.sourcelink.com"
	oEmail.MailFrom = "automation@sourcelink.com"
	oEmail.SendTo = "jdill@sourcelink.com"
	oEmail.MessageSubject = "Error - Spectrum BofA IMAC"
	oEmail.MessageText = "Log file has been updated with error info"
	oEmail.connect()
	
	NOTE: Making sure all memory is reclaimed before cancelling the program.
	RELEASE ALL
	CLOSE ALL
	CANCEL
ENDPROC

**********************************************************************
********PROCEDURE - THIS IS THE CODE FOR THE SPAWN PROGRAM THAT WE USE
**********************************************************************
PROCEDURE spawn(strProcess as String)
	#DEFINE NORMAL_PRIORITY_CLASS 32
	#DEFINE IDLE_PRIORITY_CLASS 64
	#DEFINE HIGH_PRIORITY_CLASS 128
	#DEFINE REALTIME_PRIORITY_CLASS 1600

	NOTE: Return code from WaitForSingleObject() if it timed out.
	#DEFINE WAIT_TIMEOUT 0x00000102

	NOTE: This controls how long, in milli secconds, WaitForSingleObject() waits before it times out. Change this to suit your preferences.
	#DEFINE WAIT_INTERVAL 200

	DECLARE INTEGER CreateProcess IN kernel32.DLL ;
	INTEGER lpApplicationName, ;
	STRING lpCommandLine, ;
	INTEGER lpProcessAttributes, ;
	INTEGER lpThreadAttributes, ;
	INTEGER bInheritHandles, ;
	INTEGER dwCreationFlags, ;
	INTEGER lpEnvironment, ;
	INTEGER lpCurrentDirectory, ;
	STRING @lpStartupInfo, ;
	STRING @lpProcessInformation

	DECLARE INTEGER WaitForSingleObject IN kernel32.DLL ;
	INTEGER hHandle, INTEGER dwMilliseconds

	DECLARE INTEGER CloseHandle IN kernel32.DLL ;
	INTEGER hObject

	DECLARE INTEGER GetLastError IN kernel32.DLL

	start = long2str(68) + REPLICATE(CHR(0), 64)
	process_info = REPLICATE(CHR(0), 16)
	File2Run = strProcess + CHR(0)
	RetCode = CreateProcess(0, File2Run, 0, 0, 1, ;
	NORMAL_PRIORITY_CLASS, 0, 0, @start, @process_info)

	hProcess = str2long(SUBSTR(process_info, 1, 4))

	DO WHILE .T.
		IF WaitForSingleObject(hProcess, WAIT_INTERVAL) != WAIT_TIMEOUT
			EXIT
		ELSE
			DOEVENTS
		ENDIF
	ENDDO

	RetCode = CloseHandle(hProcess)
	RETURN
ENDPROC

**********************************************************************
**********************FUNCTION - ONLY ADDED BECAUSE IT'S USED BY SPAWN
**********************************************************************
FUNCTION long2str(m.longval as String)
	PRIVATE i, m.retstr
	m.retstr = ""
	FOR i = 24 TO 0 STEP -8
		m.retstr = CHR(INT(m.longval/(2^i))) + m.retstr
		m.longval = MOD(m.longval, (2^i))
	NEXT
	RETURN m.retstr
ENDFUNC

**********************************************************************
**********************FUNCTION - ONLY ADDED BECAUSE IT'S USED BY SPAWN
**********************************************************************
FUNCTION str2long(m.longstr as Long)
	PRIVATE i, m.retval
	m.retval = 0
	FOR i = 0 TO 24 STEP 8
		m.retval = m.retval + (ASC(m.longstr) * (2^i))
		m.longstr = RIGHT(m.longstr, LEN(m.longstr) - 1)
	NEXT
	RETURN m.retval
ENDFUNC