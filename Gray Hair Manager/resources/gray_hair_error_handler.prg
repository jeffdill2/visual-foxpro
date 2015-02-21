PROCEDURE procErrorHandler
	LPARAMETERS	nErrorNum as Number, ;
				strErrorMsg as String, ;
				strCodeInError as String, ;
				nLineNumOfError as Number, ;
				strProgramWithError as String, ;
				strEmailServer as String, ;
				strProgrammer as String

	NOTE: Creating the line item of the error to place in the log file.
	strErrorLogLineItem = "- - - " + ALLTRIM(TTOC(DATETIME())) + CHR(13) + ;
						  "- Error number:" + ALLTRIM(STR(nErrorNum)) + CHR(13) + ;
						  "- Error message:" + ALLTRIM(strErrorMsg) + CHR(13) + ;
						  "- Code in error:" + ALLTRIM(strCodeInError) + CHR(13) + ;
						  "- Line number of error:" + ALLTRIM(STR(nLineNumOfError)) + CHR(13) + ;
						  "- Program with error:" + ALLTRIM(strProgramWithError)

	NOTE: Sending an email alerting me that an error has occured in the program - the reason that I am setting up the email object individually for this procedure as opposed to calling the Send Email function is that if the error occurs in the Send Email function, that's an issue!
	oEmail = CREATEOBJECT("OSSMTP.SMTPSession")
	oEmail.Server = strEmailServer
	oEmail.MailFrom = "noreply@cgraphics.com"
	oEmail.SendTo = strProgrammer + "@cgraphics.com"
	oEmail.MessageSubject = "Gray Hair Manager - Error Has Occured"
	oEmail.MessageText = strErrorLogLineItem
	oEmail.SendEmail()
		
	NOTE: Making sure all memory is reclaimed before cancelling the program.
	RELEASE ALL
	CLOSE ALL
	QUIT
ENDPROC