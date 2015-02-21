SET PROCEDURE TO "N:\foxstuff\sourceLink_resources\classes\PCode_Web_Service\pcodegetter_class.prg" ADDITIVE 
oPCodeGetter = CREATEOBJECT("pCodeGetterFS")

oPCodeGetter.setUserID("")
oPCodeGetter.setPassword("")

oPCodeGetter.setTest(.f.)

&& REQUIRED if you do not know client ID, division ID, or Vendor ID
&& Contact Marek Gozcal or Dan Browne
oPCodeGetter.setClientID(0) 
oPCodeGetter.setDivID(0) 

&& This is the job number and name
oPCodeGetter.setExternalJobNo("12345")
oPCodeGetter.setJobName("Testing New FS Job")

oPCodeGetter.setIMB(.t.)

oPCodeGetter.setByStore(.f.)
oPCodeGetter.setEntryPoint("29607")

&& The following methods are set for each drop you have 
&& each can be called more than once and the object saves
&& into an array for passing to the web service
&& -----------------------------------------------------------------------------------
&& no matter whether you are using full service or basic
&& First class must be set as 310 and standard as 311
oPCodeGetter.setMailClass("310")
oPCodeGetter.setStartDate(CTOD("05/21/2011"))
oPCodeGetter.setEndDate(CTOD("05/21/2011"))
oPCodeGetter.setInhStartDate(CTOD("05/26/2011"))
oPCodeGetter.setInhEndDate(CTOD("05/28/2011"))

oPCodeGetter.setDropQty(542)

oPCodeGetter.setVersion(1)
oPCodeGetter.setVendorID("0") 
oPCodeGetter.setDropName("Drop Name Test")
&& END DROP-SPECIFIC METHODS -----------------------------------------------------------------------

&& pass F for full service or B for basic
oPCodeGetter.setServiceOption("F")

&& ACS can be N, C or A
&& The client must be set up by Dan Browne to receive ACS in order to make this anything but N
oPCodeGetter.setACS("N")
oPCodeGetter.setOEL(.f.)

&& The client Must be set up by Dan Brown in order to utilize orgin tracking
&& this is tracking that will go on the BRE
oPCodeGetter.setOriginZip("")
oPCodeGetter.setOriginServiceType("")

&& if you are adding to an existing multitrac job,
&& you may set this to the existing multitrac job number
&& most cases will have this set to 0 to set up a new multitrac job
oPCodeGetter.setJobNo(0)

&& sends the job to Multitrac to be setup
oPCodeGetter.sendJob()

&& this will return the multitrac job number or 0 in case of an error
strMTJob = oPCodeGetter.getMTJob()
?oPCodeGetter.getMTJob()

&& if the getMTJob method returns 0, you may call these methods to 
&& retrieve error information
?oPCodeGetter.getErrorCode()
?oPCodeGetter.getErrorDesc()

&& The following methods are used to run the Multitrac Generator
&& any methods marked optional do not need to be called

&& OPTIONAL - you may leave the default or set your own XML file 
oPCodeGetter.setRunXMLFile()

&& REQUIRED - sets the job number received from the getMTJob function
oPCodeGetter.setMTJobNo(strMTJob)

&& REQUIRED - set the file type passed can be CSV or DBF
oPCodeGetter.setFileType("DBF")

&& REQUIRED - set the name of the input file (must have path included)
oPCodeGetter.setInputFile("N:\foxstuff\resources\classes\PCode_Web_Service\77454_16_encoded.dbf")

&& REQUIRED - set the name of the output file (must have path included)
oPCodeGetter.setOutputFile("N:\foxstuff\resources\classes\PCode_Web_Service\77454_16_encoded_wIMB.dbf")

&& REQUIRED - set ClientID field ClientID or name must be defined
oPCodeGetter.setClientIDField("UNIQUE_ID")

&& REQUIRED - set Name Field (Client ID or Name must be defined
oPCodeGetter.setNameField("NAME")

&& Optional - set address1 (primary address) field
oPCodeGetter.setAddress1Field("ADD1")

&& Optional - set address2 (secondary address) field
oPCodeGetter.setAddress2Field("ADD2")

&& REQUIRED - set City Field
oPCodeGetter.setCityField("CITY")

&& REQUIRED - set state field
oPCodeGetter.setStateField("STATE")

&& REQUIRED - You must set a zip, zip+4, dpbc field or a barcode field with postnet information
&& if you set barcode, you can omit zip, zip+4 and dpbc
&& if you omit barcode, you must set zip, zip+4 and dpbc
&& omitted fields do not have to have the set method called
&& set zip field
oPCodeGetter.setZipField("")

&& set plus4 field
oPCodeGetter.setZip4Field("")

&& set dpbc field
oPCodeGetter.setDpbcField("")

&& set Barcode Field
&& Barcode must be 5, 9, or 11 in length (check digit can not be included)
oPCodeGetter.setBarcodeField("BAR")

&& if this is not passed, it will assume that the version is 1
&& this is REQUIRED if you have more than one version in multitrac
oPCodeGetter.setVersionField("VERSION")

&& OPTIONAL - set phone field
oPCodeGetter.setPhoneField("")

&& OPTIONAL -set email field
oPCodeGetter.setEmailField("")

&& OPTIONAL - set store field
oPCodeGetter.setStoreField("")

&& OPTIONAL - set optional endorsement field
oPCodeGetter.setOptEndrsField("")

&& OPTIONAL - Set seed flag field
oPCodeGetter.setSeedFlagField("")

&& OPTIONAL - set Entry Field
&& RES is assumed if not entry is passed
&& REQUIRED - if you are doing drop shipping
oPCodeGetter.setEntryField("ENTRY")

&& OPTIONAL - set PRate Field
oPCodeGetter.setPrateField("")

&& OPTIONAL - set planet code start character
&& This will not typically be used at all
&&oPCodeGetter.setPCStartChar("")

&& OPTIONAL - set planet code end character
&& thiss will not typically be used at all
&&oPCodeGetter.setPCEndChar("")

&& This method will run the MT Generator
&& Returns a string with "ERROR" in it if there was an error
?oPCodeGetter.runMTGenerator()

&& This method will reset the object
oPCodeGetter.resetJob()