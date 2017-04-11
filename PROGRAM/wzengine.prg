#include "..\Include\wizard.h"

#DEFINE MAC_BUILD 0

EXTERNAL ARRAY aSortArray

**************************************************************************
**************************************************************************
* Summary of Classes:
*
*	WizEngine - base Wizard Engine
*
*		ReadSettings - reads preference settings
*		WriteSettings - writes preference settings
*		Error - common error handling
*		ErrorCleanup - stub method for use with specific wizard called by Error
*		Alert - displays MessageBox alert
*		ProcessOutput - stub for specific wizard to process output
*		AddTherm - use to get a thermometer
*		AddHandleToCloseList - adds a LLFF handle to list to be closed during destroy
*		AddAliasToPreservedList	- adds alias to list of aliases to preserve during destroy
*		CloseUnpreservedAliases - close all aliases that aren't on the preserved list
*		CloseHandles - close all handles that are on the close list
*		GetOS - returns operating system code (see #DEFINES)
*
*-  These are used by VFP for the Macintosh if MAC_BUILD is # 0
*-
*-      SetLibrary		-- open an API library
*-		LocateApp		-- returns path to app of a given signature
*-		PutPref			-- write a value to an INI-style file
*-		GetPref			-- retrieve a value from an INI-style file
*-		FindINI			-- search for a file (really, any file)
*-		GetINISection	-- populate an array with all items within a given INI file section
*-		FxStripLF		-- strip linefeeds from a file
*
*
*	WizEngineAll - subclass of WizEngine. Has common methods to use
*	
*		AScanner - does ASCAN on column of array
*		AColScan - does ASCAN on column of array
*		InsaItem - insert item into array
*		DelaItem - delete item from array
*		SaveOutFile - asks for output file name and checks if it can be used
*		JustPath - returns path of file name
*		JustStem - returns stem of file name (name only with no extension)
*		JustFName - returns file name
*		ForceExt - forces file to have certain extension
*		AddBs - adds backslash (colon for Macs) to file path if needed
*		GetStyle - returns a FoxPro font style code based on Bold, Italic, Underline, etc. 
*		AddCdxTag - adds a tag to CDX index 
*		GetTagExpr - used by AddCdxTag to get tag expression for specific data types
*		GetDBCAlias - returns DBC alias
*		GetFullTagExpr - returns full tag expression for sort array
*		CheckDBCTag - tests if DBC is opened non-exclusive so tag cannot be made 
**************************************************************************
**************************************************************************

******************************************************************************
* Used by GetOS and other methods
******************************************************************************
* Operating System codes
#DEFINE	OS_W32S				1
#DEFINE	OS_NT				2
#DEFINE	OS_WIN95			3
#DEFINE	OS_MAC				4
#DEFINE	OS_DOS				5
#DEFINE	OS_UNIX				6

******************************************************************************
* Used by AddCdxTag/GetTagExpr method
******************************************************************************

* Data types
#DEFINE DT_INTEGER 	'I'
#DEFINE DT_NUM   	'N'
#DEFINE DT_FLOAT 	'F'
#DEFINE DT_LOGIC 	'L'
#IFNDEF DT_MEMO
	#DEFINE DT_MEMO  'M'
#ENDIF
#DEFINE DT_GEN   	'G'
#DEFINE DT_CHAR  	'C'
#DEFINE DT_VARCHAR	'V'
#DEFINE DT_DATE  	'D'
#DEFINE DT_DATETIME	'T'
#DEFINE DT_CURRENCY	'Y'
#DEFINE DT_DOUBLE	'B'
#DEFINE C_FILEUSE_LOC	"File in use. Could not create index."
#DEFINE C_BADPARMS_LOC	"Incorrect number of parameters passed to indexing routine."
#DEFINE C_EXCLUSIVE_LOC	"EXCLUSIVE"

******************************************************************************
* Used by SaveOutFile method
******************************************************************************
#DEFINE C_FILEUSE2_LOC		"File is in use. Please select another."

******************************************************************************
* Used by Error method
******************************************************************************
*#define ERRORTITLE_LOC		"Upsizing Wizard"
#define NORUNTIME_LOC ;
	"Microsoft Visual FoxPro Wizards require the Standard or Professional version."
	
#define ERRORMESSAGE_LOC ;
	"Error #" + alltrim(str(m.nError)) + " in " + m.cMethod + ;
	" (" + alltrim(str(m.nLine)) + "): " + m.cMessage

* The result of the above message will look like this:
*
*		Error #1 in WIZTEMPLATE.INIT (14): File does not exist.

#define MB_ICONEXCLAMATION		48
#define MB_ABORTRETRYIGNORE		2
#define MB_OK					0

******************************************************************************
* Used by Alert method
******************************************************************************
* #define ALERTTITLE_LOC		"Upsizing Wizard"
		
******************************************************************************
* Used by ProcessOutput method
******************************************************************************
#define NOPROCESS_LOC		"No process defined."


******************************************************************************
DEFINE CLASS WizEngine AS custom
******************************************************************************
*!*		name = "wzengine"	&& this change broke object references in several wizards

    lQuiet = .F.
    ErrorMessage = ''


	* Wizard Globals
	iHelpContextID = 0		&& used as default
	cWizName = ""			&& wizard name	(e.g., Group/Total report)
	cWizClass = ""			&& wizard class	(e.g., Report)
	nTimeStamp = 0			&& time stamp
	cWizVersion = "9.0"		&& version number
	cVisFoxVersion = PADR(VAL(SUBSTR(VERSION(),ATC("FOXPRO",VERSION()) + 7)),1)	&& version number (e.g., '5')
	cWizTitle = ""			&& output document title
	cOutFile = ""			&& output file name
	cDBCName = ""			&& DBC name
	cDBCAlias = ""			&& DBC Alias name
	cDBCTable = ""			&& DBC Table name
	nWizWorkArea = 0		&& table workarea
	cWizAlias = ""			&& table alias
	nWizAction = 1			&& output action	(e.g., Save and modify)
	cWizOptions = ""		&& wizard options (e.g., NOSCRN)
	lSortAscend = .T.		&& sort tag
	lHasSortTag = .F.		&& uses index tag
	SetErrorOff = .F.		&& bypass normal Error handling
	HadError = .f.			&& error occurred
	iError = -1				&& error number
	cMessage = ''			&& error message
	lCancelled = .F.		&& was the wizard cancelled?
	
	lHasNoTask = .f.		&& whether wizard should perform output task
	lInitValue = .t.		&& used to prevent instantiation of the wizard
							&& don't modify directly--return desired value from Init2
	nCurrentOS = 0			&& operating system code
	lForceExclusive = .T.	&& force SET EXCLUSIVE ON (needed for DBC operations)
	lUseNative = !EMPTY(OS(2))	&& whether to use native product functions for file handling
	
	* This member is passed to Init by WizardTemplate when the oEngine object is 
	* created. The member contains the name of the procedure to RETURN TO
	* when an error occurs in the engine. Typical settings would be
	* WIZARD (return to Wizard.app) or MASTER (return to Command window).
	cReturnToProc = 'MASTER'
	                       
	DIMENSION aSettings[1,2]	&& array of settings to save as prefs
	aSettings = ""			
	DIMENSION aDBFList[1]		&& array of tables for table picker
	aDBFList = ""
	DIMENSION aWizTables[1,1]	&& array of tables used
	aWizTables = ""
	DIMENSION aWizFields[1,1]	&& array of selected fields
	aWizFields = ""
	DIMENSION aWizFList[1,1]	&& array of all fields (nx7 -- AFIELDS)
	aWizFList = ""
	DIMENSION aWizLabels[1,1]	&& array of selected english field labels
	aWizLabels = ""
	DIMENSION aWizSorts[1,1]	&& array of sort fields
	aWizSorts = ""
	DIMENSION aCalcFields[1,1]	&& array of Calculated fields
	aCalcFields = ""
	
	* This array is used to store the list of aliases at startup. During Destroy,
	* if an alias is not in this array, it is closed.
	dimension aAliasesToPreserve[1, 2]
	
	* This array is used to store a list of LLFF handles to be closed during Destroy.
	dimension aHandlesToClose[1]
	aHandlesToClose = -1
	
	* This array is used to store a list of libraries to be released during Destroy.
	dimension aReleaseLibraryList[1]
	aReleaseLibraryList = ''
	
	* This array is used to store environment settings saved during Setup (called by Init)
	* and restored in Cleanup (called by Destroy).
	dimension aEnvironment[1]
	
	procedure AddHandleToCloseList
		* This procedure is used to add a LLFF handle to the this.aHandlesToClose list.
		* These handles are closed during Destroy.
		parameters iHandle
		if ascan(this.aHandlesToClose, m.iHandle) = 0
			if this.aHandlesToClose[1] <> -1
				dimension this.aHandlesToClose[alen(this.aHandlesToClose) + 1]
			endif
			this.aHandlesToClose[alen(this.aHandlesToClose)] = m.iHandle
		endif
	endproc

	procedure AddAliasToPreservedList
		* This procedure is used to add an alias to the this.aAliasesToPreserve list.
		* These aliases are left open during Destroy.
		parameters cAlias
		local cExact, i
		m.cExact = set('exact')
		set exact on
		if ascan(this.aAliasesToPreserve, upper(m.cAlias)) = 0
			m.i = alen(this.aAliasesToPreserve, 1)
			if .not. empty(this.aAliasesToPreserve[1, 1])
				m.i = m.i + 1
				dimension this.aAliasesToPreserve[m.i, alen(this.aAliasesToPreserve, 2)]
			endif
			this.aAliasesToPreserve[m.i, 1] = upper(m.cAlias)
			this.aAliasesToPreserve[m.i, 2] = select(upper(m.cAlias))
		endif
	endproc
	
	procedure CloseUnpreservedAliases
		* This procedure closes aliases which were not open during Init
		* (i.e. are not found in this.aAliasesToPreserve)

		local i, cExact
		local array aAliasesInUse[1, 2]
		
		if .not. empty(aused(aAliasesInUse))
			m.cExact = set('exact')
			set exact on
			for m.i = 1 to alen(aAliasesInUse, 1)
				if ascan(this.aAliasesToPreserve, aAliasesInUse[m.i, 1]) = 0
					use in (aAliasesInUse[m.i, 1])
				endif
			endfor
			set exact &cExact
		endif
	endproc

	procedure CloseHandles
		* This procedure closes low-level file handles store in this.aHandlesToClose

		local i

		if this.aHandlesToClose[1] <> -1
			for m.i = 1 to alen(this.aHandlesToClose)
				=fclose(this.aHandlesToClose[m.i])
			endfor
		endif
	endproc
	
	procedure AddLibraryToReleaseList
		parameters cLibrary
		if empty(ascan(this.aReleaseLibraryList, m.cLibrary))
			if .not. empty(this.aReleaseLibraryList[1])
				dimension this.aReleaseLibraryList[alen(this.aReleaseLibraryList) + 1]
			endif
			this.aReleaseLibraryList[alen(this.aReleaseLibraryList)] = m.cLibrary
		endif
	endproc
	
	procedure ReleaseLibraries
		* This procedure releases the libraries listed in aReleaseLibraryList[]
		local i
		for m.i = 1 to alen(this.aReleaseLibraryList)
			if not empty(this.aReleaseLibraryList[m.i]) .and. ;
				upper(this.aReleaseLibraryList[m.i]) $ upper(set('library'))
				release library (this.aReleaseLibraryList[m.i])
			endif
		endfor
	endproc
	
	procedure Destroy
		this.Cleanup
	endproc
	
	procedure Init
		* cReturnToProc is the procedure to return to 
		* if an error occurs in the engine. cProcedure is the
		* SET('PROCEDURE') setting to restore in cleanup.
		parameters cReturnToProc, cProcedure
		
		dimension this.aEnvironment[36,1]
		this.aEnvironment[11,1] = 0

* Coment JEI RKR 2005.03.30
*!*			IF .not. C_DEBUG
*!*				IF WEXIST("Visual Foxpro Debugger")
*!*					HIDE WINDOW "Visual FoxPro Debugger"
*!*				ELSE
*!*					* Using FoxPro frame
*!*					IF WVISIBLE("Watch")
*!*						this.aEnvironment[11,1] = 2^0
*!*						RELEASE WINDOW Watch
*!*					ENDIF
*!*					IF WVISIBLE("Locals")
*!*						this.aEnvironment[11,1] = this.aEnvironment[11,1] + 2^1				
*!*						RELEASE WINDOW Locals
*!*					ENDIF
*!*					IF WVISIBLE("Call Stack")
*!*						this.aEnvironment[11,1] = this.aEnvironment[11,1] + 2^2
*!*						RELEASE WINDOW Call Stack
*!*					ENDIF
*!*					IF WVISIBLE("Debug Output")
*!*						this.aEnvironment[11,1] = this.aEnvironment[11,1] + 2^3	
*!*						RELEASE WINDOW debug output
*!*					ENDIF						
*!*					IF WVISIBLE("Trace")
*!*						this.aEnvironment[11,1] = this.aEnvironment[11,1] + 2^4
*!*						RELEASE WINDOW trace
*!*					ENDIF						
*!*				ENDIF
*!*			ENDIF
		
		if used('THIS') .or. used('THISFORM') .or. used('THISFORMSET')
			clear typeahead
			if This.lQuiet or MessageBox(C_BADALIAS_LOC, MB_YESNO, ALERTTITLE_LOC) = IDYES
				* close the file and continue
				if used('this')
					use in this
				endif
				if used('thisform')
					use in thisform
				endif
				if used('thisformset')
					use in thisformset
				endif
			else
				return .f.
			endif
		endif
		
		this.cReturnToProc = iif(!empty(m.cReturnToProc), m.cReturnToProc, '')

		this.Setup(m.cProcedure)
		
		this.lInitValue = this.Init2()

		if .not. this.lInitValue
			this.Cleanup
			return .f.
			if empty(this.cReturnToProc)
				return .f.
			endif
		endif
	endproc
	
	procedure Init2
		* This is a stub for the Init2 method which may be created in
		* subclass engines. The subclass engine may RETURN .F. to prevent
		* the instantiation of the class.
	endproc
	
	procedure Setup
		parameters cProcedure
		clear program
		
		this.aEnvironment[1,1] = SET("TALK")		
		SET TALK OFF
		
*** DH 2015-08-20: handle cProcedure not passed to Init, meaning it's .F. here
***		if parameters() = 0
		if vartype(m.cProcedure) <> 'C'
			m.cProcedure = set('procedure')
		else
			m.cProcedure = iif(empty(m.cProcedure), '', m.cProcedure)
		endif
		
		this.aEnvironment[6,1] = set("exclusive")
		IF this.lForceExclusive
			set exclusive on
		ENDIF

		this.aEnvironment[2,1] =  set("step")
		this.aEnvironment[32,1] = on('escape')
		this.aEnvironment[18,1] = set('escape')
		push key clear
		
		this.aEnvironment[10,1] = set("trbetween")
		
		if .not. C_DEBUG
			set step off
			on escape
			set escape off
***			set skip of bar _mpr_suspend of _mprog .T.
***			set skip of popup _mtools .T.
			set trbetween off
		endif
				
		=aused(this.aAliasesToPreserve)
		this.aEnvironment[5,1] = select()

		this.aEnvironment[3,1] = set("compatible")
		set compatible off noprompt
		this.aEnvironment[4,1] = m.cProcedure
		this.aEnvironment[7,1] = set("message", 1)
		set message to MESSAGE_LOC
		this.aEnvironment[8,1] = set("safety")
		set safety off
		this.aEnvironment[9,1] = set("path")
		this.aEnvironment[12,1] = set("fields")
		set fields off
		this.aEnvironment[13,1] = set("fields", 2)
		set fields local
		this.aEnvironment[14,1] = on("error")

		*- NOTE: oEngine.aEnvironment[17,1] is updated by main wizard program		
		this.aEnvironment[17,1] = set("classlib")
		
		this.aEnvironment[19,1] = set("exact")
		set exact on
		this.aEnvironment[20,1] = set("echo")
		set echo off
		this.aEnvironment[21,1] = set("memowidth")
		this.aEnvironment[22,1] = set("udfparms")
		set udfparms to value
		this.aEnvironment[23,1] = set("near")
		set near off
		this.aEnvironment[24,1] = set("unique")
		set unique off
		this.aEnvironment[25,1] = set("ansi")
		set ansi off
		this.aEnvironment[26,1] = set("carry")
		set carry off
		this.aEnvironment[27,1] = set("cpdialog")
		set cpdialog off
		this.aEnvironment[28,1] = set("status bar")
		this.aEnvironment[29,1] = sys(5) + curdir()
		this.aEnvironment[30,1] = set("deleted")
		this.aEnvironment[31,1] = set("date")
		this.aEnvironment[15,1] = set('point')
		this.aEnvironment[33,1] = SET("database")

		IF _mac
			this.aEnvironment[34,1] = SET('LIBRARY')
		ENDIF

		*- always SET MULTILOCKS ON (to handle offline views)
		THIS.aEnvironment[35,1] = SET("MULTILOCKS")
		SET MULTILOCKS ON

		this.aEnvironment[36,1] = SYS(3054)
		SYS(3054,0)

	endproc
	
	procedure Cleanup
		local iIndex, cListItem, cListItemList, cListItemAlias
		
		* Reset these in case something breaks later
***		set skip of popup _mtools .F.
***		set skip of bar _mpr_suspend of _mprog .F.

		this.CloseHandles
		this.CloseUnpreservedAliases
		this.ReleaseLibraries
		
		* copy this.aEnvironment to local aEnvironment so we can macro substitute directly
		local array aEnvironment[alen(this.aEnvironment,1), alen(this.aEnvironment,2)]
		=acopy(this.aEnvironment, aEnvironment)
		
		on key label f1
		on key
				
		set compatible &aEnvironment[3,1]
		set procedure to
		if .not. empty(aEnvironment[4,1])
			m.iIndex = 1
			do while .t.
				m.cListItem = this.ParseList(aEnvironment[4, 1], m.iIndex)
				if empty(m.cListItem)
					exit
				endif
				set procedure to (m.cListItem) additive
				m.iIndex = m.iIndex + 1
			enddo
		endif
		

		set exclusive &aEnvironment[6,1]
		select (aEnvironment[5,1])

		set message to [&aEnvironment[7,1]]
		set safety &aEnvironment[8,1]
		
		if .not. empty(aEnvironment[9,1])
*** 2015-11-17 -- thn -- macro substitution does not work here, use name substitution instead
*			set path to &aEnvironment[9, 1]
			set path to (aEnvironment[9, 1])
		else
			set path to
		endif
		
		set fields &aEnvironment[12,1]
		set fields &aEnvironment[13,1]
		on error &aEnvironment[14,1]
		
		set classlib to
		if .not. empty(aEnvironment[17,1])
			THIS.SetErrorOff = .T.
			m.iIndex = 1
			do while .t.
				m.cListItem = this.ParseList(aEnvironment[17, 1], m.iIndex)
				if empty(m.cListItem)
					exit
				endif
				if ' ALIAS ' $ m.cListItem
					m.cListItemAlias = substr(m.cListItem, at('ALIAS ', m.cListItem))
					m.cListItem = alltrim(strtran(m.cListItem, m.cListItemAlias))
					* Check for long file names with surrounding quotes - new to VFP5
					IF !FILE(m.cListItem) AND FILE(EVAL(m.cListItem))
						m.cListItem = EVAL(m.cListItem)
					ENDIF
					m.cListItemAlias = alltrim(substr(m.cListItemAlias, 6))
					set classlib to (m.cListItem) alias (m.cListItemAlias) additive
				else
					* Check for long file names with surrounding quotes - new to VFP5
					IF !FILE(m.cListItem) AND FILE(EVAL(m.cListItem))
						m.cListItem = EVAL(m.cListItem)
					ENDIF
					set classlib to (m.cListItem) additive
				endif
				m.iIndex = m.iIndex + 1
			enddo
			THIS.SetErrorOff = .F.
		endif

		set exact &aEnvironment[19,1]
		set memowidth to (aEnvironment[21,1])
		set udfparms to &aEnvironment[22,1]
		set near &aEnvironment[23,1]
		set unique &aEnvironment[24,1]
		set ansi &aEnvironment[25,1]
		set carry &aEnvironment[26,1]
		set cpdialog &aEnvironment[27,1]
		set status bar &aEnvironment[28,1]
		set default to (aEnvironment[29,1])
		set deleted &aEnvironment[30,1]
		set date to &aEnvironment[31,1]
		set point to "&aEnvironment[15,1]"
		set decimals to &aEnvironment[16,1]
		
		IF WEXIST("Visual Foxpro Debugger")
		ELSE
			IF BITAND(this.aEnvironment[11,1],2^0)#0
				ACTIVATE WINDOW Watch
			ENDIF
			IF BITAND(this.aEnvironment[11,1],2^1)#0
				ACTIVATE WINDOW Locals
			ENDIF
			IF BITAND(this.aEnvironment[11,1],2^2)#0
				ACTIVATE WINDOW "Call Stack"
			ENDIF
			IF BITAND(this.aEnvironment[11,1],2^3)#0
				ACTIVATE WINDOW "Debug Output"
			ENDIF
			IF BITAND(this.aEnvironment[11,1],2^4)#0
				ACTIVATE WINDOW "Trace"
			ENDIF
		ENDIF
		
		set trbetween &aEnvironment[10,1]
		set talk &aEnvironment[1,1]
		set step &aEnvironment[2,1]
		set escape &aEnvironment[18,1]
		on escape &aEnvironment[32,1]
		this.SetErrorOff = .t.
		set database to &aEnvironment[33,1]
		this.SetErrorOff = .f.
		this.haderror = .f. 

		IF _mac
			IF NOT EMPTY(aEnvironment[34,1])
				SET LIBRARY TO (aEnvironment[34,1])
			ELSE
				SET LIBRARY TO
			ENDIF
		ENDIF

		*- restore multilocks
		* We need to trap for error and skip here because product allows
		* a view to be opened with SET MULTI OFF, but once it is set ON
		* you cannot reset it to OFF.
		THIS.SetErrorOff = .T.
		SET MULTILOCKS &aEnvironment[35,1]
		THIS.SetErrorOff = .F.
		
		SYS(3054,INT(VAL(THIS.aEnvironment[36,1])))

		pop key
		
	endproc
	
	procedure ParseList
		parameters cList, iIndex
		local cTemp
		do case
		case m.iIndex = 1
			if occurs(',', m.cList) <> 0
				return alltrim(left(m.cList, at(',', m.cList) - 1))
			else
				return alltrim(m.cList)
			endif
		case occurs(',', m.cList) < m.iIndex - 1
			return ''
		otherwise
			m.cTemp = substr(m.cList, at(',', m.cList, m.iIndex - 1) + 1)
			if occurs(',', m.cTemp) <> 0
				return alltrim(left(m.cTemp, at(',', m.cTemp) - 1))
			else
				return alltrim(m.cTemp)
			endif
		endcase
	endproc
	
	procedure ErrorCleanup
		* This is a stub for subclasses which have cleanup to perform before
		* the engine object is released in the Error method.
	endproc
	
	PROCEDURE Error
		Parameters nError, cMethod, nLine, oObject, cMessage

		local cAction
		
		THIS.HadError = .T.
		this.iError = m.nError
		this.cMessage = iif(empty(m.cMessage), message(), m.cMessage)
		
		if this.SetErrorOff
			do case
			case inlist(lower(THIS.cWizName), 'formwizard', 'mformwizard') ;
				AND ALLTRIM(STR(m.nError))$'3/108/1705/1708/1718'
		 		* Called if file opened with EXCLUSIVE OFF
				=THIS.ALERT(C_FILEUSE_LOC)
		  	otherwise
	  		endcase
			RETURN
		endif
		
		m.cMessage = iif(empty(m.cMessage), message(), m.cMessage)
		if type('m.oObject') = 'O' .and. .not. isnull(m.oObject) .and. at('.', m.cMethod) = 0
			m.cMethod = m.oObject.Name + '.' + m.cMethod
		endif
				
		if C_DEBUG
			m.cAction = this.Alert(ERRORMESSAGE_LOC, MB_ICONEXCLAMATION + ;
				MB_ABORTRETRYIGNORE, ERRORTITLE_LOC)
			do case
			case m.cAction='RETRY'
				this.HadError = .f.
				clear typeahead
				set step on
				&cAction
			case m.cAction='IGNORE'
				this.HadError = .f.
				return
			endcase
		else
			do case
				case This.lQuiet
					This.ErrorMessage = ERRORMESSAGE_LOC
				case m.nError = 1098
				* User-defined error
					m.cAction = this.Alert(message(), MB_ICONEXCLAMATION + ;
						MB_OK, ERRORTITLE_LOC)
				otherwise
					m.cAction = this.Alert(ERRORMESSAGE_LOC, MB_ICONEXCLAMATION + ;
						MB_OK, ERRORTITLE_LOC)
			endcase
		endif
		this.ErrorCleanup
		if !empty(this.cReturnToProc)
			local cReturnToProc
			m.cReturnToProc = this.cReturnToProc
			release oEngine
			return to &cReturnToProc
		ELSE
			*{ ADD JEI RKR 31.03.2005
			* release this
			*} ADD JEI RKR 31.03.2005
		endif
	ENDPROC
	
	PROCEDURE ReadSettings
	ENDPROC

	PROCEDURE WriteSettings
	ENDPROC

	PROCEDURE Alert
		parameters m.cMessage, m.cOptions, m.cTitle, m.cParameter1, m.cParameter2

		private m.cOptions, m.cResponse

		m.cOptions = iif(empty(m.cOptions), 0, m.cOptions)

		if parameters() > 3 && a parameter was passed
			m.cMessage = [&cMessage]
		endif
		
		clear typeahead
		if !empty(m.cTitle)
			m.cResponse = MessageBox(m.cMessage, m.cOptions, m.cTitle)
		else
			m.cResponse = MessageBox(m.cMessage, m.cOptions, ALERTTITLE_LOC)
		endif

		do case
		* The strings below are used internally and should not 
		* be localized
		case m.cResponse = 1
			m.cResponse = 'OK'
		case m.cResponse = 6
			m.cResponse = 'YES'
		case m.cResponse = 7
			m.cResponse = 'NO'
		case m.cResponse = 2
			m.cResponse = 'CANCEL'
		case m.cResponse = 3
			m.cResponse = 'ABORT'
		case m.cResponse = 4
			m.cResponse = 'RETRY'
		case m.cResponse = 5
			m.cResponse = 'IGNORE'
		endcase
		return m.cResponse

	ENDPROC
	
	procedure ProcessOutput
		this.Alert(NOPROCESS_LOC)
	endproc

	procedure Tick
		public iTickTime
		iTickTime = seconds()
	endproc
	
	procedure Tock
		parameters cDescription
		local iSeconds
		iSeconds = seconds()
		activate screen
		if !empty(m.cDescription)
			?m.cDescription + ': '
		else
			* No need to localize this -- it's used for debugging.
			?'Elapsed time: '
		endif
		??str(((m.iSeconds - m.iTickTime) / 60), 5, 3)
	endproc
	
	procedure Help
		do case
		case type('_screen.ActiveForm.cmdHelp') = 'O'
			_screen.ActiveForm.cmdHelp.Click()
		case type('_screen.ActiveForm') = 'O' .and. ;
			type('_screen.ActiveForm.HelpContextID') = 'N' .and. ;
			_screen.ActiveForm.HelpContextID <> 0
			help id (_screen.ActiveForm.HelpContextID)
		case this.iHelpContextID <> 0
			help id (this.iHelpContextID)
		otherwise
			help
		endcase
	endproc		

	procedure GetOS
		DO CASE
		CASE _DOS 
			THIS.nCurrentOS = OS_DOS
		CASE _UNIX
			THIS.nCurrentOS = OS_UNIX
		CASE _MAC
			THIS.nCurrentOS = OS_MAC
		CASE ATC("Windows 3",OS(1)) # 0
			THIS.nCurrentOS = OS_W32S
		CASE ATC("Windows NT",OS(1))#0 OR VAL(OS(3))>=5
			THIS.nCurrentOS = OS_NT
		OTHERWISE
			* Some future system (Windows 95)
			THIS.nCurrentOS = OS_WIN95
		ENDCASE
	endproc

	procedure addbs
		parameters cString
		return addbs(m.cString)
	endproc
	
	procedure justpath
		parameters cString
		return justpath(m.cString)
	endproc
	
	procedure justext
		parameters cString
		return justext(m.cString)
	endproc
	
	procedure justfname
		parameters cString
		return justfname(m.cString)
	endproc
	
	procedure WizLocFile
		parameters cFilename, cPrompt
		local cTempname, cWizardPath
		local array aFile[1]
		
		* If we don't have access to JustFname, etc. (either in FoxTools or
		* in WizEngineAll), just do a LOCFILE() the best we can and return.
		this.HadError = .f.
		this.SetErrorOff = .t.
		m.cTempname = this.justfname(m.cFilename)
		this.SetErrorOff = .f.
		
		if this.HadError
			if type('m.cPrompt') <> 'C' .or. empty(m.cPrompt)
				m.cPrompt = proper(m.cFilename) + ':'
			endif
			this.HadError = .f.
			this.SetErrorOff = .t.
			m.cTempname = locfile(m.cFilename, '', m.cPrompt)
			this.SetErrorOff = .f.
			if this.HadError
				this.HadError = .f.
				m.cTempname = ''
			endif
			return m.cTempname
		endif
		
		m.cTempname = this.justfname(m.cFilename)
		if adir(aFile, sys(2004) + m.cTempname) = 1
			m.cTempname = sys(2004) + m.cTempname
		else
			IF EMPTY(_wizard)
				m.cWizardpath = sys(2004)
			ELSE
				m.cWizardPath = this.addbs(this.justpath(_wizard))
			ENDIF
			do case
			case adir(aFile, m.cWizardPath + m.cTempname) = 1
				m.cTempname = m.cWizardPath + m.cTempname
			case adir(aFile, m.cWizardPath + "WIZARDS\" + m.cTempname) = 1
				m.cTempname = m.cWizardPath + "WIZARDS\" + m.cTempname
			otherwise
				if type('m.cPrompt') <> 'C' .or. empty(m.cPrompt)
					m.cPrompt = proper(m.cTempname) + ':'
				endif
				this.HadError = .f.
				this.SetErrorOff = .t.
				m.cTempname = locfile(m.cTempname, this.justext(m.cTempname), m.cPrompt)
				this.SetErrorOff = .f.
				if this.HadError
					this.HadError = .f.
					m.cTempname = ''
				else
					if adir(aFile, m.cTempname) = 0
						this.Alert(E_FILENOTFOUND_LOC)
						m.cTempname = ''
					endif
				endif
			endcase
		endif
		return m.cTempname
	endproc

#IF MAC_BUILD

* These functions are for Mac only

	*----------------------------------
	FUNCTION GetMacCPU
	*----------------------------------
		RETURN IIF(C_MACPPC_TAG_LOC $ UPPER(VERS()),"PPC","68K")
	ENDFUNC

	*----------------------------------
	FUNCTION SetLibrary
	*----------------------------------
		*- a function to set FoxTools
		*- returns path + name of library, or empty string if failure

		PARAMETER cLibraryFName, cSpecialDir, cVersionFuncName, cReqVersion, cBadVersMessage, cBadLibMessage
		LOCAL cLibrary

		cLibrary = ""

		IF !_mac
			*- this is a Mac only function
			RETURN ""
		ENDIF

		DO CASE
			CASE cLibraryFName $ SET('library')
				IF PARAMETERS() > 2
					IF &cVersionFuncName() < cReqVersion
						THIS.Alert(cBadVersMessage + cLibraryFName)
						RETURN ""
					ENDIF
				ENDIF
				cLibrary = cLibraryFName
				
			CASE !EMPTY(cSpecialDir) AND ADIR(aFile, THIS.AddBS(cSpecialDir) + cLibraryFName) = 1
				*- look in requested folder first
				THIS.SetErrorOff = .t.
				cLibrary = THIS.AddBS(cSpecialDir) + cLibraryFName
				SET LIBRARY TO (THIS.AddBS(cSpecialDir) + cLibraryFName) ADDITIVE
				THIS.SetErrorOff = .f.
				IF THIS.HadError
					THIS.Alert(cBadLibMessage + ".")
					RETURN ""
				ENDIF
				IF PARAMETERS() > 2
					IF &cVersionFuncName() < cReqVersion
						RELEASE LIBRARY (cLibrary)
						THIS.Alert(cBadVersMessage + cLibraryFName)
						RETURN ""
					ENDIF
				ENDIF
				
			CASE ADIR(aFile, SYS(2033,1) + ":" + cLibraryFName) = 1
				*- look in extensions folder
				THIS.SetErrorOff = .t.
				cLibrary = SYS(2033,1) + ":" + cLibraryFName
				SET LIBRARY TO (cLibrary) ADDITIVE
				THIS.SetErrorOff = .f.
				IF THIS.HadError
					THIS.Alert(cBadLibMessage + ".")
					RETURN ""
				ENDIF
				IF PARAMETERS() > 2
					IF &cVersionFuncName() < cReqVersion
						RELEASE LIBRARY (cLibrary)
						THIS.Alert(cBadVersMessage + cLibraryFName)
						RETURN ""
					ENDIF
				ENDIF

			CASE ADIR(aFile, SYS(2004) + cLibraryFName) = 1
				*- look in Foxpro startup folder
				THIS.SetErrorOff = .t.
				cLibrary = SYS(2004) + cLibraryFName
				SET LIBRARY TO (SYS(2004) + cLibraryFName) ADDITIVE
				THIS.SetErrorOff = .f.
				IF THIS.HadError
					THIS.Alert(cBadLibMessage + ".")
					RETURN ""
				ENDIF
				IF PARAMETERS() > 2
					IF &cVersionFuncName() < cReqVersion
						RELEASE LIBRARY (cLibrary)
						THIS.Alert(cBadVersMessage)
						RETURN ""
					ENDIF
				ENDIF


			OTHERWISE
				THIS.Alert(cBadLibMessage + ".")
					RETURN ""

		ENDCASE

		RETURN cLibrary

	ENDFUNC

	*----------------------------------
	FUNCTION LocateApp
	*----------------------------------
		PARAMETER cSig

		LOCAL ax

		*- NOTE: Mac version of Foxtools must be loaded!
		IF !("FOXTOOL" $ SET("LIBR"))
			RETURN ""
		ENDIF

		DIMENSION ax[4]
		ax = ""
		IF FxGetCreat(cSig,0,@ax) == 0
			RETURN ax[4] + ax[1]
		ELSE
			*- error
			RETURN ""
		ENDIF

	ENDFUNC

	*-
	*- next four functions (GetPref, PutPref, FindINI and GetINISection) support working
	*- with .INI type files on the Macintosh.
	*-

	#DEFINE C_CRLF		CHR(13) + CHR(10)
	#DEFINE C_CR		CHR(13)

	*----------------------------------
	FUNCTION PutPref
	*----------------------------------
		*- write value to an INI-style file

		PARAMETERS cSection, cItem, cNewValue, cINIFile

		LOCAL cFile, iSelect, iLine, iMemoWidth, cValue, iCtr, cUItem, ;
			lWritten, iPrev, cOldSafety

		lWritten = .F.

		*- see if INI file is there
		*- first look for value that was passed
		cFile = THIS.FindINI(m.cINIFile)
		IF EMPTY(cFile)
			*- gave up
			RETURN
		ENDIF
		
		*- assume found, load into memo file
		iSelect = SELECT()
		CREATE CURSOR _temp (content m)
		APPEND BLANK
		APPEND MEMO content FROM (m.cFile)

		iMemoWidth = SET("MEMOWIDTH")
		SET MEMOWIDTH TO 254

		cValue = ""
		iLine = ATCLINE("["+ cSection + "]", content)
		IF iLine > 0
			_MLINE = 0
			iLenItem = LEN(m.cItem)
			cUItem = UPPER(m.cItem)
			cLine = MLINE(_temp.content, m.iLine)
			iPrev = _MLINE
			cLine = MLINE(_temp.content, 1, _MLINE)
			FOR iCtr = 1 TO MEMLINES(_temp.content) - iLine 
				IF LEFT(cLine,1) == "[" OR EMPTY(ALLT(cLine))
					EXIT
				ENDIF
				IF UPPER(LEFT(cLine,iLenItem)) == m.cUItem
					*- found the line
					IF "=" $ m.cLine
						REPLACE content WITH LEFT(_temp.content, m.iPrev) + cItem + ;
							"=" + cNewValue + ;
							SUBS(_temp.content, m.iPrev + LEN(cLine) + ;
							IIF(SUBS(_temp.content,iPrev,1) $ C_CRLF,1,0) + ;
							IIF(SUBS(_temp.content,iPrev + 1,1) $ C_CRLF,1,0))
						lWritten = .T.
						EXIT
					ENDIF
				ENDIF
				iPrev = _MLINE
				cLine = MLINE(_temp.content, 1, _MLINE)
			NEXT

			IF !lWritten
				*- must have hit a new section in file, or eof
				REPLACE content WITH LEFT(_temp.content,m.iPrev) + cItem + ;
					"=" + cNewValue + ;
					SUBS(_temp.content, m.iPrev)
				lWritten = .T.
			ENDIF

		ENDIF

		IF !lWritten
			*- must be no section by that name, so add it
			REPLACE content WITH _temp.content + CHR(13) + ;
				"[" + m.cSection + "]" + CHR(13) + ;
				cItem + "=" + cNewValue + CHR(13)
			lWritten = .T.
		ENDIF

		IF lWritten
			*- write out the revised file
			cOldSafety = SET('SAFETY')
			SET SAFETY OFF
			COPY MEMO content TO (cFile + '.')
			SET SAFETY &cOldSafety
		ENDIF

		SET MEMOWIDTH TO iMemoWidth
		USE IN _temp
		SELECT (iSelect)

	ENDFUNC

	*----------------------------------
	FUNCTION GetPref
	*----------------------------------
		*- read values from an INI-style file
		*- returns value

		PARAMETERS cSection, cItem, cINIFile

		LOCAL cFile, iSelect, iLine, iMemoWidth, cValue, iCtr, iFH

		cValue = ""
		cFile = THIS.FindINI(m.cINIFile)
		IF EMPTY(cFile)
			*- gave up
			RETURN ""
		ENDIF

		iFH = FOPEN(cFile)
		IF iFH == -1
			*- couldn't open the file for some reason
			RETURN ""
		ENDIF

		iLenItem = LEN(m.cItem)
		cItem = UPPER(cItem)

		DO WHILE !FEOF(m.iFH)
			cLine = FGETS(m.iFH)
			IF ATC("["+ cSection + "]", LTRIM(cLine)) == 1
				*- found the section
				DO WHILE !FEOF(m.iFH)
					cLine = FGETS(m.iFH)
					IF LEFT(cLine,1) == "[" OR EMPTY(ALLT(cLine))
						*- new section -- must have failed
						EXIT
					ENDIF
					IF UPPER(LEFT(cLine,iLenItem)) == m.cItem AND "=" $ m.cLine
						*- found the line
						cValue = SUBS(m.cLine,AT("=",m.cLine) + 1)
						EXIT
					ENDIF
				ENDDO	&& within section
			ENDIF		&& found the section header
		ENDDO			&& going through entire file

		=FCLOSE(m.iFH)

		RETURN m.cValue

	ENDFUNC

	*----------------------------------
	FUNCTION FindINI
	*----------------------------------
		*- hunt for a file -- look first in current folder, then home() folder
		*- then in system preferences folder

		PARAMETER cINIFile
		LOCAL cFile
		cFile = m.cINIFile
		IF !FILE(m.cFile)
			*- then look in FoxPro home folder
			cFile = HOME() + m.cINIFile
			IF !FILE(m.cFile)
				*- then look in preferences file
				cFile = SYS(2033,2) + ":" + m.cINIFile
				IF !FILE(m.cFile)
					*- give up
					RETURN ""
				ENDIF
			ENDIF
		ENDIF
		RETURN cFile
	ENDFUNC

	*----------------------------------
	FUNCTION GetINISection
	*----------------------------------
		*- populate an array with all items within a given section
		PARAMETERS aSections, cSection, cINIFile

		EXTERNAL ARRAY aSections

		LOCAL lDone

		#DEFINE ERROR_SUCCESS		0		&& OK
		#DEFINE ERROR_NOINIFILE		-108	&& DLL file to check INI not found
		#DEFINE ERROR_NOINIENTRY	-109	&& No entry in INI file
		#DEFINE ERROR_FAILINI		-110	&& failed to get INI entry


		cFile = THIS.FindINI(m.cINIFile)
		IF EMPTY(cFile)
			*- gave up
			RETURN ERROR_NOINIFILE
		ENDIF

		iFH = FOPEN(cFile)
		IF iFH == -1
			*- couldn't open the file for some reason
			RETURN ERROR_FAILINI
		ENDIF

		m.lDone = .F.

		DO WHILE !lDone
			cLine = FGETS(m.iFH)
			IF FEOF(m.iFH)
				m.lDone = .T.
				EXIT
			ENDIF
			IF ATC("["+ cSection + "]", LTRIM(cLine)) == 1
				*- found the section
				DO WHILE !FEOF(m.iFH)
					cLine = FGETS(m.iFH)
					IF LEFT(cLine,1) == "[" OR EMPTY(ALLT(cLine))
						*- new section
						m.lDone = .T.
						EXIT
					ENDIF
					*- insert item into array
					THIS.InsaItem(@aSections,ALLT(LEFT(cLine,AT("=",m.cLine) - 1)))
				ENDDO	&& within section
			ENDIF		&& found the section header
		ENDDO			&& going through entire file

		=FCLOSE(m.iFH)

		RETURN IIF(EMPTY(aSections[1]), ERROR_NOINIENTRY, ERROR_SUCCESS)

	ENDFUNC

	*----------------------------------
	FUNCTION FxStripLF
	*----------------------------------
		*- strip line feeds from file
		PARAMETER cFile

		LOCAL iFH, cBuffer, iSelect

		*- NOTE: If Foxtools is loaded, use it
		IF !("FOXTOOL" $ SET("LIBR"))
			=FxStripLF(cFile)
			RETURN
		ENDIF

		iSelect = SELECT()

		IF !FILE(cFILE)
			RETURN
		ENDIF

		CREATE CURSOR _temp1 (cText m)
		APPEND BLANK
		APPEND MEMO cText FROM (cFile)
		REPLACE cText WITH CHRTRAN(cText,C_CRLF,C_CR)
		COPY MEMO cText TO (cFile)
		IF JustStem(cFile) == JustFName(cfile)
			*- there was no extension, and COPY MEMO will have added one...
			IF FILE(cFile + '.TXT')
				ERASE (cFile)
				RENAME (cFile + '.TXT') TO (cFile)
			ENDIF
		ENDIF
		USE IN _temp1
		SELECT (iSelect)

		RETURN
	ENDFUNC
	
#ENDIF

ENDDEFINE
		
******************************************************************************
DEFINE CLASS WizEngineAll AS WizEngine 
******************************************************************************
	procedure IsWizEngineAll
		* This function is used by WizLocFile to determine whether the engine
		* is based on WizEngineAll.
		return .t.
	endproc
	
	procedure aScanner
		* This procedure searches an array for an expression and returns 
		* the element number of the first match. (Pass .T. for lReturnRow to
		* get the Row number.)
		* The search may be restricted to a particular column of the array.
		* This procedure makes a copy of the array received to allow it to work
		* with member arrays.
		
		parameters m.cArrayName, m.expression, m.column, m.lReturnRow, ;
			m.start, m.howmany
			
		private lSingleDimension, iElement, thearray
		
		=acopy(&cArrayName, thearray)
		
		if alen(thearray,2)=0
			dimension thearray[alen(thearray),1]
			m.lSingleDimension=.t.
		else
			m.lSingleDimension=.f.
		endif
		m.iElement=iif(type('m.start')='N',m.start-1,0)
		m.column=iif(empty(m.column),1,m.column)
		m.start=iif(empty(m.start),1,m.start)
		m.howmany=iif(empty(m.howmany),alen(thearray),m.howmany)
		do while .t.
			m.iElement=ascan(thearray,expression,m.iElement+1,m.howmany)
			if m.iElement=0 .or. asubscript(thearray,m.iElement,2)=m.column
				exit
			else
				m.howmany=m.howmany-m.iElement
			endif
		enddo
		if m.lSingleDimension
			dimension thearray[alen(thearray,1)]
		endif
		return iif(m.lReturnRow, ;
			iif(m.iElement = 0, 0, asubscript(thearray, m.iElement, 1)), m.iElement)
	endproc

	PROCEDURE insaitem
		* Inserts an array element into an array.
		* For 1-D array
		LPARAMETER aArray,sContents,iRow
		IF ALEN(aArray) = 1 AND EMPTY(aArray[1])
		  aArray[1]=m.sContents
		ELSE
		  DIMENSION aArray[ALEN(aArray)+1,1]
		  IF PARAM()=2
		    aArray[ALEN(aArray)]=m.sContents
		  ELSE
		    =AINS(aArray,m.iRow+1)
			aArray[m.iRow+1]=m.sContents
		  ENDIF	
		ENDIF
	ENDPROC
	
	PROCEDURE delaitem
		* Generic routine to delete an array element.
		LPARAMETERS aArray,wziRow
		IF ALEN(aArray)>=m.wziRow
		  IF ALEN(aArray)=1
		    aArray=''
		  ELSE
		    =ADEL(aArray,m.wziRow)
		    DIMENSION aArray[ALEN(aArray)-1]
		  ENDIF
		ENDIF
	ENDPROC

	FUNCTION acolscan
		* This function does an ASCAN for a specific row where
		* aSearch - array to scan
		* sExpr - expression to scan
		* nColumn - column to scan
		* lRetRow - return row (T) or array element (F)
		LPARAMETER aSearch,sExpr,nColumn,lRetRow
		LOCAL apos
		IF TYPE('m.nColumn')#'N'
			nColumn =1
		ENDIF
		IF TYPE('m.lRetRow')#'L'
			m.RetRow = .F.
		ENDIF
		
		m.apos = 1
		DO WHILE .T.
			m.apos = ASCAN(aSearch,m.sExpr,m.apos)
			DO CASE
			CASE m.apos=0	&& did not find match
				EXIT
			CASE ASUBSCRIPT(aSearch,m.apos,2)=m.nColumn
				EXIT
			OTHERWISE
				m.apos=m.apos+1
			ENDCASE
		ENDDO
		IF m.lRetRow AND m.aPos > 0
			RETURN ASUBSCRIPT(aSearch,m.apos,1)
		ELSE
			RETURN m.apos
		ENDIF
	ENDPROC

	PROCEDURE SaveOutFile
		LPARAMETER pMess,pDefFile,pExtn
		LOCAL cSaveFile,wziFHand
	
		IF TYPE("m.pMess")# "C"
			m.pMess = ""
		ENDIF
		IF TYPE("m.pDefFile")# "C"
			m.pDefFile = ""
		ENDIF
		IF TYPE("m.pExtn")# "C" OR EMPTY(m.pExtn)
			m.pExtn = "*"
		ENDIF
	
		DO WHILE .T.
			m.cSaveFile = PUTFILE(m.pMess,m.pDefFile,m.pExtn)
		
			IF EMPTY(m.cSaveFile)
				EXIT
			ENDIF
		
			IF m.pExtn # "*"
				m.cSaveFile = THIS.FORCEEXT(m.cSaveFile,m.pExtn)
			ENDIF
		
			IF FILE(m.cSaveFile)
				*check if file already open
				m.wziFHand=FOPEN(m.cSaveFile)
				IF m.wziFHand= -1
					THIS.Alert(C_FILEUSE2_LOC)
					LOOP
				ENDIF
				=FCLOSE(m.wziFHand)
			ENDIF
			
			EXIT
		ENDDO
		
		THIS.cOutFile = m.cSaveFile
		RETURN !EMPTY(THIS.cOutFile)
	ENDPROC

	FUNCTION JustPath
		* Returns just the pathname.
		LPARAMETERS m.filname
		IF THIS.lUseNative
			RETURN JUSTPATH(m.filname)
		ENDIF
		LOCAL cdirsep
		cdirsep = IIF(_mac,':','\')
		m.filname = SYS(2027,ALLTRIM(UPPER(m.filname)))
		IF m.cdirsep $ m.filname
		   m.filname = SUBSTR(m.filname,1,RATC(m.cdirsep,m.filname))
		   IF RIGHT(m.filname,1) = m.cdirsep AND LEN(m.filname) > 1 ;
		            AND SUBSTR(m.filname,LEN(m.filname)-1,1) <> ':'
		         filname = SUBSTR(m.filname,1,LEN(m.filname)-1)
		   ENDIF
		   RETURN m.filname
		ELSE
		   RETURN ''
		ENDIF
	ENDFUNC
	
	FUNCTION ForceExt
		* Force filename to have a particular extension.
		LPARAMETERS m.filname,m.ext
		IF THIS.lUseNative
			RETURN FORCEEXT(m.filname,m.ext)
		ENDIF
		LOCAL m.ext
		IF SUBSTR(m.ext,1,1) = "."
		   m.ext = SUBSTR(m.ext,2,3)
		ENDIF

		m.pname = THIS.justpath(m.filname)
		m.filname = THIS.justfname(UPPER(ALLTRIM(m.filname)))
		IF AT('.',m.filname) > 0
		   m.filname = SUBSTR(m.filname,1,AT('.',m.filname)-1) + '.' + m.ext
		ELSE
		   m.filname = m.filname + '.' + m.ext
		ENDIF
		RETURN THIS.addbs(m.pname) + m.filname
	ENDFUNC
	
	FUNCTION JustFname
		* Return just the filename (i.e., no path) from "filname"
		LPARAMETERS m.filname
		IF THIS.lUseNative
			RETURN JUSTFNAME(m.filname)
		ENDIF
		LOCAL clocalfname, cdirsep
		clocalfname = SYS(2027,m.filname)
		cdirsep = IIF(_mac,':','\')
		IF RATC(m.cdirsep ,m.clocalfname) > 0
		   m.clocalfname= SUBSTR(m.clocalfname,RATC(m.cdirsep,m.clocalfname)+1,255)
		ENDIF
		IF AT(':',m.clocalfname) > 0
		   m.clocalfname= SUBSTR(m.clocalfname,AT(':',m.clocalfname)+1,255)
		ENDIF
		RETURN ALLTRIM(UPPER(m.clocalfname))
	ENDFUNC

	FUNCTION AddBS
		* Add a backslash unless there is one already there.
		LPARAMETER m.pathname
		IF THIS.lUseNative
			RETURN ADDBS(m.pathname)
		ENDIF
		LOCAL m.separator
		m.separator = IIF(_MAC,":","\")
		m.pathname = ALLTRIM(UPPER(m.pathname))
		IF !(RIGHT(m.pathname,1) $ '\:') AND !EMPTY(m.pathname)
		   m.pathname = m.pathname + m.separator
		ENDIF
		RETURN m.pathname
	ENDFUNC

	FUNCTION JustStem
		* Return just the stem name from "filname"
		LPARAMETERS m.filname
		IF THIS.lUseNative
			RETURN JUSTSTEM(m.filname)
		ENDIF
		IF RATC('\',m.filname) > 0
		   m.filname = SUBSTR(m.filname,RATC('\',m.filname)+1,255)
		ENDIF
		IF RATC(':',m.filname) > 0
		   m.filname = SUBSTR(m.filname,RATC(':',m.filname)+1,255)
		ENDIF
		IF AT('.',m.filname) > 0
		   m.filname = SUBSTR(m.filname,1,AT('.',m.filname)-1)
		ENDIF
		RETURN ALLTRIM(UPPER(m.filname))
	ENDFUNC

	FUNCTION justext
		* Return just the extension from "filname"
		PARAMETERS m.filname
		IF THIS.lUseNative
			RETURN JUSTEXT(m.filname)
		ENDIF
		LOCAL m.ext
		m.filname = this.justfname(m.filname)   && prevents problems with ..\ paths
		m.ext = ""
		IF AT('.', m.filname) > 0
		   m.ext = SUBSTR(m.filname, AT('.', m.filname) + 1, 3)
		ENDIF
		RETURN UPPER(m.ext)
	ENDFUNC
	
	FUNCTION GetStyle
		PARAMETER lIsBold,lIsItalic,lIsUnder
		DO CASE
		CASE m.lIsBold AND m.lIsItalic AND m.lIsUnder
			RETURN "BIU"
		CASE m.lIsBold AND m.lIsItalic 
			RETURN "BI"
		CASE m.lIsBold AND m.lIsUnder
			RETURN "BU"
		CASE m.lIsItalic AND m.lIsUnder
			RETURN "IU"
		CASE m.lIsBold 
			RETURN "B"
		CASE m.lIsItalic 
			RETURN "I"
		CASE m.lIsUnder
			RETURN "U"
		OTHERWISE
			RETURN "N"
		ENDCASE
	ENDFUNC

	PROCEDURE AddCdxTag
		* Takes contents from THIS.aWizSorts array and creates an index TAG
		* from the fields passed in array. If an expression already exists
		* no tag is made and the tag name is returned.
		* Assume database is already selected since a new index TAG is created.
		* Parameters:
		*   aSrtArray - reference of sort fields array (e.g., aSortFields)
		*   aFieldsRef - reference of instance array (e.g., aWizFList, aGridFList)
		
		PARAMETER aSortRef,aFieldsRef
		
		IF PARAMETER() # 2
			THIS.ALERT(C_BADPARMS_LOC)
			RETURN ""
		ENDIF
		
		PRIVATE aSorts
		DIMENSION aSorts[1,1]
		STORE "" TO aSorts
		=ACOPY(THIS.&aSortRef.,aSorts)		
		
		PRIVATE sTagName,sFldExpr,sTagExpr,lHasmemo,sCurAlias,cSortFld
		PRIVATE sCdxName,i,sDBF,wzaCDX,aFileInfo,nTmpCnt,cTmpName,nbuffering
		STORE '' TO sFldExpr,sTagExpr,sTagName,cTmpName,iPos 
		STORE 1 TO nTmpCnt
		
		* Nothing to sort
		IF EMPTY(aSorts[1,1])
		  RETURN ''
		ENDIF

		* Check if cursor in use
		IF AT('.TMP',DBF()) # 0
		  RETURN ''
		ENDIF

		* Also check if read-only
		=ADIR(aFileInfo,DBF())
		IF AT('R',aFileInfo[5])#0
			RETURN ''
		ENDIF
		RELEASE aFileInfo

		m.sCurAlias=ALIAS()							&& alias name
		m.sDBF = DBF()								&& DBF stem
		m.sCdxName = THIS.FORCEEXT(m.sDBF,'CDX') 	&& CDX name

		* make sure we have variable defined
		IF TYPE('THIS.lSortAscend')#'L'
			THIS.lSortAscend = .T.
		ENDIF

		* Get index expression here
		= ACOPY(aSorts,wzaCdx)
		STORE .F. TO wzaCdx

		* Get tag expression looping through fields and type casting
		* for different data types on fields.
		FOR i = 1 TO ALEN(aSorts)

			FOR iPos = 1 TO ALEN(THIS.&aFieldsRef.,1)
			  IF UPPER(THIS.&aFieldsRef.[m.iPos,1])==UPPER(aSorts[m.i])
			    EXIT
			  ENDIF
			ENDFOR
			
			m.cSortFld = THIS.&aFieldsRef.[m.iPos,1]
			
			* check if alias used
			IF AT('.',m.cSortFld) # 0
				m.cSortFld = SUBSTR(m.cSortFld,AT('.',m.cSortFld)+1)
			ENDIF
			
			m.sFldExpr = THIS.GetTagExpr(m.cSortFld,THIS.&aFieldsRef.[m.iPos,2],THIS.&aFieldsRef.[m.iPos,3],THIS.&aFieldsRef.[m.iPos,4],(ALEN(aSorts)=1))

			IF !EMPTY(m.sFldExpr)
				m.sTagExpr = m.sTagExpr + IIF(EMPTY(m.sTagExpr),"","+") + m.sFldExpr
			ENDIF
		  
		ENDFOR
		
		* Get CDX Tag name - use WIZARD_1, WIZARD_2, etc. if expression
		IF ALEN(aSorts) = 1
			m.sTagName = LEFT(aSorts[1],10)
		ELSE
			m.sTagName = "WIZARD_1"
		ENDIF

		* Check for unique Tag name
		DO WHILE TAGNO(m.sTagName)#0
			nTmpCnt = nTmpCnt + 1
			m.sTagName = "WIZARD_"+ALLTRIM(STR(nTmpCnt))
		ENDDO
		
		* Create new index tag here
		m.wzhaderr = .F.
		m.wzisexcl = .T.

		* check if file can be locked, else try to open it exclusively
		IF !ISFLOCKED()
			m.wzisexcl=.F.
			*- set up for error handling
			THIS.SetErrorOff = .T.
			USE (m.sDBF) AGAIN ALIAS (m.sCurAlias) EXCLUSIVE
			IF EMPTY(ALIAS()) OR !ISEXCLUSIVE()	&& file in use error -- could not open exclusive
			  m.sTagName=""
			  m.wzhaderr=.T.
			ENDIF
			THIS.SetErrorOff = .F.
		ENDIF

		m.wzstagdesc=IIF(THIS.lSortAscend,'',' DESC')
		
		* create tag since we now have dbf exclusive
		IF !m.wzhaderr
		  m.nbuffering = cursorgetprop('buffering')
		  * Check if a tag already exists with same expression
		  IF FILE(m.sCdxName)
		    FOR m.i = 1 TO 256			&& max # of tags
		      IF EMPTY(TAG(m.sCdxName,m.i))
				IF m.nbuffering > 3		&& tablebuffering
					=TABLEUPDATE(.T.,.T.)
			    	=cursorsetprop('buffering',1)
			    ENDIF
		        INDEX ON &sTagExpr TAG &sTagName &wzstagdesc
		        EXIT
		      ENDIF
		      * found tag with same expr (checks for asce/desc)
		      * use NORMALIZE function to ensure that functions are not abbrev
		      IF UPPER(NORM(KEY(m.sCdxName,m.i)))=UPPER(NORM(m.sTagExpr))
		   		wzsTmpTag=TAG(m.sCdxName,m.i)
		        SET ORDER TO &wzsTmpTag
				IF (!THIS.lSortAscend AND 'DESCENDING'$SET('ORDER')) OR ;
				   (THIS.lSortAscend AND !'DESCENDING'$SET('ORDER'))
			        sTagName=TAG(m.sCdxName,m.i)
			    	EXIT
			    ENDIF
			  ENDIF
		    ENDFOR
		  ELSE
			IF m.nbuffering > 3		&& tablebuffering
				=TABLEUPDATE(.T.,.T.)
		    	=cursorsetprop('buffering',1)
		    ENDIF
		    INDEX ON &sTagExpr TAG &sTagName &wzstagdesc
		  ENDIF
		  IF m.nbuffering > 3		&& tablebuffering
		  	=cursorsetprop('buffering',m.nbuffering)
		  ENDIF
		ENDIF

		DO CASE
		CASE m.wzhaderr		&& an error occured so reset original
			USE (m.sDBF) AGAIN ALIAS (m.sCurAlias) SHARED
		CASE !m.wzisexcl	&& need to restore to original
			USE (m.sDBF) AGAIN ALIAS (m.sCurAlias) SHARED ORDER &sTagName
		OTHERWISE
			* already indexed -- do nothing
		ENDCASE

		LOCATE	&& goto top
		RETURN m.sTagName

	ENDPROC

	PROCEDURE GetTagExpr
		* Routine returns character expression for sort field
		* so that a compound index can be made.

		LPARAMETER sortfld,cDataType,nDataWid,nDataDec,lOneField

		DO CASE
		  CASE (m.cDataType = DT_CHAR OR m.cDataType = DT_VARCHAR) AND m.nDataWid>40
		  	* Note: since we allow 3 fields for sorting and Maximum Key 
		  	* length per expression is 120 (for intl collate), 
		  	* set limit to 40 per field.
		    RETURN "LEFT("+m.sortfld+",40)"
		  CASE m.cDataType = DT_CHAR OR m.cDataType = DT_VARCHAR
		    RETURN m.sortfld
		  CASE INLIST(m.cDataType,DT_NUM,DT_FLOAT,DT_INTEGER,DT_DOUBLE,DT_CURRENCY) AND m.lOneField
		    RETURN m.sortfld
		  CASE m.cDataType = DT_LOGIC
		    RETURN "IIF("+m.sortfld+",'T','F')"
		  CASE m.cDataType = DT_INTEGER
		    RETURN "STR("+m.sortfld+")"
		  CASE INLIST(m.cDataType,DT_NUM,DT_FLOAT)
		    RETURN "STR("+m.sortfld+","+ALLT(STR(m.nDataWid))+;
				","+ALLT(STR(m.nDataDec))+")"
		  CASE m.cDataType = DT_CURRENCY
		    RETURN "ALLT(STR("+m.sortfld+",16,4))"
		  CASE m.cDataType = DT_DOUBLE
		    RETURN "ALLT(STR(SIGN("+m.sortfld+")*IIF("+m.sortfld+"=0,0,LOG10(ABS("+m.sortfld+"))),20,16))"
		  CASE m.cDataType = DT_DATE
		    RETURN "DTOS("+m.sortfld+")"
		  CASE m.cDataType = DT_DATETIME
			 RETURN "DTOS(TTOD("+m.sortfld+"))+STR(HOUR("+m.sortfld+;
			 "),2)+STR(MINUTE("+m.sortfld+"),2)+STR(SEC("+m.sortfld+"),2)"
		  OTHERWISE  	&& don't index
		    RETURN ""
		ENDCASE
	ENDPROC

	PROCEDURE GetFullTagExpr
		* Get tag expression looping through fields and type casting
		* for different data types on fields.
		* This method assumes that DBF is already selected!!!
		
		LPARAMETER aSortArray
		LOCAL aFldData,i,ipos,sFldExpr,sTagExpr
		IF EMPTY(aSortArray[1])
			RETURN ""
		ENDIF
		DIMENSION aFldData[1]
		=AFIELDS(aFldData)
		sTagExpr = ""
		FOR i = 1 TO ALEN(aSortArray)

			IF EMPTY(aSortArray[m.i])
				LOOP
			ENDIF
			
			FOR iPos = 1 TO ALEN(aFldData,1)
			  IF UPPER(aFldData[m.iPos,1])==UPPER(aSortArray[m.i])
			    EXIT
			  ENDIF
			ENDFOR
			
			sFldExpr = THIS.GetTagExpr(aSortArray[m.i],aFldData[m.iPos,2],aFldData[m.iPos,3],aFldData[m.iPos,4],(ALEN(aSortArray)=1))

			IF !EMPTY(m.sFldExpr)
				sTagExpr = m.sTagExpr + IIF(EMPTY(m.sTagExpr),"","+") + m.sFldExpr
			ENDIF
		ENDFOR 
		RETURN sTagExpr 
	ENDPROC
		
	PROCEDURE GetDbcAlias
		* Takes the current DBC and gets its alias name
		* cDBC - DBC name passed if not current DBC()

		LPARAMETER cDBC

		LOCAL aDBCtmp,cGetDBC,nPos

		IF TYPE("m.cDBC") # "C"
			m.cDBC  =""
		ENDIF

		IF EMPTY(m.cDBC) AND EMPTY(DBC()) 
			RETURN ""
		ENDIF

		m.cGetDBC = IIF(EMPTY(m.cDBC),DBC(),UPPER(m.cDBC))

		DIMENSION aDBCtmp[1,2]
		=ADATA(aDBCtmp)
		m.nPos = ASCAN(aDBCtmp,m.cGetDBC)

		RETURN IIF(m.nPos = 0,"",aDBCtmp[m.nPos-1])
	ENDPROC

	PROCEDURE CheckDBCTag
		* Function returns .F. if DBC opened shared so unable to Index new tag
		* Assume that DBF is already selected when method called!
		LPARAMETER cSortExpr
		LOCAL cDBCAlias,cDBC,lIsView,cTagName,i

		cDBCAlias = ""
		lIsView = (CURSORGETPROP("SourceType") # 3)
		cDBC = CURSORGETPROP("DATABASE")

		IF m.lIsView OR EMPTY(m.cDBC)
			RETURN .T.
		ENDIF

		cDBCAlias = THIS.GetDbcAlias(m.cDBC)
		
		* Test for exclusive use of DBC
		IF !EMPTY(m.cDBCAlias) AND ISEXCLUSIVE(m.cDBCAlias,2)
			RETURN .T.
		ENDIF
		
		* RMK - 07/25/2004 - Remove tag delimiter
		IF !EMPTY(TAGDELIM) AND RIGHT(m.cSortExpr, LEN(TAGDELIM)) == TAGDELIM
			m.cSortExpr = LEFT(m.cSortExpr, LEN(m.cSortExpr) - LEN(TAGDELIM))
		ENDIF
		

		* DBC is opened shared, so test if tag exists
		* Scan for tag name, else create it
		cTagName = ""
		FOR i = 1 TO TagCount()
			IF UPPER(KEY(m.i)) == UPPER(m.cSortExpr)
				cTagName = TAG(m.i)
				EXIT
			ENDIF
		ENDFOR

		RETURN !EMPTY(m.cTagName)
	ENDPROC

	PROCEDURE StringToArray

		LPARAMETERS tcString, taTarget, tcDelim

		LOCAL lnCount, lcString, lcCurrentExp, lnX, lcDelim, lnPos

		lcDelim=IIF(EMPTY(tcDelim), ",", tcDelim)
		lcString=tcString+lcDelim
		lnCount=OCCURS(lcDelim, lcString)

		DIMENSION taTarget[lnCount]

		FOR lnX=1 TO lnCount
			lnPos=AT(lcDelim, lcString)
			taTarget[lnX]=ALLTRIM(LEFT(lcString, lnPos - 1))
			lcString=SUBSTR(lcString, lnPos + 1)
		ENDFOR

		RETURN
	ENDPROC
		
ENDDEFINE


DEFINE CLASS AutoWiz AS custom
	
	*- Minimal Set if table/view already opened:
	*- Views need to be already opened because of issues with
	*- remote connections and parameterized views.
	lUsePages = .T.					&& use pages (form wizard)
	cWizTitle = ""					&& output document title
	cOutFile = ""					&& output file name
	nWizAction = 1					&& output action	(e.g., Save and modify)
	lSortAscend = .T.				&& sort tag
	DIMENSION aWizFields[1,1]		&& array of selected fields
	aWizFields = ""
	DIMENSION aWizSorts[1,1]		&& array of sort fields
	aWizSorts = ""
	DIMENSION aWizStyles[2,2]		&& array of style settings
	aWizStyle = ""

	*- No tables open
	cWizTable = ""					&& name of table to try and open
	cWizAlias = ""					&& alias to use	

	*- used by Reports
	lTruncate = .F.					&& don't wrap fields
	lLandscape = .F.				&& portrait/landscape
	nColumns = 1					&& # of columns in report
	cStyleFile = ""					&& report style file
	nLayout = 1						&& report layout: 1 = columns, 2 = rows

	*- used by Labels
	DIMENSION aLblLines[1,1]		&& contents of label lines
	lMetric = .F.					&& metric?
	lHasSortTag = .F.				&& is sort field already a tag?
	cLblData = ""					&& label data
	
ENDDEFINE


#IF MAC_BUILD

*- a stub, to prevent build errors
*- this function is part of FoxTools for the Macintosh
PROCEDURE FxGetCreat
ENDPROC

#ENDIF
