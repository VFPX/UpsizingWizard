* This is a sample program that upsizes the VFP sample Northwind database.

#define ccCRLF chr(13) + chr(10)
local lcConnString, ;
	lnHandle, ;
	lcCurrPath, ;
	lcRunDir, ;
	loEngine, ;
	llReturn, ;
	lcSeparator, ;
	lcOutput, ;
	lcLine, ;
	lcFile

* Connect to the server. If we failed, we can't continue.

*// Customization: Change this connection string as necessary
lcConnString = 'driver=SQL Server;server=(local);trusted_connection=yes'
lnHandle     = sqlstringconnect(lcConnString)
if lnHandle < 1
	return .F.
endif lnHandle < 1

* Set some paths we'll need.

lcCurrPath = set('PATH')
lcRunDir   = addbs(justpath(sys(16)))
set path to '&lcRunDir.PROGRAM, &lcRunDir.DATA' additive

* Instantiate the UpsizeEngine object.

loEngine = newobject('UpsizeEngine', 'program\WizUsz.prg')
with loEngine

* Flag that we don't want any UI.

	.lQuiet = .T.

* Set the connection handle.

	.MasterConnHand = lnHandle
*// Customization: If lnVersion < 7, you must fill in the DeviceDBName,
*	ServerDBSize, ServerLogSize, and DeviceLogName properties.

* Set the target database name and flag whether it's new or not.

*// Customization: Specify the desired target database and set CreateNewDB to
*	.T. to create it.
	.ServerDBName = 'aaa'
	.CreateNewDB  = .T.
*// Customization: You could call .ValidName(.ServerDBName) to ensure the
*	name if valid for this server

* Set the source database, open it, and create a cursor of the tables in it.

*// Customization: Specify the desired database name and path.
	.SourceDB = _samples + 'Northwind\Northwind.dbc'
	open database (.SourceDB)
	.AnalyzeTables()
*// Customization: This flags all tables as exportable. Instead, you can call
*	the SelectTable method to flag individual tables for export. For example:
*	loEngine.SelectTable('customers')
	loEngine.SelectAllTables()
*// Customization: If you want views upsized, call loEngine.ReadViews()

* Create a cursor of fields in the tables to export.

*// Customization: This uses default data type mappings for the fields. You
*	can go through the cursor whose name is in loEngine.EnumFieldsTbl and
*	change the mappings as desired.
	.AnalyzeFields()

*// Customization: This example uses default settings for indexes. If you want
*	to customize how indexes are created, call AnalyzeIndexes and go through
*	the cursor whose name is in loEngine.EnumIndexesTbl and change the settings
*	as desired.

* Set export options.

*// Customization: Most of these are left at their default values but you can
*	change them as desired.
	.DoUpsize            = .T.	&& default
	.DoScripts           = .F.	&& default
	.DoReport            = .F.
	.ExportRelations     = .T.	&& default
	.ExportStructureOnly = .F.	&& default
	.ExportViewToRmt     = .F.
	.ExportTableToView   = .F.	&& default
	.ExportIndexes       = .T.	&& default
	.ExportDefaults      = .T.	&& default
	.ExportTimeStamp     = .T.	&& default
	.ExportValidation    = .T.	&& default
	.ExportSavePwd       = .F.	&& default
	.ExportDRI           = .F.	&& default
	.NotUseBulkInsert    = .F.	&& default
	.Overwrite           = .T.	&& overwrite existing tables; default is .F.
	.DropLocalTables     = .F.	&& default
	.ExportClustered     = .F.	&& default
	.UserUpsizeMethod    = 6	&& use Fast Export if Bulk Insert fails
*// Customization: You can set other properties as desired, including:
*	ViewPrefixOrSuffix: 1 = prefix (default), 2 = suffix, 3 = none
*	ViewNameExtension: suffix or prefix to use for views (default = "v_")
*	NullOverride: 1 = General only (default), 2 = General and Memo only,
*		3 = all fields
*	NewDir: the directory where the Upsizing Engine places its analysis and
*		other files
*	BlankDateValue: the value to use for blank dates (default = 01/01/1900)
*	DisconnectOnExit: set to .F. to keep the SQL Server connection open afterward

*// Customization: If you want to show the progress of the upsizing, use
*	something like:
*	bindevent(loEngine, 'InitProcess',     SomeObject, 'InitProcess')
*	bindevent(loEngine, 'UpdateProcess',   SomeObject, 'UpdateProcess')
*	bindevent(loEngine, 'CompleteProcess', SomeObject, 'CompleteProcess')
*	InitProcess receives two parameters: the title of the process and the
*		number of item to process
*	UpdateProcess receives two parameters: the current item number and
*		optionally the task description
*	CompleteProcess receives no parameters

* Perform the upsizing.

*// Customization: Set NormalShutdown to .F. if you want the analysis tables
*	to not be deleted after the upsizing is done.
	.NormalShutdown = .T.
	.ProcessOutput()
	llReturn = not .HadError

* Display the error message and log if it failed.

	if not llReturn
		messagebox('The upsize failed. The error message is:' + chr(13) + ;
			chr(13) + .ErrorMessage)
	endif not llReturn
	if not used(.ErrTbl)
		select 0
		use (.ErrTbl)
	else
		select (.ErrTbl)
	endif not used(.ErrTbl)
	if reccount() > 0
		lcSeparator = replicate('-', 60) + ccCRLF
		lcOutput    = ttoc(datetime()) + ccCRLF + lcSeparator + ccCRLF
		scan
			lcLine   = 'ObjectType: ' + OBJTYPE + ccCRLF + ;
				'ObjName: ' + OBJNAME + ccCRLF + ;
				'ErrNo: ' + transform(ERRNUMBER) + ccCRLF + ;
				'ErrMsg: ' + ERRMSG + ccCRLF + ;
				'WizErr: ' + WIZERR + ccCRLF + ;
				'FailedSql: ' + FAILEDSQL + ccCRLF
			lcLine   = lcLine + lcSeparator + ccCRLF
			lcOutput = lcOutput + lcLine
		endscan
		lcFile = addbs(sys(2023)) + 'Upsize_Errors_' + sys(2015) + '.log'
		strtofile(lcOutput, lcFile)
		modify file (lcFile) nowait
	endif reccount() > 0
endwith

* Clean up and exit. We don't need to disconnect from the server because
* UpsizeEngine.Destroy does that.

release loEngine
set path to &lcCurrPath
return llReturn
