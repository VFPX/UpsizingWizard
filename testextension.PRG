* This is a sample program that upsizes the VFP sample Northwind database and
* shows how to use extension methods.

local lcConnString, ;
	lnHandle, ;
	lcCurrPath, ;
	lcRunDir, ;
	loEngine, ;
	llReturn

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

loEngine = newobject('UpsizeEngine', 'WizUsz.prg')
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
	.ServerDBName = 'yyy'
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

* Instantiate an extension object and tell the upsizing engine about it.

	.oExtension = createobject('UpsizeExtension')

* Perform the upsizing.

*// Customization: Set NormalShutdown to .F. if you want the analysis tables
*	to not be deleted after the upsizing is done.
	.NormalShutdown = .T.
	.ProcessOutput()
	llReturn = not .HadError
	if not llReturn
		messagebox('The upsize failed. The error message is:' + chr(13) + ;
			chr(13) + .ErrorMessage)
	endif not llReturn
endwith

* Clean up and exit. We don't need to disconnect from the server because
* UpsizeEngine.Destroy does that.

set path to &lcCurrPath
return llReturn

define class UpsizeExtension as Custom
	function CreateDevice(tcDeviceType, tcDevicePhysicalName, tnDeviceNumber, ;
		toUpsizeEngine)
		messagebox('In CreateDevice method')
	endfunc

	function CreateTargetDB(toUpsizeEngine)
		messagebox('In CreateTargetDB method')
	endfunc

	function TruncLogOn(toUpsizeEngine)
		messagebox('In TruncLogOn method')
	endfunc

	function AnalyzeFields(tlAllTables, toUpsizeEngine)
		messagebox('In AnalyzeFields method')
	endfunc

	function CreateTableSpaces(toUpsizeEngine)
		messagebox('In CreateTableSpaces method')
	endfunc

	function CreateClusters(toUpsizeEngine)
		messagebox('In CreateClusters method')
	endfunc

	function CreateTables(toUpsizeEngine)
		messagebox('In CreateTables method')
	endfunc

	function SendData(toUpsizeEngine)
		messagebox('In SendData method')
	endfunc

	function AnalyzeIndexes(toUpsizeEngine)
		messagebox('In AnalyzeIndexes method')
	endfunc

	function BuildRICode(toUpsizeEngine)
		messagebox('In BuildRICode method')
	endfunc

	function CreateIndexes(toUpsizeEngine)
		messagebox('In CreateIndexes method')
	endfunc

	function DefaultsAndRules(toUpsizeEngine)
		messagebox('In DefaultsAndRules method')
	endfunc

	function CreateTriggers(toUpsizeEngine)
		messagebox('In CreateTriggers method')
	endfunc

	function RedirectApp(toUpsizeEngine)
		messagebox('In RedirectApp method')
	endfunc

	function CreateScript(toUpsizeEngine)
		messagebox('In CreateScript method')
	endfunc

	function BuildReport(toUpsizeEngine)
		messagebox('In BuildReport method')
	endfunc

	function TruncLogOff(toUpsizeEngine)
		messagebox('In TruncLogOff method')
	endfunc

	function UpsizeComplete(toUpsizeEngine)
		messagebox('In UpsizeComplete method')
	endfunc
enddefine
