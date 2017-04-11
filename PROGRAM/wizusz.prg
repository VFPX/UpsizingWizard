#INCLUDE ..\Include\ALLDEFS.H
#DEFINE MAX_PARAMS 254
#DEFINE MAX_ROWS 180
#DEFINE WIZARD_CLOSING	.T.
#DEFINE BOOLEAN_TRUE ".T."
#DEFINE BOOLEAN_FALSE ".F."
#DEFINE BOOLEAN_SQL_TRUE	"1"
#DEFINE	BOOLEAN_SQL_FALSE	"0"
#DEFINE VFP_BETWEEN	"BETWEEN"
#DEFINE VFP_COMMA	","

* jvf 9/1/99
#DEFINE VIEW_NAME_EXTENSION_LOC	"v_"
* New local table name for wiz report if was dropped
#DEFINE DROPPED_TABLE_STATUS_LOC "Table Dropped"
#DEFINE BULK_INSERT_FILENAME "BulkIns.out"

* Use tilde rather than comma in case commas in char fields
#DEFINE BULK_INSERT_FIELD_DELIMITER	~

*{ Add JEI -RKR - 2005.03.25
#DEFINE gcQT  CHR(34)
#DEFINE gc2QT  ['']
*} Add JEI -RKR - 2005.03.25

*******************************************
DEFINE CLASS UpsizeEngine AS WizEngineAll of WZEngine.prg
*******************************************

    *ODBC-related properties
    MasterConnHand=0
    CONNECTSTRING=""
    DataSourceName=""
    ServerType="SQL Server"
    CurrentServerType=""
    ServerVer=0
    UserConnection=""
    ViewConnection=""
    UserName=""
    USERID=0
    MyError=0
    FETCHMEMO=.F.

    *Navigation variables
    DeviceRecalc=.T.
    AnalyzeTablesRecalc=.T.
    AnalyzeFieldsRecalc=.T.
    AnalyzeIndexesRecalc=.T.
    ChooseTargetDBRecalc=.T.
    TableCboRecalc=.T.
    GetRiInfoRecalc=.T.
    EligibleRelsRecalc=.T.
    DataSourceChosen=.F.
    DSNChange=.F.
    NoDataSourceRightNow=.F.
    GetConnDefsRecalc=.T.
    DeviceLogChosen=.F.
    DeviceDBChosen=.F.
    SourceDBChosen=.F.
    GridFilled=.F.

    *Device properties
    DeviceNumbersFree=0
    DeviceDBName=""
    DeviceDBPName=""
    DeviceDBSize=0
    DeviceDBNumber=0
    DeviceDBNew=.F.
    DeviceLogName=""
    DeviceLogPName=""
    DeviceLogSize=0
    DeviceLogNumber=0
    DeviceLogNew=.F.
    MasterPath=""
    NewDeviceCount=0
    DefaultFreeSpace=0
    DeviceDBInDefa=.F.		&&used to in Device class, DeviceThingOK method
    DeviceLogInDefa=.F.
    DBonDefault=.F.

    *Server Database properties
    CreateNewDB=.F.
    ServerFreeSpace=-5
    ServerDBName=""
    ServerDBSize=0
    ServerLogSize=0
	ServerTempFolder = ''

    *Oracle Wizard properties

    *Tables Selection Step properties
    TBFoxTableSize = 0
    TBFoxIndexSize = 0

    *Tablespaces Step properties
    TSNewTableTS = .F.
    TSNewIndexTS = .F.
    TSTableTSName = ""
    TSIndexTSName = ""
    TSDefaultTSName = ""
    TSPermTablespaces =	.T.
    TSNew = .F.
    TSDone = .F.

    *Tablespace File Step properties
    TSFTableFileName = ""
    TSFIndexFileName = ""
    TSFTableFileSize = 0
    TSFIndexFileSize = 0
    TSFDone = .F.

    *Cluster Table Step properties
    CLClustersDone = .F.
    CLTablesDone = .F.
    CLKeysDone = .F.

    *Export properties
    SourceDB=""
    ExportIndexes=.T.
    ExportValidation=.T.
    ExportRelations=.T.
    ExportDRI=.F.
    ExportStructureOnly=.F.
    ExportDefaults=.T.
    ExportTimeStamp=.T.
    ExportTableToView=.F.
    ExportViewToRmt=.T.
    ExportSavePwd=.F.
    OVERWRITE=.F.					&&If .T., existing tables are overwritten
    NullOverride = 1
    ExportClustered=.F.				&& Default to Primary Keys not being Clustered
    ViewNameExtension=VIEW_NAME_EXTENSION_LOC
    ViewPrefixOrSuffix=1
    DropLocalTables=.F.

    *Names of tables created and aliases
    EnumFieldsTbl=""
    MappingTable=""					&& Same as EnumfieldsTbl, used again
    EnumTablesTbl=""
    EnumClustersTbl=""				&& Cluster table name
    EnumIndexesTbl=""
    DeviceTable=""					&& Same table, used twice
    DeviceTableAlias=""				&& Same as DeviceTable, used again by the log device screen
    ViewsTbl=""
    EnumRelsTbl=""
    ScriptTbl=""					&& Basically a big memo field for holding generated sql
    ErrTbl=""
    OraNames=""						&& Just a cursor of Oracle index names

    *Action properties
    DoUpsize=.T.
    DoScripts=.F.
    DoReport=.T.
    * jvf 08/13/99 Now can choose for all tables
    TimeStampAll=0					&& add timestamp column to all tables
    IdentityAll=0					&& add identity column to all tables

    *Permissions properties
    Perm_Device		=	.T.
    Perm_Table		=	.T.
    Perm_Database	=	.T.
    Perm_Default	=	.T.
    Perm_Sproc		=	.T.
    Perm_Index		=	.T.
    Perm_Trigger	=	.T.
    Perm_AltTS		=	.T.
    Perm_CreaTS		=	.T.
    Perm_Cluster	=	.T.
    Perm_UnlimTS	=	.T.
    
    UserUpsizeMethod = 0	&& upsize method chosen by the user.

    *Arrays of server datatypes, one for each FoxPro type
    DIMENSION C[1]
    DIMENSION N[1]
    DIMENSION B[1]
    DIMENSION L[1]
    DIMENSION M[1]
    DIMENSION Y[1]
    DIMENSION D[1]
    DIMENSION T[1]
    DIMENSION P[1]
    DIMENSION G[1]
    DIMENSION F[1]
    DIMENSION I[1]

    * rmk - 01/06/2004 - support for new types in VFP 9
    DIMENSION V[1]
    DIMENSION Q[1]  && varbinary
    DIMENSION W[1]	&& blob

    *Other
    UserInput=""			&& Inputbox always puts user input here
    ZeroDefaultCreated=.F.	&& Flag set true after zero default has been created
    ScriptTblCreated=.F.
    ProcessingOutput=.F.	&& Set to .T. after user clicks Finish button
    OldRow=1				&& Used by type mapping grid to see if row changed
    OldType=""				&& Used by type mapping grid in case user wants to undo change
    NormalShutdown=.F.		&& Flag used to prevent analysis files from getting nuked
    DataErrors=.F.
    NewProjName=""
    PwdInDef=.F.			&& See comment in page9 activate method
    RealClick=.T.			&& Flag for page9
    SaveErrors=.T.			&& Save error tables, set in BuildReport method
    ZDUsed=.F.				&& So that sql for zero default included (if appropriate) in sql script
    FiltCond=""				&& Used on type mapping page
    NewDir=""				&& Directory created to store tables etc. the upsizing wizard creates
    CreatedNewDir=.F.
    KeepNewDir=.F.
    SQLServer=.T.			&& True if we're connected to SQL Server
    TimeStampName=""		&& Name used for all timestamp fields (if any) that are added
    IdentityName = ""		&& Name used for all identity fields (if any) that are added
    TruncLog = -1			&& Status of Trunc.log on chkpt. option of database
    cFinishMsg = ""			&& message to display at end
*** DH 07/24/2013: this property isn't used anymore.
*    nSQL7CompLevel = 0		&& SQL 7.0 database compatiblity level

* Properties added from JEI
	DIMENSION aChooseViews[1]
	DIMENSION aDataErrTbls[1]
	ServerISLocal = .F.		&& .T. - SQL Server is on current machine, .F. - SQL Server is on remote machine
	NotUseBulkInsert = .f.	&& .T. - Always not use Bulk Insert

* Extension object.

	oExtension = .NULL.

* Value to use for blank Date and DateTime values.

	BlankDateValue = .NULL.

* Default folder to use for report information.

	ReportDir = ''

	function Init(tcReturnToProc, tcProcedure)
		dodefault(tcReturnToProc, tcProcedure)
		This.ReportDir = fullpath(NEW_DIRNAME_LOC)
*** 2015-11-30 DH: based on change suggested by thn: in dev mode, add upsizing
*** wizard's key folders to path to retrieve tables and others when running
*** from a custom home folder
		local cDir
		cDir = Sys(16)
		cDir = FullPath(Addbs(justpath(substr(m.cDir, atc(' ', m.cDir, 2) + 1))) + '..\')
			&& upsizing wizard root folder
		if Empty(_VFP.StartMode) and not lower(m.cDir) + 'data' $ lower(set('PATH'))
			set path to (m.cDir + 'data;' + m.cDir + 'program;') additive
		endif
*** 2015-11-30 DH: end of new code
	endfunc

* When the connection handle is set, get the connection string, set the
* connection properties, and get the server version.

	function MasterConnHand_Assign(tnHandle)
		with This
			.MasterConnHand = tnHandle
*** DH 12/06/2012: added IF as suggested by Matt Slay to prevent problem in
*** SetConnProps when tnHandle is 0.
			if tnHandle > 0
				.ServerVer      = .GetServerVersion()
				.ConnectString  = sqlgetprop(tnHandle, 'ConnectString')
				.SetConnProps()
*** DH 12/06/2012: added ENDIF to match IF
			endif tnHandle > 0
		endwith
	endfunc

* Determine the server version.

	procedure GetServerVersion
		local lnhdbc, ;
			lcName, ;
			lcbName, ;
			lnReturn
		declare short SQLGetInfo in odbc32 ;
			integer hdbc, integer fInfoType, string @cName, ;
			integer cNameMax, integer @cbName
		lnhdbc   = sqlgetprop(This.MasterConnHand, 'ODBChdbc')
		lcName   = space(128)
		lcbName  = 0
		lnReturn = sqlgetinfo(lnhdbc, SQL_DBMS_VER, @lcName, 128, @lcbName)
		if lnReturn = SQL_SUCCESS
			lnReturn = val(left(lcName, at(chr(0), lcName) - 1))
		else
			lnReturn = -1
		endif lnReturn = SQL_SUCCESS
		return lnReturn
	endproc
	
    PROCEDURE ProcessOutput
        LOCAL lcSQL, lcMsg
        public oEngine

        THIS.ProcessingOutput=.T.

        *Let user bail if they want
        oEngine = This
        ON ESCAPE OEngine.Esc_proc

		This.InitTherm('Preparing server for upsizing', 100)
		This.GetChooseView()

        IF THIS.SQLServer

            *SQL Server: create devices
			This.UpdateTherm(5, 'Setting up devices')
            THIS.DealWithDevices()

            *SQL Server: create database
			This.UpdateTherm(10, 'Creating target database')
            THIS.CreateTargetDB()

			*{ ADD JEI - RKR - 2005.03.28
			This.UpdateTherm(15, 'Checking for local server')
			This.CheckForLocalServer()
			*} ADD JEI - RKR - 2005.03.28

            *Connect to target database
			This.UpdateTherm(30, 'Connecting to database')
            lcSQL="use " + ALLTRIM(THIS.ServerDBName)
            =THIS.ExecuteTempSPT(lcSQL)

            *set option Trunc.Log on chkpt. for database on server
			This.UpdateTherm(35, 'Setting database options')
            THIS.TruncLogOn()

        ENDIF

        *Make sure everything's been analyzed that's going to be upsized
		This.UpdateTherm(90, 'Setting database options')
        THIS.AnalyzeFields()

        *create tablespaces, clusters and cluster indexes
        IF THIS.ServerType="Oracle"
            THIS.CreateTablespaces()
            THIS.CreateClusters()
        ENDIF
		raiseevent(This, 'CompleteProcess')

        *create tables
        THIS.CreateTables()

        *send data
        THIS.SendData()

        *create indexes
        THIS.AnalyzeIndexes()

        *build RI code
        IF THIS.ExportRelations THEN
            THIS.BuildRiCode()
        ENDIF

        IF THIS.ExportIndexes THEN
            THIS.CreateIndexes()
        ENDIF

        *deal with defaults and validation rules
        THIS.DefaultsAndRules()

        *create put rules and RI code into triggers
        THIS.CreateTriggers()

        *redirect app
        THIS.RedirectApp()

        *do report stuff
        THIS.CreateScript

        THIS.BuildReport()

        *done
        *reset option trunc. log on chkpt. to initial value
        IF THIS.SQLServer
            THIS.TruncLogOff()
        ENDIF

        *test if all the tables were upsized. If not display warning message.
        THIS.UpsizeComplete()
        release oEngine

    ENDPROC

* Ensures the specified table is selected for export.

	procedure SelectTable(tcTable)
		local llReturn, ;
			lnSelect
		llReturn = used(This.EnumTablesTbl)
		if llReturn
			lnSelect = select()
			select (This.EnumTablesTbl)
			locate for lower(trim(TBLNAME)) == lower(alltrim(tcTable))
			llReturn = found()
			if llReturn
				replace EXPORT with .T.
			endif llReturn
			select (lnSelect)
		endif llReturn
		return llReturn
	endproc

* Select all tables for export.

	procedure SelectAllTables(tcTable)
		local llReturn
		llReturn = used(This.EnumTablesTbl)
		if llReturn
			replace all EXPORT with .T. in (This.EnumTablesTbl)
		endif llReturn
		return llReturn
	endproc

    PROCEDURE ERROR
        PARAMETERS nError, cMethod, nLine, oObject
        LOCAL lcErrMsg, lnServerError

        =AERROR(aErrArray)
        nError=aErrArray[1]
        THIS.MyError=nError

        DO CASE
            CASE nError=1523
                *User hit cancel button in ODBC dialog
                THIS.HadError=.T.
                RETURN .F.

            CASE nError=15
                *Not a table (probably means the table is corrupt)
                THIS.HadError=.T.
                RETURN

            CASE nError=108
                *File opened exclusive by someone else
                THIS.HadError=.T.
                RETURN

            CASE nError=1705 OR nError=3
                *File access denied
                *Should be caused only when the wizard tries to open
                *all tables in database exclusively
                THIS.HadError=.T.
                RETURN

            CASE nError=1976
                *Table is not marked for current database
                *Like 1705 above, should only happen in the this.Upsizable function
                THIS.HadError=.T.
                RETURN

            CASE nError=1984
                *Table definition and DBC are out of sync
                *(Another error that could result from this.Upsizable)
                THIS.HadError=.T.
                RETURN

            CASE nError=1498
                *Attempt to create remote view failed because of bogus SQL
                THIS.HadError=.T.
                RETURN

            CASE nError=1160 and not This.lQuiet
                *Out of disk space
                lnUserChoice=MESSAGEBOX(NO_DISK_SPACE_LOC,RETRY_CANCEL+ICON_EXCLAMATION,TITLE_TEXT_LOC)
                IF lnUserChoice=RETRY_CHOICE THEN
                    RETRY
                ENDIF
            CASE nError=1577
                *Table "name" is referenced in a relation
                *Occurs when drop table that's in a relation
                THIS.HadError=.T.
                RETURN

        ENDCASE

        *Unhandled errors->We're dead
        WizEngineAll::ERROR(nError, cMethod, nLine, oObject)

    ENDPROC


    PROCEDURE Esc_proc

        *If the user hits escape when the wizard is processing output
        CLEAR TYPEAHEAD
        IF This.lQuiet or ;
        	MESSAGEBOX(ESCAPE_CONT_LOC,ICON_EXCLAMATION+YES_NO_BUTTONS,TITLE_TEXT_LOC)=USER_YES
            SET ESCAPE OFF
            THIS.Die()
            RETURN TO MASTER
        ELSE
            RETURN
        ENDIF
    ENDPROC


    PROCEDURE DESTROY
        LOCAL lcAction, lcDelDir, lcSQL, llRetVal

        *don't nuke the analysis tables if the user wants a report
        lcAction=IIF(THIS.DoReport AND THIS.ProcessingOutput,"Close","Delete")

        *Deal with tables that are part of the upsizing wizard project (that don't need to be deleted)
        THIS.DisposeTable("Keywords","Close")
        THIS.DisposeTable("ExprMap","Close")
        THIS.DisposeTable("TypeMap","Close")

        *Close this cursor (created in SQL 95 case)
        THIS.DisposeTable(THIS.OraNames,"Close")

        *Clean up device stuff if it exists
        THIS.DeviceCleanUp(WIZARD_CLOSING)

        * jvf 8/17/99 Don't close them if we've already dropped 'em.
        IF !THIS.DropLocalTables
            *Close user tables
            THIS.CloseUserTables
        ENDIF

        *Close/delete analysis tables
        THIS.AnalCleanUp(lcAction, WIZARD_CLOSING)

        * Restore SQL Server 7.0 compatibility levels
*** DH 07/24/2013: compatibility level isn't used anymore.
*        IF ATC("SQL Server",THIS.ServerType)#0 AND THIS.nSQL7CompLevel>=70
*            lcSQL=[sp_dbcmptlevel ]+ALLTRIM(THIS.ServerDBName)+[,]+TRANS(THIS.nSQL7CompLevel)
*            llRetVal=THIS.ExecuteTempSPT(lcSQL)
*        ENDIF

        *Connection cleanup
        IF THIS.MasterConnHand>0 THEN
            SQLDISCONN(THIS.MasterConnHand)
        ENDIF

        *Clean up error tables
        IF !THIS.SaveErrors THEN
            THIS.DisposeTable(THIS.ErrTbl,"Delete")
            IF !EMPTY(This.aDataErrTbls) THEN
                FOR I=1 TO ALEN(This.aDataErrTbls,1)
                    THIS.DisposeTable(This.aDataErrTbls[i,2],"Delete")
                NEXT
            ENDIF
        ENDIF

        *Nix newly created directory if appropriate
        IF THIS.CreatedNewDir AND !THIS.KeepNewDir THEN
            lcDelDir=FULLPATH(SYS(5))
*** DH 07/02/2013: CD to original folder 
*            CD ..
			cd (This.aEnvironment[29, 1])
*** DH 07/02/2013: added TRY around RD command just in case
			try
	            RMDIR (lcDelDir)
	        catch
	        endtry
        ENDIF

        *Release memory variables
        RELEASE aOpenDatabases, aDataSources, aServerDatabases, ;
            aConnDefs, aExport, aTablesToExport, ;
            aDataErrTbls, aDeviceNumbers, ;
            aClusters, aValidTables, aClusterTables, aServerTablespaces, ;
            aDataFiles, aSelectedTablespaces, aSelectList, aFiles

        * restore fetch memo option
*** DH 2015-08-20: don't do this since it isn't set in the first place
*        =CURSORSETPROP('FetchMemo', THIS.FETCHMEMO, 0)

        *- save messagebox until we've have finished clearnup
        IF !EMPTY(THIS.cFinishMsg) and not This.lQuiet
            =MESSAGEBOX(THIS.cFinishMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC)
        ENDIF

        WizEngineAll::DESTROY

    ENDPROC


    PROCEDURE AnalCleanUp
        *Called by main Cleanup proc and if the user changes the source database
        PARAMETERS lcAction, llWizardClosing

        IF !llWizardClosing
	
            *Reset flags
            THIS.AnalyzeTablesRecalc=.T.
            THIS.AnalyzeIndexesRecalc=.T.
            THIS.AnalyzeFieldsRecalc=.T.
            THIS.GetConnDefsRecalc=.T.
            THIS.GridFilled=.F.
            THIS.EligibleRelsRecalc=.T.
            THIS.GetRiInfoRecalc=.T.

        ENDIF

        IF THIS.NormalShutdown THEN
            IF THIS.DoReport AND THIS.ProcessingOutput THEN
                lcAction="Close"
            ELSE
                lcAction="Delete"
            ENDIF
        ELSE
            lcAction="Delete"
        ENDIF

        THIS.DisposeTable(THIS.MappingTable,"Close")
        THIS.DisposeTable(THIS.EnumFieldsTbl,lcAction)
        THIS.DisposeTable(THIS.EnumTablesTbl,lcAction)
        THIS.DisposeTable(THIS.EnumIndexesTbl,lcAction)
        THIS.DisposeTable(THIS.ViewsTbl,lcAction)
        THIS.DisposeTable(THIS.EnumRelsTbl,lcAction)

        *Don't want to delete error or script table unless user hit cancel or error
        IF THIS.NormalShutdown THEN
            lcAction="Close"
        ENDIF
        THIS.DisposeTable(THIS.ErrTbl,lcAction)
        THIS.DisposeTable(THIS.ScriptTbl,lcAction)

        IF !llWizardClosing
            *Reset analysis table-related variables to original default values
            THIS.EnumRelsTbl=""
            THIS.MappingTable=""
            THIS.EnumFieldsTbl=""
            THIS.EnumTablesTbl=""
            THIS.EnumIndexesTbl=""
            THIS.ViewsTbl=""
            THIS.UserConnection=""
        ENDIF


    ENDPROC


    PROCEDURE DeviceCleanUp
        PARAMETERS WizardClosing
        *Called by main Cleanup proc and if the user changes the data source


        *Clean up the device tables
        THIS.DisposeTable(THIS.DeviceTableAlias,"Close")
        THIS.DisposeTable(THIS.DeviceTable,"Delete")

        IF !WizardClosing THEN
            *Reset all Device-related properties to their original default values
            THIS.UserInput=""
            THIS.DeviceTable=""
            THIS.DeviceTableAlias=""
            THIS.DeviceNumbersFree=0
            THIS.DeviceDBName=""
            THIS.DeviceDBPName=""
            THIS.DeviceDBSize=0
            THIS.DeviceDBNumber=0
            THIS.DeviceDBNew=.F.
            THIS.DeviceDBChosen=.F.
            THIS.DeviceLogName=""
            THIS.DeviceLogPName=""
            THIS.DeviceLogSize=0
            THIS.DeviceLogNumber=0
            THIS.DeviceLogNew=.F.
            THIS.DeviceLogChosen=.F.
            THIS.MasterPath=""
            THIS.NewDeviceCount=0
            THIS.ChooseTargetDBRecalc=.T.
            THIS.DataSourceChosen=.F.
            DeviceDBInDefa=.F.
            DeviceLogInDefa=.F.
            DBonDefault=.F.

            *Set recalc flag on
            THIS.DeviceRecalc=.T.

        ENDIF

    ENDPROC


    PROCEDURE CloseUserTables
        LOCAL lcEnumTablesTbl

        *Closes user tables that weren't open before running the upsizing wizard

        lcEnumTablesTbl=THIS.EnumTablesTbl
        IF lcEnumTablesTbl=="" THEN
            RETURN
        ENDIF

        DIMENSION aTableArray[1]
        SELECT CursName FROM (lcEnumTablesTbl) ;
            WHERE Upsizable=.T. AND PreOpened=.F. ;
            INTO ARRAY aTableArray
        IF !EMPTY(aTableArray) THEN
            FOR I=1 TO ALEN(aTableArray)
                *close the table
                SELECT (RTRIM(aTableArray[i]))
                USE
            NEXT
        ENDIF

    ENDPROC


    PROCEDURE DisposeTable
        PARAMETERS lcTableName, lcAction
        *Close the table if it's open
        IF USED(lcTableName)
            SELECT (lcTableName)
            USE
        ENDIF

        *Clean up any backup files incidentally created
        IF FILE(lcTableName+".bak") THEN
            DELETE FILE lcTableName+".bak"
        ENDIF
        IF FILE(lcTableName+".tbk") THEN
            DELETE FILE lcTableName+".tbk"
        ENDIF

        *Delete file if appropriate
        IF lcAction="Delete" THEN
            IF FILE(lcTableName+".dbf") THEN
                DELETE FILE lcTableName+".dbf"
            ENDIF
            IF FILE(lcTableName+".cdx") THEN
                DELETE FILE lcTableName+".cdx"
            ENDIF
            IF FILE(lcTableName+".fpt") THEN
                DELETE FILE lcTableName+".fpt"
            ENDIF
        ENDIF

    ENDPROC


    PROCEDURE InitTherm
        PARAMETERS lcTitle, lnBasis, lnInterval

        *This routine creates the thermometer or initializes it if it already exists

        IF !THIS.ProcessingOutput THEN
            RETURN
        ENDIF
		raiseevent(This, 'InitProcess', lcTitle, lnBasis)

    ENDPROC

	function InitProcess(tcTitle, tnBasis)
	endfunc
	function UpdateProcess(tnProgress, tcTask)
	endfunc
	function CompleteProcess
	endfunc

    PROCEDURE UpDateTherm
        PARAMETERS lnProgress, lcTask
		local lcTaskDescrip

        *This routine updates the thermometer if processing output

        IF !THIS.ProcessingOutput THEN
            RETURN
        ELSE
			lcTaskDescrip = iif(pcount() = 2, lcTask, '')
			raiseevent(This, 'UpdateProcess', lnProgress, lcTaskDescrip)
        ENDIF

    ENDPROC


    PROCEDURE DealWithDevices
        *Creates devices if appropriate
        LOCAL lnRetVal, lcDevicePName,lnDeviceNumber

        IF THIS.ServerVer>=7
            THIS.DeviceDBNew = .F.
            THIS.DeviceLogNew = .F.
        ENDIF

        lcDevicePName=""
        lnDeviceNumber=0

        IF THIS.DeviceDBNew = .T. THEN
            lnRetVal=THIS.CreateDevice("DB",@lcDevicePName,@lnDeviceNumber)
            THIS.DeviceDBPName=lcDevicePName
            THIS.DeviceDBNumber=lnDeviceNumber
        ENDIF
        IF THIS.DeviceLogNew = .T. THEN
            lnRetVal=THIS.CreateDevice("Log", @lcDevicePName,@lnDeviceNumber)
            THIS.DeviceLogPName=lcDevicePName
            THIS.DeviceLogNumber=lnDeviceNumber
        ENDIF
    ENDPROC


    PROCEDURE CreateTargetDB
        LOCAL lcSQL, lcMsg, lnErr, lcErrMsg, lnCompLevel

        *Build the SQL statement
        IF THIS.CreateNewDB THEN
            IF THIS.ServerVer>=7
                lcSQL= "create database " + RTRIM(THIS.ServerDBName)
            ELSE
                lcSQL= "create database " + RTRIM(THIS.ServerDBName) + " on "
                lcSQL= lcSQL + RTRIM(THIS.DeviceDBName) + "=" + ALLTRIM(STR(THIS.ServerDBSize))
                IF THIS.ServerLogSize<> 0 THEN
                    lcSQL= lcSQL+ " log on " + RTRIM(THIS.DeviceLogName)  + "=" + ALLTRIM(STR(THIS.ServerLogSize))
                ENDIF
            ENDIF
        ELSE
            RETURN
        ENDIF

* If we have an extension object and it has a CreateTargetDB method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'CreateTargetDB', 5) and ;
			not This.oExtension.CreateTargetDB(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        *Execute if appropriate
        IF THIS.DoUpsize THEN
            lcMsg=STRTRAN(CREATING_DATABASE_LOC,'|1',RTRIM(THIS.ServerDBName))
            THIS.InitTherm(lcMsg,0,0)
            THIS.UpDateTherm(0,TAKES_AWHILE_LOC)
            SQLSETPROP(THIS.MasterConnHand,"QueryTimeOut",600)
            THIS.MyError=0
            IF !THIS.ExecuteTempSPT(lcSQL, @lnErr,@lcErrMsg) THEN
                IF lnErr=262 THEN
                    *User doesn't have CREATE DATABASE permissions
*** DH 03/23/2015: DataSourceName may be blank so get it from the connection string.
*** Suggested by Mike Potjer.
					local lcSource
					lcSource = rtrim(This.DataSourceName)
					if empty(lcSource)
						lcSource = evl(This.DataSourceName, strextract(This.ConnectString, 'server=', ';', 1, 3))
					endif empty(lcSource)
***                    lcMsg=STRTRAN(NO_CREATEDB_PERM_LOC,'|1',RTRIM(THIS.DataSourceName))
                    lcMsg=STRTRAN(NO_CREATEDB_PERM_LOC,'|1', lcSource)
                ELSE
                    *Something else went wrong
                    lcMsg=STRTRAN(CREATE_DB_FAILED_LOC,'|1',RTRIM(THIS.ServerDBName))
                ENDIF
				if This.lQuiet
					This.HadError     = .T.
					This.ErrorMessage = lcMsg
				else
	                MESSAGEBOX(lcMsg,ICON_EXCLAMATION,TITLE_TEXT_LOC)
				endif This.lQuiet
                THIS.Die
            ENDIF
            SQLSETPROP(THIS.MasterConnHand,"QueryTimeOut",30)
			raiseevent(This, 'CompleteProcess')
        ENDIF

        *Stash sql for script
        THIS.StoreSQL(lcSQL,CREATE_DBSQL_LOC)

*** DH 07/24/2013: compatibility level isn't used anymore.
*        IF THIS.ServerVer>=7
*            lnCompLevel=0
*            lcSQL = "select cmptlevel from master.dbo.sysdatabases where name='"+THIS.ServerDBName+"'"
*            IF THIS.SingleValueSPT(lcSQL,@lnCompLevel,"cmptlevel")
*                THIS.nSQL7CompLevel = lnCompLevel
*                lcSQL=[sp_dbcmptlevel ]+THIS.ServerDBName+[,65]
*                THIS.ExecuteTempSPT(lcSQL)
*            ENDIF
*        ENDIF

    ENDPROC


    PROCEDURE CreateTablespaces

* If we have an extension object and it has a CreateTableSpaces method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'CreateTableSpaces', 5) and ;
			not This.oExtension.CreateTableSpaces(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        * create new Table tablespace or add file to an existing one
        IF THIS.TSNewTableTS
            THIS.CreateTS(THIS.TSTableTSName, THIS.TSFTableFileName, THIS.TSFTableFileSize)
        ELSE
            IF !EMPTY(THIS.TSFTableFileName)
                THIS.CreateDataFile(THIS.TSTableTSName, THIS.TSFTableFileName, THIS.TSFTableFileSize)
            ENDIF
        ENDIF

        *If the tablespace for tables and indexes are the same, bail
        IF THIS.TSTableTSName=THIS.TSIndexTSName THEN
            RETURN
        ENDIF

        * create new Index tablespace or add file to an existing one
        IF THIS.TSNewIndexTS
            THIS.CreateTS(THIS.TSIndexTSName, THIS.TSFIndexFileName, THIS.TSFIndexFileSize)
        ELSE
            IF !EMPTY(THIS.TSFIndexFileName)
                THIS.CreateDataFile(THIS.TSIndexTSName, THIS.TSFIndexFileName, THIS.TSFIndexFileSize)
            ENDIF
        ENDIF

    ENDPROC


    PROCEDURE CreateTables
        LOCAL lcEnumTablesTbl, llRetVal, dummy, llTableExists, MyMessageBox, ;
            lnTableCount, lcMsg, lcSQL, lnError, lcErrMsg, lcCRLF, lcRmtTable

* If we have an extension object and it has a CreateTables method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'CreateTables', 5) and ;
			not This.oExtension.CreateTables(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        dummy = "x"
        lcCRLF = CHR(10) + CHR(13)

        * first go create a SQL statement for all tables the user chose to export
        THIS.CreateTableSQL

        * then if the user wants to upsize and not just create scripts, export the tables
        IF THIS.DoUpsize THEN
            lcEnumTablesTbl = THIS.EnumTablesTbl
            SELECT (lcEnumTablesTbl)

            * For thermometer
            SELECT COUNT(*) FROM (lcEnumTablesTbl) WHERE EXPORT = .T. INTO ARRAY aTableCount
            lnTableCount = 0
            THIS.InitTherm(CREATING_TABLES_LOC, aTableCount,0)

            SCAN FOR EXPORT = .T.
                * For thermometer
                lcMessage = STRTRAN(THIS_TABLE_LOC, "|1", RTRIM(&lcEnumTablesTbl..RmtTblName))
                THIS.UpDateTherm(lnTableCount, lcMessage)
                lnTableCount = lnTableCount + 1

                * If we're Oracle and table is part of cluster, make sure cluster was created
                IF THIS.ServerType = "Oracle" AND !EMPTY(ClustName)
                    IF !THIS.ClusterExported(RTRIM(ClustName))
                        * If cluster wasn't created, skip the table, log error message
                        lcMsg = STRTRAN(CLUST_NOT_CREATED_LOC, '|1', RTRIM(ClustName))
                        REPLACE Exported WITH .F., TblError WITH lcMsg ADDITIVE
                        lnTableCount = lnTableCount + 1
                        LOOP
                    ENDIF
                ENDIF

                *Check to see if table already exists
                lcRmtTable = RTRIM(&lcEnumTablesTbl..RmtTblName)
                IF THIS.TableExists(lcRmtTable) THEN
                    IF !THIS.OVERWRITE THEN
                        THIS.OverWriteOK(RTRIM(&lcEnumTablesTbl..RmtTblName), "Table")
                        DO CASE
                            CASE THIS.UserInput = '3'
                                * user says leave it alone ('NO') THEN
                                lcMsg = TABLE_NOT_CREATED_LOC
                                REPLACE Exported WITH .F., TblError WITH lcMsg ADDITIVE
                                lnTableCount = lnTableCount + 1
                                LOOP
                            CASE THIS.UserInput = '2'
                                * user says overwrite all
                                THIS.OVERWRITE = .T.
                                *CASE else
                                * just keep going
                        ENDCASE
                        THIS.UserInput = ""
                    ENDIF
                    * Drop the table; skip the table if it can't be dropped for some reason
                    llRetVal = THIS.DropTable(&lcEnumTablesTbl..RmtTblName)
                    IF !llRetVal THEN
                        REPLACE Exported WITH .F., TblError WITH CANT_DROP2_LOC ADDITIVE
                        lnTableCount = lnTableCount+1
                        LOOP
                    ENDIF
                ENDIF

                lcSQL = TableSQL
                llRetVal = THIS.ExecuteTempSPT(lcSQL, @lnError, @lcErrMsg)
                IF llRetVal THEN
                    REPLACE Exported WITH .T.
                ELSE
                    THIS.StoreError(lnError, lcErrMsg, lcSQL, CANT_CREATE_LOC, lcRmtTable, TABLE_LOC)
                    REPLACE Exported WITH .F., TblError WITH CANT_CREATE_LOC ADDITIVE, ;
                        TblErrNo WITH lnError
                ENDIF
            ENDSCAN
			raiseevent(This, 'CompleteProcess')
        ENDIF
    ENDPROC


    PROCEDURE CreateTableSQL
        LOCAL lcEnumTablesTbl, lcEnumFieldsTbl, lnOldArea, lcTableName, llTStamp, ;
            lnTableCount, lcCRLF, lcTimeStamp, I, lcExact
        LOCAL llAddedIdentity
*** DH 2015-09-25: added locals for new code
		local lcTable, lcField

        lcCR=CHR(10)
        lcEnumTablesTbl = THIS.EnumTablesTbl
        lcEnumFieldsTbl = THIS.EnumFieldsTbl
        lnOldArea = SELECT()
        SELECT (lcEnumTablesTbl)

        *Update thermometer
        SELECT COUNT(*) FROM (lcEnumTablesTbl) WHERE EXPORT=.T. INTO ARRAY aTableCount
        THIS.InitTherm(TABLE_SQL_LOC,aTableCount,0)
        lnTableCount=0

        *SQL Server bit fields don't support NULLs
        SELECT (lcEnumFieldsTbl)
*** DH 03/20/2015: SQL Server 2005 and later do support null bit fields
		if This.ServerVer < 9
	        REPLACE RmtNull WITH .F. FOR RmtType="bit"
		endif This.ServerVer < 9

* If we're supposed to replace blank date values with nulls, ensure D and T
* fields support nulls.

		if isnull(This.BlankDateValue)
			replace RMTNULL with .T. for inlist(DATATYPE, 'D', 'T')
		endif isnull(This.BlankDateValue)

        * Handle Null overrides
        DO CASE
            CASE THIS.NullOverride=1	&&general only
                REPLACE RmtNull WITH .T. FOR RmtType#"bit" AND INLIST(DATATYPE,"G")
            CASE THIS.NullOverride=2	&&general and memo only
                REPLACE RmtNull WITH .T. FOR RmtType#"bit" AND INLIST(DATATYPE,"G","M")
            CASE THIS.NullOverride=3	&&all fields
                REPLACE RmtNull WITH .T. FOR RmtType#"bit"
        ENDCASE

        *Make sure if we add timestamp that we don't get duplicate field names
        lcTimeStamp=LOWER(TIMESTAMP_LOC)
        lcExact=SET("EXACT")
        SET EXACT ON
        LOCATE FOR FldName=(lcTimeStamp)
        I=1
        DO WHILE LEN(lcTimeStamp)<MAX_NAME_LENGTH
            IF EOF()
                EXIT
            ENDIF
            lcTimeStamp=TIMESTAMP_LOC+LTRIM(STR(I))
            LOCATE FOR FldName=(lcTimeStamp)
            I=I+1
        ENDDO
        SET EXACT &lcExact
        THIS.TimeStampName=lcTimeStamp

        *Make sure if we add identity column that we don't get duplicate field names
        lcIdentityCol = LOWER(IDENTCOL_LOC)
        lcExact=SET("EXACT")
        SET EXACT ON
        LOCATE FOR FldName=(lcIdentityCol)
        I=1
        DO WHILE LEN(lcIdentityCol)<MAX_NAME_LENGTH
            IF EOF()
                EXIT
            ENDIF
            lcIdentityCol=IDENTCOL_LOC+LTRIM(STR(I))
            LOCATE FOR FldName=(lcIdentityCol)
            I=I+1
        ENDDO
        SET EXACT &lcExact
        THIS.IdentityName=lcIdentityCol

		
        SELECT (lcEnumTablesTbl)
        SCAN FOR EXPORT=.T.
			llAddedIdentity = .F. && set to true if we add an identity column

            lcTableName=RTRIM(&lcEnumTablesTbl..TblName)
            lcMsg=STRTRAN(THIS_TABLE_LOC,'|1',lcTableName)
            THIS.UpDateTherm(lnTableCount,lcMsg)

			*{Change JEI - RKR - 28.03.2005
            *lcCreateString="CREATE TABLE " + RTRIM(&lcEnumTablesTbl..RmtTblName) +" ("
*** DH 2015-09-25: handle brackets already in name
***			lcCreateString="CREATE TABLE [" + RTRIM(&lcEnumTablesTbl..RmtTblName) +"] ("
			*{Change JEI - RKR - 28.03.2005
			lcTable = trim(&lcEnumTablesTbl..RmtTblName)
			if left(lcTable, 1) = '['
				lcCreateString = 'CREATE TABLE ' + lcTable + ' ('
			else
				lcCreateString = 'CREATE TABLE [' + lcTable + '] ('
			endif left(lcTable, 1) = '['
*** DH 2015-09-25: end of new code
			            
            SELECT (lcEnumFieldsTbl)
            SCAN FOR RTRIM(&lcEnumFieldsTbl..TblName)==lcTableName
            
            	&&{ Change JEI - RKR 28.03.2005 : upsize field names which are reserved words in SQL
				*lcCreateString = lcCreateString + RTRIM(&lcEnumFieldsTbl..RmtFldname)
*** DH 2015-09-24: handle brackets already in name
***                lcCreateString = lcCreateString + "[" + RTRIM(&lcEnumFieldsTbl..RmtFldname) + "]"
            	&&} Change JEI - RKR 28.03.2005                
				lcField = trim(&lcEnumFieldsTbl..RmtFldname)
				if left(lcField, 1) = '['
	                lcCreateString = lcCreateString + lcField
	            else
	                lcCreateString = lcCreateString + '[' + lcField + ']'
				endif left(lcField, 1) = '['
*** DH 2015-09-24: end of new code
            	
                lcCreateString = lcCreateString + " " + RTRIM(&lcEnumFieldsTbl..RmtType)
                IF &lcEnumFieldsTbl..RmtLength<>0 THEN
                    lcCreateString = lcCreateString + "(" + ALLTRIM(STR(&lcEnumFieldsTbl..RmtLength))
*** DH 12/15/2014: use only SQL Server rather than variants of it
***                    IF THIS.ServerType="Oracle" OR THIS.ServerType=="SQL Server95" THEN
					IF THIS.ServerType="Oracle" OR left(THIS.ServerType, 10) = 'SQL Server'
                        IF &lcEnumFieldsTbl..RmtPrec<>0 THEN
                            lcCreateString = lcCreateString + "," + ALLTRIM(STR(&lcEnumFieldsTbl..RmtPrec))
                        ENDIF
                    ENDIF
                    lcCreateString = lcCreateString + ")"
                ENDIF
                lcCreateString = lcCreateString + " " + IIF(&lcEnumFieldsTbl..RmtNull=.T., "NULL ","NOT NULL ")
                * JVF 11/02/02 Since VFP 8 has autoinc, need to account for ID column at field level.
                * Add SQL Server Seed and Step VFP Increment and Step provided.
                IF THIS.SQLServer AND RmtType= "int (Ident)" AND !llAddedIdentity
                    lcCreateString = STRTRAN(lcCreateString, "(Ident)") + ;
                        " IDENTITY(" + ALLTRIM(STR(AutoInNext-1)) + "," + ;
                        ALLTRIM(STR(AutoInStep)) + ")"

					llAddedIdentity = .T.
                ENDIF
                lcCreateString = lcCreateString + ", "
            ENDSCAN
            SELECT (lcEnumTablesTbl)

            * Add timestamp and identity if appropriate
            * Account for choosing All
            * jvf: 08/13/99
            llTStamp = (THIS.SQLServer AND (TStampAdd OR THIS.TimeStampAll=1))
            
            * rmk - 01/27/2004 - don't add identify column if upsizing AutoInc column
            llIdentity = (THIS.SQLServer AND (IdentAdd OR THIS.IdentityAll=1) AND !llAddedIdentity)

            IF llTStamp
                lcCreateString = lcCreateString + lcTimeStamp + " timestamp, "
            ENDIF

            IF llIdentity
                lcCreateString = lcCreateString + lcIdentityCol + " int IDENTITY(1,1), "
            ENDIF

            * peel off the extra comma at the end and add a closing parenthesis
            lcCreateString = LEFT(lcCreateString, LEN(RTRIM(lcCreateString)) - 1) + ")"

            * Oracle only: add tablespace or cluster name if appropriate
            IF THIS.ServerType = ORACLE_SERVER
                * clustered tables are created in the tablespace of the cluster index
                IF !EMPTY(&lcEnumTablesTbl..ClustName)
                    lcClustName = TRIM(&lcEnumTablesTbl..ClustName)
                    lcClusterKey = THIS.CreateClusterKey(lcTableName, .F.)
                    lcCreateString = lcCreateString + " CLUSTER " + lcClustName + " (" + lcClusterKey + ")"
                ELSE
                    IF !EMPTY(THIS.TSTableTSName)
                        lcCreateString = lcCreateString + " TABLESPACE " + THIS.TSTableTSName
                    ENDIF
                ENDIF
            ENDIF

            REPLACE TableSQL WITH lcCreateString
            lnTableCount = lnTableCount + 1

        ENDSCAN

		raiseevent(This, 'CompleteProcess')
        SELECT (lnOldArea)

    ENDPROC



    FUNCTION TableExists
        PARAMETERS lcTableName
        LOCAL dummy, lcSQuote

        *Checks to see if a table of the same name already exists on the server

        dummy='x'
        lcSQuote=CHR(39)
        IF THIS.ServerType="Oracle" THEN
            lcSQL="SELECT TABLE_NAME FROM USER_TABLES WHERE TABLE_NAME=" + lcSQuote+ UPPER(lcTableName) + lcSQuote
            lcField="TABLE_NAME"
        ELSE
            lcSQL="select uid from sysobjects where uid = user_id() and name =" + lcSQuote + lcTableName + lcSQuote
            lcField="uid"
        ENDIF

        RETURN THIS.SingleValueSPT(lcSQL, dummy, lcField)

    ENDFUNC



    PROCEDURE AddTableToCluster
        PARAMETERS lcClustName, aKeyFields
        LOCAL lnOldArea, lcEnumFieldsTbl, lcTableName, I

        * Adds the cluster name to the table record and the ClustOrder
        * to the fields participating in the key expression
        * The <table> parameter is given by the current record in 'Tables'

        DIMENSION aKeyFields[ALEN(aKeyFields)]
        lnOldArea = SELECT()
        lcEnumFieldsTbl = THIS.EnumFieldsTbl
        lcTableName = LOWER(TRIM(TblName))
        REPLACE ClustName WITH lcClustName

        * set the ClusterOrder field for each Key column
        SELECT (lcEnumFieldsTbl)
        REPLACE ClustOrder WITH 0 FOR TblName == lcTableName
        FOR m.I = 1 TO ALEN(aKeyFields,1)
            LOCATE FOR TblName = lcTableName AND FldName = LOWER(aKeyFields[m.i])
            IF EOF()
                * we should never get here
                LOOP
            ENDIF
            REPLACE ClustOrder WITH m.I
        ENDFOR

        SELECT (lnOldArea)

    ENDFUNC


    PROCEDURE DeleteClusterInfo(lcClusterName)
        LOCAL lnOldArea, lcEnumTablesTbl, lcEnumFieldsTbl, llAll

        * removes <lcClusterName> from its tables in Tables and
        * removes ClustOrder from all corresponding field records in Fields
        * if called with no parameter, clears all cluster info from Tables and Fields

        llAll = (PARAMETERS() = 0)
        lnOldArea = SELECT()
        lcEnumTablesTbl = THIS.EnumTablesTbl
        lcEnumFieldsTbl = THIS.EnumFieldsTbl

        IF !llAll
            SELECT (lcEnumTablesTbl)
            SCAN FOR ClustName = lcClusterName
                lcTableName = LOWER(TRIM(TblName))
                SELECT (lcEnumFieldsTbl)
                REPLACE ClustOrder WITH 0 FOR TblName == lcTableName
                SELECT (lcEnumTablesTbl)
                REPLACE ClustName WITH ""
            ENDSCAN
        ELSE
            SELECT (lcEnumFieldsTbl)
            REPLACE ClustOrder WITH 0 ALL
            SELECT (lcEnumTablesTbl)
            REPLACE ClustName WITH "" ALL
        ENDIF

        SELECT (lnOldArea)
    ENDPROC



    PROCEDURE GetDefaultClusters
        LOCAL lcEnumRelsTbl, lnOldArea, lcEnumTablesTbl, lcClustName, ;
            aParentKey, aChildKey, lcChildExpr, myarray, lcMsg
        DIMENSION aParentKey[1], aChildKey[1], myarray[1]

        THIS.GetRiInfo
        SELECT COUNT(*) FROM (THIS.EnumRelsTbl) INTO ARRAY myarray
        IF EMPTY(myarray)
        	if This.lQuiet
        		This.HadError     = .T.
        		This.ErrorMessage = CANTDEFCLUSTERS_LOC
        	else
	            =MESSAGEBOX(CANTDEFCLUSTERS_LOC, ICON_EXCLAMATION, TITLE_TEXT_LOC)
        	endif This.lQuiet
            RETURN .F.
        ELSE
            THIS.DeleteClusterInfo()

            lnOldArea = SELECT()
            lcEnumRelsTbl = THIS.EnumRelsTbl
            lcEnumTablesTbl = THIS.EnumTablesTbl

            * select all parent tables in aParents
            SELECT DISTINCT  DD_PARENT,DD_PAREXPR  FROM (lcEnumRelsTbl) INTO ARRAY aParents
            FOR m.I = 1 TO ALEN(aParents,1)
                * find next available parent table
                SELECT (lcEnumTablesTbl)
                LOCATE FOR TblName = LOWER(aParents(m.I,1))
                IF EOF() OR !EMPTY(ClustName)
                    LOOP
                ENDIF

                * get cluster name and parent key fields
                lcClustName = CL_LOC + TRIM(aParents[m.i,1])
                aParentKey[1] = ""
                THIS.KeyArray(aParents[m.i,2], @aParentKey)

                * add child table to cluster
                THIS.AddTableToCluster(lcClustName, @aParentKey)

                * find available child tables and add them to the cluster
                SELECT (lcEnumRelsTbl)
                SCAN FOR DD_PARENT = aParents[m.i,1] AND DD_PAREXPR = aParents[m.i,2]
                    lcChild = TRIM(DD_CHILD)
                    lcChildExpr = DD_CHIEXPR
                    SELECT(lcEnumTablesTbl)
                    LOCATE FOR TblName = LOWER(lcChild)
                    IF EOF() OR !EMPTY(ClustName)
                        LOOP
                    ENDIF

                    * get child key fields, compare with parent key fields
                    aChildKey[1] = ""
                    THIS.KeyArray(lcChildExpr, @aChildKey)
                    IF ALEN(aParentKey) != ALEN(aChildKey)
                        LOOP
                    ENDIF

                    * add child table to cluster
                    THIS.AddTableToCluster(lcClustName, @aChildKey)

                    SELECT (lcEnumRelsTbl)
                ENDSCAN
            ENDFOR

            * initialise the aClusters array
            DIMENSION aClusters[1,5]
            aClusters[1,1] = ""
            SELECT DISTINCT ClustName FROM (lcEnumTablesTbl) INTO ARRAY aParents
            FOR m.I = 1 TO ALEN(aParents,1)
                IF (m.I = 1)
                    aClusters[1,1] = aParents[1]
                    aClusters[1,2] = "INDEX"
                    aClusters[1,3] = 0
                    aClusters[1,4] = 2
                    aClusters[1,5] = .F.
                ELSE
					This.InsArrayRow(@aClusters, aParents[m.i], "INDEX", 0, 2, .F.)
                ENDIF
            ENDFOR
        ENDIF

        SELECT (lnOldArea)
    ENDPROC


    FUNCTION GetClusterKey
        PARAMETERS lcTableName, lcClustName
        LOCAL lcEnumRelsTbl, lnOldArea

        * Returns the key of a table that's in a cluster

        lnOldArea = SELECT()
        lcEnumRelsTbl = THIS.EnumRelsTbl
        SELECT (lcEnumRelsTbl)
        LOCATE FOR RTRIM(ClustName) == lcClustName
        IF RTRIM(DD_PARENT)==lcTableName THEN
            lcClusterKey=DD_PAREXPR
        ELSE
            lcClusterKey=DD_CHIEXPR
        ENDIF

        SELECT (lnOldArea)
        RETURN ALLTRIM(lcClusterKey)

    ENDFUNC


    FUNCTION AddTimeStamp
        PARAMETERS lcTableName
        LOCAL lcEnumFieldsTbl, aCount[1]

        *This routine returns True if a table has float, real, binary,
        *varbinary, image, or text data types in them

        lcEnumFieldsTbl = THIS.EnumFieldsTbl

        SELECT COUNT(*) FROM (lcEnumFieldsTbl) ;
            WHERE TblName = lcTableName ;
            AND (DATATYPE = "M" OR ;
            DATATYPE = "G" OR ;
            DATATYPE = "P") ;
            INTO ARRAY aCount

        RETURN aCount[1] > 0
    ENDFUNC


    FUNCTION DropTable
        PARAMETERS lcTable
        LOCAL lcSQL

        IF THIS.ServerType="Oracle" THEN
            lcSQL="drop table " + RTRIM(lcTable) + " CASCADE CONSTRAINTS"
        ELSE
            lcSQL="drop table " + RTRIM(THIS.UserName) + "." + RTRIM(lcTable)
        ENDIF
        lnRetVal=THIS.ExecuteTempSPT(lcSQL)
        RETURN lnRetVal

    ENDFUNC


    PROCEDURE CreateClusters
        LOCAL lcSQL, lcEnumClustersTbl, llClusterCreated, dummy, lcClusterName, lcThermMsg, ;
            lnClustCount, lnError, lcErrMsg, aClustCount

* If we have an extension object and it has a CreateClusters method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'CreateClusters', 5) and ;
			not This.oExtension.CreateClusters(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        *In this routine, the SQL for clusters is created and (possibly) executed
        *and cluster indexes are created as appropriate

        *Bail if there aren't any clusters to create
        IF !THIS.CreateClusterSQL() THEN
            RETURN
        ENDIF

        dummy = "x"
        IF THIS.DoUpsize THEN
            lcEnumClustersTbl = THIS.EnumClustersTbl
            SELECT (lcEnumClustersTbl)
            aClustCount = RECCOUNT()
            THIS.InitTherm(CREATING_CLUSTERS_LOC, aClustCount, 0)
            lnClustCount = 0

            SCAN
                *check for existing cluster
                lcClusterName = RTRIM(ClustName)
                lcThermMsg = STRTRAN(THIS_CLUST_LOC, "|1", lcClusterName)
                THIS.UpDateTherm(lnClustCount, lcThermMsg)
                lnClustCount = lnClustCount + 1
                lcSQL = "SELECT CLUSTER_NAME FROM USER_CLUSTERS WHERE CLUSTER_NAME=" + "'" + UPPER(lcClusterName) + "'"
                IF THIS.SingleValueSPT(lcSQL, dummy, "CLUSTER_NAME") THEN
                    IF !THIS.OVERWRITE THEN
                        *Pop up dialog and ask user what they want to do
                        THIS.OverWriteOK(RTRIM(ClustName), "Cluster")
                        DO CASE
                            CASE THIS.UserInput = '3'
                                *user says leave it alone ('NO') THEN
                                REPLACE Exported WITH .F., ClustErr WITH CLUST_EXISTS_LOC
                                LOOP
                            CASE THIS.UserInput = '2'
                                *user says overwrite all
                                THIS.OVERWRITE = .T.
                                *CASE else
                                *just keep going
                        ENDCASE
                        THIS.UserInput = ""
                    ENDIF
                    * Drop the cluster and all tables; otherwise you can't know
                    * if the cluster key is going to work for the tables
                    * the wizard will try to add
                    *
                    *If dropping cluster and tables doesn't work, bag the cluster
                    *and the tables

                    IF !THIS.DropCluster(RTRIM(ClustName), @lnError, @lcErrMsg)
                        REPLACE Exported WITH .F., ClustErr WITH CANT_DROP_LOC ADDITIVE
                        THIS.StoreError(lnError, lcErrMsg, lcSQL, CANT_DROP_LOC, lcClusterName, CLUSTER_LOC)
                        LOOP
                    ENDIF
                ENDIF

                lcSQL = &lcEnumClustersTbl..ClusterSQL
                IF !THIS.ExecuteTempSPT(lcSQL, @lnError, @lcErrMsg) THEN
                    THIS.StoreError(lnError, lcErrMsg, lcSQL, CANT_CREATE_CLUST_LOC, lcClusterName, CLUSTER_LOC)
                    REPLACE Exported WITH .F., ClustErr WITH CANT_CREATE_CLUST_LOC ADDITIVE, ;
                        ClustErrNo WITH lnError
                ELSE
                    REPLACE Exported WITH .T.
                ENDIF
            ENDSCAN
			raiseevent(This, 'CompleteProcess')
        ENDIF

        * Now go create cluster indexes, if any
        THIS.CreateClusterIndexes

    ENDPROC


    PROCEDURE CreateClusterIndexes
        LOCAL lcEnumRelsTbl, lcEnumIndexesTbl, llRetVal, lnError, lcMessage, lnOldArea

        * If there are clusters, create the index sql now
        *
        * DRI on clusters requires that indexes be created before the DRI is executed
        * On the other hand, indexes can conflict with DRI.  This way the wizard
        * creates cluster indexes, does the DRI (which resolves potential index-DRI conflicts
        * and then creates non-cluster indexes
        *

        *Create table to hold index names and expressions
        IF RTRIM(THIS.EnumIndexesTbl) == ""
            THIS.EnumIndexesTbl = THIS.CreateWzTable("Indexes")
        ENDIF

        lcEnumClustersTbl = THIS.EnumClustersTbl
        lcEnumIndexesTbl = THIS.EnumIndexesTbl
        lnOldArea = SELECT()
        SELECT (lcEnumClustersTbl)

        SCAN FOR ClustType = "INDEX"
            lcClusterName = RTRIM(ClustName)
            lcTagName = CLUST_INDEX_PREFIX + LEFT(lcClusterName, MAX_NAME_LENGTH-LEN(CLUST_INDEX_PREFIX))
            lcTagName = THIS.UniqueOraName(lcTagName)
            lcSQL = "CREATE INDEX [" + lcTagName + "] ON CLUSTER " + lcClusterName

            IF !EMPTY(THIS.TSIndexTSName)
                lcSQL = lcSQL + " TABLESPACE " + LTRIM(THIS.TSIndexTSName)
            ENDIF

            IF THIS.DoUpsize THEN
                llRetVal = THIS.ExecuteTempSPT(lcSQL, @lnError, @lcMessage)
                IF !llRetVal
                    THIS.StoreError(lnError, lcMessage, lcSQL, CLUST_IDX_FAILED_LOC, lcTagName, INDEX_LOC)
                ENDIF
            ENDIF

            SELECT (lcEnumIndexesTbl)
            APPEND BLANK
            REPLACE RmtName WITH lcTagName, ;
                RmtTable WITH lcClusterName, ;
                IndexSQL WITH lcSQL, ;
                Exported WITH llRetVal, ;
                DontCreate WITH .T.

            IF !EMPTY(lcMessage) THEN
                REPLACE IdxError WITH CLUST_IDX_FAILED_LOC, ;
                    IdxErrNo WITH lnError
            ENDIF

            SELECT (lcEnumClustersTbl)
        ENDSCAN

        SELECT(lnOldArea)

    ENDPROC


    FUNCTION UniqueOraName
        PARAMETERS lcObjName, llAddToTable
        LOCAL lcSQL, lnOldArea, lcOraNames, I, lcNewName, lnN, lcM, lcFieldName

        *Returns a name that is not in use by an index, constraint, or trigger
        *Only an issue for Oracle and SQL '95
        *
        *SQL Server 4.x and '95 let you have the same index name as long as they're
        *on different tables; however must have unique name for constraints

        lnOldArea=SELECT()
        IF PARAMETERS()=1
            llAddToTable=.T.
        ENDIF

        * Get the names of all the user's indexes and constraints
        * this cursor is built once, then saved in <this.OraNames>
        IF THIS.OraNames == "" THEN
            SELECT 0
            THIS.OraNames = THIS.UniqueCursorName("OraIdx")

            IF THIS.ServerType="Oracle" THEN
                lcSQL=        "SELECT INDEX_NAME FROM USER_INDEXES "
                lcSQL=lcSQL + "UNION SELECT CONSTRAINT_NAME FROM USER_CONSTRAINTS "
                lcSQL=lcSQL + "UNION SELECT TRIGGER_NAME FROM USER_TRIGGERS"
            ELSE
                lcSQL="SELECT name FROM sysobjects"
            ENDIF

            *If it doesn't work, just send back original name; fatal
            *error is probably imminent anyway
            IF !THIS.ExecuteTempSPT(lcSQL, lnN, lcM, THIS.OraNames) THEN
                THIS.OraNames=""
                RETURN lcObjName
            ENDIF
        ENDIF

        IF THIS.ServerType = ORACLE_SERVER THEN
            lcFieldName = "INDEX_NAME"
        ELSE
            lcFieldName = "name"
        ENDIF

        *See if there's a conflict
        lcOraNames = THIS.OraNames
        SELECT (lcOraNames)
        lcNewName = lcObjName
        LOCATE FOR LOWER(RTRIM(&lcFieldName)) == LOWER(lcNewName)

        I=1
        DO WHILE FOUND()
            *If there's a duplicate name, fiddle with the new object name til it's unique
            lcNewName=LEFT(lcObjName,MAX_NAME_LENGTH-(LEN(LTRIM(STR(I))))) + LTRIM(STR(I))
            I=I+1
            LOCATE FOR LOWER(RTRIM(&lcFieldName))==LOWER(lcNewName)
        ENDDO

        *Add the new name to the list of those in use
        IF llAddToTable THEN
            APPEND BLANK
            REPLACE &lcFieldName WITH lcNewName
        ENDIF

        SELECT (lnOldArea)
        RETURN lcNewName

    ENDFUNC


    FUNCTION DropCluster
        PARAMETERS lcClusterName, lnErrNo, lcErrMsg

        *Drop cluster, its tables, and any RI on the tables
        lcSQL="DROP CLUSTER " + RTRIM(lcClusterName)+ " INCLUDING TABLES CASCADE CONSTRAINTS"
        RETURN THIS.ExecuteTempSPT(lcSQL,@lnErrNo, @lcErrMsg)

    ENDFUNC


    FUNCTION ClusterExported
        PARAMETERS lcClustName
        LOCAL lcEnumClustersTbl, lnOldArea, llExported

        * Checks to see if the cluster was actually created

        lcEnumClustersTbl = THIS.EnumClustersTbl
        lnOldArea = SELECT()
        SELECT (lcEnumClustersTbl)
        LOCATE FOR RTRIM(ClustName) == lcClustName
        llExported = Exported
        SELECT(lnOldArea)
        RETURN llExported

    ENDFUNC


    FUNCTION HashDefault
        PARAMETERS lcRel
        LOCAL lcParent, lcChild, lnDupID, lnHashKeys, lnOldArea

        *Don't change this ratio w/o letting UE know
        #DEFINE HASH_RATIO 			1.1

        #DEFINE HASHKEYS_FLOOR		500
        #DEFINE HASHKEYS_CEILING	2147483647

        *
        *Derives a number for default hashkey value
        *

        *Figures out how many records are in the parent table, multiplies it by some number

        THIS.ParseRel (lcRel, @lcParent, @lcChild, @lnDupID)

        lnOldArea=SELECT()
        SELECT (lcParent)

        lnHashKeys=RECCOUNT()*HASH_RATIO
        lnHashKeys=IIF(lnHashKeys>HASHKEYS_FLOOR,lnHashKeys,HASHKEYS_FLOOR)
        lnHashKeys=IIF(lnHashKeys>HASHKEYS_CEILING,HASHKEYS_CEILING,lnHashKeys)
        SELECT (lnOldArea)

        RETURN INT(lnHashKeys)

    ENDFUNC


    PROCEDURE OverWriteOK
        PARAMETERS lcObjectName, lcObjectType
*** DH 2015-09-14: don't prompt the user and assume No was chosen if lQuiet is .T.
		if This.lQuiet
			This.UserInput = 3
			return
		endif This.lQuiet
*** DH 2015-09-14: end of new code

        LOCAL aButtonNames, lcMessageText, MyMessageBox

        *display message box displaying Yes, Yes to all, No buttons
        IF lcObjectType="Table" THEN
            lcMessageText=STRTRAN(TABLE_EXISTS_LOC,"|1",lcObjectName)
        ELSE
            lcMessageText=STRTRAN(CLUSTER_EXISTS_LOC,"|1",lcObjectName)
        ENDIF

        DIMENSION aButtonNames[3]
        aButtonNames[1]=YES_LOC
        aButtonNames[2]=YESALL_LOC
        aButtonNames[3]=NO_LOC
        MyMessageBox = newobject('MessageBox2', 'VFPCtrls.vcx', '', ;
        	lcMessageText, @aButtonNames)
        MyMessageBox.Show()
        if vartype(MyMessageBox) = 'O'
        	This.UserInput = MyMessageBox.cChoice
        	MyMessageBox.Release()
        endif vartype(MyMessageBox) = 'O'
    ENDPROC


    FUNCTION CreateClusterSQL
        LOCAL lcClustName, lcPkey, aClusterTables, lcEnumClustersTbl, lnOldArea
        DIMENSION aClusterTables[1]

        *If the clusters table doesn't exist yet, there won't be any clusters
        IF EMPTY(THIS.EnumClustersTbl)
            RETURN .F.
        ENDIF

        * needs to be of the form:
        * CREATE CLUSTER <clustername> (<key_name> <datatype>, etc.)

        lnOldArea = SELECT()
        lcEnumClustersTbl = THIS.EnumClustersTbl
        SELECT (lcEnumClustersTbl)
        SCAN
            lcClustName = RTRIM(ClustName)
            THIS.GetEligibleClusterTables(@aClusterTables, lcClustName)
            lcPkey = THIS.CreateClusterKey(aClusterTables[1], .T.)

            lcSQL = "CREATE CLUSTER " + lcClustName + " (" + lcPkey + ")"

            IF !EMPTY(THIS.TSTableTSName)
                lcSQL = lcSQL + " TABLESPACE " + LTRIM(THIS.TSTableTSName)
            ENDIF

            IF ClustType = RTRIM("HASH")
                lcSQL = lcSQL + " HASHKEYS " + LTRIM(STR(HashKeys))
            ENDIF

            IF !EMPTY(ClustSize)
                lcSQL = lcSQL + " SIZE " + LTRIM(STR(ClustSize)) + " K"
            ENDIF

            REPLACE ClusterSQL WITH lcSQL
        ENDSCAN
        SELECT(lnOldArea)

        RETURN .T.

    ENDFUNC


    FUNCTION CreateClusterKey
        PARAMETERS lcTableName, llDataType

        * Creates the Cluster Key with data types for CreateClustersSQL
        * Creates the Cluster Key list for CreateTablesSQL

        LOCAL aKeyArray, lcClustKey, I
        DIMENSION aKeyArray[1,4]
        THIS.GetInfoTableFields(@aKeyArray, lcTableName, .T.)
        lcClustKey = ""

        FOR m.I = 1 TO ALEN(aKeyArray,1)
            lcClustKey = lcClustKey + RTRIM(aKeyArray[m.i, 1])

            IF llDataType
                lcClustKey = lcClustKey + " " + RTRIM(aKeyArray[m.i, 2])
                IF !EMPTY(aKeyArray[m.i, 3])
                    lcClustKey = lcClustKey + "(" + ALLTRIM(STR(aKeyArray[m.i, 3]))
                    IF !EMPTY(aKeyArray[m.i, 4])
                        lcClustKey = lcClustKey + "," + ALLTRIM(STR(aKeyArray[m.i, 4]))
                    ENDIF
                    lcClustKey = lcClustKey + ")"
                ENDIF
            ENDIF

            lcClustKey = lcClustKey + ", "
        ENDFOR

        *peel off the extra comma at the end
        lcClustKey = LEFT(lcClustKey, LEN(lcClustKey)-2)
        RETURN lcClustKey

    ENDFUNC


    PROCEDURE SendData
        LOCAL lcOldArea, lcEnumTables, lcSprocSQL, lnBigRows, ;
            lnSmallRows,  lnBigBlocks, lnSmallBlocks, lnExportErrors, ;
            ll255, llMaxErrExceeded, lcErrMsg, lcCursorName, lcErrTblName, lcExportType, lReturnVal
		local llDone

        IF !THIS.DoUpsize OR THIS.ExportStructureOnly THEN
            RETURN
        ENDIF

* If we have an extension object and it has a SendData method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'SendData', 5) and ;
			not This.oExtension.SendData(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        lcOldArea=SELECT()
        lcErrMsg=""
        lcEnumTables=THIS.EnumTablesTbl
        SELECT (lcEnumTables)

        *Don't send deleted records
        lcDelStatus=SET("DELETED")
        SET DELETED ON

        SCAN FOR &lcEnumTables..EXPORT=.T. AND Exported=.T.

            lcExportType = ""  && can resolve to BULKINSERT, FASTEXPORT, or JIMEXPORT
            lcTablePath=&lcEnumTables..TblPath
            lcRmtTableName=RTRIM(&lcEnumTables..RmtTblName)
            lcTableName=RTRIM(&lcEnumTables..TblName)
            lcCursorName=RTRIM(&lcEnumTables..CursName)

            lcEnumFieldsTbl=RTRIM(THIS.EnumFieldsTbl)
            IF THIS.ServerType="Oracle" THEN
                *this query eliminates long, raw, and long raw data types
                SELECT COUNT(FldName) FROM  &lcEnumFieldsTbl ;
                    WHERE (RmtType="raw" OR RmtType="long") ;
                    AND RTRIM(TblName)==lcTableName INTO ARRAY aExportType
                IF !EMPTY(aExportType)
                    lcExportType = "JIMEXPORT"
                ENDIF
            ELSE
                lReturnVal = SQLEXEC(THIS.MasterConnHand, "SET IDENTITY_INSERT " + lcRmtTableName + " ON")
            ENDIF

            *	Text and image datatypes won't work with BULKINSERT OR FASTEXPORT
            *	(can't pass these as parameters to stored procedures, nor can they
            *	reside in text files to be used as source of SQL>=7 bulk insert.
            IF EMPTY(lcExportType)
                SELECT COUNT(FldName) FROM &lcEnumFieldsTbl ;
                    WHERE (RmtType="text" OR RmtType="image") ;
                    AND RTRIM(TblName)==lcTableName INTO ARRAY aExportType
                IF !EMPTY(aExportType) OR !This.ServerISLocal
                    lcExportType = "JIMEXPORT"
                ENDIF
            ENDIF

            IF EMPTY(lcExportType) AND THIS.ServerType#"Oracle" AND THIS.ServerVer>=7
                * jvf Determine if we can use SQL >=7's BULK INSERT function:
                * 	- Can't bulk insert tables with NULL columns of non-character
                * 		datatype b/c COPY TO loses the null value by creating
                * 		empty values in the text file (eg, /  /, F).

                *{ JEI RKR 23.11.2005 Change
*!*	                SELECT COUNT(FldName) FROM  &lcEnumFieldsTbl ;
*!*	                    WHERE (RmtType <> "char" AND lclnull) ;
*!*	                    AND RTRIM(TblName)==lcTableName INTO ARRAY aExportType
                SELECT COUNT(FldName) FROM  &lcEnumFieldsTbl ;
                    WHERE lclnull ;
                    AND RTRIM(TblName)==lcTableName INTO ARRAY aExportType
                *} JEI RKR 23.11.2005
					do case
						case empty(aExportType)
						case not This.NotUseBulkInsert and ;
							not This.CheckForNullValuesInTable(lcTableName)
							lcExportType = "BULKINSERT"
						case This.NotUseBulkInsert
							lcExportType = "JIMEXPORT"
						otherwise
							lcExportType = "FASTEXPORT"
					endcase
            ENDIF

            * If survived all conditions above, we can go fast
            IF EMPTY(lcExportType)
                IF THIS.ServerType#"Oracle" AND THIS.ServerVer>=7
                	*( JEI RKR 23.11.2005 Change
                	IF This.NotUseBulkInsert = .t.
                		lcExportType = "FASTEXPORT"
                	ELSE
						if This.NotUseBulkInsert or ;
							This.CheckForNullValuesInTable(lcTableName)
                			lcExportType = "FASTEXPORT"
                		ELSE
		                    lcExportType = "BULKINSERT"  && specific to SQL vers >= 7
		                ENDIF
					ENDIF
				    *lcExportType = "BULKINSERT"  && specific to SQL vers >= 7
				    *} JEI RKR 23.11.2005 Change
                ELSE
                    lcExportType = "FASTEXPORT"
                ENDIF
            ENDIF

            * If the table has 255 fields, we have to use JimExport
            * Stored procs can only pass 255 parameters and one of them
            * is used for something other than data leaving 254.
            * jvf 1/9/01 SQL >= 7 can handle 1024 SP Parameters. Handled in Case below.
            ll255=(MAX_PARAMS=THIS.CountFields(lcCursorName))

* If we're upsizing to SQL Server 2005 or later, try to do a bulk XML import
* since it's way faster.

			if This.SQLServer and This.ServerVer >= 9
				llDone = This.BulkXMLLoad(lcTableName, lcCursorName, ;
					lcRmtTableName)
				lnExportErrors = iif(llDone, 0, 1)
			endif This.SQLServer ...

* If we didn't use bulk XML load, or if it failed, use one of the other
* mechanisms.

            DO CASE
				case llDone
                CASE lcExportType = "BULKINSERT" AND THIS.Perm_Database
                    * jvf Use SQL 7's BULK INSERT technique
                    lnExportErrors=THIS.BulkInsert(lcTableName, lcCursorName, ;
                        lcRmtTableName, @llMaxErrExceeded)
                    IF (lnExportErrors == -1) THEN
                    	RETURN
                    ENDIF
                    * jvf SQL >=7 can handle 1024 SP Parameters.
                CASE lcExportType = "FASTEXPORT" AND THIS.Perm_Sproc AND ;
                        (!ll255 OR (THIS.ServerType="Oracle" OR THIS.ServerVer<7)) AND !TStampAdd
                    *go fast if possible and user can create sprocs
                    lnExportErrors=THIS.FastExport(lcTableName, lcCursorName, lcRmtTableName, @llMaxErrExceeded)
                OTHERWISE
*** 11/20/2012: pass .T. to JimExport so it updates the thermometer
***                    lnExportErrors=THIS.JimExport(lcTableName, lcCursorName, lcRmtTableName, @llMaxErrExceeded)
                    lnExportErrors=THIS.JimExport(lcTableName, lcCursorName, lcRmtTableName, @llMaxErrExceeded, '', .T.)
            ENDCASE

            IF lnExportErrors<>0 THEN
                lcErrorTable=ERR_TBL_PREFIX_LOC + ;
                    LEFT(lcTableName,MAX_NAME_LENGTH-LEN(ERR_TBL_PREFIX_LOC))
            ELSE
                lcErrorTable=""
            ENDIF

            IF llMaxErrExceeded THEN
                lcErrMsg=DEXPORT_FAILED_LOC
                THIS.StoreError(.NULL.,"","",lcErrMsg,lcTableName,DATA_EXPORT_LOC)
            ELSE
                IF lnExportErrors<>0 THEN
                    lcErrMsg=STRTRAN(SOME_ERRS_LOC,"|1",LTRIM(STR(lnExportErrors)))
                    THIS.StoreError(.NULL.,"","",lcErrMsg,lcTableName,DATA_EXPORT_LOC)
                ENDIF
            ENDIF

            IF THIS.NormalShutdown=.F.
                EXIT
            ENDIF

            *For report
            REPLACE &lcEnumTables..DataErrs WITH lnExportErrors, ;
                &lcEnumTables..DataErrMsg WITH lcErrMsg, ;
                &lcEnumTables..ErrTable WITH lcErrorTable

            lnExportErrors=0
            llMaxErrExceeded=.F.
            lcErrMsg=""

            SQLEXEC(THIS.MasterConnHand, "SET IDENTITY_INSERT " + lcRmtTableName + " OFF")
        ENDSCAN

        SET DELETED &lcDelStatus
        SELECT (lcOldArea)

    ENDPROC


    FUNCTION CountFields
        PARAMETERS lcCursorName
        LOCAL lnFldCount, lnOldArea

        lnOldArea=SELECT()
        SELECT (lcCursorName)
        lnFldCount=AFIELDS(zoo)
        SELECT (lnOldArea)

        RETURN lnFldCount

    ENDFUNC

* SQL Server 2005 (or later) bulk XML import. We need to get the server, so
* either parse it out of the connection string or get it from the DSNs entry
* in the Registry. Parse the user name and password from the connection
* string, then call BulkXMLLoad.PRG.

	function BulkXMLLoad(tcTableName, tcCursorName, tcRmtTableName)
		local lcServer, ;
			loRegistry, ;
			laValues[1], ;
			lnPos, ;
			lcUserName, ;
			lcPassword, ;
			lcMsg, ;
			llReturn
		if empty(This.DataSourceName)
			lcServer = strextract(This.ConnectString, 'server=', ';', 1, 3)
		else
			loRegistry = newobject('ODBCReg', home() + 'ffc\registry.vcx')
			loRegistry.EnumODBCData(@laValues, This.DataSourceName)
			lnPos = ascan(laValues, 'Server', -1, -1, 1, 15)
			if lnPos > 0
				lcServer = laValues[lnPos, 2]
			endif lnPos > 0
		endif empty(This.DataSourceName)
		if not 'TRUSTED_CONNECTION=YES' $ upper(This.ConnectString) and ;
			not 'INTEGRATED SECURITY=SSPI' $ upper(This.ConnectString)
			lcUserName = strextract(This.ConnectString, 'uid=', ';', 1, 3)
			lcPassword = strextract(This.ConnectString, 'pwd=', ';', 1, 3)
		endif not 'TRUSTED_CONNECTION=YES' $ upper(This.ConnectString) ...
		This.InitTherm(SEND_DATA_LOC, 100)
		lcMsg = strtran(THIS_TABLE_LOC, '|1', tcTableName)
		This.UpDateTherm(0, lcMsg)
		llReturn = empty(BulkXMLLoad(tcCursorName, tcRmtTableName, ;
			This.BlankDateValue, This.ServerDBName, lcServer, lcUserName, ;
			lcPassword, alltrim(This.ServerTempFolder)))
		if llReturn
			raiseevent(This, 'CompleteProcess')
		endif llReturn
		return llReturn
	endfunc

    FUNCTION JimExport
        *Thanks Jim Lewallen for this code (or the important and bug-free parts of it anyway)

*** 11/20/2012: added tlUpdateProgress parameter
***        PARAMETERS lcTableName, lcCursorName, lcRmtTableName, llMaxErrExceeded, tcDataErrTable &&  JEI RKR 24.11.2005 Add tcDataErrTable 
        PARAMETERS lcTableName, lcCursorName, lcRmtTableName, llMaxErrExceeded, tcDataErrTable, tlUpdateProgress

        LOCAL lnOldArea, lnFieldCount, lcInsertString, lcInsertFinal, llRetVal, lnRecs, ;
            lnRecordsCopied, lcMsg, I, lnExportErrors, lcDataErrTable, lnMaxErrors, ;
            aTblFields, lcSQLErrMsg, lnSQLErrno
*** DH 2015-09-25: added local for new code
		local lcField
		
		*{ JEI RKR 24.11.2005 Add
		LOCAL lcEmptyDateValue, lcEmptyDateValueEnd as String
		*} JEI RKR 24.11.2005 Add
		*{ JEI RKR 05.01.2006 Add
		LOCAL lnNumberOfPar as Integer
		lnNumberOfPar = PARAMETERS()
		IF lnNumberOfPar > 4
			lcDataErrTable = tcDataErrTable
		ENDIF
		*} JEI RKR 05.01.2006 Add
		
        #DEFINE STAGGER_COUNT		5

        *Create array of local field names and their remote equivalents
        lcEnumFields=THIS.EnumFieldsTbl
        DIMENSION aTblFields[1]
        *{  JEI RKR Add field DataType, RmtNull
		SELECT FldName, RmtFldname, DataType, RmtNull, RmtType FROM (lcEnumFieldsTbl) WHERE RTRIM(TblName)==lcTableName ;
		    INTO ARRAY aTblFields

        lnOldArea=SELECT()
        SELECT (lcCursorName)
        lcRemoteName=THIS.RemotizeName(lcTableName)

        *Thermometer stuff
        lnRecs=RECCOUNT()
        lnRecordsCopied=0
*** 11/20/2012: wrap this code in IF so we only update the progress if we're supposed to
        if tlUpdateProgress
	        THIS.InitTherm(SEND_DATA_LOC,lnRecs,0)
	        lcMsg=STRTRAN(THIS_TABLE_LOC,'|1',lcTableName)
	        THIS.UpDateTherm(lnRecordsCopied,lcMsg)
        endif tlUpdateProgress

        *Use the remote field name here
*** DH 2015-09-25: handle brackets already in name
***        lcInsertString = 'INSERT INTO ['+ALLTRIM(lcRmtTableName) + '] ('
		if left(lcRmtTableName, 1) = '['
			lcInsertString = 'INSERT INTO ' + alltrim(lcRmtTableName) + ' ('
		else
			lcInsertString = 'INSERT INTO [' + alltrim(lcRmtTableName) + '] ('
		endif left(lcRmtTableName, 1) = '['
        lnFieldCount = ALEN(aTblFields,1)
        FOR ii = 1 TO lnFieldCount
*** DH 2015-09-25: handle brackets already in name
***            lcInsertString = lcInsertString + "[" + ALLTRIM(aTblFields[ii,2]) + "]" ;
                + IIF( ii < lnFieldCount, ', ',' )')
			lcField = alltrim(aTblFields[ii,2])
			if left(lcField, 1) = '['
				lcInsertString = lcInsertString + lcField ;
					+ iif(ii < lnFieldCount, ', ',' )')
			else
				lcInsertString = lcInsertString + '[' + lcField + ']' ;
					+ iif(ii < lnFieldCount, ', ',' )')
			endif left(lcField, 1) = '['
        ENDFOR

        lcInsertString = lcInsertString + ' VALUES ('
        *{ JEI RKR 24.11.2005 Add

        lcEmptyDateValue = ""
        lcEmptyDateValueEnd = ""
        *} JEI RKR 24.11.2005 Add
        *Use the local field name here
        lcInsertFinal = lcInsertString
        FOR ii = 1 TO lnFieldCount
            SELECT (lcCursorName)
            *{ JEI RKR 24.11.2005 Add  
			do case
				case upper(alltrim(aTblFields[ii, 3])) $ 'DT' and aTblFields[ii, 4]
*** DH 2015-08-10: Jim Nelson reported this fails if BlankDateValue is null and
*** SET NULLDISPLAY is blank so handle it explicitly.
***					lcEmptyDateValue = "IIF(EMPTY(" + ALLT(lcCursorName)+'.'+ALLT(aTblFields[ii,1])+ "), " + ;
						transform(This.BlankDateValue) + ","
					lcEmptyDateValue = "IIF(EMPTY(" + ALLT(lcCursorName)+'.'+ALLT(aTblFields[ii,1])+ "), " + ;
						iif(isnull(This.BlankDateValue), '.null.', transform(This.BlankDateValue)) + ","
        			lcEmptyDateValueEnd = ")"    	
				case upper(alltrim(aTblFields[ii, 3])) $ 'YBFIN'
        			IF aTblFields[ii,4] = .t. && Allow Null
        				lcEmptyDateValue = "IIF([*] $  TRANSFORM(" + ALLT(lcCursorName)+'.'+ALLT(aTblFields[ii,1]) + "), .Null.,"
        			ELSE
        				lcEmptyDateValue = "IIF([*] $  TRANSFORM(" + ALLT(lcCursorName)+'.'+ALLT(aTblFields[ii,1]) + "), 0,"
        			ENDIF
        			lcEmptyDateValueEnd = ")"
				case lower(aTblFields[ii, 5]) = 'varchar'
			    	lcEmptyDateValue = 'trim('
					lcEmptyDateValueEnd = ')'
			endcase
            *{ JEI RKR 24.11.2005 Add  
            lcInsertFinal = lcInsertFinal + RMT_OPERATOR + lcEmptyDateValue + ALLT(lcCursorName) ;
                + '.'+ALLT(aTblFields[ii,1])  + lcEmptyDateValueEnd ;
                + IIF(ii < lnFieldCount,' , ',' )')

            *{ JEI RKR 24.11.2005 Add   
          	lcEmptyDateValue = ""
	        lcEmptyDateValueEnd = ""
	        *} JEI RKR 24.11.2005 Add
        ENDFOR

        *Set maximum number of errors allowed so user's disk doesn't fill up if
        *something goes wrong over and over
        lnMaxErrors=lnRecs*DATA_ERR_FRACTION
        IF lnMaxErrors<DATA_ERR_MIN THEN
            lnMaxErrors=DATA_ERR_MIN
        ENDIF

        I=0
        lnExportErrors=0

        *undone: need to set number of inserts per batch

        SCAN
            IF !THIS.ExecuteTempSPT(lcInsertFinal,@lnSQLErrno,@lcSQLErrMsg) THEN
                COPY TO ARRAY aErrData NEXT 1
                THIS.DataExportErr(@aErrData, lcTableName, @lcDataErrTable, lnSQLErrno, lcSQLErrMsg)
                lnExportErrors=lnExportErrors+1
            ENDIF

            IF lnExportErrors>lnMaxErrors THEN
*** 11/20/2012: wrap this code in IF so we only update the progress if we're supposed to
		        if tlUpdateProgress
        	        THIS.UpDateTherm(lnRecordsCopied,CANCELED_LOC)
		        endif tlUpdateProgress
                llMaxErrExceeded=.T.
                EXIT
            ENDIF

            IF I=STAGGER_COUNT THEN
                lnRecordsCopied=lnRecordsCopied+STAGGER_COUNT
*** 11/20/2012: wrap this code in IF so we only update the progress if we're supposed to
		        if tlUpdateProgress
	                IF lnExportErrors<>0 THEN
	                    THIS.UpDateTherm(lnRecordsCopied,lcMsg + ", " + LTRIM(STR(lnExportErrors))+ " " + ERROR_COUNT_LOC)
	                ELSE
	                    THIS.UpDateTherm(lnRecordsCopied, lcMsg)
	                ENDIF
		        endif tlUpdateProgress
                I=1
            ELSE
                I=I+1
            ENDIF
        ENDSCAN

*** 11/20/2012: wrap this code in IF so we only update the progress if we're supposed to
***        IF !llMaxErrExceeded THEN
        IF !llMaxErrExceeded and tlUpdateProgress
			raiseevent(This, 'CompleteProcess')
        ENDIF

        *Close the errors table if it exists
        IF lnExportErrors<>0  AND lnNumberOfPar < 5 THEN && JEI RKR 06.01.2006 Add lnNumberOfPar < 5
            SELECT (lcDataErrTable)
            USE
        ENDIF
        SELECT (lnOldArea)

		tcDataErrTable = lcDataErrTable

        RETURN lnExportErrors

    ENDFUNC


    PROCEDURE DataExportErr
        PARAMETERS aErrData, lcTableName, lcDataErrTable, lnSQLErrno, lcSQLErrMsg
        LOCAL lnAlen, lnArrayPos, lnOldArea, lcExact
	
        *If record(s) is not exported successfully, it's placed in an error table

        lnOldArea=SELECT()

        IF EMPTY(lcDataErrTable) THEN
            IF !THIS.DataErrors THEN
                DIMENSION This.aDataErrTbls [1,3]
                aDatErrTbls=.F.
                THIS.DataErrors=.T.
            ELSE
                DIMENSION This.aDataErrTbls [ALEN(This.aDataErrTbls,1)+1,3]
            ENDIF
            lnAlen=ALEN(This.aDataErrTbls,1)
            This.aDataErrTbls[lnAlen,1]=lcTableName
            lcDataErrTable=THIS.UniqueTableName(ERRT_LOC + LTRIM(STR(lnAlen)))
            This.aDataErrTbls[lnAlen,2]=lcDataErrTable

            *Create table with same structure
           	*{ JEI RKR 24.11.2005 Change
           	LOCAL lcTempTableName as String
           	lcTempTableName = THIS.UniqueTableName(ERRT_LOC + "TableStructure")
           	COPY STRUCTURE EXTENDED TO &lcTempTableName
            SELECT 0
            USE (lcTempTableName)
            Replace FIELD_NEXT WITH 0,;
					FIELD_STEP WITH 0 ALL IN (ALIAS())
 
			USE
			CREATE &lcDataErrTable. FROM &lcTempTableName.
*** DH 07/02/2013: ensure DBF extension is included or the file isn't erased.
*			ERASE (lcTempTableName)
			ERASE (forceext(lcTempTableName, 'dbf'))
			ERASE (FORCEEXT(lcTempTableName, "fpt"))
*            COPY STRUCTURE TO &lcDataErrTable
            *{ JEI RKR 24.11.2005 Change
            USE (lcDataErrTable) EXCLUSIVE

            *May not be able to store the error if the table has 255 fields already
            IF THIS.CountFields(lcDataErrTable) < MAX_FIELDS THEN
                This.aDataErrTbls[lnAlen,3]=.T.		&&this column says whether it's <255 or not
                ALTER TABLE (lcDataErrTable) ADD COLUMN SQLErrNo N(10,0)
                ALTER TABLE (lcDataErrTable) ADD COLUMN SQLErrMsg M
            ENDIF
        ELSE
            SELECT (lcDataErrTable)
        ENDIF

        *If a field has numeric overflow condition, can get an error here; ignore it
        THIS.SetErrorOff=.T.
        APPEND FROM ARRAY aErrData
        THIS.SetErrorOff=.F.

        lcExact=SET('EXACT')
        SET EXACT ON
        IF This.aDataErrTbls[ASCAN(This.aDataErrTbls,lcDataErrTable)+1] THEN
            REPLACE SQLErrNo WITH lnSQLErrno, SQLErrMsg WITH lcSQLErrMsg
        ENDIF
        SET EXACT &lcExact

        SELECT (lnOldArea)

        *The data export routine will close the error table, so it's not closed here


    ENDPROC



    FUNCTION FastExport
        PARAMETERS lcTableName, lcCursorName, lcRmtTableName, llMaxErrExceeded
        LOCAL lnBigRows, lnSmallRows, lnBigBlocks, lnSmallBlocks, llTStamp, lnExportErrors

        *
        *This thing creates stored procedures that slam data in batches
        *(i.e. the sproc has multiple insert statements
        *

        *lnBigRows is the number of insert statements to put into the sproc; this
        *depends on how many fields are in the table
        *
        *lnBigBlocks is the number of times to execute the sproc
        *
        *lnSmallRows is basically the number of rows left after the sproc has
        *been executed lnBigBlocks number of times
        *
        *lnSmallBlocks is either 0 (if lnBigRows divides evenly into the number of
        *records in a table) or 1
        *
        lnSmallRows=0
        lnBigBlocks=0
        lnBigRows=0
        lnSmallBlocks=0

        *See if the table has a timestamp field
        *(Note: this proc is getting called from within a SCAN/ENDSCAN so we
        *know we're on the right record already
        llTStamp=TStampAdd

        *Figure out how many rows to send at a time
        THIS.RowHeuristics(lcCursorName, @lnBigRows, ;
            @lnSmallRows, @lnBigBlocks, @lnSmallBlocks)

        *Create stored procedures that do the inserts
        THIS.CreateSQLServerSproc(lcTableName,lcRmtTableName, lnBigRows, lnSmallRows, llTStamp)

        *Execute the stored procedures
        lnExportErrors=THIS.ExecuteSproc(lcTableName, lcCursorName, lcRmtTableName, lnBigBlocks, ;
            lnSmallBlocks, lnBigRows, lnSmallRows, llTStamp, @llMaxErrExceeded)

        RETURN lnExportErrors

    ENDFUNC


    FUNCTION BulkInsert
        * Assuption: We won't get here if this table has image, text, or non-char nulls
        * columns. Also assume we are upsizing to SQL 7+.
        LPARAMETERS tcTableName, tcCursorName, tcRmtTableName, tlMaxErrExceeded
*** DH 11/20/2013: added llRetVal
*        LOCAL lnServerError, lcErrMsg, lcBulkInsertString, lcBulkFileNameWithPath, lnUserDecision
        LOCAL lnServerError, lcErrMsg, lcBulkInsertString, lcBulkFileNameWithPath, lnUserDecision, llRetVal

        * See if the table has a timestamp and identity field
        * (Note: this proc is getting called from within a SCAN/ENDSCAN so we
        * know we're on the right record already.

        lnServerError = 0
        lcErrMsg = ""
*** DH 11/20/2013: removed duplicate assignment
*        lnServerError = 0

        * 06/23/01 jvf Account for bulk inserting into a remote machine. Use UNC to specify source path.
        * lcBulkFileNameWithPath = ;
        *	"\\"+ ADDBS(ALLTRIM(LEFT(sys(0), AT("#",sys(0))-1)))+StrTran(SYS(5),":","$") + ;
        *	CURDIR() + BULK_INSERT_FILENAME

        * JVF 10/30/02 Issue 16653: Invalid path of filename on COPY TO (tcTargetTextFile).
        * On 6/23/01 when I tried to fix the bug that occurs on COPY TO when your default directory is a network drive
        * I made the assumption that the network drive will be a share. The following much simpler code seems to work in
        * all situations, namingly: default dir is local, default dir is a mapped network drive, default dir is an unmapped
        * network drive (UNC). (6/23/01 code commented out.)

        lcBulkFileNameWithPath = ADDBS(FULLPATH(CURDIR())) + BULK_INSERT_FILENAME

        * Clean up possible previous bulkins.out file.
        IF FILE(lcBulkFileNameWithPath)
            DELETE FILE (lcBulkFileNameWithPath)
        ENDIF

        IF THIS.GenBulkInsertTextFile(tcTableName, tcCursorName, lcBulkFileNameWithPath)
            lcBulkInsertString = THIS.GetBulkInsertString(tcRmtTableName, lcBulkFileNameWithPath)
            * Execute the BULK INSERT into SQL 7 database
            IF THIS.ExecuteTempSPT(lcBulkInsertString, @lnServerError, @lcErrMsg)
                llRetVal = .T.
*** DH 12/06/2012: handle GenBulkInsertTextFile failing like ExecuteTempSPT failing: ask the
*** user for another mechanism, so comment out ELSE and add ENDIFs to close the above two IFs.
*            ELSE
			endif
		endif

*** DH 07/02/2013: erase the bulk insert file after we're done with it.

        IF FILE(lcBulkFileNameWithPath)
            DELETE FILE (lcBulkFileNameWithPath)
        ENDIF

*** DH 07/02/2013: added check for llRetVal or tells users bulk insert failed even if worked
		if not llRetVal
            	IF (this.UserUpsizeMethod == 0) THEN
*** DH 2015-09-14: don't prompt user if lQuiet is .T.
					if This.lQuiet
						lnUserDecision = 6
					else
						lnUserDecision = messagebox('The Upsizing Wizard was ' + ;
							'unable to upsize the data using the bulk insert ' + ;
							'technique. However, the data still may be ' + ;
							'upsized using the fast export method.' + chr(13) + ;
							 chr(13) + 'Yes : Use fast export for all tables' + ;
							 chr(13) + 'No : Skip data export for current table' + ;
							 chr(13) + 'Cancel : Skip exporting data', 35, 'Upsize Wizard')
					endif This.lQuiet
            	ELSE
            		lnUserDecision = this.UserUpsizeMethod
            	ENDIF
            	
            	IF (lnUserDecision == 7) THEN
            		RETURN lnServerError
            	ENDIF
            	
            	IF (lnUserDecision == 2) THEN
            		RETURN -1
            	ENDIF
            	
            	IF (lnUserDecision == 6) THEN
            		This.UserUpsizeMethod = 6
            	ENDIF
            	
            	* JEI RKR 18.11.2005 Change: Skip all Bulk insert error
*            	IF (lnServerError <> 4861) THEN
*** DH 12/06/2012: handle lnServerError = 0
*				IF !BETWEEN(lnServerError , 4860, 4882)
				IF lnServerError <> 0 and !BETWEEN(lnServerError , 4860, 4882)
            		RETURN lnServerError
            	ENDIF
            	&& chandsr added code here.
            	&& If bulkinsert failed due to the fact that the SQL server is on a remote box and the bulk insert file is on the local box
            	&& we retry the operation with fastexport. This would ensure that we export data albeit slower than bulkinsert.
            	lnServerError = this.FastExport (tcTableName, tcCursorName, tcRmtTableName, @tlMaxErrExceeded)
            	&& chandsr added code here
*** DH 12/06/2012: comment out the ENDIFs because ENDIFs were added above
*            ENDIF
*** DH 07/02/2013: uncommented one of the ENDIFs due to added IF above.
        ENDIF
        RETURN lnServerError
    ENDFUNC

    FUNCTION GenBulkInsertTextFile
        * This function generate a text file that will be the source of the BULK INSERT statement
        * The fastest technique is to use COPY TO, and then clean it up to support BULK INSERT.
        * So far, that means removing double quotes and filling in empty dates.
        LPARAMETERS tcTableName, tcCursorName, tcTargetTextFile
        LOCAL lcFieldList, llRetVal, lnOldSele, lcMsg, lnRecordsCopied, lnRecs, lnxx, lnBytes, ;
            lcFileStr, lcEnumTablesTbl, lcCursor, llHasTimeStamp, llHasIdentity, lnCodePage
		local laFieldNames[1], ;
			lnI, ;
			lcField, ;
			lnField
        lnOldSele = SELECT()

        SELECT (tcTableName)

        *Thermometer stuff
        lnRecs=RECCOUNT()
        lnRecordsCopied=0
        THIS.InitTherm(SEND_DATA_LOC,lnRecs,0)
        lcMsg=STRTRAN(THIS_TABLE_LOC,'|1',tcTableName)
        THIS.UpDateTherm(lnRecordsCopied,lcMsg)

        * If adding timestamp or identity, add additional placeholders (more delimiters)
        * To accomplish this we add additional fields to the source cursor so that COPY TO will
        * automatically add additional placeholders and BULK INSERT will succeed.

        lcEnumTablesTbl = RTRIM(THIS.EnumTablesTbl)

* Create an array of fields being upsized to Varchar.

		select FLDNAME from (THIS.EnumFieldsTbl) ;
			into array laFieldNames ;
			where trim(TblName) == lcTableName and RmtType = 'varchar'

* Create a list of fields to output, trimming any Varchar values.

		lcFieldList = ''
		for lnI = 1 to afields(laFields)
			lcField = laFields[lnI, 1]
			lnField = ascan(laFieldNames, lcField, -1, -1, 1, 15)
			if lnField > 0
				lcField = 'cast(trim(' + lcField + ') as V(' + ;
					transform(laFields[lnI, 3]) + ')) as ' + lcField
			endif lnField > 0
			lcFieldList = lcFieldList + iif(empty(lcFieldList), '', ',') + ;
				lcField
		next lnI
        IF &lcEnumTablesTbl..TStampAdd OR (THIS.TimeStampAll=1)
            lcFieldList = lcFieldList + ", '' as tstampcol "
            llHasTimeStamp = .T.
        ENDIF
        IF &lcEnumTablesTbl..IdentAdd OR (THIS.IdentityAll=1)
            lcFieldList = lcFieldList + ", '' as identcol "
            llHasIdentity = .T.
        ENDIF

        lcCursor = THIS.UniqueCursorName(tcCursorName)
        SELECT &lcFieldList FROM (tcTableName) INTO CURSOR (lcCursor)

		*{ JEI RKR - 2005.04.02 Add Code page
		* Get VFP code page
		lnCodePage = CPCURRENT(1)
        * Create the text file
*** MJP 07/15/2013: don't put delimiters around the character fields.
*        COPY TO (tcTargetTextFile) DELIMITED WITH CHARACTER BULK_INSERT_FIELD_DELIMITER AS lnCodePage
        COPY TO (tcTargetTextFile) DELIMITED WITH "" WITH CHARACTER BULK_INSERT_FIELD_DELIMITER AS lnCodePage
   		*{ JEI RKR - 2005.04.02
        USE IN (lcCursor)
        
*** DH 12/06/2012: ensure the file isn't too large
		adir(laFiles, tcTargetTextFile)
		if laFiles[1, 2] > 30000000
			llRetVal = .F.
			erase (tcTargetTextFile)
		else
*** DH 12/06/2012: end of new code

	        * Clean up bulk insert file:
	        * - Remove double quotes - BULK INSERT treats them literally
*** MJP 07/15/2013: if we don't add delimiters in the first place, there's no need to
***	remove them.  But we still need to read the file into a variable.
*	        lcFileStr = STRTRAN(FILETOSTR(tcTargetTextFile), ["])
            lcFileStr = FILETOSTR(tcTargetTextFile)

	        * - COPY TO leaves blank dates as such: /  /, so
	        *   make empty dates 01/01/1900 - the SQL 7 default
	        lcFileStr = STRTRAN(lcFileStr, "/  /", SQL_SERVER_EMPTY_DATE_CHAR)
	        * 1/6/01 jvf Convert VFP logicals to 0 and 1 for SQL Server.
	        * - COPY TO makes VFP logicals as such: F and T, so
	        *   make them 0 and 1, resp. for SQL Server
	        lcFileStr = THIS.ConvertVFPLogicalToSQLServerBit(tcTableName, lcFileStr, llHasTimeStamp, llHasIdentity)

	        lnBytes = STRTOFILE(lcFileStr, tcTargetTextFile)
	        llRetVal = (lnBytes > 0)
*** DH 12/06/2012: ENDIF to match IF added above
		endif laFiles[1, 2] > 30000000

        SELECT (lnOldSele)
        RETURN llRetVal
    ENDFUNC

    FUNCTION ConvertVFPLogicalToSQLServerBit
        LPARAMETERS tcTableName, lcFileStr, llHasTimeStamp, llHasIdentity
*** DH 07/11/2013: rewrote this to be more efficient and handle files > 16 MB in size
*!*	        LOCAL ii, jj, lnLines, lnPos, lcValue, llPosAdj
*!*	        lnLines = ALINES(laLines, lcFileStr)
*!*	        lnSele = SELECT()
*!*	        SELECT (tcTableName)
*!*	        llPosAdj = 0
*!*	        llPosAdj = llPosAdj + IIF(llHasTimeStamp, 1, 0)
*!*	        llPosAdj = llPosAdj + IIF(llHasIdentity, 1, 0)
*!*	        FOR ii = 1 TO FCOUNT()
*!*	            IF (TYPE(FIELD(ii)) = "L")
*!*	                FOR jj = 1 TO lnLines
*!*	                    IF jj = 1
*!*	                        lnPos = AT("~", lcFileStr, (ii-1)*jj)+1
*!*	                    ELSE
*!*	                        llNewPosAdj = (llPosAdj + FCOUNT()-1) * (jj-1)
*!*	                        lnPos = AT("~", lcFileStr, (ii-1+llNewPosAdj))+1
*!*	                    ENDIF
*!*	                    lcValue = SUBSTR(lcFileStr, lnPos, 1)
*!*	                    lcFileStr = STUFF(lcFileStr, lnPos, 1, IIF(lcValue = "T", "1", "0"))
*!*	                ENDFOR
*!*	            ENDIF
*!*	        ENDFOR
*!*	        SELECT (lnSele)
		local laLines[1], lnLines, laFields[1], lnFields, ii, jj, lcLine, lnPos, lcValue
		lnLines  = alines(laLines, lcFileStr)
		lnFields = afields(laFields, tcTableName)
		for ii = 1 to lnFields
			if laFields[ii, 2] = 'L'
				for jj = 1 to lnLines
					lcLine      = laLines[jj] + '~'
					lnPos       = at('~', lcLine, ii)
					lcValue     = substr(lcLine, lnPos - 1, 1)
					laLines[jj] = left(stuff(lcLine, lnPos - 1, 1, ;
						iif(lcValue = 'T', '1', '0')), len(lcLine) - 1)
				next jj
			endif laFields[ii, 2] = 'L'
        next ii
		lcFileStr = ''
		for jj = 1 to lnLines
			lcFileStr = lcFileStr + laLines[jj] + chr(13) + chr(10)
		next jj
*** DH 07/11/2013: end of new code
        RETURN lcFileStr
    ENDFUNC

    FUNCTION GetBulkInsertString
        * jvf 9/1/99 Build bulk insert string per export table
        * Used COPY TO (textfile) DELI to create source file
        * (while extracting binary and memo fields).
        LPARAMETERS tcRmtTableName, tcSourceTextFile
        lcBulkInsStr = "BULK INSERT " + tcRmtTableName + ;
            " FROM '" + tcSourceTextFile + "' WITH " + ;
            " (CODEPAGE = 'ACP', " + ;
            " FIELDTERMINATOR = '" + [BULK_INSERT_FIELD_DELIMITER] + "', " + ;
            " TABLOCK, KEEPIDENTITY, KEEPNULLS)"

        RETURN lcBulkInsStr
    ENDFUNC

    PROCEDURE CreateSQLServerSproc
        PARAMETERS lcTableName, lcRmtTableName, lnBigRows, lnSmallRows, llTStamp

        LOCAL lcEnumFields, lcSQL, lcParamString, lcInsertString, ;
            lcEnumTables, lcCR, lcLF, lcParamName, lcFieldName, llSmallCondition, ;
            lcType, lcLength, I, ii, j, lcSmallParamString, lcSmallSQL, llRetVal, ;
            lnOldArea, LCAS, lcParamChar, lcDecimals, lcColumnsList

        #DEFINE SQL_PARAM_CHAR			"@a"
        #DEFINE ORACLE_PARAM_CHAR		"x"
        #DEFINE BIG_PARAM "Big"

        *If the table has no data, don't create the sproc
        IF lnBigRows=0 	THEN
            RETURN
        ENDIF

        lnOldArea=SELECT()
        lcEnumTables=THIS.EnumTablesTbl

        lcInsertString=""
        lcParamString=""
        lcCR=CHR(10)
        lcLF=CHR(13)
        lnOldArea=SELECT()

        lcEnumFields=THIS.EnumFieldsTbl
        SELECT (lcEnumFields)
        COPY TO ARRAY aFieldNames FIELDS RmtType, RmtLength, RmtPrec, fldname FOR RTRIM(TblName)==lcTableName

        lcSQL="CREATE PROCEDURE " + DATA_PROC_NAME

        *Oracle () around the whole parameter string

        IF THIS.ServerType="Oracle" THEN
            *Oracle doesn't allow params to begin with "@"
            lcParamString=" (" + ORACLE_PARAM_CHAR + BIG_PARAM + " CHAR, "
            lcParamChar=ORACLE_PARAM_CHAR
        ELSE
            *SQL Server requires that params begin with "@"
            lcParamString=" " + SQL_PARAM_CHAR + BIG_PARAM + " char(4), "
            lcParamChar=SQL_PARAM_CHAR
        ENDIF

        *@Big parameter is used to execute all (for "big" inserts) or part (for "small")
        *inserts) of the insert statements in the sprocs that get created.  This way
        *we don't need two sprocs; all other parameters are for data.  This means
        *we can only send 254 data parameters at a time so this routine only works
        *for tables with 254 or less fields
        
        lcColumnNames = " ("
        
        FOR iRow = 1 TO ALEN(aFieldNames) / 4 STEP 1				&& compute the number of rows by dividing the length by the number of columns in the array
        	* JEI RKR 18.11.2005 Add "[" and "]" for field with name like User
        	lcColumnNames = lcColumnNames + "[" + ALLTRIM (aFieldNames(iRow, 4)) + "]"

        	IF (iRow < ALEN (aFieldNames) / 4) THEN
	        	lcColumnNames = lcColumnNames + ","
	        ELSE
	        	lcColumnNames = lcColumnNames + ") "
	        ENDIF
        ENDFOR

        j=0
        
        FOR I=1 TO lnBigRows
			*{ JEI RKR 18.11.2005 Add "[" and "]"
			lcRmtTableName = ALLTRIM(lcRmtTableName)
			IF LEFT(lcRmtTableName, 1) <> "["
				lcRmtTableName = "[" + lcRmtTableName + "]"
			ENDIF
			*} JEI RKR 18.11.2005 Add "[" and "]"
            lcInsertString=lcInsertString + "INSERT INTO " + RTRIM(lcRmtTableName) + lcColumnNames + ;
                lcCR + "VALUES " + "("

            FOR ii=1 TO ALEN(aFieldNames,1)
                j=j+1
                lcParamName=lcParamChar + LTRIM(STR(j))
                lcType=RTRIM(aFieldNames[ii,1])
                lcLength=LTRIM(STR(aFieldNames[ii,2]))
                lcDecimals = LTRIM(STR(aFieldNames[ii,3]))
                IF lcLength="0" OR THIS.ServerType="Oracle" THEN
                    *Note: Oracle parameters take a datatype but no length or precision
                    lcLength=""
                ELSE
                    lcLength= "(" + lcLength + IIF(lcDecimals # '0',"," + lcDecimals,"") + ")"
                ENDIF

                *builds sproc param string like "@a1 char(4), @a2 text, " etc. or
                *"x1 varchar2, x2 number, " etc. for Oracle
                
                && chandsr added code. 
                lctype = SUBSTR(lctype, 1, IIF(AT('(', lcType) = 0, LEN(lcType), AT ('(', lcType) - 1) ) 
                && Complicated expression just for fun (just kidding).
                && if expression  is of the form int (ident) the result of executing this expression would be int.
                && this has been done to ensure that the datatype definition such as the one mentioned above do not cause errors on SQL server.
                && chandsr added code.
                
                lcParamString=lcParamString + lcParamName + " " + lcType + lcLength + ", "

                *builds list of values for INSERT string like "@1, @2, " etc.
                lcInsertString=lcInsertString  + lcParamName + ", "

            NEXT ii

            *If column has a timstamp column, add parameter for that
            *IF llTStamp THEN
            *	lcInsertString=lcInsertString + "NULL)" + lcCR +lcLF
            *ELSE
            *otherwise, peel off last comma, add closing paren
            lcInsertString=LEFT(lcInsertString,LEN(lcInsertString)-2)
            *Oracle likes a semicolon after each SQL command
            IF THIS.ServerType="Oracle" THEN
                lcInsertString=lcInsertString  + ");" + lcCR +lcLF
            ELSE
                lcInsertString=lcInsertString  + ")" + lcCR +lcLF
            ENDIF
            *ENDIF

            *if the sproc code includes enough row inserts for the "small"
            *insert sproc, tack on code that will skip over the remaining inserts
            IF I=lnSmallRows THEN
                lcParamName=lcParamChar + BIG_PARAM
                lcInsertString=lcInsertString  + "IF " + lcParamName + " = " + ;
                    "'" + "TRUE" + "'" + lcCR + lcLF
                IF THIS.ServerType="Oracle" THEN
                    lcInsertString=lcInsertString  + "THEN" + lcCR
                ELSE
                    lcInsertString=lcInsertString  + "BEGIN" + lcCR
                ENDIF
                llSmallCondition=.T.
            ENDIF

        NEXT I

        *peel off last comma, put whole string together
        lcParamString=LEFT(lcParamString,LEN(lcParamString)-2)

        IF THIS.ServerType="Oracle" THEN
            *Add closing paren, BEGIN statement
            lcParamString=lcParamString + ")"
            LCAS= " AS BEGIN"
        ELSE
            LCAS= " AS "
        ENDIF

        lcSQL=lcSQL + lcParamString + LCAS + lcCR + lcLF + lcInsertString
        IF llSmallCondition OR THIS.ServerType="Oracle" THEN
            IF THIS.ServerType="Oracle" THEN
                *The Oracle IF..THEN construct requires an "END IF"
                IF llSmallCondition THEN
                    lcSQL=lcSQL + "END IF;" + lcCR + lcLF
                ENDIF
                *Oracle sprocs have to be ended explicitly
                lcSQL=lcSQL + "END " + DATA_PROC_NAME + ";"
            ELSE
                lcSQL=lcSQL + "END"
            ENDIF
        ENDIF

        *drop any existing sproc
        =THIS.ExecuteTempSPT("drop procedure " + DATA_PROC_NAME)

        *Create the stored procedure
        llRetVal=THIS.ExecuteTempSPT(lcSQL)

        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE RowHeuristics
        PARAMETERS lcCursorName, lnBigRows, lnSmallRows,  lnBigBlocks, lnSmallBlocks
        LOCAL lnReccount, lnFieldCount, lnTotalParams, lnOldArea

        *Figures out how many rows stored procedures should insert at a time

        lnOldArea=SELECT()
        SELECT (lcCursorName)

        *used to be lnRecCount=reccount() but that counts deleted records too
        SELECT COUNT(*) FROM (lcTableName) WHERE !DELETED() INTO ARRAY aReccount
        lnReccount=aReccount
        lnFieldCount=AFIELDS(aTemp)
        lnTotalParams=aReccount*lnFieldCount
        SELECT (lnOldArea)

        IF lnTotalParams<=MAX_PARAMS THEN
            IF lnReccount<=MAX_ROWS
                lnBigRows=lnReccount
                lnBigBlocks=1
                lnSmallRows=0
                lnSmallBlocks=0
            ELSE
                lnBigRows=MAX_ROWS
                lnBigBlocks=1
                lnSmallRows=lnReccount-MAX_ROWS
                lnSmallBlocks=1
            ENDIF
        ELSE
            IF INT(MAX_PARAMS/lnFieldCount)> MAX_ROWS
                lnBigRows=MAX_ROWS
                lnBigBlocks=INT(lnReccount/lnBigRows)
                lnSmallRows=lnReccount-(lnBigBlocks*lnBigRows)
                lnSmallBlocks=1
            ELSE
                lnBigRows=INT(MAX_PARAMS/lnFieldCount)
                lnBigBlocks=INT(lnReccount/lnBigRows)
                lnSmallRows=lnReccount-(lnBigBlocks*lnBigRows)
                lnSmallBlocks=1
            ENDIF
        ENDIF

    ENDPROC


    FUNCTION ExecuteSproc
        PARAMETERS lcTableName, lcCursorName, lcRmtTableName, lnBigBlocks, lnSmallBlocks, ;
            lnBigRows, lnSmallRows, llTStamp, llMaxErrExceeded

        LOCAL lnOldArea, lnNumberOfFields, lnLoopMax, lnRecs, lcMsg, ;
            lcNumberofRows, lcArrayType, llRetVal, lcLoopLimiter, lnRowsToCopy, ;
            lnRecordsCopied, lcSQL, lnExportErrors, lcDataErrTable, lnMaxErrors, ;
            llBail, lcData, lnSQLErrno, lcSQLErrMsg, I, ii
		local laFieldNames[1], ;
			lcField, ;
			lnField, ;
			lcValue
*** DH 07/02/2013: added more variables to LOCAL
		local laFields[1], lnFields

		*{ JEI RKR  23.11.2005 ADD 
		LOCAL lcDateFieldList, lcNumericFieldsListAllowNull, lcNumericFieldsListNotAllowNull, lcCurentField, lcTempCursorName as String
		LOCAL lnCurrentField, lnCurrentRow, lnCurrentRowInArray, lnRes, lnOldWorkArea as Integer
		*} JEI RKR  23.11.2005 ADD 

        PRIVATE aBigArray, aSmallArray

        *If the table has no data, bail
        IF lnBigRows=0 THEN
            RETURN 0
        ENDIF

        lnOldArea=SELECT()
        SELECT (lcCursorName)
        GO TOP

        lnNumberOfFields=AFIELDS(aTemp)

        *Thermometer stuff
        lnRecs=RECCOUNT()
        lnRecordsCopied=0
        THIS.InitTherm(SEND_DATA_LOC,lnRecs,0)
        lcMsg=STRTRAN(THIS_TABLE_LOC,'|1',lcTableName)
        THIS.UpDateTherm(lnRecordsCopied,lcMsg)

        *Set maximum number of errors allowed so user's disk doesn't fill up if
        *something goes wrong over and over
        lnMaxErrors=lnRecs*DATA_ERR_FRACTION
        IF lnMaxErrors<DATA_ERR_MIN THEN
            lnMaxErrors=DATA_ERR_MIN
        ENDIF
        lnExportErrors=0

        DIMENSION aBigArray[lnBigRows,lnNumberOfFields]

        lcArrayType="aBigArray"
        lcNumberofRows="lnBigRows"

        *Create big strings that look like this:
        *:aBigArray[1,1],:aBigArray[1,2],:aBigArray[1,3],:aBigArray[1,4]

*** DH 11/20/2012: moved this block of code inside main loop because lcNumberOfRows can change
*        lcData=""
*        FOR I=1 TO &lcNumberofRows
*            FOR ii=1 TO lnNumberOfFields
*                IF !lcData=="" THEN
*                    lcData=lcData + " , "
*                ENDIF
*                lcData=lcData+ RMT_OPERATOR + "&lcArrayType[" + LTRIM(STR(I)) + "," + ;
*                    LTRIM(STR(ii)) + "]"
*            NEXT ii
*        NEXT I

        *Set variables for first pass through loop
        lcArrayType="aBigArray"
        lcLoopLimiter="lnBigBlocks"
        lnRowsToCopy=lnBigRows
        lnRecordsCopied=1
        lcBigInsert="TRUE"		&&determines if all inserts in sproc are executed

		*{ JEI RKR  23.11.2005 ADD 
		lcDateFieldList = This.GetColumnNum(lcTableName, "DT", .T.)
		lcNumericFieldsListAllowNull = This.GetColumnNum(lcTableName, "YBFIN", .T.)
		lcNumericFieldsListNotAllowNull = This.GetColumnNum(lcTableName, "YBFIN", .F.)
		*} JEI RKR  23.11.2005 ADD

* Create an array of fields being upsized to Varchar.

		select FLDNAME from (THIS.EnumFieldsTbl) ;
			into array laFieldNames ;
			where trim(TblName) == lcTableName and RmtType = 'varchar'

*** DH 07/02/2013: Used CurrBlock instead of I because I is used for an inner loop as well
*        FOR I=1 TO IIF(lnSmallRows=0,1,2)
        FOR CurrBlock=1 TO IIF(lnSmallRows=0,1,2)
            *In case massive data export errors occurred
            IF lnExportErrors>lnMaxErrors THEN
                THIS.UpDateTherm(lnRecordsCopied,CANCELED_LOC)
                llMaxErrExceeded=.T.
                EXIT FOR
            ENDIF

*** DH 11/20/2012: moved this block of code inside main loop because lcNumberOfRows can change
            *Build the array string
*            IF THIS.ServerType="Oracle" THEN
*                lcSQL="BEGIN " + DATA_PROC_NAME + " (" + "'" + lcBigInsert + "'" + ;
*                    ", " + lcData+ 	"); END;"
*            ELSE
*                lcSQL="EXECUTE " + DATA_PROC_NAME + "'" + lcBigInsert + "' " + ;
*                    ", " + lcData
*            ENDIF

            *Fill it and send it
            FOR ii= 1 TO &lcLoopLimiter
*** DH 11/20/2012: release the array because at the end of the table, we may
*** have fewer records left than existing rows in the array, which means there
*** are superfluous rows in the array.
				release &lcArrayType
                *get a load of data
                COPY TO ARRAY &lcArrayType NEXT lnRowsToCopy
                skip
*** DH 11/20/2012: recalculate lcNumberOfRows so we only work with the number
*** of rows actually in the array. Since lcNumberOfRows may change, we'll also
*** determine lcSQL here rather than earlier.
				lcNumberOfRows = transform(alen(&lcArrayType., 1))
				lcData=""
				FOR I=1 TO &lcNumberofRows
				    FOR ij=1 TO lnNumberOfFields
				        IF !lcData=="" THEN
				            lcData=lcData + " , "
				        ENDIF
				        lcData=lcData+ RMT_OPERATOR + "&lcArrayType[" + LTRIM(STR(I)) + "," + ;
				            LTRIM(STR(ij)) + "]"
				    NEXT ij
				NEXT I
*Build the array string
				IF THIS.ServerType="Oracle" THEN
				    lcSQL="BEGIN " + DATA_PROC_NAME + " (" + "'" + lcBigInsert + "'" + ;
				        ", " + lcData+ 	"); END;"
				ELSE
				    lcSQL="EXECUTE " + DATA_PROC_NAME + "'" + lcBigInsert + "' " + ;
				        ", " + lcData
				ENDIF
*** DH 11/20/2012: end of moved code

				*{ JEI RKR 23.11.2005 Add
				IF !EMPTY(STRTRAN(lcDateFieldList + ;
					lcNumericFieldsListAllowNull + ;
					lcNumericFieldsListNotAllowNull , ",", "")) or ;
					not empty(laFieldNames[1])
					FOR lnCurrentField = 1 TO lnNumberOfFields
						lcCurentField = "," + TRANSFORM(lnCurrentField) + ","
						IF lcCurentField $ lcDateFieldList
							FOR lnCurrentRow = 1 TO &lcNumberofRows.
								IF EMPTY(&lcArrayType.[lnCurrentRow, lnCurrentField])
*** DH 07/02/2013: use the chosen blank date value rather than assuming null
*									&lcArrayType.[lnCurrentRow, lnCurrentField] = .Null.
									&lcArrayType.[lnCurrentRow, lnCurrentField] = This.BlankDateValue
								ENDIF
							ENDFOR
						ELSE
							IF lcCurentField $ "," + lcNumericFieldsListAllowNull + "," + lcNumericFieldsListNotAllowNull + ","	
								FOR lnCurrentRow = 1 TO &lcNumberofRows.				
									IF "*" $ TRANSFORM(&lcArrayType.[lnCurrentRow, lnCurrentField])
										IF lcCurentField $ "," + lcNumericFieldsListAllowNull
											&lcArrayType.[lnCurrentRow, lnCurrentField] = .NULL.
										Else
											&lcArrayType.[lnCurrentRow, lnCurrentField] = 0
										ENDIF
									ENDIF
								ENDFOR
							ENDIF
						ENDIF

* If this field is being upsized to Varchar, trim its values.

						lcField = field(lnCurrentField)
						lnField = ascan(laFieldNames, lcField, -1, -1, 1, 15)
						if lnField > 0
							for lnCurrentRow = 1 to &lcNumberofRows
								lcValue = &lcArrayType.[lnCurrentRow, lnCurrentField]
								&lcArrayType.[lnCurrentRow, lnCurrentField] = trim(lcValue)
							next lnCurrentRow
						endif lnField > 0
					ENDFOR
				ENDIF
				*} JEI RKR 23.11.2005 Add
				
                *send it to the server
                IF !THIS.ExecuteTempSPT(lcSQL, @lnSQLErrno, @lcSQLErrMsg) THEN
                		*{ JEI RKR 05.01.2006 Add and coment
                		lnOldWorkArea = SELECT()
               			lcTempCursorName = SYS(2015)
						lcTempCursorName = This.UniqueTableName(lcTempCursorName)
*** DH 07/02/2013: create a cursor that allows nulls in all fields to avoid an
*** error if the target allows nulls but the source doesn't.
*						SELECT * FROM (lcCursorName) WHERE .F. INTO CURSOR (lcTempCursorName) READWRITE
						lnFields = afields(laFields, lcCursorName)
						for lnField = 1 to lnFields
							laFields[lnField,  5] = .T.
*** DH 07/10/2013: blank out the other values so they don't cause problems (e.g.
*** field rules, triggers, etc.)
							laFields[lnField,  7] = ''
							laFields[lnField,  8] = ''
							laFields[lnField,  9] = ''
							laFields[lnField, 10] = ''
							laFields[lnField, 11] = ''
							laFields[lnField, 12] = ''
							laFields[lnField, 13] = ''
							laFields[lnField, 14] = ''
							laFields[lnField, 15] = ''
							laFields[lnField, 16] = ''
							laFields[lnField, 17] = 0
							laFields[lnField, 18] = 0
						next lnField
						create cursor (lcTempCursorName) from array laFields
*** DH 07/02/2013: end of new code
						SELECT (lcTempCursorName)
						APPEND FROM ARRAY &lcArrayType.

						lnRes = THIS.JimExport(lcTableName, lcTempCursorName, lcRmtTableName, @llMaxErrExceeded, @lcDataErrTable)
						IF lnRes <> 0
							lnExportErrors = lnExportErrors + lnRes
						ENDIF
						IF USED(lcTempCursorName)
							USE IN (lcTempCursorName)
						ENDIF
						SELECT(lnOldWorkArea)
*!*	                    SKIP -1*lnRowsToCopy
*!*	                    COPY TO ARRAY aDatErr NEXT lnRowsToCopy
*!*	                    SKIP
*!*	                    THIS.DataExportErr(@aDatErr, lcTableName, @lcDataErrTable, lnSQLErrno, lcSQLErrMsg)
*!*	                    lnExportErrors=lnExportErrors+lnRowsToCopy
						*} JEI RKR 05.01.2006
                ENDIF

                *Thermometer
                lnRecordsCopied=lnRecordsCopied+lnRowsToCopy

                *If massive export errors, bail
                IF lnExportErrors>lnMaxErrors THEN
                    THIS.UpDateTherm(lnRecordsCopied,CANCELED_LOC)
                    llMaxErrExceeded=.T.
                    EXIT FOR
                ENDIF

                IF lnExportErrors<>0 THEN
                    THIS.UpDateTherm(lnRecordsCopied,lcMsg + ", " + LTRIM(STR(lnExportErrors))+ " " + ERROR_COUNT_LOC)
                ELSE
                    THIS.UpDateTherm(lnRecordsCopied, lcMsg)
                ENDIF

            NEXT ii

            *if there are leftover rows after the "big" copies, do the inner loop again with
            *the smaller array
            *lcArrayType="aSmallArray"
            *The extra parameters not used in the small insert are set to null here

            aBigArray=.NULL.
            lcLoopLimiter="lnSmallBlocks"
            lnRowsToCopy=lnSmallRows
            THIS.BitifyArray(lcTableName, lnSmallRows)
            lcBigInsert="FALS"		&&only takes four characters

        NEXT CurrBlock

        *If export errors occurred, close the error table
        IF lnExportErrors<>0 THEN
			IF USED(lcDataErrTable) && JEI RKR 05.01.2006 ADD
				SELECT (lcDataErrTable)
				USE
			ENDIF
        ENDIF

        IF !llMaxErrExceeded THEN
			raiseevent(This, 'CompleteProcess')
        ENDIF

        *drop the stored procedure
        lcSQL="DROP PROCEDURE rwf_insert_"
        llRetVal=THIS.ExecuteTempSPT(lcSQL)

        SELECT (lnOldArea)

        RETURN lnExportErrors

    ENDFUNC

    PROCEDURE BitifyArray
        PARAMETERS lcTableName, lnRowsToCopy
        LOCAL lnOldArea, lcEnumTablesTbl, lnOldRecno, I, j

        *When ExecuteSproc does the "small" insert, the extra (empty) rows in
        *the array sent to the server have to be acceptable as parameters.
        *Null is fine for everything but bit fields
        lnOldArea=SELECT()

        *Grab the sql (table definition) of the table
        lcEnumTablesTbl=THIS.EnumTablesTbl
        SELECT (lcEnumTablesTbl)
        lnOldRecno=RECNO()
        LOCATE FOR LOWER(RTRIM(TblName))==LOWER(RTRIM(lcTableName))
        lcSQL=TableSQL

        *Parse it looking for bit fields
        lcSQL=SUBSTR(lcSQL,AT("(",lcSQL))

        * jvf 2/4/01 Bug ID 192139 - subscipt outside defined range on: aBigArray[j,i]=.F.
        * Add space(1) after comma as search string to differentiate
        * field separator vs. numeric column struc with precision (eg, numeric(8,2)).
        lcSQL=STUFF(lcSQL,RAT(")",lcSQL),1,", ")
        I=1
        DO WHILE ", " $ lcSQL
            lcSubStr=LEFT(lcSQL,AT(", ",lcSQL))
            IF " bit " $ lcSubStr
                FOR j=lnRowsToCopy +1 TO ALEN(aBigArray,1)
                    aBigArray[j,i]=.F.
                NEXT j
            ENDIF

            * jvf 2/04/01
            * Money wouldn't upsize in FastExport for same reason bit don't.
            * Added this extra loop to prefill array with 0.00 instead of null
            * for money columns.
            IF " money " $ lcSubStr
                FOR j=lnRowsToCopy +1 TO ALEN(aBigArray,1)
                    aBigArray[j,i]=0.00
                NEXT j
            ENDIF

            lcSQL=SUBSTR(lcSQL,AT(", ",lcSQL)+1)
            I=I+1
        ENDDO

        GO lnOldRecno
        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE CreateIndexes
        LOCAL lnOldArea, lcEnum_Indexes, lcSQL, lcScanCondition, I, lnError, ;
            llRetVal, lcClusterName, lnLoopLimiter, lnIndexCount, lcDel, lcErrMsg, ;
            lcTagName, llTableUpsized, lnOldTO, lcMsg

* If we have an extension object and it has a CreateIndexes method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'CreateIndexes', 5) and ;
			not This.oExtension.CreateIndexes(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        lnOldArea=SELECT()
        THIS.InitTherm(STARTING_COMMENT_LOC,0,0)

        *Generate all the sql for the indexes
        THIS.BuildIndexSQL

        *Create the indexes
        IF THIS.DoUpsize AND THIS.Perm_Index  THEN
            *Make sure we don't create indexes with logical/bit fields
            IF THIS.ServerType<>"Oracle" THEN
                THIS.MarkBitIndexes
            ENDIF

            lcEnum_Indexes=THIS.EnumIndexesTbl
            SELECT (lcEnum_Indexes)

            *Create indexes on tables (clustered indexes first, otherwise existing
            *indexes automatically are regenerated)

            IF THIS.ServerType="Oracle" THEN
                lcScanCondition="DontCreate=.F. AND Exported=.F."
                lnLoopLimiter=1
            ELSE
                lcScanCondition="Clustered=.T. AND DontCreate=.F. AND Exported=.F."
                lnLoopLimiter=2
            ENDIF

            *note: if the user selected a table for export, moved ahead a few
            *pages causing the indexes to be analyzed, then deselected that table,
            *then the thermometer reading won't be right--just means that it will
            *jump up fast when some indexes are skipped
            SELECT COUNT(*) FROM (lcEnum_Indexes) WHERE DontCreate=.F. ;
                AND Exported=.F. INTO ARRAY aIndexCount
            lnIndexCount=0
            THIS.InitTherm(STARTING_COMMENT_LOC,aIndexCount,0)

            *Never timeout
            lnOldTO=SQLGETPROP(THIS.MasterConnHand,"QueryTimeOut")
            =SQLSETPROP(THIS.MasterConnHand,"QueryTimeOut",0)

            FOR I=1 TO lnLoopLimiter
                SCAN FOR &lcScanCondition
                    *Make sure table was upsized
                    lcTableName=RTRIM(&lcEnum_Indexes..IndexName)
                    lcErrMsg=""
                    lcMsg=STRTRAN(THIS_TABLE_LOC,"|1",lcTableName)
                    IF THIS.TableUpsized(lcTableName, @llTableUpsized) THEN
                        *for thermometer
                        THIS.UpDateTherm(lnIndexCount,lcMsg)
                        lcSQL=&lcEnum_Indexes..IndexSQL
					
						IF !("DELETED()" $ UPPER(ALLTRIM(&lcEnum_Indexes..Lclexpr))) && JEI RKR 21.11.2005 Add 
	                        llRetVal=THIS.ExecuteTempSPT(lcSQL,@lnError,@lcErrMsg)
	                    ELSE
	                    	&&{ JEI RKR 21.11.2005 Add 
	                    	llRetVal = .t.
	                    	lnError = 0
	                    	&&} JEI RKR 21.11.2005 Add 
	                    ENDIF

                        IF !llRetVal THEN
                            REPLACE &lcEnum_Indexes..IdxErrNo WITH lnError
                            lcTagName=&lcEnum_Indexes..TagName
                            THIS.StoreError(lnError,lcErrMsg,lcSQL,INDEX_FAILED_LOC,lcTagName,INDEX_LOC)
                        ENDIF
                    ELSE
                        llRetVal=.F.
                        lcErrMsg=TABLE_NOT_EXPORTED_LOC
                    ENDIF
                    REPLACE &lcEnum_Indexes..Exported WITH llRetVal, ;
                        &lcEnum_Indexes..IdxError WITH lcErrMsg, ;
                        &lcEnum_Indexes..TblUpszd WITH llTableUpsized

                    lnIndexCount=lnIndexCount+1
                    THIS.UpDateTherm(lnIndexCount, lcMsg)
                ENDSCAN
                lcScanCondition="Clustered=.F. AND DontCreate=.F. AND Exported=.F."
            NEXT

            *Put this back
            IF TYPE('lnOldTO') = 'N' AND lnOldTO >= 0
                =SQLSETPROP(THIS.MasterConnHand,"QueryTimeOut",lnOldTO)
            ENDIF

			raiseevent(This, 'CompleteProcess')
        ENDIF

        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE MarkBitIndexes
        LOCAL lcEnumIndexesTbl, lcEnumFieldsTbl, aLogicals, lnOldArea, llDontCreate, ;
            lcTableName

        *
        *Marks indexes that have logical fields in them as un-createable
        *

        lcEnumFieldsTbl=THIS.EnumFieldsTbl
        lcEnumIndexesTbl=THIS.EnumIndexesTbl

        lnOldArea=SELECT()
        SELECT (lcEnumIndexesTbl)
        SCAN
            lcTableName=RTRIM(IndexName)
            *Find out which (if any) fields in the table are logical fields
            DIMENSION aLogicals[1]
            aLogicals=.F.
            SELECT RmtFldname FROM (lcEnumFieldsTbl) WHERE RTRIM(TblName)==lcTableName AND ;
                DATATYPE="L" INTO ARRAY aLogicals
            llDontCreate=.F.

            *See if any of the logical fields are part of the index expression
            IF !EMPTY(aLogicals) THEN
                FOR jj=1 TO ALEN(aLogicals,1)
                    IF RTRIM(aLogicals[jj]) $ &lcEnumIndexesTbl..RmtExpr THEN
                        llDontCreate=.T.
                        EXIT
                    ENDIF
                NEXT jj
            ENDIF

            IF llDontCreate THEN
                lcMsg=CANT_CREATE_INDEX_LOC
            ELSE
                lcMsg=""
            ENDIF

            REPLACE DontCreate WITH llDontCreate, IdxError WITH lcMsg

        ENDSCAN

        SELECT (lnOldArea)

    ENDPROC


    FUNCTION TableUpsized
        PARAMETERS lcTableName, llChosenForExport
        LOCAL lcEnumTablesTbl, lnOldArea, llExport

        *Returns whether a table was actually created on the server successfully or not
        *If the user is generating a script but not upsizing, then the function
        *returns whether the table was selected for upsizing

        lnOldArea=SELECT()
        lcEnumTablesTbl=RTRIM(THIS.EnumTablesTbl)
        SELECT (lcEnumTablesTbl)
        LOCATE FOR TblName=lcTableName
        IF THIS.DoUpsize THEN
            llExport=&lcEnumTablesTbl..Exported
        ELSE
            llExport=&lcEnumTablesTbl..EXPORT
        ENDIF
        llChosenForExport=&lcEnumTablesTbl..EXPORT
        SELECT (lnOldArea)
        RETURN llExport

    ENDFUNC


    PROCEDURE AnalyzeIndexes
        LOCAL lnOldArea, lcEnum_Tables, lcEnum_Indexes, I, ii, lcExprLeftover, ;
            lcTablePath, lcEnum_Fields, lcExpression, lcTagName, lcRemoteExpression, lcRemoteTagName, ;
            lcSQL, lcRemoteTable, lcClustered, llUserTableOpened, lcTableName, lcMsg, llCreateIndexes, ;
            aKeyFields, llDontCreate, lcErrMsg, lcLclIdxType

        *Don't do this routine if not necessary
        IF THIS.ProcessingOutput THEN
            IF  !THIS.ExportIndexes AND ;
                    !THIS.ExportRelations AND ;
                    !THIS.ExportTableToView AND ;
                    !THIS.ExportViewToRmt THEN
                RETURN
            ENDIF
        ENDIF

* If we have an extension object and it has a AnalyzeIndexes method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'AnalyzeIndexes', 5) and ;
			not This.oExtension.AnalyzeIndexes(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        lnOldArea=SELECT()
        lcEnum_Fields=THIS.EnumFieldsTbl
        lcEnum_Tables=THIS.EnumTablesTbl
        SELECT COUNT(*) FROM (lcEnum_Tables) WHERE EXPORT=.T. AND EMPTY(ClustName)=.T. ;
            AND CDXAnald=.F. INTO ARRAY aTableCount

        *Thermometer stuff
        IF aTableCount=0 THEN
            RETURN
        ENDIF
        lnTableCount=0
        IF THIS.ProcessingOutput or This.lQuiet
            lcMsg=STRTRAN(ANALYZING_INDEXES_LOC,"...")
            THIS.InitTherm(lcMsg,aTableCount,0)
        ELSE
            WAIT ANALYZING_INDEXES_LOC WINDOW NOWAIT
        ENDIF

        *Create table to hold index names and expressions
        IF RTRIM(THIS.EnumIndexesTbl)==""
            lcEnum_Indexes=THIS.CreateWzTable("Indexes")
            THIS.EnumIndexesTbl=lcEnum_Indexes
            llCreateIndexes=.T.
        ELSE
            lcEnum_Indexes=THIS.EnumIndexesTbl
            llCreateIndexes=.F.
        ENDIF

        *read tables-to-be-exported one at a time and see if they have .CDXs
        SELECT (lcEnum_Tables)

        SCAN FOR EXPORT=.T. AND CDXAnald=.F. AND EMPTY(ClustName)=.T.
            lcCursorName=RTRIM(&lcEnum_Tables..CursName)
            lcTableName=RTRIM(&lcEnum_Tables..TblName)
            lcRemoteTable=RTRIM(&lcEnum_Tables..RmtTblName)
            lcTablePath=RTRIM(&lcEnum_Tables..TblPath)

            *Therm stuff
            lcMsg=STRTRAN(THIS_TABLE_LOC,'|1',lcTableName)
            THIS.UpDateTherm(lnTableCount,lcMsg)
            lnTableCount=lnTableCount+1

            *Read information for each tag

            SELECT (lcCursorName)
            FOR I=1 TO TAGCOUNT(STRTRAN(lcTablePath,".DBF",".CDX"))
                lcRemoteExpression=""
                lcClustered=.F.
                lcExpression=LOWER(SYS(14,I))		&&tag expression
                lcTagName=TAG(I,lcCursorName)		&&tag name

                *Figure out index type
                DO CASE
                    CASE PRIMARY(I)
                        lcLclIdxType="Primary key"
*** DH 12/15/2014: use only SQL Server rather than variants of it
***                        IF THIS.ServerType="Oracle" OR THIS.ServerType=="SQL Server95" THEN
						IF THIS.ServerType="Oracle" OR left(THIS.ServerType, 10) = 'SQL Server'
                            lcTagType="PRIMARY KEY"
                            IF THIS.ServerType="Oracle" THEN
                                lcClustered=.F.
                            ELSE
                                lcClustered=.T.
                            ENDIF
                        ELSE
                            lcTagType="UNIQUE CLUSTERED"
                            lcClustered=.T.
                        ENDIF
                    CASE CANDIDATE(I)
                        lcLclIdxType="Candidate"
                        *same for both Oracle and SQL Server
                        lcTagType="UNIQUE"
                    OTHERWISE
                        IF UNIQUE() THEN
                            lcLclIdxType="Unique"
                        ELSE
                            lcLclIdxType="Regular"
                        ENDIF
                        *if UNIQUE() or just a regular index
                        lcTagType=""
                ENDCASE

                *pull the field names out of each expression into comma separated list
                lcRemoteExpression=THIS.ExtractFieldNames(lcExpression,lcTableName)
                DIMENSION aKeyFields[1]
                aKeyFields=.F.
                THIS.KeyArray(lcRemoteExpression, @aKeyFields)
                IF ALEN(aKeyFields,1)>MAX_INDEX_FIELDS THEN
                    llDontCreate=.T.
                    lcErrMsg=TOO_MANY_FIELDS_LOC
                ELSE
                    llDontCreate=.F.
                    lcErrMsg=""
                ENDIF

                *Write info into index analysis table
                SELECT (lcEnum_Indexes)
                APPEND BLANK
                REPLACE IndexName WITH lcTableName, ;
                    TagName WITH LOWER(lcTagName), ;
                    LclExpr WITH lcExpression, 	;
                    LclIdxType WITH lcLclIdxType, ;
                    RmtExpr WITH lcRemoteExpression, ;
                    RmtName WITH LOWER(THIS.RemotizeName(lcTagName)), ;
                    RmtType WITH lcTagType, ;
                    Clustered WITH lcClustered, ;
                    RmtTable WITH lcRemoteTable, ;
                    Exported WITH .F., ;
                    DontCreate WITH llDontCreate,;
                    IdxError WITH lcErrMsg

                *Fox will let you have multiple "primary keys" on a table.  Oracle and SQL 95
                *declarative RI won't.  Consequently, the wizard needs to be
                *sure that the primary key on a table is the one used
                *for RI; the other "primary keys" can just be regular indexes.

                *Here the pkey expression gets stored for comparison later
                IF lcTagType="PRIMARY KEY" THEN
                    SELECT (lcEnum_Tables)
                    REPLACE PkeyExpr WITH lcRemoteExpression, ;
                        PKTagName WITH lcTagName
                ENDIF

                SELECT (lcCursorName)

            NEXT I

            SELECT (lcEnum_Tables)
            REPLACE CDXAnald WITH .T.

        ENDSCAN

        SELECT (lcEnum_Indexes)
        SELECT (lnOldArea)
        THIS.AnalyzeIndexesRecalc=.F.

        IF THIS.ProcessingOutput or This.lQuiet
			raiseevent(This, 'CompleteProcess')
        ELSE
            WAIT CLEAR
        ENDIF

    ENDPROC


    PROCEDURE BuildIndexSQL
        LOCAL lcEnum_Indexes, lcSQL, lcRmtTable, lcRmtIdxName, lcRmtType, lcConstraint, lcTSClause, lcClustered
*** DH 2015-09-25: added locals for new code
		local lcName, lcTable, lcExpr

        *Build the CREATE INDEX sql string
        lcEnum_Indexes = THIS.EnumIndexesTbl
        SELECT (lcEnum_Indexes)

        SCAN FOR Exported=.F.
            *Get remote table name
            lcRmtTable = LEFT(THIS.RemotizeName(RTRIM(&lcEnum_Indexes..IndexName)),30)
            lcRmtType = TRIM(&lcEnum_Indexes..RmtType)
            lcRmtExpr = TRIM(&lcEnum_Indexes..RmtExpr)
            lcRmtName = TRIM(&lcEnum_Indexes..RmtName)
            IF THIS.ServerType = ORACLE_SERVER
                lcRmtName = THIS.UniqueOraName(lcRmtName)
            ENDIF
*** DH 2015-09-25: handle brackets already in name
			if left(lcRmtName, 1) = '['
				lcName = lcRmtName
			else
				lcName = '[' + lcRmtName + ']'
			endif left(lcRmtName, 1) = '['
			if left(lcRmtTable, 1) = '['
				lcTable = lcRmtTable
			else
				lcTable = '[' + lcRmtTable + ']'
			endif left(lcRmtTable, 1) = '['
			if left(lcRmtExpr, 1) = '['
				lcExpr = lcRmtExpr
			else
				lcExpr = '[' + lcRmtExpr + ']'
			endif left(lcRmtExpr, 1) = '['
*** DH 2015-09-25: end of new code

            *check index type; deal with Oracle and Primary differently
*** DH 12/15/2014: use only SQL Server rather than variants of it
***            IF lcRmtType = "PRIMARY KEY" OR ;
                    (lcRmtType = "UNIQUE" AND THIS.ServerType = "Oracle") ;
                    OR (lcRmtType = "UNIQUE" AND THIS.ServerType == "SQL Server95") THEN
            IF lcRmtType = "PRIMARY KEY" OR ;
                    (lcRmtType = "UNIQUE" AND THIS.ServerType = "Oracle") ;
                    OR (lcRmtType = "UNIQUE" AND left(THIS.ServerType, 10) = 'SQL Server') THEN

                * Oracle and SQL95 Unique and Primary Key indexes implemented via ALTER TABLE
                IF THIS.ServerType = ORACLE_SERVER
                    lcTSClause = IIF(!EMPTY(THIS.TSIndexTSName), " USING INDEX TABLESPACE " + THIS.TSIndexTSName, "")
*** DH 2015-09-25: handle brackets already in name
***                    lcSQL = "ALTER TABLE " + lcRmtTable + " ADD (CONSTRAINT " + lcRmtName + ;
                        " " + lcRmtType + " (" + lcRmtExpr + ")" + lcTSClause + ")"
                    lcSQL = 'ALTER TABLE ' + lcTable + ' ADD (CONSTRAINT ' + lcName + ;
                        ' ' + lcRmtType + ' (' + lcExpr + ')' + lcTSClause + ')'
                ELSE
                    * jvf 08/13/99
                    * Use this convention for unique db objects...
                    * PRIMARY KEY: "PK_" + lcRmtName,3
                    * CANDIDATE KEYS: "UQ_" + lcRmtName

                    * ## Add NONCLUSTERED clause when creating a PRIMARY KEY b/c SS
                    * defaults to CLUSTERED INDEX.
                    
                    *{ JEI RKR 14.07.2005 Change: Add check for CLUSTERED clause in lcRmtType 
                    lcClustered = ""
                    IF !(" CLUSTERED" $ lcRmtType)
	                    lcClustered = " NONCLUSTERED "
                    ENDIF
                    *} JEI RKR 14.07.2005
                    IF lcRmtType = "UNIQUE"
                        * Can have multiple, so get unique name
*** DH 2015-09-25: strip brackets in name
***                        lcRmtName = "UQ_" + THIS.UniqueTableName(lcRmtTable)
                        lcRmtName = "UQ_" + THIS.UniqueTableName(chrtran(lcRmtTable, '[]', ''))
                    ELSE &&Primary Key
*** DH 2015-09-25: strip brackets in name
***                        lcRmtName = "PK_" + lcRmtTable
                        lcRmtName = "PK_" + chrtran(lcRmtTable, '[]', '')
                        IF !EMPTY(THIS.ExportClustered)
                            lcClustered = " CLUSTERED "
                        ENDIF
                    ENDIF
*** DH 2015-09-25: handle brackets already in name
***                    lcSQL = "ALTER TABLE [" + lcRmtTable + "] ADD CONSTRAINT " + lcRmtName + ;
                        " " + lcRmtType + lcClustered + "(" + lcRmtExpr + ")"
					lcSQL = 'ALTER TABLE ' + lcTable + ' ADD CONSTRAINT ' + lcRmtName + ;
						' ' + lcRmtType + lcClustered + '(' + lcExpr + ')'
                    *!*End Mark
                ENDIF
            ELSE
                * All other index types (including regular Oracle indexes) are handled here
*** DH 2015-09-25: handle brackets already in name
***                lcSQL = "CREATE " + lcRmtType + " INDEX [" + lcRmtName + "] ON [" + lcRmtTable + "] (" + lcRmtExpr + ")"
				lcSQL = 'CREATE ' + lcRmtType + ' INDEX ' + lcName + ' ON ' + lcTable + ' (' + lcExpr + ')'

                IF THIS.ServerType = ORACLE_SERVER
                    lcTSClause = IIF(!EMPTY(THIS.TSIndexTSName), " TABLESPACE " + THIS.TSIndexTSName, "")
                    lcSQL = lcSQL + lcTSClause
                ENDIF
            ENDIF

            REPLACE &lcEnum_Indexes..IndexSQL WITH lcSQL, &lcEnum_Indexes..RmtName WITH lcRmtName
        ENDSCAN
    ENDPROC



    FUNCTION CreateDevice
        PARAMETERS lcDeviceType, lcDevicePhysName,lnDeviceNumber

        LOCAL lcDeviceLogicalName, lcSQL1, lcSQL2, lcSQL3, I, lnRetVal, ;
            lcRoot, lcMasterPath, lnServerErr, lnUserChoice,lcErrMsg, ;
            lcDevicePhysPath, lcSQT
*** This used to be set in GetDeviceNumbers
        local aDeviceNumbers[1]

* If we have an extension object and it has a CreateDevice method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'CreateDevice', 5) and ;
			not This.oExtension.CreateDevice(lcDeviceType, lcDevicePhysName, ;
			lnDeviceNumber, This)
			return
		endif vartype(This.oExtension) = 'O' ...

        lcSQT=CHR(39)
        *Fill local variables from global
        IF lcDeviceType="Log" THEN
            *If the log and database names are the same, bail
            IF THIS.DeviceLogName=THIS.DeviceDBName THEN
                RETURN
            ENDIF
            *Put up thermometer
            THIS.InitTherm(CREATING_LOGDEVICE_LOC,0,0)
            lcDeviceLogicalName=RTRIM(THIS.DeviceLogName)
            lnDeviceSize=THIS.DeviceLogSize
        ELSE
            *Put up thermometer
            THIS.InitTherm(CREATING_DBDEVICE_LOC,0,0)
            lcDeviceLogicalName=RTRIM(THIS.DeviceDBName)
            lnDeviceSize=THIS.DeviceDBSize
        ENDIF
        THIS.UpDateTherm(0,TAKES_AWHILE_LOC)

        *convert from megabytes to 2k pages
        lnDeviceSize=lnDeviceSize*512

        * Build path for physical device based on location of Master database
        IF THIS.MasterPath=="" THEN
            lcMasterPath=""
            lcSQL = "select phyname from sysdevices where name = " + lcSQT + "master" + lcSQT
            lnRetVal = THIS.SingleValueSPT(lcSQL, @lcMasterPath, "phyname")
            THIS.MasterPath=RTRIM(lcMasterPath)
        ENDIF
        lcDevicePhysPath = THIS.JUSTPATH(THIS.MasterPath)+ "\"

        *Build device physical name
        lcRoot = LEFT(lcDeviceLogicalName,6)
        lcDevicePhysName = lcDevicePhysPath + lcRoot + ".DAT"

        *Get a device number and mark it as taken
        lnDeviceNumber=ASCAN(aDeviceNumbers,.F.)
        aDeviceNumbers[lnDeviceNumber]=.T.

        *Build sql string
        lcSQL1 = "disk init name=" + lcSQT+ lcDeviceLogicalName + lcSQT+ ", "
        lcSQL2 = "physname=" + lcSQT+ lcDevicePhysName + lcSQT+ ", "
        lcSQL3 = "vdevno=" + ALLTRIM(STR(lnDeviceNumber)) + ", "
        lcSQL3 = lcSQL3 + "size=" + ALLTRIM(STR(lnDeviceSize))
        lcSQL = lcSQL1 + lcSQL2 + lcSQL3

        *
        * Create device, incrementing physical name if already taken
        * Since error SQS_ERR_DISK_INIT_FAIL can result from other problems, only try 20 times
        *

        IF THIS.DoUpsize THEN
            *Set query time out to huge value while this is running
            =SQLSETPROP(THIS.MasterConnHand,"QueryTimeOut",600)

            lnServerErr = 0
            lcErrMsg = ""

            *- if problems creating devices ("DISK command not allowed within multi-statement transactions"),
            *- uncomment this line
            *- this.ExecuteTempSPT("COMMIT TRANSACTION",@lnServerErr,@lcErrMsg)

            *Try 20 times because physical name may be taken
            FOR I=1 TO 20
                lnServerErr=0
                IF !THIS.ExecuteTempSPT(lcSQL,@lnServerErr,@lcErrMsg) THEN
                    lcDevicePhysName = lcDevicePhysPath + lcRoot + ALLTRIM(STR(I)) + ".DAT"
                    lcSQL2 = "physname=" + lcSQT + lcDevicePhysName + lcSQT+ ", "
                    lcSQL = lcSQL1 + lcSQL2 + lcSQL3
                ELSE
                    EXIT
                ENDIF
            NEXT
            *If the device creation fails, quit
            IF lnServerErr<>0 THEN
	        	if This.lQuiet
	        		This.HadError     = .T.
	        		This.ErrorMessage = CANT_CREATE_DEVICE_LOC
				else
	                MESSAGEBOX(CANT_CREATE_DEVICE_LOC,ICON_EXCLAMATION,TITLE_TEXT_LOC)
	        	endif This.lQuiet
                THIS.Die
            ELSE
                *Set the query timeout back to a more reasonable figure
                =SQLSETPROP(THIS.MasterConnHand,"QueryTimeOut",30)
            ENDIF

        ENDIF

        *Store sql for script
        THIS.StoreSQL(lcSQL,DEVICE_SCRIPT_COMMENT_LOC)
		raiseevent(This, 'CompleteProcess')

    ENDFUNC


    FUNCTION ExecuteTempSPT
        PARAMETERS lcSQL, lnServerError, lcErrMsg, lcCursor
        LOCAL nRetVal, lnButtons, lcMsg, lcNewSQL, lcEscape

        lcEscape=IIF(THIS.ProcessingOutput,"ON",SET ("ESCAPE"))
        SET ESCAPE OFF

        IF PARAMETERS()=4 THEN
            nRetVal=SQLEXEC(THIS.MasterConnHand,lcSQL, lcCursor)
        ELSE
            nRetVal=SQLEXEC(THIS.MasterConnHand,lcSQL)
        ENDIF

        SET ESCAPE &lcEscape

        DO CASE
                *It worked
            CASE nRetVal=1
                lnServerError=0
                lcErrMsg=""
                RETURN .T.

                *Server error occurred
            CASE nRetVal=-1
                =AERROR(aErrArray)
                lnServerError=aErrArray[1]
                lcErrMsg=aErrArray[2]

                IF lnServerError=1526 AND !ISNULL(aErrArray[5])THEN
                    lnServerError=aErrArray[5]
                ENDIF

                DO CASE
                    CASE lnServerError=1105
                        *If Log is full, try to dump it (but only if upsizing to existing db)
			        	if This.lQuiet
			        		This.HadError     = .T.
			        		This.ErrorMessage = LOG_FULL_LOC
						else
                        	MESSAGEBOX(LOG_FULL_LOC,ICON_EXCLAMATION,TITLE_TEXT_LOC)
			        	endif This.lQuiet
                    CASE  lnServerError=1101 OR lnServerError=1510
                        *Device full
			        	if This.lQuiet
			        		This.HadError     = .T.
			        		This.ErrorMessage = DEVICE_FULL_LOC
						else
                        	MESSAGEBOX(DEVICE_FULL_LOC,ICON_EXCLAMATION,TITLE_TEXT_LOC)
			        	endif This.lQuiet

                    CASE lnServerError=1804
                        *SQL Server bug having to do with dropped device
			        	if This.lQuiet
			        		This.HadError     = .T.
			        		This.ErrorMessage = SQL_BUG_LOC
						else
                        	MESSAGEBOX(SQL_BUG_LOC,ICON_EXCLAMATION,TITLE_TEXT_LOC)
			        	endif This.lQuiet

                        *this error happens when dropping sproc that doesn't exist
                    CASE lnServerError=3701
                        RETURN .F.

                    CASE lnServerError=102
                        *syntax error in sql
                        RETURN .F.

                    CASE lnServerError=2615
                        *duplicate record entered
                        *should be caused only by running sp_foreignkey on a relation
                        *with the same foreign key twice
                        RETURN .F.

                    OTHERWISE
                        *unknown error
                        RETURN .F.
                ENDCASE

                *Connection level error occurred
            CASE nRetVal=-2
                *This is trouble; continue to generate script if user wants; otherwise bail
                lcMsg=STRTRAN(CONNECT_FAILURE_LOC,"|1",LTRIM(STR(lnServerErr)))
	        	if This.lQuiet
	        		This.HadError     = .T.
	        		This.ErrorMessage = lcMsg
				else
	                MESSAGEBOX(lcMsg,ICON_EXCLAMATION,TITLE_TEXT_LOC)
	        	endif This.lQuiet

        ENDCASE

        THIS.Die

    ENDFUNC

    PROCEDURE Die
        IF THIS.SQLServer
            THIS.TruncLogOff
        ENDIF
        THIS.NormalShutdown=.F.
        release oEngine
*** DH 03/20/2015: replaced CANCEL with block of code following as suggested
*** by Mike Potjer. Otherwise, if a problem occurs, the code just stops and the
*** caller can't determine what happened.
***		cancel
		if not empty(This.cReturnToProc)
			local lcReturnToProc
			lcReturnToProc = This.cReturnToProc
			return to &lcReturnToProc
		else
			cancel
		endif not empty(This.cReturnToProc)

    ENDPROC


    FUNCTION SingleValueSPT
        PARAMETERS lcSQL, lcReturnValue, lcFieldName, llReturnedOneValue
        LOCAL lcMsg, lcErrMsg, llRetVal, lcCursor, lnOldArea, lnServerError

        *
        *Executes a server query and sees if it return one value or not
        *If it returns one value, that value gets placed in a variable passed by reference
        *

        lnOldArea=SELECT()
        lcCursor=THIS.UniqueCursorName("_spt")
        SELECT 0
        IF THIS.ExecuteTempSPT(lcSQL,@lnServerError,@lcErrMsg,lcCursor) THEN
            IF RECCOUNT(lcCursor)=0 THEN
                llReturnedOneValue= .F.
            ELSE
                lcReturnValue=&lcCursor..&lcFieldName
                llReturnedOneValue=.T.
            ENDIF
            USE
        ELSE
            lcMsg=STRTRAN(QUERY_FAILURE_LOC,"|1",LTRIM(STR(lnServerError)))
        	if This.lQuiet
        		This.HadError     = .T.
        		This.ErrorMessage = lcMsg
			else
            	MESSAGEBOX(lcMsg,ICON_EXCLAMATION,TITLE_TEXT_LOC)
        	endif This.lQuiet
            THIS.Die
            RETURN
        ENDIF

        SELECT (lnOldArea)
        RETURN llReturnedOneValue

    ENDFUNC


    PROCEDURE AnalyzeTables
        LOCAL lcEnum_Tables, lcTableName, llUTableOpen, lcSourceDB, lnOldArea, ;
            llUpsizable, lcTablePath, llAlreadyOpened, llWarnUser, lcCursorName, ;
            aOpenTables, I, ctmpTblName

        lcSourceDB=THIS.SourceDB
        lnOldArea=SELECT()
        SET DATABASE TO (lcSourceDB)
*** DH 2015-08-10: set SourceDB to DBC() because CreateWzTable calls CreateNewDir,
*** which changes CURDIR, so if SourceDB contains a relative path, the DBC won't be found
*** when used in later code.
		This.SourceDB = dbc()
*** DH 2015-08-10: end of new code
        lcEnum_Tables=THIS.CreateWzTable("Tables")
        THIS.EnumTablesTbl=lcEnum_Tables

        IF ADBOBJECTS(aTblArray,"table")=0 THEN
            RETURN
        ENDIF

        DIMENSION aOpenTables[1,2]
        =AUSED(aOpenTables)
        FOR I=1 TO ALEN(aTblArray,1)
            APPEND BLANK
            lcTablePath= FULL(DBGETPROP(aTblArray(I),"TABLE","PATH"),DBC())
            llUpsizable=THIS.Upsizable(aTblArray(I), lcTablePath, @llAlreadyOpened, @lcCursorName, @aOpenTables)

            REPLACE TblName WITH LOWER(aTblArray(I)), ;
                CursName WITH lcCursorName,  ;
                TblPath WITH lcTablePath, ;
                Upsizable WITH llUpsizable, ;
                PreOpened WITH llAlreadyOpened, ;
                Type	WITH "T" && Add JEI RKR 2005.05.09

            * Check for DBCS
            ctmpTblName = ALLTRIM(TblName)
            IF LEN(m.ctmpTblName)#LENC(m.ctmpTblName)
                ctmpTblName=STRTRAN(ctmpTblName,CHR(32),"_")
                IF LEN(m.ctmpTblName)>29
                    IF ISLEADBYTE(SUBSTR(m.ctmpTblName,30,1))
                        ctmpTblName = LEFT(m.ctmpTblName,29)
                    ENDIF
                ENDIF
            ENDIF
            REPLACE RmtTblName WITH THIS.RemotizeName(m.ctmpTblName)

            IF !llUpsizable THEN
                llWarnUser=.T.
            ENDIF
            lcCursorName=""
        NEXT

        IF llWarnUser and not This.lQuiet
            =MESSAGEBOX(NO_OPEN_EXCLU_LOC,ICON_EXCLAMATION,TITLE_TEXT_LOC)
        ENDIF

        THIS.AnalyzeTablesRecalc=.F.
        SELECT (lnOldArea)

    ENDPROC


    FUNCTION AnalyzeClusters
        PARAMETERS aClusters
        LOCAL lcEnumClusters, lnOldArea, lcTableName, I, lnTableCount, aClusterCount

        * this creates the cluster table

        lnOldArea = SELECT()
        aClusterCount = ALEN(aClusters, 1)
        DIMENSION aClusters(ALEN(aClusters,1), 5)	&& won't compile without dimension
        IF EMPTY(aClusters[1,1]) OR aClusterCount = 0
            RETURN
        ENDIF

        *Create table for clusters if it doesn't exist yet
        IF RTRIM(THIS.EnumClustersTbl) == ""
            lcEnumClusters = THIS.CreateWzTable("Clusters")
            THIS.EnumClustersTbl = lcEnumClusters
        ELSE
            lcEnumClusters = RTRIM(THIS.EnumClustersTbl)
            SELECT &lcEnumClusters
            ZAP
        ENDIF

        * copy data from aClusters to Clusters table
        lnTableCount = 0
        SELECT (lcEnumClusters)
        FOR m.I = 1 TO aClusterCount
            lcClusterName = LOWER(aClusters[m.i, 1])
            lcMsg = STRTRAN(THIS_TABLE_LOC, "|1", lcClusterName)
            THIS.UpDateTherm(lnTableCount, lcMsg)

            APPEND BLANK
            REPLACE ClustName WITH  lcClusterName, ;
                ClustType WITH aClusters[m.i, 2], ;
                HashKeys WITH aClusters[m.i, 3], ;
                ClustSize WITH aClusters[m.i, 4], ;
                EXPORT WITH .T., ;
                Exported WITH .F.

        ENDFOR

        SELECT (lnOldArea)
    ENDPROC


    PROCEDURE AnalyzeFields
    	LPARAMETERS llAnalizeAllTables as Logical && JEI RKR 06.12.2005 Add for analize fields
    
        LOCAL lcEnum_Fields, lnOldArea, lcEnum_Tables, lcTableName, lcSourceDB, ;
            I, lnTableCount, lnListIndex, lnListIndexTemp, lcCursorName, lcErrMsg

* If we have an extension object and it has a AnalyzeFields method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'AnalyzeFields', 5) and ;
			not This.oExtension.AnalyzeFields(llAnalizeAllTables, This)
			return
		endif vartype(This.oExtension) = 'O' ...

        *
        * this creates the enumerates all the fields, their types etc.
        * in the tables selected to be upsized
        *
        * gets called when the tables selected for upsizing have changed

        lcEnum_Tables=THIS.EnumTablesTbl
        lnOldArea=SELECT()

        SELECT COUNT(*) FROM (lcEnum_Tables) WHERE FldsAnald=.F. AND EXPORT=.T. ;
            INTO ARRAY aTableCount
        IF aTableCount=0 THEN
            RETURN
        ENDIF

        *Tell the user what's going on
        IF THIS.ProcessingOutput or This.lQuiet
            lcMsg=STRTRAN(ANALYZING_FIELDS_LOC,"...")
            THIS.InitTherm(lcMsg,aTableCount,0)
        ELSE
            WAIT ANALYZING_FIELDS_LOC WINDOW NOWAIT
        ENDIF
        lnTableCount=0

        *Create table for fields if it doesn't exist yet
        IF RTRIM(THIS.EnumFieldsTbl)=="" THEN
            lcEnum_Fields=THIS.CreateWzTable("Fields")
            THIS.EnumFieldsTbl=lcEnum_Fields
            llCreateIndexes=.T.
        ELSE
            lcEnum_Fields=RTRIM(THIS.EnumFieldsTbl)
            llCreateIndexes=.F.
        ENDIF

        *only look at tables that haven't been crunched through this procedure before
        SELECT (lcEnum_Tables)
        SCAN FOR FldsAnald=.F. AND (EXPORT=.T. OR llAnalizeAllTables)
            lcCursorName=RTRIM(&lcEnum_Tables..CursName)
            lcTableName=RTRIM(&lcEnum_Tables..TblName)
            lcMsg=STRTRAN(THIS_TABLE_LOC,"|1",lcTableName)
            THIS.UpDateTherm(lnTableCount,lcMsg)
            SELECT (lcCursorName)
            =AFIELDS(aFldArray)

            * test for number of fields and record size(for SQL Server)
            IF THIS.SQLServer
                lcErrMsg = ""
                
                * rmk - 01/06/2004 - adjust limits for SQL Server 7.0 and above
		        IF THIS.ServerVer>=7
	                IF ALEN(aFldArray,1)> 1024
	                    lcErrMsg = STRTRAN(CANT_UPSIZE_LOC ,"|1", lcTableName)+ EXCEED_FIELDS_LOC
	                ENDIF
	                IF RECSIZE() > 8060
	                    IF !EMPTY(lcErrMsg)
	                        lcErrMsg = lcErrMsg + " and " + EXCEED_RECSIZE_LOC
	                    ELSE
	                        lcErrMsg = STRTRAN(CANT_UPSIZE_LOC ,"|1", lcTableName)+ EXCEED_RECSIZE_LOC
	                    ENDIF
	                ENDIF
		        ELSE
	                IF ALEN(aFldArray,1)> 250
	                    lcErrMsg = STRTRAN(CANT_UPSIZE_LOC ,"|1", lcTableName)+ EXCEED_FIELDS_LOC
	                ENDIF
	                IF RECSIZE() > 1962
	                    IF !EMPTY(lcErrMsg)
	                        lcErrMsg = lcErrMsg + " and " + EXCEED_RECSIZE_LOC
	                    ELSE
	                        lcErrMsg = STRTRAN(CANT_UPSIZE_LOC ,"|1", lcTableName)+ EXCEED_RECSIZE_LOC
	                    ENDIF
	                ENDIF
	            ENDIF
                IF !EMPTY(lcErrMsg) and not This.lQuiet
                    =MESSAGEBOX(lcErrMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC)
                ENDIF
            ENDIF

            * 11/02/02 JVF Added: AutoInNext and AutoInStep o account
            * for VFP 8.0 autoinc attrib. Using Afields(i,18) Step Value > 0 to determine
            * if auto inc column.
            SELECT (lcEnum_Fields)
            FOR I=1 TO ALEN(aFldArray,1)
                APPEND BLANK

                lcFldName=LOWER(aFldArray(I,1))
                REPLACE TblName WITH lcTableName, ;
                    FldName WITH lcFldName,;
                    DATATYPE WITH aFldArray(I,2), ;
                    LENGTH WITH aFldArray(I,3) ;
                    PRECISION WITH aFldArray(I,4), ;
                    RmtFldname WITH THIS.RemotizeName(lcFldName), ;
                    RmtLength WITH aFldArray(I,3) ;
                    RmtPrec WITH aFldArray(I,4),;
                    lclnull WITH aFldArray(I,5), ;
                    RmtNull WITH aFldArray(I,5) AND !((aFldArray(I,17) <> 0) OR (aFldArray(I,18) <> 0)), ; && JEI RKR 2005.03.28 From aFldArray(I,5)
                    NOCPTRANS WITH aFldArray(I,6), ;
                    AutoInNext WITH aFldArray(I,17), ;
                    AutoInStep WITH aFldArray(I,18)
            NEXT I

            SELECT (lcEnum_Tables)
            * set default mapping
            THIS.DefaultMapping(lcTableName)
            * set TimeStamp default (M, G, P)
            REPLACE TStampAdd WITH (THIS.SQLServer AND THIS.AddTimeStamp(TblName))
            REPLACE FldsAnald WITH .T.

            IF THIS.ProcessingOutput THEN
                lnTableCount=lnTableCount+1
                THIS.UpDateTherm(lnTableCount)
            ENDIF
        ENDSCAN

        SELECT (lcEnum_Fields)

        *Only do this the first time through
        *IF llCreateIndexes THEN
        *	INDEX ON TblName TAG TblName
        *	INDEX ON FldName TAG FldName
        *	SET ORDER TO
        *ENDIF

        *Deal with thermometer or wait window
        IF THIS.ProcessingOutput or This.lQuiet
			raiseevent(This, 'CompleteProcess')
        ELSE
            WAIT CLEAR
        ENDIF

        THIS.AnalyzeFieldsRecalc=.F.
        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE DefaultMapping
        PARAMETERS lcTableName
        LOCAL lnOldArea, aDefaultMapping, lcEnum_Fields, lnLength, lnPrecision, ;
            lcTypeString, llSkippedFirst, lcNocpType, lcRemoteType

        lnOldArea=SELECT()
        lcEnum_Fields=RTRIM(THIS.EnumFieldsTbl)
        *grab array of data types for setting default datatypes
        DIMENSION aDefaultMapping(11,4)
        THIS.GetDefaultMapping(@aDefaultMapping)
        *Columns in aDefaultMapping
        *Column 1: LocalType (local FoxPro data type)
        *Column 2: RemoteType (default server data type)
        *Column 3: VarLength (whether the data type is variable length or not)
        *Column 4: FullLocal	(full name of FP data type, e.g. "C"->"character")

        *Go through the fields and stick in default data type

        SELECT (lcEnum_Fields)
        FOR I=1 TO ALEN(aDefaultMapping,1)

            *If the remote datatype doesn't take a length argument, put 0 in there.
            *Otherwise, put a 1 in and then transfer the length and precision values of the local type.
            *The 0 will cause the field to be ignored by the CreateTableSQL routine

            IF aDefaultMapping(I,3)=.T. THEN
                lnLength=1
                lnPrecision=1
            ELSE
                lnLength=0
                lnPrecision=0
            ENDIF

            REPLACE RmtType WITH aDefaultMapping(I,2), ;
                FullType WITH aDefaultMapping(I,4) ;
                RmtLength WITH lnLength, RmtPrec WITH lnPrecision ;
                FOR DATATYPE = aDefaultMapping(I,1) AND RTRIM(TblName) == lcTableName
        NEXT

        * Set up ComboType field of the type mapping grid
        * create a string like "character (14)" or "numeric (3,2)" or "memo" or "memo-nocp"
        SCAN FOR RTRIM(TblName) == lcTableName
            * change default FullType for NoCptrans

            * rmk - 01/06/2004
            * lcTypeNocp = IIF(DATATYPE = 'C', 'char_nocp', 'memo_nocp')
            DO CASE
            CASE DATATYPE = 'C'
	            lcTypeNocp = 'char_nocp'
            CASE DATATYPE = 'V'
	            lcTypeNocp = 'varchar_nocp'
            OTHERWISE
            	lcTypeNocp = 'memo_nocp'
			ENDCASE

            lcTypeString = IIF(NOCPTRANS, lcTypeNocp, RTRIM(FullType))

            * add field length and decimals for Fox variable length types
            IF INLIST(DATATYPE, 'C', 'V', 'Q', 'N', 'F')
                lcTypeString = lcTypeString + " (" + LTRIM(STR(LENGTH))
                IF PRECISION <> 0 THEN
                    lcTypeString = lcTypeString + "," + LTRIM(STR(PRECISION))
                ENDIF
                lcTypeString = lcTypeString+")"
            ELSE
                * JVF 11/02/02 Add AutoInc string to Local type
                IF DATATYPE = "I" AND THIS.SQLServer
                    IF AutoInStep > 0
                        lcTypeString = lcTypeString + " (AutoInc)"

                        * JVF 11/02/02 # 1478 Add "(Identity)" string to RemoteType if AutoInc,
                        REPLACE RmtType WITH "int (Ident)"
                    ELSE
                        * Since we added record it to typemap, strip it off if necc.
                        REPLACE RmtType WITH "int"
                    ENDIF
                ENDIF
            ENDIF
            REPLACE ComboType WITH lcTypeString
        ENDSCAN

        * Set up the RemLength, RemPrec fields of the type mapping grid
        * Replace all the 1s in the Rmtlength field with local length and precision values
        * jvf 08/16/99
        SCAN FOR RmtLength <> 0 AND RTRIM(TblName) == lcTableName
            IF PRECISION <> 0 AND DATATYPE = 'N'
                REPLACE RmtLength WITH LENGTH-1, RmtPrec WITH PRECISION
            ELSE
                REPLACE RmtLength WITH LENGTH, RmtPrec WITH PRECISION
            ENDIF
        ENDSCAN

        * implement the server-specific cases for RmtType, Rmtlength, RmtPrec
        IF THIS.SQLServer THEN
            REPLACE ALL RmtType WITH "char" ;
                FOR DATATYPE = "C" AND NOCPTRANS AND RTRIM(TblName)==lcTableName
            * jvf 1/8/01 RmtType s/b Text, not Image, for Binary Memos - Bug ID 45752
            REPLACE ALL RmtType WITH "text" ;
                FOR DATATYPE = "M" AND NOCPTRANS AND RTRIM(TblName)==lcTableName
        ENDIF

        * If we're upsizing to Oracle, need to put default values in for money and logical types
        * Same for int when converted to numeric
        IF THIS.ServerType == ORACLE_SERVER THEN

            REPLACE ALL RmtType WITH "raw" ;
                FOR DATATYPE = "C" AND NOCPTRANS AND RTRIM(TblName)==lcTableName

            REPLACE ALL RmtType WITH "long raw" ;
                FOR DATATYPE = "M" AND NOCPTRANS AND RTRIM(TblName)==lcTableName

            REPLACE ALL RmtLength WITH 19, RmtPrec WITH 4 ;
                FOR DATATYPE="Y" AND RTRIM(TblName)==lcTableName

            REPLACE ALL RmtLength WITH 1, RmtPrec WITH 0 ;
                FOR DATATYPE="L" AND RTRIM(TblName)==lcTableName

            REPLACE ALL RmtLength WITH 11, RmtPrec WITH 0 ;
                FOR DATATYPE="I" AND RTRIM(TblName)==lcTableName

            *Oracle only allows one LONG or LONG RAW column per table; change all
            *but the first LONG RAW, ie General, to RAW(255); change extra LONG fields
            *to VARCHAR2(2000)

            *This handles general fields
            LOCATE FOR RTRIM(RmtType)=="long raw" AND RTRIM(TblName)==lcTableName
            DO WHILE FOUND()
                IF llSkippedFirst THEN
                    REPLACE RmtType WITH "raw", RmtLength WITH 255
                ELSE
                    llSkippedFirst=.T.
                ENDIF
                CONTINUE
            ENDDO

            *This handles memo fields
            LOCATE FOR RTRIM(RmtType)=="long" AND RTRIM(TblName)==lcTableName
            DO WHILE FOUND()
                IF llSkippedFirst THEN
                    REPLACE RmtType WITH "varchar2", RmtLength WITH 2000
                ELSE
                    llSkippedFirst=.T.
                ENDIF
                CONTINUE
            ENDDO
        ENDIF

        SELECT (lnOldArea)

    ENDPROC


    FUNCTION GetEligibleTables
        PARAMETERS aTableArray
        LOCAL lcEnumTables, lnOldArea, I

        lnOldArea=SELECT()
        lcEnumTables=THIS.EnumTablesTbl
        SELECT (lcEnumTables)
        SET FILTER TO Upsizable=.T.
        GO TOP
        I=1
        IF EOF()
            RETURN 0
        ELSE
            SCAN
                IF !EMPTY(aTableArray[1]) THEN
                    DIMENSION aTableArray[ALEN(aTableArray,1)+1]
                ENDIF
                aTableArray[i]=TblName
                I=I+1
            ENDSCAN
        ENDIF
        SELECT (lnOldArea)
        RETURN

    ENDFUNC


    FUNCTION GetEligibleClusterTables
        PARAMETERS aTableArray, lcSelectClustName
        LOCAL lcEnumTables, lnOldArea, I, lcFilter

        * if EOF() the array is unchanged
        DIMENSION aTableArray[1]
        aTableArray[1] = ""
        lnOldArea=SELECT()
        lcEnumTables=THIS.EnumTablesTbl
        SELECT (lcEnumTables)

        lcSelectClustName = TRIM(lcSelectClustName)
        IF lcSelectClustName = ""
            lcFilter = "Export = .T. AND EMPTY(ClustName)"
        ELSE
            lcFilter = "Export = .T. AND TRIM(ClustName) = lcSelectClustName"
        ENDIF

        GO TOP
        I=1
        IF EOF()
            RETURN 0
        ELSE
            SCAN FOR &lcFilter
                IF !EMPTY(aTableArray[1]) THEN
                    DIMENSION aTableArray[ALEN(aTableArray,1)+1]
                ENDIF
                aTableArray[i] = TRIM(TblName)
                I=I+1
            ENDSCAN
        ENDIF
        SELECT (lnOldArea)
        RETURN

    ENDFUNC


    FUNCTION GetEligibleTableFields
        PARAMETERS aFieldArray, lcTableName, llKeys
        LOCAL lcEnumFields, lnOldArea, I, lcFilter

        DIMENSION aFieldArray[1]
        aFieldArray[1] = ""
        lnOldArea = SELECT()
        lcEnumFields = THIS.EnumFieldsTbl
        SELECT (lcEnumFields)

        lcTableName = TRIM(lcTableName)
        IF llKeys
            lcFilter = "TblName = lcTableName  AND !EMPTY(ClustOrder)"
        ELSE
            lcFilter = "TblName = lcTableName  AND EMPTY(ClustOrder)"
        ENDIF

        GO TOP
        I = 1
        IF EOF()
            RETURN 0
        ELSE
            SCAN FOR &lcFilter
                IF !EMPTY(aFieldArray[1]) THEN
                    DIMENSION aFieldArray[ALEN(aFieldArray,1)+1]
                ENDIF
                aFieldArray[i] = TRIM(FldName)
                I = I + 1
            ENDSCAN
        ENDIF
        SELECT (lnOldArea)
        RETURN

    ENDFUNC

    FUNCTION GetInfoTableFields
        PARAMETERS aInfoFieldArray, lcTableName, llKeys
        LOCAL lcEnumFields, lnOldArea, I, lcFilter

        DIMENSION aInfoFieldArray[1,5]
        aInfoFieldArray[1,1] = ""
        lnOldArea = SELECT()
        lcEnumFields = THIS.EnumFieldsTbl
        SELECT (lcEnumFields)
        lcTableName = TRIM(lcTableName)
        IF llKeys
            lcFilter = "TblName = lcTableName  AND !EMPTY(ClustOrder)"
        ELSE
            lcFilter = "TblName = lcTableName  AND EMPTY(ClustOrder)"
        ENDIF

        GO TOP
        I = 1
        IF EOF()
            RETURN 0
        ELSE
            SCAN FOR &lcFilter
                IF !EMPTY(aInfoFieldArray[1,1])
                    DIMENSION aInfoFieldArray[ALEN(aInfoFieldArray,1)+1,5]
                ENDIF
                aInfoFieldArray[i,1] = TRIM(FldName)
                aInfoFieldArray[i,2] = TRIM(RmtType)
                aInfoFieldArray[i,3] = RmtLength
                aInfoFieldArray[i,4] = RmtPrec
                IF llKeys
                    aInfoFieldArray[i,5] = ClustOrder
                ELSE
                    aInfoFieldArray[i,5] = .T.
                ENDIF
                I = I + 1
            ENDSCAN
        ENDIF
        =ASORT(aInfoFieldArray, 5)
        SELECT (lnOldArea)
        RETURN

    ENDFUNC


    * verify if the tables selected in cluster have common fields

    PROCEDURE VerifyClusterTablesFields
        PARAMETERS aClusterTables
        LOCAL llValid, aInfoTable1Fields[1,5], aInfoTableFields[1,5]

        * False if no table selected
        IF EMPTY(aClusterTables[1,1])
            RETURN .F.
        ENDIF

        THIS.GetInfoTableFields(@aInfoTable1Fields, aClusterTables[1], .F.)

        * True if we have a single cluster with at least one key
        IF ALEN(aClusterTables,1) = 1
            RETURN IIF(EMPTY(aClusterTables[1,1]), .F., .T.)
        ENDIF

        * we have at least two tables here and first table has at least a field selected
        FOR m.I = 2 TO ALEN(aClusterTables,1)
            THIS.GetInfoTableFields(@aInfoTableFields, aClusterTables[m.i],.F.)
            FOR m.j = 1 TO ALEN(aInfoTable1Fields,1)
                IF aInfoTable1Fields[m.j,5]
                    FOR m.k = 1 TO ALEN(aInfoTableFields,1)
                        llValid = .F.
                        IF (aInfoTable1Fields[m.j,2] = aInfoTableFields[m.k,2] AND ;
                                aInfoTable1Fields[m.j,3] = aInfoTableFields[m.k,3] AND ;
                                aInfoTable1Fields[m.j,4] = aInfoTableFields[m.k,4])
                            llValid = .T.
                            EXIT
                        ENDIF
                    ENDFOR
                    aInfoTable1Fields[m.j,5] = llValid
                ENDIF
            ENDFOR
        ENDFOR

        RETURN ASCAN(aInfoTable1Fields,.T.) > 0
    ENDPROC



    PROCEDURE VerifyClusterKeyFields
        PARAMETERS aClusterTables
        LOCAL llValid, aInfoTable1Fields[1,5], aInfoTableFields[1,5]

        * verify if the fields selected in each table(part of the cluster) are the same

        * False if no clusters
        IF EMPTY(aClusterTables[1,1])
            RETURN .F.
        ENDIF

        THIS.GetInfoTableFields(@aInfoTable1Fields, aClusterTables[1,1], .T.)

        * True if we have a single table with a valid key
        IF ALEN(aClusterTables,1) = 1
            RETURN IIF(EMPTY(aInfoTable1Fields[1,1]), .F., .T.)
        ENDIF

        * we have at least two tables here and first table has a valid key
        FOR m.I = 2 TO ALEN(aClusterTables,1)
            THIS.GetInfoTableFields(@aInfoTableFields, aClusterTables[m.i], .T.)

            * keys for both clusters should match in number of fields, type and size
            IF EMPTY(aInfoTableFields[1,1]) OR ;
                    ALEN(aInfoTableFields,1) != ALEN(aInfoTable1Fields,1)
                RETURN .F.
            ENDIF

            FOR m.j = 1 TO ALEN(aInfoTable1Fields,1)
                IF (aInfoTable1Fields[m.j,2] = aInfoTableFields[m.j,2] AND ;
                        aInfoTable1Fields[m.j,3] = aInfoTableFields[m.j,3] AND ;
                        aInfoTable1Fields[m.j,4] = aInfoTableFields[m.j,4])
                    LOOP
                ELSE
                    RETURN .F.
                ENDIF
            ENDFOR
        ENDFOR

        RETURN .T.
    ENDPROC


    PROCEDURE GetDefaultMapping
        PARAMETERS aPassedArray
        LOCAL lnOldArea, lcServerConstraint

*** DH 2015-09-14: handle TypeMap not existing
		if not file('TypeMap.dbf')
			This.HadError = .T.
			This.ErrorMessage = 'TypeMap.dbf cannot be found'
			This.Die()
		endif not file('TypeMap.dbf')
*** DH 2015-09-14: end of new code
        lnOldArea=SELECT()
        IF NOT USED("TypeMap")
            SELECT 0
*** DH 09/05/2013: don't open exclusively
***            USE TypeMap EXCLUSIVE
            USE TypeMap
        ELSE
            SELECT TypeMap
        ENDIF

        *Didn't foresee a problem, thus this cheezy snippet
*** DH 09/05/2013: use only SQL Server rather than variants of it
***        IF THIS.ServerType=="SQL Server95" THEN
		IF left(THIS.ServerType, 10) = 'SQL Server'
            lcServerConstraint="SQL Server"
        ELSE
***            IF THIS.ServerType=="SQL Server" THEN
***                lcServerConstraint="SQL Server4x"
***            ELSE
                lcServerConstraint=RTRIM(THIS.ServerType)
***            ENDIF
        ENDIF

        SELECT LocalType, RemoteType, VarLength, FullLocal FROM TypeMap ;
            WHERE  TypeMap.DEFAULT=.T. AND TypeMap.SERVER=lcServerConstraint ;
            INTO ARRAY aPassedArray

        SELECT(lnOldArea)

    ENDFUNC


    #IF SUPPORT_ORACLE
    PROCEDURE DealWithTypeLong
        LOCAL lcEnumFieldsTbl, lnOldArea

        *
        *Oracle tables only allow one field to be Long or LongRaw; this warns
        *the user about the problem.  The DefaultMapping routine deals with it
        *

        THIS.AnalyzeFields

        lcEnumFieldsTbl=THIS.EnumFieldsTbl
        lnOldArea=SELECT()
        SELECT 0
        lcCursor=THIS.UniqueCursorName("_foo")
        SELECT COUNT(*) FROM (lcEnumFieldsTbl) ;
            WHERE RTRIM(DATATYPE)=="M" OR ;
            RTRIM(DATATYPE)=="G" OR ;
            RTRIM(DATATYPE)=="P" ;
            GROUP BY TblName ;
            INTO CURSOR &lcCursor
        SELECT COUNT(*) FROM (lcCursor) WHERE CNT>1 INTO ARRAY aMemoCount
        USE

        IF aMemoCount>0 THEN
            lcMsg=STRTRAN(LONG_TYPE_LOC,'|1',LTRIM(STR(aMemoCount)))
            lcMsg=STRTRAN(lcMsg,'|2',IIF(aMemoCount>1,TABLES_HAVE_LOC,TABLE_HAS_LOC))
        	if not This.lQuiet
            	MESSAGEBOX(lcMsg,ICON_EXCLAMATION,TITLE_TEXT_LOC)
        	endif not This.lQuiet
        ENDIF

        SELECT (lnOldArea)

    ENDPROC
#ENDIF


#IF SUPPORT_ORACLE
    FUNCTION TwoLongs
        PARAMETERS lcTableName, lcFirstLong, lcOtherLong
        LOCAL lcEnumFieldsTbl, lnOldArea

        *Checks to see if a field in a table already has a type of Long or Long Raw
        *Returns the name of the field if there is one

        lcEnumFieldsTbl=THIS.EnumFieldsTbl
        lnOldArea=SELECT()
        SELECT (lcEnumFieldsTbl)
        LOCATE FOR RTRIM(TblName)==lcTableName AND RmtType="long"
        IF RTRIM(FldName)==lcFirstLong THEN
            CONTINUE
            IF FOUND() THEN
                lcOtherLong=FldName
            ENDIF
        ELSE
            lcOtherLong=FldName
        ENDIF

        SELECT (lnOldArea)
        RETURN !EMPTY(lcOtherLong)

    ENDFUNC
#ENDIF


    FUNCTION Upsizable
        PARAMETERS lcTableName, lcTablePath, llAlreadyOpen, lcCursorName, aOpenTables
        LOCAL lnOldArea, I, lcNewCursName

        *
        *This function checks to see that a table is actually marked as part of the
        *selected database
        *
        *It also opens all tables exclusively if they aren't already
        *

        *Substitute underscores for any spaces (as FoxPro does)

        *See if the table is already open, possibly with an alias different from the table name
        IF !EMPTY(aOpenTables) THEN
            FOR I=1 TO ALEN(aOpenTables,1)
                IF DBF(aOpenTables[i,2])==RTRIM(UPPER(lcTablePath)) THEN
                    lcCursorName=aOpenTables[i,1]
                    EXIT
                ENDIF
            NEXT
        ENDIF

        *If it's not open already, handle table names with spaces
        IF EMPTY(lcCursorName) THEN
            lcCursorName=RTRIM(STRTRAN(lcTableName,CHR(32),"_"))
            *Handle the case of table name being an important Fox keyword
            *Note the base wizard class ensures that no tables are already open
            *with these keywords, so we only worry about opening them here
            IF INLIST(UPPER(lcCursorName),"THIS","THISFORMSET")
                lcCursorName=LEFT(lcCursorName,MAX_FIELDNAME_LEN-1)+"_"
            ENDIF
        ENDIF

        lnOldArea=SELECT()
        IF !FILE(lcTablePath) THEN
            SELECT (lnOldArea)
            RETURN .F.
        ENDIF
        IF !USED(lcCursorName) THEN
            THIS.SetErrorOff=.T.
            THIS.HadError=.F.
            llAlreadyOpened=.F.
            SELECT 0
            USE (lcTableName) ALIAS (lcCursorName) EXCLUSIVE
            THIS.SetErrorOff=.F.
            IF THIS.HadError
                SELECT (lnOldArea)
                RETURN .F.
            ENDIF
        ELSE
            *Make sure that if a table's open, it belongs to the database
            *to be upsized
            SELECT (lcCursorName)
            IF !LOWER(CURSORGETPROP('database'))==LOWER(ALLTRIM(THIS.SourceDB)) THEN
                lcCursorName=THIS.UniqueTorVName("Namewithmanycharacters")
                THIS.SetErrorOff=.T.
                THIS.HadError=.F.
                SELECT 0
                USE (lcTableName) EXCLUSIVE ALIAS (lcCursorName)
                THIS.SetErrorOff=.F.
                IF THIS.HadError
                    USE
                    SELECT (lnOldArea)
                    RETURN .F.
                ENDIF
            ELSE
                llAlreadyOpened=.T.
            ENDIF

            *If it's open, make sure it's open exclusive
            SELECT (lcCursorName)
            IF !ISFLOCKED()
                USE
                THIS.SetErrorOff=.T.
                THIS.HadError=.F.
                USE (lcTableName) EXCLUSIVE
                THIS.SetErrorOff=.F.
                IF THIS.HadError
                    USE (lcTableName) SHARED
                    SELECT (lnOldArea)
                    RETURN .F.
                ENDIF
            ENDIF

        ENDIF

        SELECT (lnOldArea)

    ENDFUNC


    #IF SUPPORT_ORACLE
    PROCEDURE RemoveCluster
        PARAMETERS lcClusterName
        LOCAL lcClusterNamesTbl, lcClusterKeysTbl, lcEnumTablesTbl, lnOldArea

        *Called by "Remove" button on Create Cluster screen and by Table Selection screen

        lcClusterNamesTbl=THIS.ClusterNamesTbl
        lcClusterKeysTbl=THIS.ClusterKeysTbl
        lcEnumTablesTbl=THIS.EnumTablesTbl
        lnOldArea=SELECT()

        *delete the record for the cluster from the cluster names table
        SELECT (lcClusterNamesTbl)
        DELETE ALL FOR &lcClusterNamesTbl..ClustName=lcClusterName

        *delete any related records in the cluster keys table
        SELECT (lcClusterKeysTbl)
        DELETE ALL FOR &lcClusterKeysTbl..ClustName=lcClusterName

        *Mark any tables that were in the cluster as available
        SELECT (lcEnumTablesTbl)
        REPLACE ALL &lcEnumTablesTbl..lcClusterName WITH ""

        SELECT (lnOldArea)

    ENDPROC
#ENDIF


#IF SUPPORT_ORACLE
    FUNCTION ChangeClusterStatus
        PARAMETERS lcRel,lcClustName, lcClustType, lnHashKeys
        LOCAL lcEnumTables, lnOldArea, lcEnumIndexesTbl, aClustTables, lcParent, ;
            lcChild, lnDupeID

        *Called from cluster creation page when user adds or removes a cluster

        lnOldArea=SELECT()
        lcEnumTables=THIS.EnumTablesTbl
        lcEnumIndexesTbl=THIS.EnumIndexesTbl
        lcEnumRelsTbl=THIS.EnumRelsTbl
        IF EMPTY(lnHashKeys) THEN
            lnHashKeys=0
        ENDIF

        *Parse the relation
        lcParent=""
        lcChild=""
        lnDupeID=0
        THIS.ParseRel(lcRel,@lcParent,@lcChild,@lnDupeID)
        SELECT (lcEnumRelsTbl)

        *If the clustertype wasn't passed, the cluster is being added or deleted
        *so see if the name is already in use
        IF TYPE("lcClustType")="L" THEN
            IF !lcClustName=="" THEN
                SET ORDER TO ClustName
                SEEK RTRIM(lcClustName)
                IF FOUND() THEN
                    *Give error message if cluster name already exists
                    lcMessage=STRTRAN(DUP_CLUSTNAME_LOC,"|1",RTRIM(THIS.UserInput))
		        	if This.lQuiet
		        		This.HadError     = .T.
		        		This.ErrorMessage = lcMessage
					else
	                    MESSAGEBOX(lcMessage,ICON_EXCLAMATION,TITLE_TEXT_LOC)
		        	endif This.lQuiet
                    SET ORDER TO
                    RETURN .F.
                ENDIF
                SET ORDER TO
            ENDIF
        ENDIF

        *If it's not in use or we're just changing the cluster type, find the right record
        *and toss the values in

        LOCATE FOR DD_CHILD=lcChild AND DD_PARENT=lcParent AND Duplicates=lnDupeID
        IF TYPE("lcClustType")="L" THEN
            lcClustType=IIF(lcClustName=="","","INDEX")
            REPLACE ClustName WITH lcClustName, ClustType WITH lcClustType, ;
                HashKeys WITH lnHashKeys

            *Now associate the cluster name ("" if table is being removed)
            *with the tables in the cluster and store default cluster type of " INDEX"
            SELECT (lcEnumTables)
            REPLACE &lcEnumTables..ClustName WITH lcClustName FOR TblName=lcParent OR TblName=lcChild
        ELSE
            REPLACE ClustType WITH lcClustType, HashKeys WITH lnHashKeys
        ENDIF

        SELECT (lnOldArea)


    ENDFUNC
#ENDIF


#IF SUPPORT_ORACLE
    PROCEDURE RemoveTableFromClust
        PARAMETERS lcClustName

        *Called by page 6 when a user deselects a table that was going to be exported

        LOCAL lcEnumRelsTbl, lcEnumTablesTbl, lnOldArea, lnRecNo
        lcEnumTablesTbl=THIS.EnumTablesTbl
        lcEnumRelsTbl=THIS.EnumRelsTbl
        lnOldArea=SELECT()

        SELECT (lcEnumTablesTbl)
        lnRecNo=RECNO()
        REPLACE ClustName WITH "" FOR RTRIM(ClustName)==RTRIM(lcClustName)
        GO lnRecNo

        SELECT (lcEnumRelsTbl)
        REPLACE ClustName WITH "", ClustType WITH "" FOR RTRIM(ClustName)==RTRIM(lcClustName)

        SELECT (lnOldArea)

    ENDPROC
#ENDIF


#IF SUPPORT_ORACLE
    FUNCTION InOneCluster
        PARAMETERS lcRel
        LOCAL lcParent, lcChild, lnDupeID, lcEnumTablesTbl, lcMsg

        *
        *Checked when the user creates a cluster; ensures that a given table is only
        *in one relation
        *

        *Parse the relation
        lcParent=""
        lcChild=""
        lnDupeID=0
        THIS.ParseRel(lcRel,@lcParent,@lcChild,@lnDupeID)

        *See if the tables are already in a cluster
        lcEnumTablesTbl=THIS.EnumTablesTbl
        SELECT (lcEnumTablesTbl)

        FOR I=1 TO 2
            LOCATE FOR RTRIM(TblName)==lcParent
            IF !EMPTY(ClustName) THEN
                lcMsg=STRTRAN(ONE_CLUSTER_LOC,"|1",lcParent)
                lcMsg=STRTRAN(lcMsg,"|2",RTRIM(ClustName))
	        	if This.lQuiet
	        		This.HadError     = .T.
	        		This.ErrorMessage = lcMsg
				else
	                MESSAGEBOX(lcMsg,48,TITLE_TEXT_LOC)
	        	endif This.lQuiet
                RETURN .F.
            ENDIF
            lcParent=lcChild
        NEXT

    ENDFUNC
#ENDIF


    FUNCTION InIndex
        PARAMETERS aIndexes, lcFldName, lcTableName
        LOCAL lcEnumIndexesTbl, lnOldArea, lcTagName, llInIndex

        *
        *Returns array of tag names where a given field is part of the tag expression
        *

        lcEnumIndexesTbl=THIS.EnumIndexesTbl
        lnOldArea=SELECT()
        SELECT (lcEnumIndexesTbl)
        LOCATE FOR RTRIM(IndexName)=lcTableName
        llInIndex=.F.
        DO WHILE FOUND()
            IF lcFldName $ LclExpr THEN
                lcTagName=RTRIM(TagName)
                THIS.InsaItem(@aIndexes,lcTagName)
                llInIndex=.T.
            ENDIF
            CONTINUE
        ENDDO

        SELECT (lnOldArea)
        RETURN llInIndex

    ENDFUNC


    FUNCTION INKEY
        PARAMETERS lcRmtFieldName, lcTableName
        LOCAL lcEnumRelsTbl, lnOldArea, aRels

        *
        *Checks to see if a given field is in a key (primary or foreign)
        *

        lcTableName=LOWER(lcTableName)
        lcEnumRelsTbl=THIS.EnumRelsTbl
        lnOldArea=SELECT()
        SELECT (lcEnumRelsTbl)
        LOCATE FOR RTRIM(DD_PARENT)==lcTableName OR RTRIM(DD_CHILD)==lcTableName
        llInKey=.F.
        DO WHILE FOUND()
            IF lcRmtFieldName $ DD_CHIEXPR
                lcRelatedTable=IIF(RTRIM(DD_PARENT)=lcTableName,RTRIM(DD_CHILD),RTRIM(DD_PARENT))
                llInKey=.T.
                EXIT
            ENDIF
            CONTINUE
        ENDDO

        SELECT (lnOldArea)
        RETURN llInKey

    ENDFUNC


    PROCEDURE DontIndex
        PARAMETERS lcFieldName, lcTableName
        LOCAL lnOldArea, lcEnumIndexsTbl

        *
        *If user changed data type to something unindexable, don't create the indexes
        *that include the unindexable field
        *

        lnOldArea=SELECT()
        lcEnumIndexsTbl=THIS.EnumIndexesTbl
        SELECT (lcEnumIndexsTbl)
        LOCATE FOR RTRIM(IndexName)==RTRIM(lcTableName) AND lcFieldName $ LclExpr
        DO WHILE FOUND()
            REPLACE IdxError WITH STRTRAN(IDX_NOT_CREATED_LOC,"|1",lcFieldName), ;
                Exported WITH .F., ;
                DontCreate WITH .T.
            CONTINUE
        ENDDO

        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE DefaultsAndRules

        *
        *This proc converts FoxPro defaults and rules to server equivalents
        *
        *In the case of SQL Server, defaults are converted to defaults and rules to stored
        *procedures which are then called from insert and update triggers
        *
        *In the case of Oracle, defaults become ALTER TABLE statements and rules are
        *converted to SQL statements that will wind up in one trigger that executes on
        *the update or insert event
        *
        * If llOraFieldRules, create Oracle field rules to become part of CREATE TABLE

        LOCAL lcEnumTables, lcTableName, lcTableRule, lcRmtTableName, ;
            lcRuleExpression, lcRemoteRule, lcRuleText, lcFldName, ;
            lcDefaultExpression, lcRemoteDefault, lcEnumFields, llRuleSprocCreated, ;
            llDefaultCreated, llDefaultBound, llTableExported, lcRemoteRuleName, ;
            lcRemoteDefaultName, lcFldType, lcRuleError, lcDefaError, lcRmtFldName, ;
            lcConstName, lnTableCount, lcThermMsg, lnError, llShowTherm, llRuleCreated

* If we have an extension object and it has a DefaultsAndRules method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'DefaultsAndRules', 5) and ;
			not This.oExtension.DefaultsAndRules(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        lcEnumTables = THIS.EnumTablesTbl
        SELECT (lcEnumTables)

        * thermometer
        SELECT COUNT(*) FROM (lcEnumTables) WHERE EXPORT = .T. INTO ARRAY aTableCount
        IF THIS.ExportValidation THEN
            lcThermMsg =  CONVERT_RULE_LOC
        ENDIF
        IF THIS.ExportDefaults THEN
            IF EMPTY(lcThermMsg) THEN
                lcThermMsg = CONVERT_DEFAS_LOC
            ELSE
                lcThermMsg = lcThermMsg + AND_LOC + CONVERT_DEFAS_LOC
            ENDIF
        ENDIF
        IF !THIS.ExportDefaults AND !THIS.ExportValidation THEN
            IF THIS.ServerType = "Oracle" THEN
                RETURN
            ELSE
                lcThermMsg = BIND_DEFAS_LOC
            ENDIF
        ELSE
            lcThermMsg = CONVERT_STEM_LOC + lcThermMsg
        ENDIF
        THIS.InitTherm(lcThermMsg,aTableCount,0)
        lnTableCount = 0

        SCAN FOR &lcEnumTables..EXPORT = .T.
            llTableExported = &lcEnumTables..Exported
            lcTableName = RTRIM(&lcEnumTables..TblName)
            lcRmtTableName = RTRIM(&lcEnumTables..RmtTblName)

            lcThermMsg = STRTRAN(THIS_TABLE_LOC, '|1', lcTableName)
            THIS.UpDateTherm(lnTableCount, lcThermMsg)
            lnTableCount = lnTableCount + 1

            lcRuleError=""
            lcDefaultError=""
            lcRemoteRuleName=""
            lcRemoteRule=""
            lcRuleText=""
            lcRemoteDefault=""
            llRuleSprocCreated=.F.
            llDefaultCreated=.F.
            llDefaultBound=.F.
            lcRemoteDefaultName=""
            lcDefaError=""
            lnError=.F.

            *Turn table rules into stored procedures (SQL Server) or constraints (Oracle)
            *Grab the rule

            IF THIS.ExportValidation
                lcRuleExpression = DBGETPROP(lcTableName, "Table", "RuleExpression")
                lcRuleText = DBGETPROP(lcTableName, "Table", "RuleText")

                *If there is a rule, convert it
                IF !EMPTY(lcRuleExpression) THEN

                    *For Oracle, convert rules to trigger code
                    IF THIS.ServerType = ORACLE_SERVER THEN
                        lcConstName = ORA_CONST_TAB_PREFIX + LEFT(lcRmtTableName, MAX_NAME_LENGTH - LEN(ORA_CONST_TAB_PREFIX))
                        lcConstName = THIS.UniqueOraName(lcConstName, .T.)
                        lcRemoteRule = THIS.ConvertToConstraint(lcRuleExpression, lcTableName, lcRmtTableName, lcConstName)

                        * Create table constraint if user is upsizing and has create constraint permissions
                        IF THIS.DoUpsize AND llTableExported && AND this.Perm_default
                            IF THIS.ExecuteTempSPT(lcRemoteRule, @lnError, @lcDefaError)
                                lcRemoteRuleName = lcConstName
                            ELSE
                                REPLACE RuleErrNo WITH lnError
                                THIS.StoreError(lnError, lcRuleError, lcRemoteRule, CONVTO_TRIG_ERR_LOC, lcTableName, TABLE_RULE_LOC)
                            ENDIF
                        ENDIF

                        REPLACE &lcEnumTables..LocalRule WITH lcRuleExpression, ;
                            &lcEnumTables..RRuleName WITH lcRemoteRuleName, ;
                            &lcEnumTables..RmtRule WITH lcRemoteRule, ;
                            &lcEnumTables..RuleError WITH lcRuleError
                    ELSE

                        *For SQL Server, create a sproc
                        lcRemoteRule=THIS.ConvertToSproc(lcRuleExpression,lcRuleText,lcTableName,;
                            lcRmtTableName, "Table", @lcRemoteRuleName)

                        *Create the rule if user is upsizing and has sproc permission
                        IF THIS.DoUpsize AND llTableExported AND THIS.Perm_Sproc THEN
                            *Might have to drop existing sproc
                            IF THIS.MaybeDrop(lcRemoteRuleName,"procedure") THEN
                                llRuleSprocCreated=THIS.ExecuteTempSPT(lcRemoteRule, @lnError, @lcRuleError)
                                IF !llRuleSprocCreated THEN
                                    REPLACE &lcEnumTables..RuleErrNo WITH lnError
                                    THIS.StoreError(lnError,lcRuleError,lcRemoteRule, SPROC_ERR_LOC,lcTableName,TABLE_RULE_LOC)
                                ENDIF
                            ELSE
                                lcRuleError=CANT_DROP_SPROC_LOC
                            ENDIF
                        ENDIF

                        *Store the information for the upsizing report
                        REPLACE &lcEnumTables..LocalRule WITH lcRuleExpression, ;
                            &lcEnumTables..RRuleName WITH lcRemoteRuleName, ;
                            &lcEnumTables..RmtRule WITH lcRemoteRule, ;
                            &lcEnumTables..RuleExport WITH llRuleSprocCreated, ;
                            &lcEnumTables..RuleError WITH lcRuleError

                        lcRuleError=""
                        lnError=.F.

                    ENDIF

                ENDIF

            ENDIF

            *
            *Deal with field rules: change to sprocs (SQL Server) or constraints (Oracle)
            lcEnumFields=THIS.EnumFieldsTbl

            SELECT (lcEnumFields)

            IF THIS.ExportValidation THEN

                SCAN FOR RTRIM(TblName)==lcTableName
                    lcFldName = RTRIM(&lcEnumFields..FldName)
                    lcRmtFldName = RTRIM(&lcEnumFields..RmtFldname)
                    lcRuleExpression = DBGETPROP(lcTableName+"."+lcFldName,"Field","RuleExpression")
                    lcRuleText = DBGETPROP(lcTableName+"."+lcFldName,"Field","RuleText")

                    *do nothing if there's no local rule
                    IF !EMPTY(lcRuleExpression) THEN

                        * For Oracle, convert rules to table constraints
                        IF THIS.ServerType = ORACLE_SERVER
                            lcConstName = ORA_CONST_COL_PREFIX + LEFT(lcRmtFldName, MAX_NAME_LENGTH - LEN(ORA_CONST_COL_PREFIX))
                            lcConstName = THIS.UniqueOraName(lcConstName, .T.)
                            lcRemoteRule = THIS.ConvertToConstraint(lcRuleExpression, lcTableName, lcRmtTableName, lcConstName)

                            * Create table constraint if user is upsizing and has create constraint permissions
                            IF THIS.DoUpsize AND llTableExported && AND this.Perm_default
                                IF THIS.ExecuteTempSPT(lcRemoteRule, @lnError, @lcDefaError)
                                    lcRemoteRuleName = lcConstName
                                ELSE
                                    REPLACE &lcEnumFields..RuleErrNo WITH lnError
                                    THIS.StoreError(lnError, lcRuleError, lcRemoteRule, CONVTO_TRIG_ERR_LOC, lcTableName+"."+lcFldName, FIELD_RULE_LOC)
                                ENDIF
                            ENDIF

                            REPLACE &lcEnumFields..LocalRule WITH lcRuleExpression, ;
                                &lcEnumFields..RRuleName WITH lcRemoteRuleName, ;
                                &lcEnumFields..RmtRule WITH lcRemoteRule, ;
                                &lcEnumFields..RuleError WITH lcRuleError
                        ELSE

*** DH 2015-09-08: pass VFP field name instead of remote field name to ConvertToSproc
*                            lcRemoteRule=THIS.ConvertToSproc(lcRuleExpression,lcRuleText,lcTableName,;
                                lcRmtTableName, "Field", @lcRemoteRuleName, lcRmtFldName)
                            lcRemoteRule=THIS.ConvertToSproc(lcRuleExpression,lcRuleText,lcTableName,;
                                lcRmtTableName, "Field", @lcRemoteRuleName, lcFldName)

                            *Create the sprocs if the user is actually upsizing and has permissions
                            IF THIS.DoUpsize AND llTableExported AND THIS.Perm_Sproc THEN

                                *Create the sproc if there's a rule
                                IF !EMPTY(lcRuleExpression) THEN
                                    *Might have to drop existing sproc
                                    IF THIS.MaybeDrop(lcRemoteRuleName,"procedure") THEN
                                        llRuleSprocCreated=THIS.ExecuteTempSPT(lcRemoteRule, @lnError, @lcRuleError)
                                        IF !llRuleSprocCreated THEN
                                            REPLACE &lcEnumFields..RuleErrNo WITH lnError
                                            THIS.StoreError(lnError,lcRuleError,lcRemoteRule,SPROC_ERR_LOC,lcTableName+"."+lcFldName,FIELD_RULE_LOC)
                                        ENDIF
                                    ELSE
                                        lcRuleError=CANT_DROP_SPROC_LOC
                                    ENDIF
                                ENDIF

                            ENDIF

                            *Store all this stuff
                            REPLACE &lcEnumFields..LocalRule WITH lcRuleExpression, ;
                                &lcEnumFields..RmtRule WITH lcRemoteRule, ;
                                &lcEnumFields..RRuleName WITH lcRemoteRuleName, ;
                                &lcEnumFields..RuleExport WITH llRuleSprocCreated, ;
                                &lcEnumFields..RuleError WITH lcRuleError

                            lcRuleError=""
                            lnError=.F.

                        ENDIF

                    ENDIF

                ENDSCAN

            ENDIF

            * Deal with defaults (depends on server type)
            *
            * Unlike the above code, the difference between Oracle and SQL Server is handled
            * in the procedure ConvertToDefault rather than here

            SCAN FOR RTRIM(TblName) == lcTableName
                lcFldName = RTRIM(&lcEnumFields..FldName)
                llBitType = IIF(RTRIM(&lcEnumFields..RmtType) = "bit", .T., .F.)
                lcDefaultExpression = DBGETPROP(lcTableName + "." + lcFldName, "Field", "DefaultValue")

                IF (THIS.ExportDefaults AND !EMPTY(lcDefaultExpression)) OR llBitType THEN

                    *Convert Fox default to server default
                    do case
						case This.ExportDefaults and ;
                    		not empty(lcDefaultExpression)
*** DH 2015-09-08: pass remote field name to ConvertToDefault
*							lcRemoteDefault = This.ConvertToDefault(lcDefaultExpression, ;
								lcFldName, lcTableName, lcRmtTableName, ;
								@lcRemoteDefaultName)
							lcRmtFldName = RTRIM(&lcEnumFields..RmtFldName)
							lcRemoteDefault = This.ConvertToDefault(lcDefaultExpression, ;
								lcFldName, lcTableName, lcRmtTableName, lcRmtFldName, ;
								@lcRemoteDefaultName)
*** DH 2015-09-08: end of new code
						case This.SQLServer and This.ServerVer >= 9
						otherwise
							lcRemoteDefault = "0"
					endcase

                    *If the default expression is 0 or.F., bind the zero default to field
					if This.SQLServer and This.ServerVer < 9 and ;
						(alltrim(lcRemoteDefault) == '0' or ;
						lcDefaultExpression = '0' or ;
						lcDefaultExpression = '.F.')
                        llZD_field=.T.
                        lcRemoteDefault="0"
                        lcRemoteDefaultName=ZERO_DEFAULT_NAME
                        THIS.ZDUsed=.T.
                    ELSE
                        llZD_field=.F.
                    ENDIF

                    IF !THIS.ExportDefaults AND !llZD_field AND EMPTY(lcDefaultExpression) THEN
                        LOOP
                    ENDIF

                    * Create the default if user is upsizing and has create default permissions
                    IF THIS.DoUpsize AND llTableExported AND THIS.Perm_Default THEN
                        IF llZD_field THEN
                            llDefaultCreated = THIS.ZeroDefault()
                        ELSE
                            * If we're Oracle or a non-zero default, just create the default
                            * (for SQL Server) or alter the table (Oracle)
                            * Might have to drop existing default
                            IF THIS.ServerType = "Oracle" OR THIS.MaybeDrop(lcRemoteDefaultName, "default") THEN
                                llDefaultCreated = THIS.ExecuteTempSPT(lcRemoteDefault, @lnError, @lcDefaError)
                                IF !llDefaultCreated THEN
                                    REPLACE &lcEnumFields..DefaErrNo WITH lnError
                                    THIS.StoreError(lnError,lcDefaError,lcRemoteDefault, DEFA_ERR_LOC,lcTableName+"."+lcFldName,DEFAULT_LOC)
                                ENDIF
                            ELSE
                                lcDefaError = CANT_DROP_DEFA_LOC
                            ENDIF
                        ENDIF

                        *If we're upsizing to SQL Server, need to bind default if successfully created
						if llDefaultCreated and This.SQLServer and ;
							This.ServerVer < 9
                            llDefaultBound = THIS.BindDefault(lcRemoteDefaultName, lcRmtTableName, lcFldName)
                        ELSE
                            llDefaultBound=.F.
                        ENDIF
                    ENDIF

                    REPLACE &lcEnumFields..DEFAULT WITH lcDefaultExpression, ;
                        &lcEnumFields..RmtDefault WITH lcRemoteDefault, ;
                        &lcEnumFields..RDName WITH lcRemoteDefaultName, ;
                        &lcEnumFields..DefaExport WITH llDefaultCreated, ;
                        &lcEnumFields..DefaBound WITH llDefaultBound, ;
                        &lcEnumFields..DefaError WITH lcDefaError

                    lcDefaultError=""
                    lnError=.F.
                ENDIF
            ENDSCAN
            SELECT (lcEnumTables)
        ENDSCAN
		raiseevent(This, 'CompleteProcess')

    ENDPROC


    FUNCTION MungeXbase
        PARAMETERS lcLocalExpression, lcObjectType, lcLocalTableName, lcRemoteTableName

        *Takes an Xbase expression and replaces as many mappable keywords as possible
        *This leaves tons of potential keywords that will not work on SQL Server or Oracle

        LOCAL lcServerSQL, lcRemoteExpression, lnOldArea, lcSetTalk, lcDelimiter, ;
            lnPos, lnPos1, lnPos2, lnPosMax

        *select expression mapping table
        lnOldArea = SELECT()
        lcSetTalk = SET('TALK')
        SET TALK OFF

        IF !USED("ExprMap") THEN
            SELECT 0
            USE ExprMap EXCLUSIVE
            IF THIS.ServerType = "Oracle" THEN
                SET FILTER TO !EMPTY(ExprMap.ORACLE)
            ELSE
                SET FILTER TO !EMPTY(ExprMap.SQLServer)
            ENDIF
        ELSE
            SELECT ExprMap
        ENDIF

        lcRemoteExpression = ''
        DO WHILE !EMPTY(lcLocalExpression)

            * find next language string (i.e. smallest positive lnPos or 0)
            lnPos  = AT("'", lcLocalExpression)
            lnPos1 = AT('"', lcLocalExpression)
            lnPos2 = AT('[', lcLocalExpression)
            lnMax  = LEN(lcLocalExpression) + 1
            lnPos = MIN(IIF(lnPos > 0, lnPos, lnMax), IIF(lnPos1 > 0, lnPos1, lnMax), IIF(lnPos2 > 0, lnPos2, lnMax))
            lnPos = IIF(lnPos = lnMax, 0, lnPos)

            IF lnPos = 0
                lcLanguageString = lcLocalExpression
                lcLocalExpression = ''
            ELSE
                lcLanguageString = LEFT(lcLocalExpression, lnPos - 1)
                lcDelimiter = SUBSTR(lcLocalExpression, lnPos, 1)
                lcLocalExpression = SUBSTR(lcLocalExpression, lnPos + 1)
            ENDIF

            * convert language string to server native syntax
            IF (!EMPTY(lcLanguageString))
                lcRemoteExpression = lcRemoteExpression + THIS.ConvertLanguageString(lcLanguageString, lcObjectType, ;
                    lcLocalTableName, lcRemoteTableName)
            ENDIF

            * find next constant string
            IF !EMPTY(lcLocalExpression)
                lcDelimiter = IIF(lcDelimiter == '[', ']', lcDelimiter)
                lnPos = AT(lcDelimiter, lcLocalExpression)
                IF lnPos = 0
                    lcConstantString = lcLocalExpression
                    lcLocalExpression = ''
                ELSE
                    lcConstantString = LEFT(lcLocalExpression, lnPos - 1)
                    lcLocalExpression = SUBSTR(lcLocalExpression, lnPos + 1)
                ENDIF

                * convert language string to server native syntax
                IF (!EMPTY(lcConstantString))
                    lcRemoteExpression = lcRemoteExpression + "'" + THIS.ConvertConstantString(lcConstantString) + "'"
                ENDIF
            ENDIF

        ENDDO

        SELECT (lnOldArea)
        SET TALK &lcSetTalk
        RETURN ALLTRIM(lcRemoteExpression)

    ENDFUNC

    FUNCTION ConvertLanguageString
        PARAMETERS lcLocalExpression, lcObjectType, lcLocalTableName, lcRemoteTableName

        * Takes an Xbase expression and replaces mappable keywords, functions, names and date constants.
        * Fox function replacement is case-insensitive; Fox name replacement is case-sensitive.
        * This leaves some potential keywords that will not work on SQL Server or Oracle.

        LOCAL lcServerSQL, lnOldArea, lcServerType, lcXbase, lcEnumFields, ;
            aFieldNames, lcExpression, lnOccurence, lnPos, I

        * go through all the keywords for the appropriate server type
        lcLocalExpression = LOWER(lcLocalExpression)
        lcLocalTableName = LOWER(RTRIM(lcLocalTableName))
        lcRemoteTableName = LOWER(RTRIM(lcRemoteTableName))
        lcMapField = IIF(THIS.ServerType = ORACLE_SERVER, "Oracle", "SqlServer")

        lcExpression = CHR(1)+lcLocalExpression
        SCAN
            IF !ExprMap.PAD
                lcServerSQL = RTRIM(ExprMap.&lcMapField.)
            ELSE
                lcServerSQL = " "+ RTRIM(ExprMap.&lcMapField.) + " "
            ENDIF
            lcXbase = LOWER(RTRIM(ExprMap.FoxExpr))
            lcExpression = STRTRAN(lcExpression, lcXbase, LOWER(lcServerSQL))
        ENDSCAN
        lcExpression = SUBSTR(lcExpression,2)

        *create an array of local and remote field names (where the two are different)
        lcEnumFields = RTRIM(THIS.EnumFieldsTbl)
        DIMENSION aFieldNames[1,2]
        SELECT LOWER(FldName), LOWER(RmtFldname) FROM (lcEnumFields) WHERE ;
            RTRIM(&lcEnumFields..TblName)=lcLocalTableName AND ;
            &lcEnumFields..FldName<>&lcEnumFields..RmtFldname ;
            INTO ARRAY aFieldNames

        *replace field names with remotized names in table validation rules
        IF !EMPTY(aFieldNames) THEN
            FOR I=1 TO ALEN(aFieldNames,1)
                lcExpression = STRTRAN(lcExpression, RTRIM(aFieldNames[i,1]), RTRIM(aFieldNames[i,2]))
            NEXT
        ENDIF

        * replace table name with remotized table name
        lcExpression = STRTRAN(lcExpression, lcLocalTableName, lcRemoteTableName)

        * Convert and format date constants
        lcExpression = THIS.ConvertDates(lcExpression)

        RETURN ALLTRIM(lcExpression)

    ENDFUNC

    FUNCTION ConvertConstantString
        PARAMETERS lcString
        LOCAL lnPos, lnOccurrence

        * Takes a constant string expression that should be delimited by single quotes
        * in the remote expression. Replaces all internal single quotes (') with two single quotes ('')

        lnOccurence=1
        DO WHILE AT("'",lcString, lnOccurence) <> 0
            *Just add another chr(39) in front of each one we find
            lnPos = AT("'", lcString, lnOccurence)
            lcString = STUFF(lcString, lnPos, 1, CHR(39)+CHR(39))
            lnOccurence = lnOccurence+2
        ENDDO

        RETURN lcString

    ENDFUNC

    FUNCTION HandleQuotes
        PARAMETERS lcExpr, llNoContractions
        LOCAL lnPos, lnOccurrence

        *Chr(39) is always a contraction in validation rule expressions, default
        *expressions, and validation messages.
        *Start with : "don't" ==> "'don''t'"
        *"create default foo_defa2 as 'don''t'"

        *If I know the string has no contractions, just replace doubles with singles

        IF PARAMETERS()=1
            lnOccurence=1
            DO WHILE ATCC("'",lcExpr,lnOccurence)<>0
                *Just add another chr(39) in front of each one we find
                lnPos=ATCC("'",lcExpr,lnOccurence)
                lcExpr=STUFF(lcExpr,lnPos,1,CHR(39)+CHR(39))
                lnOccurence=lnOccurence+2
            ENDDO
        ENDIF

        lcExpr=STRTRAN(lcExpr,CHR(34),CHR(39))
        RETURN lcExpr

    ENDFUNC

    FUNCTION ConvertDates
*** DH 03/20/2015: rewrote this method for several reasons:
***		- Avoids CTOT which causes compile error if SET STRICTDATE TO 2 (Mike Potjer)
***		- Doesn't work with more than one date in expression
***		- Minimize work inside loop

        PARAMETERS lcExpression
*!*	        LOCAL lcOccurence, lnPos1, lnPos2, ltDateTime, lcLocalDTString, lcRemoteDTString, ;
*!*	            lcCentury, lcDate, lnHour, lcSeconds, lcMark

*!*	        lnOccurence = 1
*!*	        DO WHILE .T.

*!*	            * find next date string
*!*	            lnPos1 = AT("{", lcExpression, lnOccurence)
*!*	            lnPos2 = AT("}", lcExpression, lnOccurence)
*!*	            IF (lnPos1 = 0 OR lnPos2 = 0 OR lnPos1 > lnPos2)
*!*	                EXIT
*!*	            ENDIF

*!*	            lcLocalDTString = SUBSTR(lcExpression, lnPos1, lnPos2 - lnPos1 + 1)
*!*	            ltDateTime = CTOT(SUBSTR(lcExpression, lnPos1 + 1, lnPos2 - lnPos1 -1))

*!*	            lcCentury = SET('CENTURY')
*!*	            lcDate = SET('DATE')
*!*	            lnHour = SET('HOUR')
*!*	            lcSeconds = SET('SECONDS')
*!*	            lcMark = SET('MARK')

*!*	            SET CENTURY ON
*!*	            SET DATE TO AMERICAN
*!*	            SET HOURS TO 12
*!*	            SET SECONDS ON
*!*	            SET MARK TO '/'

*!*	            lcRemoteDTString = DTOC(ltDateTime)

*!*	            SET CENTURY &lcCentury
*!*	            SET DATE TO &lcDate
*!*	            SET HOURS TO lnHour
*!*	            SET SECONDS &lcSeconds
*!*	            SET MARK TO lcMark

*!*	            * need exact format for Oracle
*!*	            IF (THIS.ServerType == ORACLE_SERVER)
*!*	                lcRemoteDTString = "TO_DATE('" + lcRemoteDTString + "','MM/DD/YYYY HH:MI:SS AM')"
*!*	            ENDIF

*!*	            * need quotes for SqlServer
*!*	            IF (THIS.SQLServer)
*!*	                lcRemoteDTString = "'" + lcRemoteDTString + "'"
*!*	            ENDIF

*!*	            lcExpression = STRTRAN(lcExpression, lcLocalDTString, lcRemoteDTString)
*!*	            lnOccurence	= lnOccurence + 1
*!*	        ENDDO
        LOCAL lnPos1, lnPos2, ltDateTime, lcLocalDTString, lcRemoteDTString, ;
            lcCentury, lcDate, lnHour, lcSeconds, lcMark

        lcCentury = SET('CENTURY')
        lcDate = SET('DATE')
        lnHour = SET('HOUR')
        lcSeconds = SET('SECONDS')
        lcMark = SET('MARK')

        SET CENTURY ON
        SET DATE TO AMERICAN
        SET HOURS TO 12
        SET SECONDS ON
        SET MARK TO '/'
        DO WHILE .T.

            * find next date string
            lnPos1 = AT("{", lcExpression)
            lnPos2 = AT("}", lcExpression)
            IF (lnPos1 = 0 OR lnPos2 = 0 OR lnPos1 > lnPos2)
                EXIT
            ENDIF

            lcLocalDTString = SUBSTR(lcExpression, lnPos1, lnPos2 - lnPos1 + 1)
            ltDateTime = evaluate(lcLocalDTString)

            lcRemoteDTString = DTOC(ltDateTime)

            * need exact format for Oracle
            IF (THIS.ServerType == ORACLE_SERVER)
                lcRemoteDTString = "TO_DATE('" + lcRemoteDTString + "','MM/DD/YYYY HH:MI:SS AM')"
            ENDIF

            * need quotes for SqlServer
            IF (THIS.SQLServer)
                lcRemoteDTString = "'" + lcRemoteDTString + "'"
            ENDIF

            lcExpression = STRTRAN(lcExpression, lcLocalDTString, lcRemoteDTString)
        ENDDO

        SET CENTURY &lcCentury
        SET DATE TO &lcDate
        SET HOURS TO lnHour
        SET SECONDS &lcSeconds
        SET MARK TO lcMark

        RETURN lcExpression

    ENDFUNC

    #IF SUPPORT_ORACLE
    FUNCTION ConvertToTrigger
        PARAMETERS lcRemoteExpression,lcRuleText,lcTableName,lcRmtTableName,lcRmtFldName
        LOCAL lcCRLF, lcEnumFieldsTbl

        lcCRLF=CHR(10)+CHR(13)

        *Try to make expression Oracle-ish
        lcRemoteExpression=THIS.MungeXbase(lcRemoteExpression, "foo", lcTableName, lcRmtTableName)

        *make sure the user doesn't have stuff of the form <tblname>.<fldname> in there
        lcRemoteExpression=STRTRAN(lcRemoteExpression,lcRmtTableName+".")

        IF EMPTY(lcRuleText) THEN
            IF !EMPTY(lcRmtFldName) THEN
                *Oracle wants error messages to look like this:
                *'Value entered violates rule for field 'cust' in table 'customer'."
                *Note the use of single and double quotes is the exact opposite of SQL Server
                lcRuleText=STRTRAN(DEFLT_FLDMSG_LOC,"'|1'", gc2QT + lcRmtFldName + gc2QT)
                lcRuleText=STRTRAN(lcRuleText,"'|2'", gc2QT + lcRmtTableName + gc2QT)
            ELSE
                lcRuleText=STRTRAN(DEFLT_TBLMSG_LOC,"'|1'", gc2QT + lcRmtTableName + gc2QT)
            ENDIF
            lcRuleText="'"+lcRuleText+"'"
        ELSE

            *Replace existing single quotes with two single quotes (they should then appear as one single quote mark)
            lcSingle=CHR(39)
            lcRuleText=STRTRAN(lcRuleText,"'",lcSingle+lcSingle)
            *Replace all double quote marks with single quote marks
            lcRuleText=STRTRAN(lcRuleText, gcQT, "'")
        ENDIF

        *Build comment
        IF !EMPTY(lcRmtFldName) THEN
            *We're dealing with a field validation rule
            *Need to replace <table>.<fldname> name with ":new.<fldname>"
            lcRemoteExpression=STRTRAN(lcRemoteExpression,lcRmtFldName,":new." + lcRmtFldName)
            lcSQL=THIS.BuildComment(FTRIG_COMMENT_LOC,lcRmtFldName)
        ELSE
            *Table validation rule
            lcSQL=THIS.BuildComment(TTRIG_COMMENT_LOC,"blah blah blah")

            *Get array of field names
            lcEnumFieldsTbl=THIS.EnumFieldsTbl
            SELECT RmtFldname FROM (lcEnumFieldsTbl) WHERE RTRIM(TblName)==lcTableName ;
                INTO ARRAY aFieldNames

            *Need to replace <fldname> name with ":new.<fldname>"
            FOR I=1 TO ALEN(aFieldNames,1)
                lcRemoteExpression=STRTRAN(lcRemoteExpression,RTRIM(aFieldNames[i]),":new." + RTRIM(aFieldNames[i]))
            NEXT

        ENDIF

        *Complete SQL string
        lcSQL=lcSQL +	"IF NOT (" + lcRemoteExpression + ")" + lcCRLF
        lcSQL=lcSQL + 	"     THEN raise_application_error(" + ERR_SVR_RULEVIO_ORA + ", " + lcRuleText + ");" + lcCRLF
        lcSQL=lcSQL + 	"END IF ; " + lcCRLF

        RETURN lcSQL

    ENDFUNC
#ENDIF


#IF SUPPORT_ORACLE
    FUNCTION TestTrigger
        PARAMETERS lcTrigger, lcTable, lcFieldName, lnError, lcErrMsg

        *
        *Checks to see if a rule can be successfully converted to an Oracle trigger
        *

        *If user is just generating a script, just return .T.
        IF !THIS.DoUpsize THEN
            RETURN .T.
        ENDIF

        *Put the trigger together
        lcSQL="CREATE TRIGGER " + lcTable + TRIG_NAME
        lcSQL=lcSQL+" BEFORE INSERT OR UPDATE "
        IF !EMPTY(lcFieldName) THEN
            lcSQL=lcSQL+" OF " + lcFieldName
        ENDIF
        lcSQL=lcSQL + " ON " + lcTable + " FOR EACH ROW BEGIN "
        lcSQL=lcSQL + lcTrigger + "END;"

        *Run it up the flag pole...

        IF !THIS.ExecuteTempSPT(lcSQL, @lnError, @lcErrMsg) THEN
            RETURN .F.
        ELSE
            *Drop the trigger
            lcSQL="DROP TRIGGER " + lcTable + TRIG_NAME
            =THIS.ExecuteTempSPT(lcSQL)
            RETURN .T.
        ENDIF

    ENDFUNC
#ENDIF


    FUNCTION ConvertToSproc
        PARAMETERS lcExpression, lcMessage, lcTableName, lcRmtTableName,  lcObjectType, lcSprocName, lcFldName

        *
        *Takes an Xbase rule and turns it into a stored procedure
        *

        *Do what you can to make the Xbase expression work on the server
        lcExpression=THIS.MungeXbase(lcExpression, lcObjectType, lcTableName, lcRmtTableName)

        LOCAL lcSQL, lcCRLF
        lcCR=CHR(10)
        lcCRLF=CHR(10)+CHR(13)

        *if there's no error message, build a default one
        IF EMPTY(lcMessage) THEN
            IF lcObjectType="Field" THEN
                lcMessage=STRTRAN(DEFLT_FLDMSG_LOC,'|1',lcFldName)
                lcMessage=STRTRAN(lcMessage,'|2',lcTableName)
            ELSE
                lcMessage=STRTRAN(DEFLT_TBLMSG_LOC,'|1',lcTableName)
            ENDIF
        ENDIF

        *need to operate on error message to make sure quotes are right
        lcMessage=THIS.HandleQuotes(lcMessage)
        IF LEFT(lcMessage,1)<>CHR(39)
            lcMessage=CHR(39)+ lcMessage + CHR(39)
        ENDIF

        *Operate on the sproc name
        IF lcObjectType="Field" THEN
            *For field validation sprocs, the aim is to get something
            *like "vrf_customer_lname"
            lcSprocName=THIS.NameObject(lcRmtTableName,lcFldName,FLD_SPROC_PREFIX,MAX_NAME_LENGTH)
        ELSE
            lcSprocName=TBL_SPROC_PREFIX + LEFT(lcRmtTableName,MAX_NAME_LENGTH-LEN(FLD_SPROC_PREFIX))
        ENDIF

        lcSQL=        "CREATE PROCEDURE " + lcSprocName + " @status char(10) output AS" + lcCRLF
        lcSQL=lcSQL + "/*" + lcCR
        lcSQL=lcSQL + " * TABLE VALIDATION RULE FOR " + gcQT + lcRmtTableName + gcQT + lcCR
        lcSQL=lcSQL + " */"  + lcCRLF + lcCRLF
        lcSQL=lcSQL + "IF @status='Failed'"+lcCR
        lcSQL=lcSQL + "      RETURN" +lcCRLF
        lcSQL=lcSQL + "IF (SELECT Count(*) FROM " + lcRmtTableName + " WHERE NOT (" + lcExpression + ")) > 0" + lcCR
        lcSQL=lcSQL + "      BEGIN" + lcCR
*** DH 2015-09-08: use syntax for SQL Server 2005+ for RAISERROR
*        lcSQL=lcSQL + "           RAISERROR " + ERR_SVR_RULEVIO_SQL +  " " + lcMessage +  lcCR
        lcSQL=lcSQL + "           RAISERROR (" + lcMessage + ", -1, -1)" + lcCR
        lcSQL=lcSQL + "           SELECT @status='Failed'" + lcCR
        lcSQL=lcSQL + "      END" + lcCR
        lcSQL=lcSQL + "ELSE" + lcCR
        lcSQL=lcSQL + "      BEGIN" + lcCR
        lcSQL=lcSQL + "          SELECT @status='Succeeded'" + lcCR
        lcSQL=lcSQL + "      END" + lcCRLF

        RETURN lcSQL

    ENDFUNC


    FUNCTION ConvertToDefault
*** DH 2015-09-08: accept remote field name
*        PARAMETERS lcDefaultExpression, lcFieldName, lcTableName, lcRemTableName, lcRemoteDefaultName
        PARAMETERS lcDefaultExpression, lcFieldName, lcTableName, lcRemTableName, lcRemFieldName, lcRemoteDefaultName
        LOCAL lcSQL

        * Defaults become ALTER TABLE statements for Oracle and Defaults on SQL Server
        * Try to make Xbase expression more server like

        lcDefaultExpression = THIS.MungeXbase(lcDefaultExpression, "Field", lcTableName, "")

        * 1/8/01 jvf, Bug ID 155006, VFP7:Upsizing Wizard Does not Create Default Value
        * If the defaul tvalue = [""] (without brackets) then an error occurs
        * because 'lcDefaultExpression' will be empty...
        * (eg, "CREATE DEFAULT <lcRemoteDefaultName> AS <lcDefaultExpression>")
        * so checking for this and setting 'lcDefaultExpression' to "''" will solve the problem.
        IF EMPTY(lcDefaultExpression) THEN
            lcDefaultExpression = "''"
        ENDIF

        IF THIS.ServerType = ORACLE_SERVER THEN
*** DH 2015-09-08: use remote field name instead of VFP field name
*            lcSQL = "ALTER TABLE " + lcRemTableName + " MODIFY (" + lcFieldName + " DEFAULT " + lcDefaultExpression + ")"
            lcSQL = "ALTER TABLE " + lcRemTableName + " MODIFY (" + lcRemFieldName + " DEFAULT " + lcDefaultExpression + ")"
        ENDIF

        IF THIS.SQLServer THEN
            lcRemoteDefaultName = THIS.NameObject(lcTableName, lcFieldName, DEFAULT_PREFIX,MAX_NAME_LENGTH)
			if This.ServerVer >= 9
*** DH 2015-09-08: use remote field name instead of VFP field name
*				lcSQL = 'ALTER TABLE ' + lcRemTableName + ;
					' ADD CONSTRAINT [' + lcRemoteDefaultName + ;
					'] DEFAULT ' + lcDefaultExpression + ;
					' FOR ' + lcFieldName
				lcSQL = 'ALTER TABLE ' + lcRemTableName + ;
					' ADD CONSTRAINT [' + lcRemoteDefaultName + ;
					'] DEFAULT ' + lcDefaultExpression + ;
					' FOR ' + lcRemFieldName
			else
            	lcSQL = "CREATE DEFAULT [" + lcRemoteDefaultName + ;
            		"] AS " + lcDefaultExpression
			endif This.ServerVer >= 9
        ENDIF

        RETURN lcSQL

    ENDFUNC


    FUNCTION ConvertToConstraint
        PARAMETERS lcRuleExpression, lcTableName, lcRemTableName, lcConstName
        LOCAL lcSQL

        * Convert rules into table and field constraints for Oracle and SQL Server
        lcRuleExpression = THIS.MungeXbase(lcRuleExpression, "Field", lcTableName, "")

        IF THIS.ServerType = ORACLE_SERVER
            lcSQL = "ALTER TABLE " + lcRemTableName + ;
                " ADD (CONSTRAINT " + lcConstName + " CHECK(" + lcRuleExpression + "))"
        ENDIF

        * For future use
        IF THIS.SQLServer THEN
        ENDIF

        RETURN lcSQL
    ENDFUNC


    FUNCTION BindDefault
        PARAMETERS lcRemoteDefaultName, lcRmtTableName,lcFldName
        LOCAL lcSQL, llRetVal

        *bind a default to a field

        lcSQL="sp_bindefault " + lcRemoteDefaultName + ", " ;
            + "'" + lcRmtTableName + "." + lcFldName + "'"

        llRetVal=THIS.ExecuteTempSPT(lcSQL)

        RETURN llRetVal

    ENDFUNC


    FUNCTION ZeroDefault
        LOCAL lcSQL, llRetVal, llDefaultExists, dummy,lcSQT

        IF THIS.ZeroDefaultCreated THEN
            RETURN .T.
        ELSE
            lcSQT=CHR(39)
            lcSQL="select uid from sysobjects where name =" + lcSQT + ZERO_DEFAULT_NAME + lcSQT
            llDefaultExists=THIS.SingleValueSPT(lcSQL, dummy, "uid")
            *Only create it if it doesn't exist already
            IF !llDefaultExists THEN
                lcSQL= "CREATE DEFAULT " + ZERO_DEFAULT_NAME + " AS 0"
                llRetVal=THIS.ExecuteTempSPT(lcSQL)
                THIS.ZeroDefaultCreated=llRetVal
                RETURN llRetVal
            ELSE
                RETURN .T.
            ENDIF
        ENDIF

    ENDFUNC


    PROCEDURE BuildRiCode

        *
        * Generates an ALTER TABLE statement for Oracle, code for trigger for SQL Server
        * for all relations between tables that are being upsized
        *

        LOCAL lcEnumTables, lnOldArea, lcTableName, aFldNames, lcCRLF, ;
            aNewForeign, aNewPrimary, llRetVal, lcErr, llSkipChildTbl, ;
            lcEnumRelsTbl,lcEnum_Indexes, lcUpdateType, lcDeleteType, ;
            lcInsertType, lnTableCount, lcThermMsg, lcXPkey, ;
            lcParentLoc, lcChildLoc, lnError, lcErrMsg, lcOParen, lcCParen

* If we have an extension object and it has a BuildRICode method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'BuildRICode', 5) and ;
			not This.oExtension.BuildRICode(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        lcCRLF = CHR(13)
        lnOldArea = SELECT()
        lcEnumTables = THIS.EnumTablesTbl

        * Go grab all the relation information for tables in the source database
        THIS.GetRiInfo
        lcEnumRelsTbl = THIS.EnumRelsTbl
        lcEnum_Indexes = THIS.EnumIndexesTbl

        * We only want to deal with relations where both tables were successfully upsized or,
        * if we're generating a script, relations where both tables were selected to upsize
        SELECT (lcEnumRelsTbl)
        SCAN
            IF THIS.TableUpsized(RTRIM(&lcEnumRelsTbl..DD_PARENT)) ;
                    AND THIS.TableUpsized(RTRIM(&lcEnumRelsTbl..DD_CHILD)) THEN
                REPLACE EXPORT WITH .T.
            ELSE
                REPLACE EXPORT WITH .F.
            ENDIF
        ENDSCAN

        SELECT COUNT(*) FROM (lcEnumRelsTbl) WHERE EXPORT=.T. INTO ARRAY aTableCount
        lnTableCount=0
        THIS.InitTherm(BUILDING_RI_LOC,aTableCount,0)

        SCAN FOR EXPORT = .T.
            lcParent=RTRIM(&lcEnumRelsTbl..Dd_RmtPar)
            lcChild=RTRIM(&lcEnumRelsTbl..Dd_RmtChi)
            lcParentLoc=RTRIM(&lcEnumRelsTbl..DD_PARENT)
            lcChildLoc=RTRIM(&lcEnumRelsTbl..DD_CHILD)
            lcNewPrimary=RTRIM(&lcEnumRelsTbl..DD_PAREXPR)
            lcNewForeign=RTRIM(&lcEnumRelsTbl..DD_CHIEXPR)
            llClustIdxOK=.T.

            * Therm stuff
            lcThermMsg = STRTRAN(RI_THIS_LOC,'|1',lcParentLoc)
            lcThermMsg = STRTRAN(lcThermMsg,'|2',lcChildLoc)
            THIS.UpDateTherm(lnTableCount,lcThermMsg)
            lnTableCount = lnTableCount+1

            * Pick up what kind of relation type this is
            lcUpdateType = &lcEnumRelsTbl..dd_update
            lcDeleteType = &lcEnumRelsTbl..dd_delete
            lcInsertType = &lcEnumRelsTbl..dd_insert

            * For report
            IF lcUpdateType = "I" AND lcDeleteType = "I" AND lcInsertType = "I" ;
                    AND EMPTY(&lcEnumRelsTbl..ClustName) THEN
                REPLACE &lcEnumRelsTbl..Exported WITH .F.
            ELSE
                REPLACE &lcEnumRelsTbl..Exported WITH .T.
            ENDIF

            * Turn fields in keys into an array (used later all over the place)
            DIMENSION aNewForeign[1], aNewPrimary[1]
            aNewForeign[1]=""
            aNewPrimary[1]=""
            THIS.KeyArray(lcNewForeign,@aNewForeign)
            THIS.KeyArray(lcNewPrimary,@aNewPrimary)

            * do simple comparison of keys in case they won't match up
            IF ALEN(aNewForeign,1)<>ALEN(aNewPrimary,1) THEN
                *mark the relation as unupsizable and move to the next relation
                THIS.StoreError(.NULL.,"", "", KEYS_MISMATCH_LOC,lcParent+":"+lcChild,RI_LOC)
                REPLACE RIError WITH KEYS_MISMATCH_LOC, Exported WITH .F.
                LOOP
            ENDIF

            * Make sure keyfields are less than 17
            IF ALEN(aNewForeign,1)>MAX_INDEX_FIELDS THEN
                *mark the relation as unupsizable and move to the next relation
                REPLACE RIError WITH TOO_MANY_FIELDS_LOC, Exported WITH .F.
                THIS.StoreError(.NULL.,"", "", TOO_MANY_FIELDS_LOC,lcParent+":"+lcChild,RI_LOC)
                LOOP
            ENDIF

            * Here's the RI plan:
            *SQL Server: use triggers for everything
            *Oracle: create two triggers that enforces RI via SQL *OR*
            *use DRI (which won't cascade updates)
            *SQL '95: create triggers for everything *OR*
            *use DRI (which won't cascade updates or deletes)
            *note: this code currently is not aware of SQL '95

            *
            * This block of code handles DRI for SQL '95 and Oracle
            * For SQL '95 it implements Update-restrict and Delete-restrict
            * For Oracle, it also implements Delete-cascades
            *

*** DH 12/15/2014: use only SQL Server rather than variants of it
***            IF THIS.ExportDRI AND ;
                    (THIS.ServerType = "Oracle" OR THIS.ServerType = "SQL Server95")
            IF THIS.ExportDRI AND ;
                    (THIS.ServerType = "Oracle" OR left(THIS.ServerType, 10) = 'SQL Server')
                *Implement RI constraints at table level since foreign key may be compound

                *Deal with parent table first (or child constraints will fail)

                *See if the table already has a primary key that's correct for RI purposes
                *(i.e. the same as the one we'd create anyway)

                SELECT (lcEnumTables)
                LOCATE FOR RTRIM(TblName)==lcParent

                *Here are the cases handled below:
                *has no primary key: add pkey, convert old non-pkey index to type pkey,
                *mark as already created so it doesn't get recreated
                *has primary key, it's right: create the pkey now and mark the index
                *as created
                *has primary key, it's wrong: log error

                lcXPkey=RTRIM(&lcEnumTables..PkeyExpr)
                lcTagName=RTRIM(&lcEnumTables..PKTagName)
                IF lcXPkey==lcNewPrimary OR EMPTY(lcXPkey)THEN

                    IF lcTagName=="" THEN
                        lcConstraintName=CHRTRAN(lcNewPrimary,",[]","")
                        lcConstraintName=PRIMARY_KEY_PREFIX + LEFT(lcConstraintName, MAX_NAME_LENGTH-LEN(PRIMARY_KEY_PREFIX))
                    ELSE
                        lcConstraintName=lcTagName
                    ENDIF
                    lcConstraintName=THIS.UniqueOraName(lcConstraintName)

                    IF THIS.ServerType="Oracle" THEN
                        lcOParen="("
                        lcCParen=")"
                    ELSE
                        lcOParen=""
                        lcCParen=""
                    ENDIF

                    *Add primary key constraint
                    lcSQL="ALTER TABLE [" + lcParent + "]"
                    lcSQL=lcSQL + " ADD " + lcOParen + "CONSTRAINT "
                    lcSQL=lcSQL + "[" + lcConstraintName + "] PRIMARY KEY"
                    lcSQL=lcSQL + " (" + lcNewPrimary + ")" + lcCParen

                    *Execute the statement if appropriate

                    IF THIS.DoUpsize THEN
                        llRetVal=THIS.ExecuteTempSPT(lcSQL, @lnError, @lcErrMsg)
                        IF !llRetVal THEN
                            THIS.StoreError(lnError, lcErrMsg, lcSQL, DRI_ERR_LOC,lcParent,RI_LOC)
                        ENDIF
                    ENDIF

                    SELECT (lcEnum_Indexes)
                    LOCATE FOR UPPER(RTRIM(IndexName))==UPPER(lcParentLoc) ;
                        AND UPPER(RTRIM(RmtExpr))==UPPER(lcNewPrimary)
                    IF EMPTY(lcXPkey) THEN
                        *If the table had no true primary key, change the index acting
                        *as primary key to a real primary key (even though we just created it)
                        *This is done just for the purposes of the report and the SQL script
                        REPLACE &lcEnum_Indexes..RmtType WITH "PRIMARY KEY", ;
                            &lcEnum_Indexes..Exported WITH llRetVal, ;
                            &lcEnum_Indexes..IndexSQL WITH lcSQL, ;
                            &lcEnum_Indexes..RmtName WITH lcConstraintName
                    ELSE
                        *Mark the index as already created
                        REPLACE &lcEnum_Indexes..Exported WITH llRetVal, ;
                            &lcEnum_Indexes..IndexSQL WITH lcSQL
                    ENDIF
                    SELECT (lcEnumTables)
                    llSkipChildTbl=.F.

                ELSE

                    *log error that this rel couldn't be created because table has
                    *more than one primary key
                    lcErr=STRTRAN(MULTIPLE_PKEYS_LOC,"|1",lcParent)
                    SELECT (lcEnumRelsTbl)
                    REPLACE RIError WITH lcErr ADDITIVE
                    llSkipChildTbl=.T.

                ENDIF

                *Deal with child table
                IF !llSkipChildTbl THEN

                    lcConstraintName=CHRTRAN(lcNewForeign,",[]","")
                    lcConstraintName=FOREIGN_KEY_PREFIX + LEFT(lcConstraintName, MAX_NAME_LENGTH-LEN(FOREIGN_KEY_PREFIX))
                    lcConstraintName=THIS.UniqueOraName(lcConstraintName)

                    lcSQL="ALTER TABLE [" + lcChild + "] WITH NOCHECK" &&  NOCHECK Add JEI RKR 2005.03.29
                    lcSQL=lcSQL + " ADD " +lcOParen + "CONSTRAINT "
                    lcSQL=lcSQL + "[" + lcConstraintName + "]" +  " FOREIGN KEY "
                    lcSQL=lcSQL + " (" + lcNewForeign + ")"
                    lcSQL=lcSQL + " REFERENCES [" + lcParent + "] (" + lcNewPrimary +")"

                    *SQL95 does not support cascading deletes
                    *Neither Oracle nor SQL95 support cascading updates via DRI
                    *{Change JEI - RKR 2005.03.28
*                    IF lcDeleteType=CASCADE_CHAR_LOC AND THIS.ServerType="Oracle" THEN
*** DH 12/15/2014: use only SQL Server rather than variants of it
***                    IF (lcDeleteType=CASCADE_CHAR_LOC AND THIS.ServerType="Oracle") OR ;
                    	((THIS.ServerType = "SQL Server95" AND This.ServerVer >= 8) AND ;
                    	(lcUpdateType = CASCADE_CHAR_LOC OR  lcDeleteType = CASCADE_CHAR_LOC))THEN
                    IF (lcDeleteType=CASCADE_CHAR_LOC AND THIS.ServerType="Oracle") OR ;
                    	((left(THIS.ServerType, 10) = 'SQL Server' AND This.ServerVer >= 8) AND ;
                    	(lcUpdateType = CASCADE_CHAR_LOC OR  lcDeleteType = CASCADE_CHAR_LOC))THEN

                    	IF THIS.ServerType="Oracle"
                    		lcSQL=lcSQL + " ON DELETE CASCADE"
                    	ELSE
                    		* SQL Server
                    		IF lcDeleteType = CASCADE_CHAR_LOC
                    			lcSQL=lcSQL + " ON DELETE CASCADE"
                    		ENDIF                    		
                    		IF lcUpdateType = CASCADE_CHAR_LOC
                    			lcSQL=lcSQL + " ON UPDATE CASCADE"
                    		ENDIF
                    	ENDIF
*                        lcSQL=lcSQL + " ON DELETE CASCADE)"                       
                    *}Change JEI - RKR 2005.03.28
                    ELSE
                        lcSQL=lcSQL + lcCParen
                    ENDIF
                    
                    *Execute the statement if appropriate
                    IF THIS.DoUpsize AND llRetVal THEN
                        llRetVal=THIS.ExecuteTempSPT(lcSQL,@lnError, @lcErrMsg)
                        IF !llRetVal THEN
                            THIS.StoreError(lnError, lcErrMsg, lcSQL, DRI_ERR_LOC,lcChild,RI_LOC)
                        ENDIF
                    ENDIF

                    *Add comment and tack on to existing table definition sql
                    lcSQL = lcCRLF + lcCRLF + ORA_FKEY_COMMENT_LOC + lcCRLF + lcSQL
                    THIS.StoreRiCode(lcChild,"TableSQL",lcSQL,"FKeyCrea",llRetVal)

                    SELECT (lcEnumRelsTbl)

                ENDIF

                LOOP	&& continue after DRI code

            ENDIF
            *End of DRI code

            * PARENT DELETE RI
            * Prevents deleting a PARENT record for which CHILD records exist,
            * or deletes dependent CHILD records (cascading).

            * PARENT DELETE for SQL 4.x or SQL '95 and cascade delete
            IF (THIS.SQLServer AND lcDeleteType <> IGNORE_CHAR_LOC) THEN

                lcRestr = THIS.BuildRestr(@aNewPrimary, "deleted", @aNewForeign, lcChild, "AND")
                IF lcDeleteType = CASCADE_CHAR_LOC THEN
                    lcSQL = THIS.BuildComment(CASCADE_DELETES_LOC, lcChild)
                    lcSQL = lcSQL + "DELETE " + lcChild + " FROM deleted, " + lcChild +  " WHERE " + lcRestr + lcCRLF
                ELSE
                    lcErrMsg = THIS.HandleQuotes(gcQT + STRTRAN(DEPENDENT_ROWS_LOC,"|1",lcChild) + gcQT)
                    lcSQL= THIS.BuildComment(PREVENT_DELETES_LOC, lcChild)
                    lcSQL= lcSQL + "IF (SELECT COUNT(*) FROM deleted, " + lcChild + " WHERE (" + lcRestr + ")) > 0" + lcCRLF
                    lcSQL= lcSQL + "    BEGIN" + lcCRLF
*** DH 2015-09-08: use syntax for SQL Server 2005+ for RAISERROR
*                    lcSQL= lcSQL + "    RAISERROR " + ERR_SVR_DELREFVIO + " " + lcErrMsg + lcCRLF
                    lcSQL= lcSQL + "    RAISERROR (" + lcErrMsg + ", -1, -1)" + lcCRLF
                    lcSQL= lcSQL + "    SELECT @status='Failed'" + lcCRLF
                    lcSQL= lcSQL + "    END" + lcCRLF
                ENDIF

                *save this code
                THIS.StoreRiCode(lcParent,"DeleteRI",lcSQL)
            ENDIF

            * PARENT DELETE for Oracle
            IF THIS.ServerType = "Oracle" AND lcDeleteType <> IGNORE_CHAR_LOC

                lcRestr = THIS.BuildRestr(@aNewForeign, lcChild, @aNewPrimary, ":old", "AND")
                IF lcDeleteType = CASCADE_CHAR_LOC THEN
                    lcSQL = THIS.BuildComment(CASCADE_DELETES_LOC, lcChild)
                    lcSQL = lcSQL + "IF DELETING THEN " + lcCRLF
                    lcSQL = lcSQL + "    DELETE FROM " + lcChild +  " WHERE " + lcRestr + ";" + lcCRLF
                    lcSQL = lcSQL + "END IF;" + lcCRLF
                ELSE
                    lcErrMsg = "'" + STRTRAN(DEPENDENT_ROWS_LOC, "'|1'", gc2QT + lcChild + gc2QT) + "'"
                    lcSQL = THIS.BuildComment(PREVENT_DELETES_LOC, lcChild)
                    lcSQL = lcSQL + "IF DELETING THEN " + lcCRLF
                    lcSQL = lcSQL + "    SELECT COUNT(*) INTO " + REC_COUNT_VAR + " FROM " + lcChild + " WHERE (" + lcRestr + ");" + lcCRLF
                    lcSQL = lcSQL + "    IF " + REC_COUNT_VAR + " > 0 THEN " + lcCRLF
                    lcSQL = lcSQL + "        raise_application_error(" + ERR_SVR_DELREFVIO_ORA + ", " + lcErrMsg + ");" + lcCRLF
                    lcSQL = lcSQL + "    END IF;" + lcCRLF
                    lcSQL = lcSQL + "END IF;" + lcCRLF
                ENDIF

                * save this code
                THIS.StoreRiCode(lcParent, "DeleteRI", lcSQL)

            ENDIF

            * PARENT UPDATE trigger
            * Prevents changing a PARENT key for which CHILD records exist,
            * or keeps CHILD keys in sync with PARENT keys (cascading).
            * Executed only if SQL Server or if Oracle or SQL '95 when updates are cascaded
            * Handle SQL Server (4.x or '95) case here

            * PARENT UPDATE for Sql Server
            IF THIS.SQLServer AND lcUpdateType <> IGNORE_CHAR_LOC

                lcRestr = THIS.BuildRestr(@aNewPrimary, "deleted", @aNewForeign, lcChild, "AND")
                IF lcUpdateType = CASCADE_CHAR_LOC THEN
                    lcSQL = THIS.BuildComment(CASCADE_UPDATES_LOC,lcChild)
                ELSE
                    lcSQL = THIS.BuildComment(PREVENT_M_UPDATES_LOC,lcChild)
                ENDIF
                lcSQL = lcSQL + "IF " + THIS.BuildUpdateTest(@aNewPrimary)
                lcSQL = lcSQL + " AND @status<>'Failed'" + lcCRLF
                lcSQL = lcSQL + "    BEGIN" + lcCRLF
                IF lcUpdateType=CASCADE_CHAR_LOC THEN
                    lcSetKeys = THIS.BuildRestr(@aNewForeign, lcChild, @aNewPrimary, "inserted", ",")
                    lcSQL = lcSQL + "         UPDATE " + lcChild + lcCRLF
                    lcSQL = lcSQL + "         SET " + lcSetKeys + lcCRLF
                    lcSQL = lcSQL + "         FROM " + lcChild + ", deleted, inserted" + lcCRLF
                    lcSQL = lcSQL + "         WHERE " + lcRestr + lcCRLF
                ELSE
                    lcErrMsg=THIS.HandleQuotes(gcQT+STRTRAN(DEPENDENT_ROWS_LOC,"|1",lcChild)+ gcQT)
                    lcSQL = lcSQL + "    IF (SELECT COUNT(*) FROM deleted, " + lcChild + " WHERE (" + lcRestr + ")) > 0" + lcCRLF
                    lcSQL = lcSQL + "        BEGIN" + lcCRLF
*** DH 2015-09-08: use syntax for SQL Server 2005+ for RAISERROR
*                    lcSQL= lcSQL + "    RAISERROR " + ERR_SVR_DELREFVIO + " " + lcErrMsg + lcCRLF
                    lcSQL= lcSQL + "    RAISERROR (" + lcErrMsg + ", -1, -1)" + lcCRLF
                    lcSQL = lcSQL + "            SELECT @status='Failed'" + lcCRLF
                    lcSQL = lcSQL + "            END" + lcCRLF
                ENDIF
                lcSQL = lcSQL + "    END" + lcCRLF

                *save this code
                THIS.StoreRiCode(lcParent, "UpdateRI", lcSQL)
            ENDIF

            * PARENT UPDATE for Oracle
            IF THIS.ServerType = "Oracle" AND lcUpdateType <> IGNORE_CHAR_LOC

                lcRestr = THIS.BuildRestr(@aNewForeign, lcChild, @aNewPrimary, ":old", "AND")
                lcUpdTest = THIS.BuildRestr(@aNewForeign, ":old", @aNewPrimary, ":new", "OR")
                lcUpdTest = STRTRAN(lcUpdTest,"=","!=")
                IF lcUpdateType = CASCADE_CHAR_LOC THEN
                    lcSQL = THIS.BuildComment(CASCADE_UPDATES_LOC, lcChild)
                ELSE
                    lcSQL = THIS.BuildComment(PREVENT_M_UPDATES_LOC, lcChild)
                ENDIF
                IF lcUpdateType = CASCADE_CHAR_LOC THEN
                    lcSetKeys = THIS.BuildRestr(@aNewForeign, lcChild, @aNewPrimary, ":new", ",")
                    lcSQL = lcSQL + "IF UPDATING AND " + lcUpdTest + " THEN" + lcCRLF
                    lcSQL = lcSQL + "    UPDATE " + lcChild + lcCRLF
                    lcSQL = lcSQL + "    SET " + lcSetKeys + lcCRLF
                    lcSQL = lcSQL + "    WHERE " + lcRestr + ";" + lcCRLF
                    lcSQL = lcSQL + "END IF;" + lcCRLF
                ELSE
                    lcErrMsg = "'" + STRTRAN(DEPENDENT_ROWS_LOC, "'|1'", gc2QT + lcChild + gc2QT) +  "'"
                    lcSQL = lcSQL + "IF UPDATING AND " + lcUpdTest + " THEN" + lcCRLF
                    lcSQL = lcSQL + "    SELECT COUNT(*) INTO " + REC_COUNT_VAR + " FROM " + lcChild + " WHERE (" + lcRestr + ");" + lcCRLF
                    lcSQL = lcSQL + "    IF " + REC_COUNT_VAR + " > 0 THEN" + lcCRLF
                    lcSQL = lcSQL + "        raise_application_error(" + ERR_SVR_UPDREFVIO_ORA + ", " + lcErrMsg + ");" + lcCRLF
                    lcSQL = lcSQL + "    END IF;" + lcCRLF
                    lcSQL = lcSQL + "END IF;" + lcCRLF
                ENDIF

                THIS.StoreRiCode(lcParent, "UpdateRI", lcSQL)

            ENDIF

            * CHILD UPDATE trigger
            * Prevents changing or adding a CHILD record to a key not in the PARENT table
            * CHILD INSERT trigger
            * Prevents adding a CHILD record for which no PARENT record exists

            * CHILD UPDATE AND INSERT for SQL Server 4.x and 95
            IF THIS.SQLServer THEN

                * CHILD UPDATE trigger
                IF lcUpdateType = RESTRICT_CHAR_LOC THEN
                    lcRestr = THIS.BuildRestr(@aNewPrimary, lcParent, @aNewForeign, "inserted", "AND")
                    lcErrMsg = THIS.HandleQuotes(gcQT + STRTRAN(CANT_ORPHAN_LOC, "|1",lcParent) + gcQT)

                    lcSQL = THIS.BuildComment(PREVENT_C_UPDATES_LOC, lcParent)
                    lcSQL = lcSQL + "IF " + THIS.BuildUpdateTest(@aNewForeign)
                    lcSQL =lcSQL + " AND @status<>'Failed'" + lcCRLF
                    lcSQL = lcSQL + "    BEGIN" + lcCRLF
                    lcSQL = lcSQL + "IF (SELECT COUNT(*) FROM inserted) !=" + lcCRLF
                    lcSQL = lcSQL + "           (SELECT COUNT(*) FROM " + lcParent + ", inserted WHERE (" + lcRestr + "))" + lcCRLF
                    lcSQL = lcSQL + "            BEGIN" + lcCRLF
*** DH 2015-09-08: use syntax for SQL Server 2005+ for RAISERROR
*                    lcSQL = lcSQL + "                RAISERROR " + ERR_SVR_UPDREFVIO + " " + lcErrMsg + lcCRLF
                    lcSQL= lcSQL + "                RAISERROR (" + lcErrMsg + ", -1, -1)" + lcCRLF
                    lcSQL = lcSQL + "                SELECT @status = 'Failed'" + lcCRLF
                    lcSQL = lcSQL + "            END" + lcCRLF
                    lcSQL = lcSQL + "    END" + lcCRLF

                    *save this code
                    THIS.StoreRiCode(lcChild,"UpdateRI",lcSQL)

                ENDIF

                * CHILD INSERT trigger
                IF lcInsertType = RESTRICT_CHAR_LOC
                    lcErrMsg = THIS.HandleQuotes(gcQT+ STRTRAN(CANT_ORPHAN_LOC,"|1",lcParent) + gcQT)
                    lcRestr = THIS.BuildRestr(@aNewPrimary, lcParent, @aNewForeign, "inserted", "AND")
                    lcSQL = THIS.BuildComment(PREVENT_INSERTS_LOC, lcParent)
                    lcSQL = lcSQL + "IF @status<>'Failed'" + lcCRLF
                    lcSQL = lcSQL + "    BEGIN" + lcCRLF
                    lcSQL = lcSQL + "    IF(SELECT COUNT(*) FROM inserted) !=" + lcCRLF
                    lcSQL = lcSQL + "   (SELECT COUNT(*) FROM " + lcParent + ", inserted WHERE (" + lcRestr + "))" + lcCRLF
                    lcSQL = lcSQL + "        BEGIN" + lcCRLF
*** DH 2015-09-08: use syntax for SQL Server 2005+ for RAISERROR
*                    lcSQL = lcSQL + "                RAISERROR " + ERR_SVR_UPDREFVIO + " " + lcErrMsg + lcCRLF
                    lcSQL= lcSQL + "                RAISERROR (" + lcErrMsg + ", -1, -1)" + lcCRLF
                    lcSQL = lcSQL + "            SELECT @status='Failed'" + lcCRLF
                    lcSQL = lcSQL + "        END" + lcCRLF
                    lcSQL = lcSQL + "    END" + lcCRLF

                    *save this code
                    THIS.StoreRiCode(lcChild,"InsertRI",lcSQL)

                ENDIF
            ENDIF

            * CHILD UPDATE AND INSERT for Oracle
            IF THIS.ServerType = "Oracle" AND ;
                    (lcUpdateType = RESTRICT_CHAR_LOC OR lcInsertType = RESTRICT_CHAR_LOC)

                lcUpdTest = THIS.BuildRestr(@aNewForeign, ":old", @aNewPrimary, ":new", "OR")
                lcUpdTest = STRTRAN(lcUpdTest, "=", "!=")

                lcRestr = THIS.BuildRestr(@aNewPrimary, lcParent, @aNewForeign, ":new", "AND")
                lcErrMsg = "'" + STRTRAN(CANT_ORPHAN_LOC, "'|1'", gc2QT + lcParent + gc2QT) + "'"
                lcSQL = THIS.BuildComment(PREVENT_SELF_O_LOC, lcParent)
                lcSQL = lcSQL + "IF (UPDATING AND " + lcUpdTest + ") OR INSERTING THEN" + lcCRLF
                lcSQL = lcSQL + "    SELECT COUNT(*) INTO " + REC_COUNT_VAR + " FROM " + lcParent + " WHERE (" + lcRestr + ");" + lcCRLF
                lcSQL = lcSQL + "    IF " + REC_COUNT_VAR + " = 0 THEN" + lcCRLF
                lcSQL = lcSQL + "        raise_application_error(" + ERR_SVR_UPDREFVIO_ORA + ", " + lcErrMsg + ");" + lcCRLF
                lcSQL = lcSQL + "    END IF;" + lcCRLF
                lcSQL = lcSQL + "END IF;" + lcCRLF

                *save this code
                THIS.StoreRiCode(lcChild,"UpdateRI",lcSQL)
            ENDIF

            *If we're dealing with SQL Server, run sp_primarykey, sp_foreignkey
            IF THIS.ServerType <> "Oracle" AND ALEN(aNewPrimary,1) <= 8 AND !THIS.ExportDRI THEN
                *Check if the table is in multiple rels
                SELECT COUNT(*) FROM (lcEnumRelsTbl) WHERE RTRIM(DD_PARENT)==lcParent ;
                    AND !DD_CHIEXPR=="" AND !DD_PAREXPR=="" INTO ARRAY aDupeCount
                IF aDupeCount>1 THEN
                    *check to see if the local table has a primary key index
                    SELECT RmtExpr FROM (lcEnum_Indexes) WHERE RTRIM(IndexName)==lcParent ;
                        AND LclIdxType="Primary key" INTO ARRAY aIndexExpr
                    *If primary key index expression is same as RI primary key, run sp_primary key
                    IF RTRIM(aIndexExpr)==lcNewPrimary THEN
                        THIS.SetPKey(lcParent,lcNewPrimary)
                        THIS.SetFKey(lcChild,lcNewForeign,lcParent)
                    ELSE
                        THIS.SetCommonKey(lcParent,@aNewPrimary,lcChild,@aNewForeign)
                    ENDIF
                ELSE
                    THIS.SetPKey(lcParent,lcNewPrimary)
                    THIS.SetFKey(lcChild,lcNewForeign,lcParent)
                ENDIF
            ENDIF

        ENDSCAN
		raiseevent(This, 'CompleteProcess')
        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE SetCommonKey
        PARAMETERS lcTable1, aNewPrimary, lcTable2, aNewForeign
        LOCAL lcSQL
        lcSQL="sp_commonkey " + lcTable1 + ", " + lcTable2
        FOR I=1 TO ALEN(aNewPrimary,1)
            lcSQL=lcSQL + ", " + aNewPrimary[i] + ", " + aNewForeign[i]
        NEXT
        =THIS.ExecuteTempSPT(lcSQL)
    ENDPROC


    PROCEDURE SetPKey
        PARAMETERS lcTable, lcKey
        LOCAL lcSQL
        lcSQL="sp_primarykey " + lcTable + ", " + lcKey
        =THIS.ExecuteTempSPT(lcSQL)
    ENDPROC


    PROCEDURE SetFKey
        PARAMETERS lcChild, lcChildKey, lcParent
        LOCAL lcSQL
        lcSQL="sp_foreignkey " + lcChild +", " + lcParent + ", " + lcChildKey
        =THIS.ExecuteTempSPT(lcSQL)
    ENDPROC


    FUNCTION BuildComment
        PARAMETERS lcMainComment, lcToInsert
        LOCAL lcCRLF
        #DEFINE SEARCH_TOKEN "|1"

        lcCRLF = CHR(13)
        lcMainComment = STRTRAN(lcMainComment, SEARCH_TOKEN, lcToInsert)
        lcComment= lcCRLF + "/* " + lcMainComment + " */" + lcCRLF
        RETURN lcComment

    ENDFUNC


    PROCEDURE KeyArray
        PARAMETERS lcKeyString, aKeyArray

        *Takes comma separated list of fields in a key and converts it to an array

        lcKeyString=lcKeyString+ ","
        DO WHILE !lcKeyString==""
            THIS.InsaItem(@aKeyArray,ALLTRIM(LEFT(lcKeyString,AT(",",lcKeyString)-1)))
            lcKeyString=SUBSTR(lcKeyString,AT(",",lcKeyString)+1)
        ENDDO

    ENDPROC


    FUNCTION BuildRestr
        PARAMETERS aNewPrimary, lcParent, aNewForeign, lcChild, lcConjunction
        LOCAL lcBuildRestr

        *Creates a restriction string, e.g. "customer.cust.id=order.cust_id"

        lcBuildRestr=""
        FOR I=1 TO ALEN(aNewPrimary,1)
            IF I>1 THEN
                lcBuildRestr=lcBuildRestr + " " + lcConjunction + " "
            ENDIF
            lcBuildRestr=lcBuildRestr + lcParent + "." + aNewPrimary[i] + " = " + 	;
                lcChild + "." + aNewForeign[i]
        NEXT
        RETURN lcBuildRestr

    ENDFUNC


    FUNCTION BuildUpdateTest
        PARAMETERS aKeyFields
        LOCAL lcSQL, I
        lcSQL=""

        FOR I = 1 TO ALEN(aKeyFields,1)
            IF I > 1 THEN
                lcSQL= lcSQL + " OR "
            ENDIF
            lcSQL= lcSQL + "UPDATE(" + aKeyFields[i] + ")"
        NEXT
        RETURN lcSQL

    ENDFUNC


    PROCEDURE StoreRiCode
        PARAMETERS lcRmtTblName, lcFldName1, lcSQL, lcFldName2, llRetVal
        LOCAL lnOldArea

        lnOldArea=SELECT()
        lcEnumTablesTbl=THIS.EnumTablesTbl
        SELECT (lcEnumTablesTbl)
        LOCATE FOR RTRIM(RmtTblName)==RTRIM(lcRmtTblName)

        REPLACE &lcEnumTablesTbl..&lcFldName1 WITH lcSQL ADDITIVE

        IF PARAM()=4 THEN
            REPLACE &lcEnumTablesTbl..&lcFldName2 WITH llRetVal
        ENDIF

        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE CreateTriggers
        LOCAL lcTableName, lcCRLF, lcDeleteRI,lcInsertRI,lcUpdateRI, lcSproc, ;
            lcEnumFields, lnOldArea, lnTableCount, lcTrigName, lnError, lcErrMsg, ;
            lcUpdateType, lcDeleteType, lcInsertType
*** DH 2015-09-25: added locals for new code
		local lcTriggerTableName
        lnOldArea = SELECT()
        llRetVal = .F.
        lcCRLF = CHR(13)

        lcEnumTables = RTRIM(THIS.EnumTablesTbl)
        lcEnumRelsTbl = RTRIM(THIS.EnumRelsTbl)

        SELECT (lcEnumTables)
        lnTableCount = 0

        IF !THIS.ExportRelations AND !THIS.ExportDefaults AND !THIS.ExportValidation THEN
            RETURN
        ENDIF

* If we have an extension object and it has a CreateTriggers method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'CreateTriggers', 5) and ;
			not This.oExtension.CreateTriggers(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        *Thermometer stuff
        SELECT COUNT(*) FROM (lcEnumTables) WHERE &lcEnumTables..EXPORT=.T. ;
            INTO ARRAY aTableCount
        THIS.InitTherm(CREA_TRIGGERS_LOC,aTableCount,0)
        lnTableCount=0

        IF THIS.ServerType = "Oracle"
            #IF SUPPORT_ORACLE
                * Oracle: Create up to two row triggers
                * before update and insert: insert ri, restrict update ri, restrict delete ri (AT)
                * after update or delete: cascade update ri, cascade delete ri (AT)

                SCAN FOR EXPORT = .T. AND (!EMPTY(InsertRI) OR !EMPTY(UpdateRI) OR !EMPTY(DeleteRI))
                    lcTableName = RTRIM(&lcEnumTables..RmtTblName)

                    * Thermometer stuff
                    lcThermMsg = STRTRAN(THIS_TABLE_LOC,'|1',lcTableName)
                    THIS.UpDateTherm(lnTableCount,lcThermMsg)
                    lnTableCount = lnTableCount+1

                    * Grab RI and validation rule code (Note that all the validation rule
                    * SQL has already been placed in the InsertRI field)
                    lcInsertRI = &lcEnumTables..InsertRI
                    lcUpdateRI = &lcEnumTables..UpdateRI
                    lcDeleteRI = &lcEnumTables..DeleteRI

                    *if update code is restrict, need to
                    *toss in a variable declaration at the top of the trigger
                    *(it can't come inside of the BEGIN...END commands)
                    IF AT(REC_COUNT_VAR, lcDeleteRI) <> 0 OR AT(REC_COUNT_VAR, lcUpdateRI) <> 0 THEN
                        lcDecl= "DECLARE " + REC_COUNT_VAR + " NUMBER;" + lcCRLF
                    ELSE
                        lcDecl=""
                    ENDIF

                    *Assemble before trigger
                    IF !EMPTY(lcInsertRI) OR !EMPTY(lcUpdateRI) OR !EMPTY(lcDeleteRI)
                        lcTrigName = ORA_BIUD_TRIG_PREFIX + LEFT(lcTableName, MAX_NAME_LENGTH-LEN(ORA_BIUD_TRIG_PREFIX))
                        lcTrigName = THIS.UniqueOraName(lcTrigName)

                        lcSQL = "CREATE TRIGGER " + lcTrigName + lcCRLF
                        lcSQL = lcSQL + "BEFORE INSERT OR UPDATE OR DELETE"
                        lcSQL = lcSQL + " ON " + lcTableName + " FOR EACH ROW " + lcCRLF
                        lcSQL = lcSQL + lcDecl + "BEGIN " + lcCRLF
                        lcSQL = lcSQL + lcInsertRI + lcUpdateRI + lcDeleteRI + lcCRLF + "END;"

                        IF THIS.DoUpsize AND THIS.Perm_Trigger THEN
                            llRetVal = THIS.ExecuteTempSPT(lcSQL, @lnError, @lcErrMsg)
                            IF !llRetVal THEN
                                THIS.StoreError(lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC,lcTableName,TRIGGER_LOC)
                                REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                    RIErrNo WITH lnError
                            ENDIF

                        ENDIF
                        REPLACE &lcEnumTables..InsertRI WITH lcSQL, ;
                            &lcEnumTables..ItrigName WITH lcTrigName, ;
                            &lcEnumTables..InsertX WITH llRetVal

                        * Trigger sql is appended in the script, so clean it up
                        REPLACE &lcEnumTables..UpdateRI WITH "", ;
                            &lcEnumTables..DeleteRI WITH ""
                    ENDIF

                    *Assemble after trigger
                    IF 	.F.
                        lcTrigName = ORA_AUD_TRIG_PREFIX + LEFT(lcTableName, MAX_NAME_LENGTH-LEN(ORA_AUD_TRIG_PREFIX))
                        lcTrigName = THIS.UniqueOraName(lcTrigName)
                        lcSQL =         "CREATE TRIGGER " + lcTrigName + lcCRLF
                        lcSQL = lcSQL + "AFTER UPDATE OR DELETE"
                        lcSQL = lcSQL + " ON " + lcTableName + " FOR EACH ROW " + lcCRLF
                        lcSQL  =lcSQL + lcDecl + "BEGIN " + lcCRLF
                        lcSQL = lcSQL + lcUpdateRI + lcDeleteRI + lcCRLF + "END;"

                        IF THIS.DoUpsize AND THIS.Perm_Trigger THEN
                            llRetVal = THIS.ExecuteTempSPT(lcSQL, @lnError, @lcErrMsg)
                            IF !llRetVal THEN
                                THIS.StoreError(lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC,lcTableName,TRIGGER_LOC)
                                REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                    RIErrNo WITH lnError
                            ENDIF
                        ENDIF
                        REPLACE &lcEnumTables..DeleteRI WITH lcSQL, ;
                            &lcEnumTables..DtrigName WITH lcTrigName, ;
                            &lcEnumTables..DeleteX WITH llRetVal, ;
                            &lcEnumTables..UpdateRI WITH ""
                    ENDIF

                ENDSCAN

            #ENDIF

        ELSE

            *SQL Server: Create up to three triggers
            *update trigger: updateRI, rules/sprocs
            *insert trigger: insertRI, rules/sprocs
            *delete trigger: deleteRI

            SCAN FOR &lcEnumTables..EXPORT=.T.
                lcSproc=""
                lcSQL=""
                lcTableName=RTRIM(&lcEnumTables..RmtTblName)

                lcThermMsg=STRTRAN(THIS_TABLE_LOC,'|1',lcTableName)
                THIS.UpDateTherm(lnTableCount,lcThermMsg)
                lnTableCount=lnTableCount+1

                *Build sproc string which will be used in Insert and Update triggers

                *Grab table validation rules (i.e. sprocs) that were successfully created
                *Grab them regardless if the user is just generating a script
                IF &lcEnumTables..RuleExport =.T. ;
                        OR (!THIS.DoUpsize AND !EMPTY(&lcEnumTables..RmtRule)) THEN
                    lcSproc=TBLRULE_COMMENT_LOC
                    lcSproc=lcSproc+ "execute "+RTRIM(&lcEnumTables..RRuleName)  ;
                        + " @status output" + lcCRLF
                ENDIF

                *Grab RI code
                lcInsertRI=&lcEnumTables..InsertRI
                lcUpdateRI=&lcEnumTables..UpdateRI
                lcDeleteRI=&lcEnumTables..DeleteRI

                *Grab field validation sprocs
                lcEnumFields=RTRIM(THIS.EnumFieldsTbl)
                SELECT (lcEnumFields)
                SCAN FOR RTRIM(&lcEnumFields..TblName)==lcTableName
                    lcFieldName=RTRIM(&lcEnumFields..RmtFldname)

                    *Only add it to the string if the sproc was successfully created

                    IF &lcEnumFields..RuleExport =.T. ;
                            OR (!THIS.DoUpsize AND !EMPTY(&lcEnumFields..RmtRule)) THEN
                        lcSproc=lcSproc + STRTRAN(FLDRULE_COMMENT_LOC,"|1",lcFieldName)
                        lcSproc=lcSproc + "execute " + RTRIM(&lcEnumFields..RRuleName) + " @status output" + lcCRLF
                    ENDIF

                ENDSCAN

                SELECT (lcEnumTables)

                *Strings used in all the triggers:
                lcStatus="DECLARE @status char(10)  " + STATUS_COMMENT_LOC + lcCRLF
                lcStatus=lcStatus + "SELECT @status='Succeeded'" + lcCRLF
                lcRollBack=ROLLBACK_LOC + lcCRLF + "IF @status='Failed'" + lcCRLF +;
                    "ROLLBACK TRANSACTION" + lcCRLF
*** DH 2015-09-25: strip brackets from table name
				lcTriggerTableName = chrtran(lcTableName, '[]', '')

                IF !EMPTY(lcInsertRI) OR !EMPTY(lcSproc) THEN

                    *Create insert trigger
*** DH 2015-09-25: use stripped table name
***                    lcTrigName=ITRIG_PREFIX + LEFT(lcTableName,MAX_NAME_LENGTH-LEN(ITRIG_PREFIX))
                    lcTrigName=ITRIG_PREFIX + LEFT(lcTriggerTableName,MAX_NAME_LENGTH-LEN(ITRIG_PREFIX))
                    lcSQL="CREATE TRIGGER " + lcTrigName
                    lcSQL=lcSQL + " ON " + lcTableName + " FOR INSERT AS " + lcCRLF
                    lcSQL=lcSQL + lcStatus + lcSproc + lcInsertRI
                    lcSQL=lcSQL + lcRollBack
                    IF THIS.DoUpsize THEN
                        llRetVal=THIS.ExecuteTempSPT(lcSQL, @lnError,@lcErrMsg)
                        IF !llRetVal THEN
                            THIS.StoreError(lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC,lcTableName,TRIGGER_LOC)
                            REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                RIErrNo WITH lnError
                        ENDIF
                    ENDIF
                    REPLACE &lcEnumTables..InsertRI WITH lcSQL, ;
                        &lcEnumTables..ItrigName WITH lcTrigName, ;
                        &lcEnumTables..InsertX WITH llRetVal
                ENDIF

                IF !EMPTY(lcUpdateRI) OR !EMPTY(lcSproc) THEN

                    *Create update trigger
*** DH 2015-09-25: use stripped table name
***                    lcTrigName=UTRIG_PREFIX + LEFT(lcTableName,MAX_NAME_LENGTH-LEN(UTRIG_PREFIX))
                    lcTrigName=UTRIG_PREFIX + LEFT(lcTriggerTableName,MAX_NAME_LENGTH-LEN(UTRIG_PREFIX))
                    lcSQL="CREATE TRIGGER " + lcTrigName
                    lcSQL=lcSQL + " ON " + lcTableName + " FOR UPDATE AS " + lcCRLF
                    lcSQL=lcSQL + lcStatus + lcSproc + lcUpdateRI
                    lcSQL=lcSQL + lcRollBack
                    IF THIS.DoUpsize THEN
                        llRetVal=THIS.ExecuteTempSPT(lcSQL, @lnError,@lcErrMsg)
                        IF !llRetVal THEN
                            THIS.StoreError(lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC,lcTableName,TRIGGER_LOC)
                            REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                RIErrNo WITH lnError
                        ENDIF
                    ENDIF
                    REPLACE &lcEnumTables..UpdateRI WITH lcSQL, ;
                        &lcEnumTables..UtrigName WITH lcTrigName, ;
                        &lcEnumTables..UpdateX WITH llRetVal
                ENDIF

                IF !EMPTY(lcDeleteRI) THEN

                    *Create delete trigger
*** DH 2015-09-25: use stripped table name
***                    lcTrigName=DTRIG_PREFIX + LEFT(lcTableName,MAX_NAME_LENGTH-LEN(DTRIG_PREFIX))
                    lcTrigName=DTRIG_PREFIX + LEFT(lcTriggerTableName,MAX_NAME_LENGTH-LEN(DTRIG_PREFIX))
                    lcSQL="CREATE TRIGGER " + lcTrigName
                    lcSQL=lcSQL + " ON " + lcTableName + " FOR DELETE AS " + lcCRLF
                    lcSQL=lcSQL + lcStatus + lcDeleteRI
                    lcSQL=lcSQL + lcRollBack
                    IF THIS.DoUpsize THEN
                        llRetVal=THIS.ExecuteTempSPT(lcSQL, @lnError,@lcErrMsg)
                        IF !llRetVal THEN
                            THIS.StoreError(lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC,lcTableName,TRIGGER_LOC)
                            REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                RIErrNo WITH lnError
                        ENDIF
                    ENDIF
                    REPLACE &lcEnumTables..DeleteRI WITH lcSQL, ;
                        &lcEnumTables..DtrigName WITH lcTrigName, ;
                        &lcEnumTables..DeleteX WITH llRetVal
                ENDIF

            ENDSCAN

        ENDIF
		raiseevent(This, 'CompleteProcess')
        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE GetRiInfo
        LOCAL lcEnumRelsTbl,lnOldArea, p_dbc, l_rmtchitable, l_rmtpartable, lcTableName
        PRIVATE aDupeCount

        *Thanks, George Goley, for the heart of this code
        IF !THIS.GetRiInfoRecalc THEN
            RETURN
        ENDIF

        #DEFINE KEY_CHILDTAG		13	&& For RELATION objects: name of child (from) index tag
        #DEFINE KEY_RELTABLE		18	&& For RELATION objects: name of related table
        #DEFINE KEY_RELTAG			19	&& For RELATION objects: name of related index tag
        #DEFINE d_updatespot		1
        #DEFINE d_deletespot		2
        #DEFINE d_insertspot		3

        *Make sure that all tables selected for upsizing have had their
        *indexes analyzed

        lnOldArea=SELECT()
        *Make sure index info is up-to-date
        IF THIS.AnalyzeIndexesRecalc THEN
            THIS.AnalyzeIndexes
        ENDIF

        p_dbc=RTRIM(THIS.SourceDB)

        lcEnumIndexesTbl=THIS.EnumIndexesTbl
        SELECT (lcEnumIndexesTbl)
        lcEnumRelsTbl=THIS.CreateWzTable("Relation")
        THIS.EnumRelsTbl=lcEnumRelsTbl

        USE (p_dbc) IN 0 AGAIN ALIAS mydbc
        USE (p_dbc) IN 0 AGAIN ALIAS mydbcpar
        SELECT mydbc
        lcExact=SET("EXACT")
        SET EXACT ON
        LOCATE FOR LOWER(objecttype)="table" AND TRIM(LOWER(objectname))=="ridd"

        SCAN FOR UPPER(objecttype)="RELATION"

            GOTO (mydbc.parentid) IN mydbcpar
            l_chitable=LOWER(mydbcpar.objectname)
            l_start=1
            DO WHILE l_start<=LEN(property)
                l_size=ASC(SUBSTR(property,l_start,1))+;
                    (ASC(SUBSTR(property,l_start+1,1))*256)+;
                    (ASC(SUBSTR(property,l_start+2,1))*256^2)+;
                    (ASC(SUBSTR(property,l_start+3,1))*256^3)
                l_key=SUBSTR(property,l_start+6,1)
                l_value=SUBSTR(property,l_start+7,l_size-8)
                DO CASE
                    CASE l_key==CHR(KEY_CHILDTAG)
                        l_chitag=l_value
                    CASE l_key==CHR(KEY_RELTABLE)
                        l_partable=LOWER(l_value)
                    CASE l_key==CHR(KEY_RELTAG)
                        l_partag=l_value
                ENDCASE
                l_start=l_start+l_size
            ENDDO

            l_area=SELECT(1)

            *Grab tag expression
            SELECT (lcEnumIndexesTbl)

            *Parent table
            LOCATE FOR IndexName=RTRIM(l_partable) AND TagName=RTRIM(l_partag)
            l_parexpr=&lcEnumIndexesTbl..RmtExpr

            *Child table
            LOCATE FOR IndexName=RTRIM(l_chitable) AND TagName=RTRIM(l_chitag)
            l_chiexpr=&lcEnumIndexesTbl..RmtExpr

            *Translate ri characters (RCI) into words (Restrict, Cascade, Ignore)
            l_delete=UPPER(SUBSTR(mydbc.riinfo,d_deletespot,1))
            l_delete=IIF(l_delete<>CASCADE_CHAR_LOC AND l_delete<>RESTRICT_CHAR_LOC,;
                IGNORE_CHAR_LOC,l_delete)
            l_update=UPPER(SUBSTR(mydbc.riinfo,d_updatespot,1))
            l_update=IIF(l_update<>CASCADE_CHAR_LOC AND l_update<>RESTRICT_CHAR_LOC,;
                IGNORE_CHAR_LOC,l_update)
            l_insert=UPPER(SUBSTR(mydbc.riinfo,d_insertspot,1))
            l_insert=IIF(l_insert<>CASCADE_CHAR_LOC AND l_insert<>RESTRICT_CHAR_LOC,;
                IGNORE_CHAR_LOC,l_insert)

            l_rmtpartable=THIS.RemotizeName(l_partable)
            l_rmtchitable=THIS.RemotizeName(l_chitable)

            *See if there are multiple relations between the same two tables
            SELECT COUNT(*) FROM (lcEnumRelsTbl) ;
                WHERE RTRIM(Dd_RmtPar)==RTRIM(l_rmtpartable) ;
                AND RTRIM(Dd_RmtChi)==RTRIM(l_rmtchitable) ;
                INTO ARRAY aDupeCount

            INSERT INTO &lcEnumRelsTbl (DD_CHILD,Dd_RmtChi,DD_PARENT,Dd_RmtPar,;
                DD_CHIEXPR,DD_PAREXPR,Duplicates, dd_update, dd_insert, dd_delete) ;
                VALUES (l_chitable,l_rmtchitable, l_partable, l_rmtpartable, ;
                l_chiexpr,l_parexpr, aDupeCount, l_update, l_insert, l_delete)

            SELECT mydbc

        ENDSCAN

        *clean up
        USE
        SELECT mydbcpar
        USE
        SET EXACT &lcExact
        THIS.GetRiInfoRecalc=.F.

        SELECT(lnOldArea)

    ENDPROC


    #IF SUPPORT_ORACLE
    PROCEDURE GetEligibleRels
        PARAMETERS aEligibleRels, llChoices
        LOCAL lcEnumRelsTbl, lcEnumTablesTbl, lnOldArea, lcDupeString, lcConstraint, ;
            aExportTables, lcExact

        lnOldArea=SELECT()
        lcEnumRelsTbl=RTRIM(THIS.EnumRelsTbl)
        lcEnumTablesTbl=RTRIM(THIS.EnumTablesTbl)
        *This determines if we return relations that have cluster names or that don't
        *If we want the choices of rels for clusters, we want the rels w/o cluster names
        IF llChoices THEN
            lcConstraint="Export =.T. AND EMPTY(ClustName)"
        ELSE
            lcConstraint="Export =.T. AND !EMPTY(ClustName)"
        ENDIF

        DIMENSION aExportTables[1]
        SELECT LOWER(TblName) FROM (lcEnumTablesTbl) WHERE &lcConstraint INTO ARRAY aExportTables

        IF !EMPTY(aExportTables) THEN
            SELECT (lcEnumRelsTbl)
            I=1
            lcExact=SET('EXACT')
            SET EXACT ON
            SCAN
                *check to see if both parent and child table are being exported
                *also make sure it's not a self-join
                IF ASCAN(aExportTables,RTRIM(DD_CHILD))<>0 ;
                        AND ASCAN(aExportTables,RTRIM(DD_PARENT))<>0 ;
                        AND !DD_PARENT==DD_CHILD THEN
                    *Make sure number of fields in primary and foreign keys is the same
                    DIMENSION aParentKeys[1], aChildKeys[1]
                    THIS.KeyArray(DD_PAREXPR,@aParentKeys)
                    THIS.KeyArray(DD_CHIEXPR,@aChildKeys)

                    IF ALEN(aParentKeys,1)=ALEN(aChildKeys,1) THEN
                        DIMENSION aEligibleRels [i,2]
                        IF &lcEnumRelsTbl..Duplicates <> 0 THEN
                            lcDupeString = "(" + LTRIM(STR(Duplicates+1)) + ")"
                        ELSE
                            lcDupeString=""
                        ENDIF
                        aEligibleRels[i,1]=RTRIM(DD_PARENT) + ":" + RTRIM(DD_CHILD) + lcDupeString
                        aEligibleRels[i,2]=aEligibleRels[i,1]
                        I=I+1
                    ENDIF
                ENDIF
            ENDSCAN
            SET EXACT &lcExact
        ENDIF
        SELECT (lnOldArea)

    ENDPROC
#ENDIF



#IF SUPPORT_ORACLE
    PROCEDURE ParseRel
        PARAMETERS lcRelName, lcParent, lcChild, lnDupID
        LOCAL lnOPos, lnCPos, lnColPos

        *
        *Takes a string of the form "customer:orders" and parses it into table names
        *Handles case where two tables may have multiple relations between each other
        *

        *Check if the rel is one of several with duplicate parent and child tables
        lnOPos=AT("(",lcRelName)
        lnCPos=AT(")",lcRelName)
        IF lnOPos<>0 THEN
            lnDupID=VAL(SUBSTR(lcRelName,lnOPos+1,lnCPos-lnOPos-1))
            lcRelName=LEFT(lcRelName,lnOPos-1)
        ELSE
            lnDupID=0
        ENDIF

        *Separate parent and child tablenames out of relation name
        lnColPos=AT(":",lcRelName)
        lcParent=LEFT(lcRelName,lnColPos-1)
        lcChild=SUBSTR(lcRelName,lnColPos+1)

    ENDPROC
#ENDIF


    PROCEDURE RedirectApp
        LOCAL lnDispLogin, lcOldConnString, lcRenamedUserConn, llRenamed

* If we have an extension object and it has a RedirectApp method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'RedirectApp', 5) and ;
			not This.oExtension.RedirectApp(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        *All the renaming gyrations are to avoid the problem of the user
        *getting a login dialog when remote views are created

		IF !EMPTY(This.UserConnection) AND !This.PwdInDef
            *create new conndef; name is placed into this.ViewConnection
            THIS.CreateConnDef()
            DBSETPROP(THIS.ViewConnection,"connection","ConnectString",This.CONNECTSTRING)
            lcRenamedUserConn=THIS.UniqueConnDefName()
            *rename old connection
            RENAME CONNECTION (This.UserConnection) TO (lcRenamedUserConn)
            DBSETPROP(lcRenamedUserConn,"connection","ConnectString",This.CONNECTSTRING)
            *give new connection the original's name
            RENAME CONNECTION (This.ViewConnection) TO (This.UserConnection)
            This.ViewConnection=This.UserConnection
            llRenamed=.T.
        ELSE
            llRenamed=.F.
        ENDIF

        IF THIS.ExportViewToRmt AND THIS.DoUpsize THEN
            THIS.RemotizeViews()
        ENDIF

        IF THIS.ExportTableToView AND THIS.DoUpsize THEN
            THIS.CreateRmtViews()
        ENDIF

        *need to take out password if conndef created and user doesn't want pwd saved
        IF !THIS.ViewConnection=="" AND !THIS.ExportSavePwd AND !llRenamed THEN
            lcRConnString=DBGETPROP(THIS.ViewConnection,"connection","connectstring")
            lcRPwd=THIS.ParseConnectString(lcRConnString,"pwd=")
            lcRConnString=STRTRAN(lcRConnString,"PWD="+lcRPwd+";")
            lcRConnString=STRTRAN(lcRConnString,"PWD="+lcRPwd)
            lcRConnString=STRTRAN(lcRConnString,"pwd="+lcRPwd+";")
            lcRConnString=STRTRAN(lcRConnString,"pwd="+lcRPwd)
            =DBSETPROP(THIS.ViewConnection,"connection","connectstring",lcRConnString)
        ENDIF

        IF llRenamed
            *Delete temporary connection definition
            DELETE CONNECTION (This.ViewConnection)
            RENAME CONNECTION (lcRenamedUserConn) TO (This.UserConnection)
        ENDIF

    ENDPROC

    PROCEDURE TrimDBCNameFromTables(lcTables)
        * trim off "tastrade!" from "tastrade!customer,tastrade!orders"
        LOCAL P,q,m.RetVal
        m.RetVal = ""
        DO WHILE "!" $ m.lcTables
            m.P = AT("!",m.lcTables)
            m.q = AT(",",m.lcTables)
            IF m.q > 0
                m.RetVal = m.RetVal + SUBSTRC(m.lcTables,m.P+1,m.q - m.P)
                m.lcTables = SUBSTRC(m.lcTables,m.q+1)
            ELSE
                m.RetVal = m.RetVal + SUBSTRC(m.lcTables,m.P+1)
                m.lcTables = ""
            ENDIF
        ENDDO
        RETURN m.RetVal
    ENDPROC

    PROCEDURE ProcessFromClause
        * parses the FROM clause, adds the odbc oj escape to the Sql statement and
        * returns the list of tables participating in this view
        PARAMETERS cStr, cTables
        LOCAL lcTalk, lcPath, llResult, aTables[1], lcLen

        * disable errors
        THIS.SetErrorOff = .T.
        THIS.HadError = .F.

        * needed by substr() to return empty strings at eos
        lcTalk = SET('talk')
        SET TALK OFF
        lcProc = SET('proc')
        SET PROCEDURE TO "wizjoin.prg"
        llResult = ParseFromClause(@cStr, @aTables)
        SET PROCEDURE TO &lcProc
        SET TALK &lcTalk

        * bail on error
        THIS.SetErrorOff = .F.
		IF This.HadError
            RETURN .F.
        ENDIF

        * build table list
        IF llResult
            cTables = ""
            lcLen = ALEN(aTables, 1)
            FOR m.I = 1 TO lcLen
                cTables = cTables + aTables[m.i] + IIF(m.I < lcLen, ", ", "")
            ENDFOR
        ENDIF

        RETURN llResult
    ENDPROC

    PROCEDURE RemotizeViews
        LOCAL aTableNames, lcEnumTables, lcTablesInView, lcTablesNotUpsized, ;
            lcNewTableString, lcViewSQL, lcRmtViewSQL, lnOldArea, lcViewsTbl, ;
            llTableUpsized, aWhip, lcEnumFieldsTbl, aFldArray, lcNewViewName, lcDBCAlias, ;
            aFieldsInView, mm, ii, jj, I, llSendUpdates, lcDelStatus, lcViewErr, lcCRLF, ;
            lcErrString, lnParentRec, lnViewCount, lnServerError, lcErrMsg, lcRCursor, ;
            lcRCursor,lcViewParms,llShareConnection,lcDatatype, lcRmtViewSQL2

        PRIVATE aViews
        *
        *Views that consist of tables which were upsized are modified so they point to
        *remote data and are executed on the back end
        *

        *If the user isn't upsizing, bail
        IF !THIS.DoUpsize THEN
            RETURN
        ENDIF

        lnOldArea=SELECT()
        lcCRLF=CHR(10)+CHR(13)
        *Check each view; see if all its tables were upsized
        *Get array of views; if no views, bail
        lnViewCount=ADBOBJECTS(aViews,"View")
        IF lnViewCount=0 THEN
            RETURN
        ENDIF
        FOR I =1 TO ALEN(aViews,1)
            aViews[i] = LOWER(aViews[i])
        ENDFOR
        THIS.InitTherm(RMTZING_VIEW_LOC,lnViewCount,0)
        lnViewCount=0

        *Create array of tables that were successfully exported, local names and remote names)
        *Sort them so the longest fields are first, otherwise string replacements could
        *get messed up
        lcEnumTables=THIS.EnumTablesTbl
        DIMENSION aTableNames[1]
        SELECT LOWER(TblName), ;
            LOWER(RmtTblName), ;
            ISNULL(TblName), ;
            1/LEN(ALLTRIM(TblName)) AS foo ;
            FROM &lcEnumTables WHERE Exported=.T. ;
            ORDER BY foo ;
            INTO ARRAY aTableNames
        IF EMPTY(aTableNames)
			raiseevent(This, 'CompleteProcess')
            RETURN
        ENDIF

        lcViewsTbl=THIS.CreateWzTable("Views")
        THIS.ViewsTbl=lcViewsTbl
        llViewConnCreated=.F.

        llShareConnection = .F.
        oReg = NEWOBJECT("FoxReg", home() + "ffc\registry.vcx")
        cOptionValue = ""
        cOptionName = "CrsShareConnection"
        m.nErrNum = oReg.GetFoxOption(m.cOptionName,@cOptionValue)
        IF nErrNum=0 AND cOptionValue="1"
            llShareConnection = .T.
        ENDIF

        FOR I=1 TO ALEN(aViews,1)

            *Thermometer stuff
            lcMessage=STRTRAN(THIS_VIEW_LOC,"|1",LOWER(RTRIM(aViews[i])))
            THIS.UpDateTherm(lnViewCount,lcMessage)
            lnViewCount=lnViewCount+1

            *If this is a remote view, skip it
            IF DBGETPROP(aViews[i],"View","SourceType")=2 THEN
                LOOP
            ENDIF

			*{Add JEI RKR 2005.03.25
			* Check for current View is Choosen for Upsizing
			IF !This.CheckUpsizeView(aViews[i])
				LOOP
			ENDIF
			*}Add JEI RKR 2005.03.25
			
            *flag set true if at least one table in a view was upsized;
            *we can ignore the view if this stays .F.
            llTableUpsized=.F.

            *Grab the view's SQL string (may need to change table names to remote versions)
            lcViewSQL=LOWER(DBGETPROP(aViews[i],"View","SQL"))

			&&chandsr added comments
			&&SANITIZE THE SQL STATEMENT BY REPLACING .F. and .T. WITH 0 and 1 RESPECTIVELY.
			&& Go through the FOX statements and make the statement ANSI92 compatible. This means replacing the occurrences of boolean expression with
			&& ANSI92 compatible boolean expressions.
            lcViewSQL = THIS.MakeStringANSI92Compatible (lcViewSQL)
			&&chandsr added comments end.

            * Get parameter list
            lcViewParms=LOWER(DBGETPROP(aViews[i],"View","Parameterlist"))

            * try parsing the from clause to get the tables and remotize oj
            * on error get table list from Tables property
            IF !THIS.ProcessFromClause(@lcViewSQL, @lcTablesInView)
                lcTablesInView = LOWER(DBGETPROP(aViews[i],"View","Tables"))
            ENDIF

            lcTablesInView = lcTablesInView + ","
            * add comma so there's no confusing strings like "dupe" and "dupe1"

            *will not remove free tables from the view list
            lcDBName=LOWER(THIS.JustStem2(THIS.SourceDB))
            lcTablesNotUpsized = STRTRAN(m.lcTablesInView,lcDBName+"!")

            *Remove local database name from sql string
            *If it has a space in it, it will be enclosed in quotes
            IF AT(lcDBName," ")<>0 THEN
                lcRmtViewSQL=STRTRAN(lcViewSQL,gcQT+lcDBName+"!"+gcQT)
            ELSE
                lcRmtViewSQL=STRTRAN(lcViewSQL,lcDBName+"!")
            ENDIF

            *Check which (if any) tables were upsized
            lcNewTableString=""

            DIMENSION aFieldsInView[1,2]
            aFieldsInView=.F.

            FOR ii=1 TO ALEN(aTableNames,1)
                IF LOWER(RTRIM(aTableNames[ii,1])) + "," $ lcTablesInView THEN
                    *set flag
                    llTableUpsized=.T.

                    *replace all field names changed by remotizing names
                    lcEnumFieldsTbl=THIS.EnumFieldsTbl
                    DIMENSION aFldArray[1,2]
                    aFldArray=.F.
                    SELECT FldName, RmtFldname FROM (lcEnumFieldsTbl) ;
                        WHERE TblName=RTRIM(aTableNames[ii,1]) ;
                        AND FldName<>RmtFldname INTO ARRAY aFldArray

                    IF !EMPTY(aFldArray) THEN
                        FOR jj=1 TO ALEN(aFldArray,1)
                            IF RTRIM(aFldArray[jj,1]) $ lcRmtViewSQL THEN
                                lcRmtViewSQL=STRTRAN(lcRmtViewSQL,RTRIM(aFldArray[jj,1]),RTRIM(aFldArray[jj,2]))
                            ENDIF
                        NEXT jj
                    ENDIF

                    *Replace table names with remotized table names
                    *First look for stuff of the form "from me" (change to "from_me")
                    *then look for "from" (change to "from_").  This prevent "from me"
                    *becoming "from__me".

                    lcTablesNotUpsized=STRTRAN(lcTablesNotUpsized,gcQT+RTRIM(aTableNames[ii,1])+","+gcQT)
                    lcTablesNotUpsized=STRTRAN(lcTablesNotUpsized,RTRIM(aTableNames[ii,1])+",")
                    lcRmtViewSQL=STRTRAN(lcRmtViewSQL,gcQT+RTRIM(aTableNames[ii,1])+gcQT,;
                        RTRIM(aTableNames[ii,2]))
                    lcRmtViewSQL=STRTRAN(lcRmtViewSQL," " + RTRIM(aTableNames[ii,1]),;
                        " " + RTRIM(aTableNames[ii,2]))

                    IF !lcNewTableString=="" THEN
                        lcNewTableString=lcNewTableString+","
                    ENDIF
                    lcNewTableString=lcNewTableString + RTRIM(aTableNames[ii,2])

                    aTableNames[ii,3]=.T.	&&Table is part of view

                ENDIF
            NEXT ii

            *Go on to the next view if this one has no tables upsized in it
            IF !llTableUpsized THEN
                LOOP
            ENDIF

            *If all the tables were upsized, then remotize the view
            IF ALLTRIM(STRTRAN(lcTablesNotUpsized,","))=="" THEN

                *Need a connection to associate with remote view
                *create one if user didn't connect with one
                *or if they created a new database
                IF THIS.ViewConnection=="" THEN
                    IF THIS.UserConnection=="" OR THIS.CreateNewDB THEN
                        THIS.CreateConnDef()
                    ELSE
                        THIS.ViewConnection=THIS.UserConnection
                    ENDIF
                ENDIF

                *Rename local view
                lcNewViewName=THIS.UniqueTorVName(aViews[i])
                RENAME VIEW (RTRIM(aViews[i])) TO (lcNewViewName)

                * Check for "==" used in SQL Server which is not allowed and replace with "="
                lcRmtViewSQL2 = lcRmtViewSQL
                IF THIS.SQLServer AND ATCC("==",lcRmtViewSQL)#0
                    lcRmtViewSQL2 = STRTRAN(lcRmtViewSQL,"==","=")
                ENDIF

                *Create remote view with same name
                THIS.SetErrorOff=.T.
                THIS.HadError=.F.
                CREATE SQL VIEW RTRIM(aViews[i]) REMOTE CONNECT RTRIM(THIS.ViewConnection) ;
                    AS &lcRmtViewSQL2
                THIS.SetErrorOff=.F.
                IF THIS.HadError THEN
                    *If the view creation fails, store the error and loop
                    =AERROR(aErrArray)
                    lnServerError=aErrArray[1]
                    lcErrMsg=aErrArray[2]
                    IF lnServerError=1526 THEN
                        lnServerError=aErrArray[5]
                    ENDIF

                    *Put things back and store error info
                    RENAME VIEW (lcNewViewName) TO (RTRIM(aViews[i]))
                    lcViewErr=aErrArray[2]
                    SELECT (lcViewsTbl)
                    APPEND BLANK
                    REPLACE ViewName WITH aViews[i], NewName WITH aViews[i], ;
                        ViewSQL WITH lcViewSQL, RmtViewSQL WITH lcRmtViewSQL, ;
                        Remotized WITH .F., ViewErr WITH lcErrMsg, ;
                        ViewErrNo WITH lnServerError
                    THIS.StoreError(lnServerError,lcErrMsg,lcRmtViewSQL,RMTIZE_VIEW_FAILED_LOC,aViews[i],VIEW_LOC)
                    LOOP I
                ENDIF

                *Set various and sundry view properties
                llSendUpdates=DBGETPROP(lcNewViewName,"view","SendUpdates")
                lcViewErr=""
                IF !DBSETPROP(RTRIM(aViews[i]),"view","SendUpdates",llSendUpdates)
                    THIS.AddToError(@lcViewErr,UPDATE_PROP_FAILED_LOC)
                ENDIF
                IF !DBSETPROP(RTRIM(aViews[i]),"view","Tables",lcNewTableString)
                    THIS.AddToError(@lcViewErr,TABLES_PROP_FAILED_LOC)
                ENDIF
                IF !EMPTY(lcViewParms) AND !DBSETPROP(RTRIM(aViews[i]),"view","Parameterlist",lcViewParms)
                    THIS.AddToError(@lcViewErr,TABLES_PROP_FAILED_LOC)
                ENDIF

                IF llShareConnection AND !DBSETPROP(RTRIM(aViews[i]),"view","ShareConnection",llShareConnection)
                    THIS.AddToError(@lcViewErr,TABLES_PROP_FAILED_LOC)
                ENDIF

                *Get properties of fields in local view; set remote field properties to same

                *Get cursors of fields in local view and for remote view
                lcDBCAlias=THIS.UniqueCursorName("dbcalias")
                lcDelStatus=SET("deleted")
                SET DELETED ON
                USE RTRIM(THIS.SourceDB) IN 0 AGAIN ALIAS &lcDBCAlias
                SELECT (lcDBCAlias)

                LOCATE FOR RTRIM(LOWER(objectname))==RTRIM(LOWER(lcNewViewName))
                lcParent=&lcDBCAlias..ObjectID
                lcLCursor=THIS.UniqueCursorName("localfields")
                SELECT 0
                SELECT objectname, RECNO() FROM (lcDBCAlias) WHERE parentid=lcParent AND objecttype="Field" ;
                    INTO CURSOR &lcLCursor
                SELECT(lcLCursor)
                lnLFields=RECCOUNT()

                *Find the record for the remote view
                SELECT (lcDBCAlias)
                LOCATE FOR LOWER(RTRIM(objectname))==LOWER(RTRIM(aViews[i]))
                lcParent=&lcDBCAlias..ObjectID
                lcRCursor=THIS.UniqueCursorName("remotefields")
                SELECT 0
                *Only get fields which aren't timestamps added by the wizard
                SELECT objectname, RECNO() FROM (lcDBCAlias) ;
                    WHERE parentid=lcParent ;
                    AND objecttype="Field" ;
                    AND ATC(IDENTCOL_LOC, RTRIM(objectname)) == 0 ;
                    AND ATC(TIMESTAMP_LOC, RTRIM(objectname)) == 0 ;
                    INTO CURSOR &lcRCursor
                SELECT(lcRCursor)
                lnRFields=RECCOUNT()

                IF lnRFields<>lnLFields
                    THIS.AddToErr(@lcViewErr,FIELDS_UNEQUAL_LOC)
                ELSE
                    SELECT (lcLCursor)
                    SCAN
                        llKeyField=DBGETPROP(lcNewViewName+"."+RTRIM(objectname),"field","keyfield")
                        llUpdatable=DBGETPROP(lcNewViewName+"."+RTRIM(objectname),"field","updatable")
                        lcUpdateName=DBGETPROP(lcNewViewName+"."+RTRIM(objectname),"field","updatename")
                        lcDatatype=DBGETPROP(lcNewViewName+"."+RTRIM(objectname),"field","datatype")

                        SELECT (lcRCursor)
                        IF !DBSETPROP(RTRIM(aViews[i])+"."+ RTRIM(objectname),"field","keyfield",llKeyField)
                            lcRmtViewField=RTRIM(objectname)
                            lcErrString=STRTRAN(KEYFIELD_PROP_FAILED_LOC,'|1',lcRmtViewField)
                            THIS.AddToErr(@lcViewErr,lcErrString)
                        ENDIF
                        IF !DBSETPROP(RTRIM(aViews[i])+"."+ RTRIM(objectname),"field","updatable",llUpdatable)
                            lcRmtViewField=RTRIM(objectname)
                            lcErrString=STRTRAN(UPDATABLE_PROP_FAILED_LOC,'|1',lcRmtViewField)
                            THIS.AddToErr(@lcViewErr,lcErrString)
                        ENDIF

                        *Since we map Date->DateTime field in SQL Server, we need to make sure the new
                        *remote view resets DateTime->Date.
                        IF lcDatatype="D" AND !DBSETPROP(RTRIM(aViews[i])+"."+ RTRIM(objectname),"field","datatype",lcDatatype)
                            THIS.AddToErr(@lcViewErr,lcErrString)
                        ENDIF
                        SKIP
                        SELECT (lcLCursor)
                    ENDSCAN
                ENDIF
                USE
                SET DELETED &lcDelStatus
                USE IN (lcDBCAlias)
                IF USED(m.lcRCursor)
                    USE IN (m.lcRCursor)
                ENDIF

                *store all this stuff
                SELECT (lcViewsTbl)
                APPEND BLANK
                REPLACE ViewName WITH aViews[i], NewName WITH lcNewViewName, ;
                    ViewSQL WITH lcViewSQL, RmtViewSQL WITH lcRmtViewSQL, ;
                    CONNECTION WITH THIS.ViewConnection, Remotized WITH .T., ;
                    ViewErr WITH lcViewErr, TblsUpszd WITH lcNewTableString
                IF !lcViewErr=="" THEN
                    THIS.StoreError(.NULL.,"","",VIEW_PROPS_FAILED_LOC,aViews[i],VIEW_LOC)
                ENDIF
            ELSE
                *create remote views of tables that were upsized and leave the current view alone
                *Don't create remote view here if the user wants all upsized tables
                *to have remote views created for them
                IF !THIS.ExportTableToView THEN
                    FOR zz=1 TO ALEN(aTableNames,1)
                        IF aTableNames[zz,3]=.T. THEN
                            *Function expects array even though we'll always pass just one element from here
                            DIMENSION aWhip[1,2]
                            aWhip[1,1]=aTableNames[zz,1]
                            aWhip[1,2]=aTableNames[zz,2]
                            THIS.CreateRmtViews(@aWhip)
                        ENDIF
                    NEXT zz
                ENDIF

                *Store stuff for report
                SELECT (lcViewsTbl)
                APPEND BLANK
                REPLACE ViewName WITH aViews[i], NewName WITH "", ;
                    Remotized WITH .F., TblsUpszd WITH lcNewTableString, ;
                    NotUpszd WITH LTRIM(STRTRAN(lcTablesNotUpsized,","))

            ENDIF

        NEXT I
		raiseevent(This, 'CompleteProcess')

        *Get rid of table if no records were added; report will not be added either
        IF RECCOUNT()=0 THEN
            THIS.DisposeTable(lcViewsTbl,"Delete")
            THIS.ViewsTbl=""
        ENDIF

        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE AddToErr
        PARAMETERS lcErrString, lcAddToString

        IF lcErrString==""
            lcErrString=lcAddToString
        ELSE
            lcErrString=CHR(10)+CHR(13)+lcAddToString
        ENDIF

    ENDPROC


    FUNCTION UniqueConnDefName
        LOCAL lcConnName, I,lcExact

        *Make sure name is unique
        lcConnName=CONN_NAME_LOC
        
        *{Add JEI RKR 2005.03.28
        * IF in Database no existing connection
        LOCAL aConnDefs[1,1]
        aConnDefs[1,1] = SYS(2015)
        *}Add JEI RKR 2005.03.28 
        
        =ADBOBJECTS(aConnDefs,"connection")
        I=1
        lcExact=SET('EXACT')
        SET EXACT ON
        DO WHILE ASCAN(aConnDefs,UPPER(lcConnName))<>0
            lcConnName=CONN_NAME_LOC+LTRIM(STR(I))
            I=I+1
        ENDDO
        SET EXACT &lcExact
        RETURN lcConnName

    ENDFUNC


    PROCEDURE CreateConnDef
        LOCAL I, llFound, lcConnString, lcConnName, lcNewConnString

        lcConnName=THIS.UniqueConnDefName()
        THIS.ViewConnection=lcConnName
        IF THIS.UserConnection=="" THEN

            *If the user didn't connect with a connection definition, we
            *need to create one from scratch based on the connection string
            *created when they logged into ODBC

            lcConnString=RTRIM(THIS.CONNECTSTRING)

            *parse connection string to get user id, password, and database
            IF !EMPTY(THIS.DataSourceName) && Add JEI RKR 2005.03.28
	            lcNewConnString="dsn="+RTRIM(THIS.DataSourceName)+";"

	            llFound=.F.
	            lcUserID=THIS.ParseConnectString(THIS.CONNECTSTRING,"uid=",@llFound)

	            IF llFound THEN
	                lcNewConnString=lcNewConnString+"uid="+RTRIM(lcUserID)+";"
	            ENDIF

	            lcPassword=THIS.ParseConnectString(THIS.CONNECTSTRING,"pwd=",llFound)
	            IF llFound THEN
	                lcNewConnString=lcNewConnString+"pwd="+RTRIM(lcPassword)+";"
	            ENDIF

	            *add database name only for SQL Server
	            IF THIS.ServerType <> ORACLE_SERVER
	                lcNewConnString=lcNewConnString+"database="+RTRIM(THIS.ServerDBName)+";"
	            ENDIF
	            *Create new connection and set its connection string property
	            CREATE CONNECTION &lcConnName DATASOURCE RTRIM(THIS.DataSourceName)
	            =DBSETPROP(lcConnName,"connection","ConnectString",lcNewConnString)
	        ELSE
	        	*{ Add RKR 2005.03.28
	        	IF EMPTY(THIS.ParseConnectString(lcConnString,"database=",llFound)) And !EMPTY(THIS.ServerDBName)
	        		 lcConnString = lcConnString + IIF(RIGHT(lcConnString,1)<> ";",";","") + "database="+RTRIM(THIS.ServerDBName)+";"
	        	ENDIF	        	
				CREATE CONNECTION &lcConnName CONNSTRING (lcConnString)        
	        	*} Add RKR 2005.03.28
			ENDIF
        ELSE

            *user specified a connection def so new conndef should be just like it
            *except for the database specified

            *Get all properties (except datasource which we already know)
            llAsynchronous=		DBGETPROP(THIS.UserConnection,"connection","Asynchronous")
            llBatchMode=		DBGETPROP(THIS.UserConnection,"connection","BatchMode")
            lcConnectString=	DBGETPROP(THIS.UserConnection,"connection","ConnectString")
            lcConnectTimeOut=	DBGETPROP(THIS.UserConnection,"connection","ConnectTimeOut")
            lnDispLogin=		DBGETPROP(THIS.UserConnection,"connection","DispLogin")
            llDispWarnings=		DBGETPROP(THIS.UserConnection,"connection","DispWarnings")
            lnIdleTimeOut=		DBGETPROP(THIS.UserConnection,"connection","IdleTimeOut")
            lcPassword=			DBGETPROP(THIS.UserConnection,"connection","PassWord")
            lnQueryTimeout=		DBGETPROP(THIS.UserConnection,"connection","QueryTimeout")
            lnTransactions=		DBGETPROP(THIS.UserConnection,"connection","Transactions")
            lcUserID=			DBGETPROP(THIS.UserConnection,"connection","UserID")
            lnWaitTime=			DBGETPROP(THIS.UserConnection,"connection","WaitTime")

            *Create bare-bones connection
            CREATE CONNECTION &lcConnName DATASOURCE RTRIM(THIS.DataSourceName)

            *Hack connection string so it points at new database
            lcOldDatabase=THIS.ParseConnectString(lcConnectString,"database=")
            IF !lcOldDatabase=="" THEN
                lcConnectString=STRTRAN(lcConnectString,lcOldDatabase,RTRIM(THIS.ServerDBName))
            ENDIF

            *Make new connection's properties just like this.UserConnection's
            =DBSETPROP(THIS.ViewConnection,"connection","Asynchronous",llAsynchronous)
            =DBSETPROP(THIS.ViewConnection,"connection","BatchMode",llBatchMode)
            =DBSETPROP(THIS.ViewConnection,"connection","ConnectString",lcConnectString)
            =DBSETPROP(THIS.ViewConnection,"connection","ConnectTimeOut",lcConnectTimeOut)
            =DBSETPROP(THIS.ViewConnection,"connection","DispLogin",lnDispLogin)
            =DBSETPROP(THIS.ViewConnection,"connection","DispWarnings",llDispWarnings)
            =DBSETPROP(THIS.ViewConnection,"connection","IdleTimeOut",lnIdleTimeOut)
            =DBSETPROP(THIS.ViewConnection,"connection","PassWord",lcPassword)
            =DBSETPROP(THIS.ViewConnection,"connection","QueryTimeout",lnQueryTimeout)
            =DBSETPROP(THIS.ViewConnection,"connection","Transactions",lnTransactions)
            =DBSETPROP(THIS.ViewConnection,"connection","UserID",lcUserID)
            =DBSETPROP(THIS.ViewConnection,"connection","WaitTime",lnWaitTime)

        ENDIF

        THIS.ViewConnection=lcConnName

    ENDFUNC


    FUNCTION ParseConnectString
        PARAMETERS lcConnString, lcStringPart, llFound
        LOCAL lcLowConnString, lnStartPos

        *Takes a connection string; returns the part of string asked for

        lcStringPart=LOWER(lcStringPart)
        lcLowConnString=LOWER(lcConnString)
        lnStartPos=AT(lcStringPart,lcLowConnString)
        IF lnStartPos<>0 THEN
            lcFoundString=SUBSTR(lcConnString,lnStartPos+LEN(lcStringPart))
            IF AT(";",lcFoundString)<>0 THEN
                lcFoundString=LEFT(lcFoundString,AT(";",lcFoundString)-1)
            ENDIF
            llFound=.T.
        ELSE
            lcFoundString=""
            llFound=.F.
        ENDIF

        RETURN lcFoundString

    ENDFUNC


    PROCEDURE CreateRmtViews
        PARAMETERS aTableNames

        *
        *Creates remote views for tables passed in or all tables that were upsized
        *

        LOCAL lcEnumTablesTbl, lcNewTblName, lnOldArea, lcSQL, lcEnumIndexesTbl, ;
            aPkey, lcEnumFieldsTbl, lcViewErr, lcTblError, lnErrNo, llShowTherm, lnTableCount, I, ;
            lcCRLF, lcViewExtension, lcViewName
        lnOldArea=SELECT()
        lcEnumTablesTbl=THIS.EnumTablesTbl
        lcEnumIndexesTbl=THIS.EnumIndexesTbl
        lcEnumFieldsTbl=THIS.EnumFieldsTbl
        lcCRLF=CHR(10)+CHR(13)

        IF EMPTY(aTableNames)
            DIMENSION aTableNames[1]
            SELECT TblName, RmtTblName FROM (lcEnumTablesTbl) WHERE Exported=.T. ;
                INTO ARRAY aTableNames

            IF EMPTY(aTableNames)
                RETURN	&& no tables were actually upsized
            ENDIF

            lnTableCount=ALEN(aTableNames,1)
            llShowTherm=.T.
        ENDIF

        IF llShowTherm THEN
            *Only display thermometer if this method is called by ProcessOutput;
            *otherwise the thermometer will be showing the progress
            *for the RemotizeView method already
            THIS.InitTherm(RMTZING_TABLE_LOC,lnTableCount,0)
            lnTableCount=0
        ENDIF

        IF THIS.ViewConnection=="" THEN
            IF THIS.UserConnection=="" OR THIS.CreateNewDB THEN
                THIS.CreateConnDef
            ELSE
                THIS.ViewConnection=THIS.UserConnection
            ENDIF
        ENDIF

        SELECT (lcEnumTablesTbl)
        FOR I=1 TO ALEN(aTableNames,1)
            IF llShowTherm THEN
                lcMessage=STRTRAN(THIS_TABLE_LOC,"|1",RTRIM(aTableNames[i,1]))
                THIS.UpDateTherm(lnTableCount,lcMessage)
                lnTableCount=lnTableCount+1
            ENDIF

            THIS.HadError=.F.
            THIS.MyError=0
            lcViewErr=""
            lcTblError="" && jvf 9/25/99 for table drop
            lnTblErrNo=0 && jvf 9/25/99 for drop table
            lcViewExtension=IIF(THIS.ViewPrefixOrSuffix=3, "", RTRIM(THIS.ViewNameExtension))

            * rename or drop the original table
            * jvf: 08/16/99 Added drop functionality
            lcNewTblName=RTRIM(aTableNames[i,1])
            IF THIS.DropLocalTables
                * jvf 9/15/99
                * Can't drop table unless we drop its relation first.
                * But we can't easily tell if it's in a relation. If the table to drop is on the
                * child side, easy to see that in the dbc and drop if needed. But if the table's the
                * parent side of the relation, like customer is to orders, it's hard to know what relation
                * to delete. You actually have to delete the relation of the order on cust_id. So, we'll just
                * trap for error on drop, log it, and move on.
                DROP TABLE (RTRIM(aTableNames[i,1]))
                IF THIS.HadError && probably error 1577: table is in a relation
                    lcTblError=CANT_DROP2_LOC+" "+MESSAGE()
                    lnTblErrNo=THIS.MyError
                ELSE
                    lcNewTblName = DROPPED_TABLE_STATUS_LOC  && Table Dropped
                ENDIF
            ENDIF
            * If couldn't drop table, rename it if they didn't choose a prefix/suffix for
            * the new remote view name.
            * If the user chose not to drop tables, but did not to give the view a suffix/prefix,
            * need to add _local to the table name to prevent duplicate object names in dbc.
            IF (NOT THIS.DropLocalTables OR lnTblErrNo>0) AND EMPTY(lcViewExtension)
                lcNewTblName=THIS.UniqueTorVName(RTRIM(aTableNames[i,1]))
                RENAME TABLE (RTRIM(aTableNames[i,1])) TO (lcNewTblName)
            ENDIF

            lcSQL="SELECT * FROM " + aTableNames[i,2]
            *create the view
            * jvf 08/16/99 Add prefix/suffix to new view name

            lcViewName = IIF(THIS.ViewPrefixOrSuffix=1,lcViewExtension+RTRIM(aTableNames[i,1]),;
                RTRIM(aTableNames[i,1])+lcViewExtension)

            CREATE SQL VIEW &lcViewName REMOTE CONNECT RTRIM(THIS.ViewConnection) ;
                AS &lcSQL

            *See if table has primary key or candidate
            DIMENSION aPkey[1]
            aPkey=.F.
            SELECT RmtExpr FROM (lcEnumIndexesTbl) WHERE RTRIM(IndexName)==RTRIM(aTableNames[i,1]) ;
                AND LclIdxType="Primary key" INTO ARRAY aPkey

            *If not but unique index is available, use that
            IF EMPTY(aPkey) THEN
                SELECT RmtExpr FROM (lcEnumIndexesTbl) WHERE RTRIM(IndexName)==RTRIM(aTableNames[i,1]) ;
                    AND RmtType="UNIQUE" ;
                    INTO ARRAY aPkey
            ENDIF

            IF !EMPTY(aPkey) THEN
                *Make the whole thing updatable
                * jvf 08/16/99 just using var here now instead of array element
                IF !DBSETPROP(lcViewName,"view","SendUpdates",.T.) THEN
                    lcViewErr=UPDATE_PROP_FAILED_LOC
                ENDIF

                *Set the keyfields
                DIMENSION aKeyFields[1]
                aKeyFields=.F.
                THIS.KeyArray(aPkey,@aKeyFields)
                FOR ii=1 TO ALEN(aKeyFields,1)
                    * jvf 08/16/99 just using lcViewName var here now instead of array element
                    * JEI RKR 2005.07.14 add ChrTran
                    IF !DBSETPROP(lcViewName+"."+CHRTRAN(RTRIM(aKeyFields[ii]),"[]",""),"field","keyfield",.T.) THEN
                        lcErrString=STRTRAN(KEYFIELD_PROP_FAILED_LOC,'|1',RTRIM(aKeyFields[ii]))
                        lcViewErr=lcViewErr+lcCRLF+lcErrString
                    ENDIF
                NEXT ii

                *Make all the fields updatable
                DIMENSION aFldNames[1]
                aFldNames=.F.
                SELECT RmtFldname FROM (lcEnumFieldsTbl) WHERE RTRIM(TblName)==RTRIM(aTableNames[i,1]) ;
                    INTO ARRAY aFldNames
                IF !EMPTY(aFldNames) THEN
                    FOR ii=1 TO ALEN(aFldNames,1)
                        * jvf 08/16/99 just using lcViewName var here now instead of array element
                        IF !DBSETPROP(lcViewName+"."+RTRIM(aFldNames[ii]),"field","updatable",.T.) THEN
                            lcErrString=STRTRAN(UPDATABLE_PROP_FAILED_LOC,'|1',RTRIM(aKeyFields[ii]))
                            lcViewErr=lcViewErr+lcCRLF+lcErrString
                        ENDIF
                    NEXT ii
                ENDIF

            ELSE

                lcViewErr=NO_UNIQUEKEY_LOC

            ENDIF

            *store these table and view names somewhere
            LOCATE FOR LOWER(RTRIM(&lcEnumTablesTbl..TblName))==LOWER(RTRIM(aTableNames[i,1]))
            * jvf 08/16/99 just using lcNewTblName var here now instead of array element
            REPLACE NewTblName WITH lcNewTblName, RmtView WITH lcViewName, ;
                ViewErr WITH lcViewErr, TblError WITH lcTblError, TblErrNo WITH lnTblErrNo

            lcViewErr=""
            lcTblError=""
            lnTblErrNo=0

        NEXT I

        SELECT (lnOldArea)

		raiseevent(This, 'CompleteProcess')


    ENDPROC


    FUNCTION UniqueTorVName
        PARAMETERS lcTableName
        LOCAL lcNewTableName, lcOldTableName, I, lcExact

        *Need to make sure that when renaming tables and views that we don't overwrite
        *existing ones; this function returns a unique name

        DIMENSION aViewsArr[1]
        =ADBOBJECTS(aViewsArr,'view')
        =ADBOBJECTS(aTablesArr,'table')
        lcOldTableName=lcTableName
        lcNewTblName=LEFT(RTRIM(lcTableName),MAX_FIELDNAME_LEN-LEN(LOCAL_SUFFIX_LOC))+LOCAL_SUFFIX_LOC
        I=1
        lcExact = SET('EXACT')
        SET EXACT ON
        DO WHILE ASCAN(aViewsArr,UPPER(lcNewTblName))<>0 ;
                OR ASCAN(aTablesArr,UPPER(lcNewTblName))<>0
            IF LEN(lcNewTblName)+ LEN(LTRIM(STR(I)))>=MAX_FIELDNAME_LEN THEN
                lcOldTableName=LEFT((lcTableName),LEN(LTRIM(STR(I))))
                lcNewTblName=RTRIM(lcOldTableName)+LTRIM(STR(I))+ LOCAL_SUFFIX_LOC
            ELSE
                *Just stick a number on the end
                lcNewTblName=RTRIM(lcOldTableName)+LOCAL_SUFFIX_LOC+LTRIM(STR(I))
            ENDIF
            I=I+1
        ENDDO
        SET EXACT &lcExact

        RETURN lcNewTblName

    ENDFUNC


    PROCEDURE BuildReport
        LOCAL I, lcNewDB, lcNewProj, lcMisc, lcErrTblName, lcPath, lcPathAndFile, ;
            aRepArray
        *
        *Creates a project with report or just export error tables if the user
        *didn't ask for a report
        *
        IF (THIS.DataErrors OR !THIS.ErrTbl=="") AND !THIS.DoReport THEN
            IF This.lQuiet or !MESSAGEBOX(DATA_ERRORS_LOC,ICON_EXCLAMATION+YES_NO_BUTTONS,TITLE_TEXT_LOC)=USER_YES
                THIS.SaveErrors=.F.
            ELSE
                THIS.SaveErrors=.T.
            ENDIF
        ELSE
            IF THIS.DoReport THEN
                THIS.SaveErrors=.T.
            ELSE
                THIS.SaveErrors=.F.
            ENDIF
        ENDIF

        *Bail if nothing to do
        IF !THIS.DoReport AND !THIS.SaveErrors AND !THIS.DoScripts THEN
            RETURN
        ENDIF

* If we have an extension object and it has a BuildReport method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'BuildReport', 5) and ;
			not This.oExtension.BuildReport(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        IF THIS.DoReport THEN
            THIS.InitTherm(BUILDING_REPORT_LOC,0,0)

            *Can't use table name as primary and foreign keys because it's too long
            *Stuff integers into fields
            THIS.PurgeTable(THIS.EnumTablesTbl,"Export=.F. OR Upsizable = .F.")
            THIS.ReorderTable()
            THIS.Integerize

            *Close analysis tables
            THIS.DisposeTable(THIS.MappingTable,"close")
            THIS.DisposeTable(THIS.ViewsTbl,"close")

            *Get rid of records where stuff wasn't chosen for upsizing by user
            THIS.PurgeTable2	&&Deals with this.EnumFieldsTbl
            USE IN (THIS.EnumTablesTbl)

            IF !THIS.EnumRelsTbl=="" THEN
                THIS.PurgeTable(THIS.EnumRelsTbl,"Exported=.F.")
            ENDIF
            IF !THIS.EnumIndexesTbl=="" THEN
                THIS.PurgeTable(THIS.EnumIndexesTbl,"TblUpszd=.F.")
            ENDIF

        ELSE
            IF THIS.DoScripts THEN
                THIS.InitTherm(SCRIPT_INDB_LOC,0,0)
            ELSE
                THIS.InitTherm(PREP_ERR_LOC,0,0)
            ENDIF
        ENDIF

* Create the reports folder if it doesn't exist, copy all of our tables to it,
* then change to that folder so all files are created there.

		if not sys(5) + upper(curdir()) == upper(addbs(This.ReportDir))
			if not directory(This.ReportDir)
				md (This.ReportDir)
			endif not directory(This.ReportDir)
*** DH 2015-08-18: avoid a "file in use" error by closing the table before copying it
			This.DisposeTable(This.EnumFieldsTbl, 'close')
			copy file (This.EnumFieldsTbl + '.*') to ;
				forcepath(This.EnumFieldsTbl + '.*', This.ReportDir)
			erase (This.EnumFieldsTbl + '.*')
*** DH 2015-08-18: avoid a "file in use" error by closing the table before copying it
			This.DisposeTable(This.EnumTablesTbl, 'close')
			copy file (This.EnumTablesTbl + '.*') to ;
				forcepath(This.EnumTablesTbl + '.*', This.ReportDir)
			erase (This.EnumTablesTbl + '.*')
			if This.ExportIndexes
*** DH 2015-08-18: avoid a "file in use" error by closing the table before copying it
				This.DisposeTable(This.EnumIndexesTbl, 'close')
				copy file (This.EnumIndexesTbl + '.*') to ;
					forcepath(This.EnumIndexesTbl + '.*', This.ReportDir)
				erase (This.EnumIndexesTbl + '.*')
			endif This.ExportIndexes
			if This.ExportRelations
*** DH 2015-08-18: avoid a "file in use" error by closing the table before copying it
				This.DisposeTable(This.EnumRelsTbl, 'close')
				copy file (This.EnumRelsTbl + '.*') to ;
					forcepath(This.EnumRelsTbl + '.*', This.ReportDir)
				erase (This.EnumRelsTbl + '.*')
			endif This.ExportRelations
			if not empty(This.ViewsTbl)
*** DH 2015-08-18: avoid a "file in use" error by closing the table before copying it
				This.DisposeTable(This.ViewsTbl, 'close')
				copy file (This.ViewsTbl + '.*') to ;
					forcepath(This.ViewsTbl + '.*', This.ReportDir)
				erase (This.ViewsTbl + '.*')
			endif not empty(This.ViewsTbl)
			if This.DoScripts
				This.DisposeTable(This.ScriptTbl, 'close')
				copy file (This.ScriptTbl + '.*') to ;
					forcepath(This.ScriptTbl + '.*', This.ReportDir)
				erase (This.ScriptTbl + '.*')
			endif This.DoScripts
			if This.SaveErrors and not empty(This.ErrTbl)
*** DH 2015-08-18: avoid a "file in use" error by closing the table before copying it
				This.DisposeTable(This.ErrTbl, 'close')
				copy file (This.ErrTbl + '.*') to ;
					forcepath(This.ErrTbl + '.*', This.ReportDir)
				erase (This.ErrTbl + '.*')
			endif This.SaveErrors ...
			if This.SaveErrors and not empty(This.aDataErrTbls)
				for I = 1 to alen(This.aDataErrTbls, 1)
					lcFile = This.aDataErrTbls[I, 2] + '.*'
*** DH 2015-08-18: avoid a "file in use" error by closing the table before copying it
					This.DisposeTable(This.aDataErrTbls[I, 2], 'close')
					copy file (lcFile) to forcepath(lcFile, This.ReportDir)
					erase (lcFile)
			    next I
			endif This.SaveErrors ...
			cd (This.ReportDir)
		endif not sys(5) + upper(curdir()) == upper(addbs(This.ReportDir))

        *Create new database
        lcNewDB = This.UniqueFileName(NEWDB_NAME_LOC, 'dbc')
        CREATE DATABASE &lcNewDB
        SET DATABASE TO &lcNewDB

        IF THIS.DoReport THEN
            *Create table that contains 'one time only' data and put the data in it
            lcMisc=THIS.CreateWzTable(MISC_NAME_LOC)
            THIS.PutDataInMisc(lcMisc)

            *Add analysis tables to database
            ADD TABLE (RTRIM(THIS.EnumFieldsTbl)) NAME FIELD_NAME_LOC
            ADD TABLE (RTRIM(THIS.EnumTablesTbl)) NAME TABLE_NAME_LOC
            IF THIS.ExportIndexes THEN
                ADD TABLE (RTRIM(THIS.EnumIndexesTbl)) NAME INDEX_NAME_LOC
            ENDIF

            IF THIS.ExportRelations THEN
                ADD TABLE (RTRIM(THIS.EnumRelsTbl)) NAME REL_NAME_LOC
            ENDIF
            IF !THIS.ViewsTbl=="" THEN
                ADD TABLE (RTRIM(THIS.ViewsTbl)) NAME VIEW_NAME_LOC
            ENDIF

            *Set relations between them

* For some reason, the ALTER TABLE commands sometimes fail the first time,
* so try a second time if so.
			try
	            ALTER TABLE TABLE_NAME_LOC ALTER COLUMN TblID I PRIMARY KEY
			catch
            	ALTER TABLE TABLE_NAME_LOC ALTER COLUMN TblID I PRIMARY KEY
			endtry
            *This "tables" table needs to be open for cleanup later on
            USE
            USE TABLE_NAME_LOC ALIAS (THIS.EnumTablesTbl)
            SELECT 0
			try
            	ALTER TABLE FIELD_NAME_LOC ALTER COLUMN TblID I REFERENCES TABLE_NAME_LOC TAG TblID
			catch
	            ALTER TABLE FIELD_NAME_LOC ALTER COLUMN TblID I REFERENCES TABLE_NAME_LOC TAG TblID
			endtry
            USE
            IF THIS.ExportIndexes
				try
                	ALTER TABLE INDEX_NAME_LOC ALTER COLUMN TblID I REFERENCES TABLE_NAME_LOC TAG TblID
				catch
                	ALTER TABLE INDEX_NAME_LOC ALTER COLUMN TblID I REFERENCES TABLE_NAME_LOC TAG TblID
				endtry
                USE
            ENDIF

        ENDIF

        IF THIS.DoScripts THEN
            THIS.DisposeTable(THIS.ScriptTbl,"close")
            ADD TABLE (RTRIM(THIS.ScriptTbl)) NAME SCRIPT_NAME_LOC
        ENDIF

        *Toss in error table
        IF THIS.SaveErrors AND !THIS.ErrTbl=="" THEN
            THIS.DisposeTable(THIS.ErrTbl,"close")
            ADD TABLE (RTRIM(THIS.ErrTbl)) NAME ERROR_NAME_LOC
        ENDIF

        *Add tables that contain failed data exports
        IF THIS.SaveErrors AND !EMPTY(This.aDataErrTbls) THEN
            FOR I=1 TO ALEN(This.aDataErrTbls,1)
                lcErrTblName=ERR_TBL_PREFIX_LOC + ;
                    LEFT(This.aDataErrTbls[i,1],MAX_NAME_LENGTH-LEN(ERR_TBL_PREFIX_LOC))
                ADD TABLE (RTRIM(This.aDataErrTbls[i,2])) NAME (lcErrTblName)
            NEXT
        ENDIF

        *
        *Create new project
        *

        lcNewProj = This.UniqueFileName(NEWPROJ_NAME_LOC, 'pjx')
        THIS.NewProjName=lcNewProj
        USE Project1.PJX
        COPY TO lcNewProj+".pjx"
        USE lcNewProj+".pjx"

        *change project path to its new directory
        lcPath=DBC()
        lcPath=STRTRAN(lcPath,SET('DATA')+".DBC")
        lcPathAndFile=lcPath+lcNewProj+".PJX" + CHR(0)
        lcPath=lcPath+CHR(0)

        REPLACE NAME WITH lcPathAndFile, ;
            HOMEDIR WITH lcPath, ;
            OBJECT WITH lcPath, ;
            Reserved1 WITH lcPathAndFile

        *Add database to project
        LOCATE FOR LOWER(TYPE)="d"
        REPLACE NAME WITH LOWER(lcNewDB)+".dbc", KEY WITH UPPER(lcNewDB)

        *Create reports and add report records to project
        IF THIS.DoReport
            DIMENSION aRepArray[1,2]
            THIS.AddToRepArray(FIELDS_REPORT_LOC,@aRepArray)
            THIS.AddToRepArray(TABLES_REPORT_LOC,@aRepArray)
            IF !THIS.ErrTbl=="" AND THIS.SaveErrors THEN
                THIS.AddToRepArray(ERR_REPORT_LOC,@aRepArray)
            ENDIF
            IF THIS.ExportIndexes THEN
                THIS.AddToRepArray(INDEX_REPORT_LOC,@aRepArray)
            ENDIF
            IF THIS.ExportRelations THEN
                THIS.AddToRepArray(RELS_REPORT_LOC,@aRepArray)
            ENDIF
            IF THIS.ExportViewToRmt AND !THIS.ViewsTbl=="" THEN
                THIS.AddToRepArray(VIEWS_REPORT_LOC,@aRepArray)
            ENDIF

            FOR I=1 TO ALEN(aRepArray,1)

                *Create copy of each upsizing report
                aRepArray[i,2]=THIS.UniqueFileName(aRepArray[i,1],"frx")
                SELECT 0
                USE (aRepArray[i,1]) + ".frx"
                COPY TO aRepArray[i,2] + ".frx"
                USE (aRepArray[i,2]) + ".frx"
                LOCATE FOR NAME="cursor"
                *Need to fiddle with the DE
                DO WHILE FOUND()
                    REPLACE EXPR WITH STRTRAN(EXPR,"upsize1",DBC())
                    CONTINUE
                ENDDO
                USE

                *Add to project table
                SELECT (lcNewProj)
                APPEND BLANK
                REPLACE NAME WITH LOWER(aRepArray[i,2])+".frx", ;
                    KEY WITH UPPER(aRepArray[i,2]), ;
                    EXCLUDE WITH .F., ;
                    TYPE WITH "R"

            ENDFOR

        ENDIF
        USE

		raiseevent(This, 'CompleteProcess')

        lcNewProj = ADDBS(FULLPATH(SYS(5))) + lcNewProj
        if not This.lQuiet
	        _SHELL=[MODIFY PROJECT "&lcNewProj" NOWAIT]
        endif not This.lQuiet
        THIS.KeepNewDir=.T.

    ENDPROC


    PROCEDURE Integerize

        LOCAL lcEnumTablesTbl, lcEnumFieldsTbl, lcEnumIndexesTbl, lcTemp, lcTemp1

        *Add unique IDs to table names and propagate to child tables
        lcEnumTablesTbl=THIS.EnumTablesTbl
        lcEnumFieldsTbl=THIS.EnumFieldsTbl
        lcEnumIndexesTbl=THIS.EnumIndexesTbl

        SELECT (lcEnumTablesTbl)
        REPLACE ALL TblID WITH RECNO()
        SCAN
            lcTemp =&lcEnumTablesTbl..TblID
            lcTemp1 = &lcEnumTablesTbl..TblName
            UPDATE (lcEnumFieldsTbl);
                SET &lcEnumFieldsTbl..TblID=lcTemp ;
                WHERE &lcEnumFieldsTbl..TblName==lcTemp1

            IF THIS.ExportIndexes
                lcTemp =&lcEnumTablesTbl..TblID
                lcTemp1 =&lcEnumTablesTbl..TblName
                UPDATE (lcEnumIndexesTbl);
                    SET &lcEnumIndexesTbl..TblID=lcTemp ;
                    WHERE &lcEnumIndexesTbl..IndexName==lcTemp1
            ENDIF

        ENDSCAN

    ENDPROC

    PROCEDURE AddToRepArray
        PARAMETERS lcRepName, aRepArray
        LOCAL lnArrayLen
        *This method assumes that the passed array is 2D; the string passed is placed
        *in the first column

        IF EMPTY(aRepArray) THEN
            aRepArray[1,1]=lcRepName
        ELSE
            lnArrayLen=ALEN(aRepArray,1)
            DIMENSION aRepArray[lnArrayLen+1,2]
            aRepArray[lnArrayLen+1,1]=lcRepName
        ENDIF

    ENDPROC


    PROCEDURE PutDataInMisc
        PARAMETERS lcMisc

        INSERT INTO &lcMisc ;
            (SvType, ;
            DataSourceName, ;
            UserConnection, ;
            ViewConnection, ;
            DeviceDBName, ;
            DeviceDBPName, ;
            DeviceDBSize,;
            DeviceDBNumber, ;
            DeviceLogName, ;
            DeviceLogPName, ;
            DeviceLogSize, ;
            DeviceLogNumber, ;
            ServerDBName, ;
            ServerDBSize, ;
            ServerLogSize, ;
            SourceDB, ;
            ExportIndexes, ;
            ExportValidation, ;
            ExportRelations, ;
            ExportStructureOnly, ;
            ExportDefaults, ;
            ExportTimeStamp, ;
            ExportTableToView, ;
            ExportViewToRmt, ;
            ExportDRI, ;
            ExportSavePwd, ;
            DoUpsize, ;
            DoScripts, ;
            DoReport) ;
            VALUES ;
            (THIS.ServerType, ;
            THIS.DataSourceName, ;
            THIS.UserConnection, ;
            THIS.ViewConnection, ;
            THIS.DeviceDBName, ;
            THIS.DeviceDBPName, ;
            THIS.DeviceDBSize, ;
            THIS.DeviceDBNumber, ;
            THIS.DeviceLogName, ;
            THIS.DeviceLogPName, ;
            THIS.DeviceLogSize, ;
            THIS.DeviceLogNumber, ;
            THIS.ServerDBName, ;
            THIS.ServerDBSize, ;
            THIS.ServerLogSize, ;
            THIS.SourceDB, ;
            THIS.ExportIndexes, ;
            THIS.ExportValidation, ;
            THIS.ExportRelations, ;
            THIS.ExportStructureOnly, ;
            THIS.ExportDefaults, ;
            THIS.ExportTimeStamp, ;
            THIS.ExportTableToView, ;
            THIS.ExportViewToRmt, ;
            THIS.ExportDRI, ;
            THIS.ExportSavePwd, ;
            THIS.DoUpsize, ;
            THIS.DoScripts, ;
            THIS.DoReport)

        USE

    ENDPROC


    PROCEDURE ReorderTable
        LOCAL lcNewName, lcAlias

        *Copy to new table sorted by table name (can't index on 128 character field
        *using general code page)

        *		lcNewName=SUBSTR(SYS(2015),3,10)
        lcNewName="_"+SUBSTR(SYS(2015),4,10)
        RENAME (THIS.EnumTablesTbl)+".dbf" TO (lcNewName)+".dbf"
        RENAME (THIS.EnumTablesTbl)+".fpt" TO (lcNewName)+".fpt"
        *{ JEI RKR 18.11.2005 Add
        IF FILE(THIS.EnumTablesTbl+".cdx")
        	RENAME (THIS.EnumTablesTbl)+".cdx" TO (lcNewName)+".cdx"
        ENDIF
        *{ JEI RKR 18.11.2005 Add
        SELECT 0
        USE (lcNewName)
        lcAlias=ALIAS()

        SELECT * FROM (lcNewName);
            ORDER BY TblName;
            INTO TABLE (THIS.EnumTablesTbl)

        USE IN (lcAlias)

        DELETE FILE (lcNewName)+".dbf"
        DELETE FILE (lcNewName)+".fpt"
        DELETE FILE (lcNewName)+".cdx"

    ENDPROC


    PROCEDURE PurgeTable
        PARAMETERS lcTableName, lcCondition
        *This removes objects which were analyzed but not selected for upsizing

        SELECT (lcTableName)
        SET FILTER TO
        DELETE FOR &lcCondition
        PACK
        USE

    ENDPROC


    PROCEDURE PurgeTable2
        LOCAL lcTableName

        *Gets rid of field info related to tables not selected for upsizing
        *(This info gets added if the user selects a table, changes pages, and
        *then deselects the table)

        SELECT (THIS.EnumTablesTbl)
        SCAN FOR EXPORT=.F.
            lcTableName=RTRIM(TblName)
            SELECT (THIS.EnumFieldsTbl)
            DELETE FOR RTRIM(TblName)==lcTableName
            SELECT (THIS.EnumTablesTbl)
        ENDSCAN

        SELECT (THIS.EnumFieldsTbl)
        PACK
        USE

    ENDPROC


    PROCEDURE CreateScript
        LOCAL lcEnumRelsTbl, lcEnumTablesTbl, lcCRLF, lcComment, lcEnumIndexesTbl, ;
            lcEnumFieldsTbl, lnTableCount, lcEnumClustersTbl

        *Device and database sql (if any) are already in the script memo field by now
        IF !THIS.DoScripts THEN
            RETURN
        ENDIF

* If we have an extension object and it has a CreateScript method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'CreateScript', 5) and ;
			not This.oExtension.CreateScript(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        lcCRLF = CHR(13)
        lcEnumClustersTbl = THIS.EnumClustersTbl
        lcEnumTablesTbl = THIS.EnumTablesTbl
        lcEnumIndexesTbl = THIS.EnumIndexesTbl
        lcEnumRelsTbl = THIS.EnumRelsTbl
        lcEnumFieldsTbl = THIS.EnumFieldsTbl

        SELECT COUNT(*) FROM (lcEnumTablesTbl) WHERE EXPORT=.T. INTO ARRAY aTableCount
        THIS.InitTherm(BUILDING_SCRIPT_LOC, aTableCount,0)
        lnTableCount = 0
        THIS.UpDateTherm(lnTableCount, "")

        IF THIS.ServerType = "Oracle"
            * Grab cluster SQL (if any)
            IF !EMPTY(lcEnumClustersTbl)
                SELECT (lcEnumClustersTbl)
                SCAN FOR !EMPTY(ClustName)

                    * Get cluster SQL
                    lcClustName = RTRIM(ClustName)
                    lcSQL = THIS.BuildComment(CLUST_COMMENT_LOC, lcClustName)
                    lcSQL = lcSQL + ClusterSQL
                    THIS.StoreSQL(lcSQL, "")
                    lcSQL = ""

                    *Grab cluster index SQL
                    SELECT (lcEnumIndexesTbl)
                    LOCATE FOR RTRIM(RmtTable) == lcClustName
                    IF FOUND()
                        lcSQL = THIS.BuildComment(CLUST_INDEX_LOC, RmtName)
                        lcSQL = lcSQL + IndexSQL
                        THIS.StoreSQL(lcSQL,"")
                    ENDIF
                    lcSQL = ""

                    * Grab SQL of tables (and their triggers) in cluster
                    SELECT (lcEnumTablesTbl)
                    SCAN FOR RTRIM(ClustName) == lcClustName
                        lcTableName = RTRIM(TblName)
                        lcRmtTblName = RTRIM(RmtTblName)
                        THIS.OracleScript(lcTableName, lcRmtTblName)
                        lnTableCount = lnTableCount + 1
                        THIS.UpDateTherm(lnTableCount)
                    ENDSCAN
                    SELECT (lcEnumClustersTbl)
                ENDSCAN
            ENDIF
            * Deal with tables not in clusters
            * Grab SQL of tables (and their triggers) in cluster
            SELECT (lcEnumTablesTbl)
            SCAN FOR EMPTY(ClustName) AND EXPORT=.T.
                lcTableName = RTRIM(TblName)
                lcRmtTblName = RTRIM(RmtTblName)
                THIS.OracleScript(lcTableName, lcRmtTblName)
                lnTableCount = lnTableCount + 1
                THIS.UpDateTherm(lnTableCount)
            ENDSCAN
        ELSE
            * Deal with SQL Server
            SELECT (lcEnumTablesTbl)
            IF THIS.ZDUsed THEN
                lcSQL = ZD_DESC_LOC + lcCRLF
                lcSQL=lcSQL + "CREATE DEFAULT " + ZERO_DEFAULT_NAME + " AS 0" + lcCRLF
                THIS.StoreSQL(lcSQL,"")
            ENDIF

            SCAN FOR EXPORT=.T.

                *Grab table SQL
                lcTableName=RTRIM(TblName)
                lcRmtTblName=RTRIM(RmtTblName)
                lcSQL = THIS.BuildComment(TABLE_COMMENT_LOC,lcRmtTblName) + TableSQL
                THIS.StoreSQL(lcSQL,"")

                *Triggers
                lcSQL=""
                IF !EMPTY(InsertRI) THEN
                    lcSQL = InsertRI
                ENDIF
                IF !EMPTY(UpdateRI) THEN
                    lcSQL = IIF(EMPTY(lcSQL), UpdateRI, lcSQL + UpdateRI + lcCRLF)
                ENDIF
                IF !EMPTY(DeleteRI) THEN
                    lcSQL = IIF(EMPTY(lcSQL), DeleteRI, lcSQL + DeleteRI + lcCRLF)
                ENDIF
                IF !lcSQL=="" THEN
                    THIS.StoreSQL(lcSQL, TRIGGER_COMMENT_LOC)
                ENDIF

                *Grab index SQL
                IF !EMPTY(lcEnumIndexesTbl)
                    SELECT (lcEnumIndexesTbl)
                    lcSQL = lcCRLF + INDEX_COMMENT_LOC + lcCRLF
                    SCAN FOR RTRIM(IndexName)==lcTableName AND DontCreate=.F.
                        lcSQL=lcSQL+ IndexSQL
                        THIS.StoreSQL(lcSQL,"")
                        lcSQL=""
                    ENDSCAN
                ENDIF

                *Grab default SQL
                SELECT (lcEnumFieldsTbl)
                lcSQL = lcCRLF + DEFAULT_COMMENT_LOC + lcCRLF
                SCAN FOR RTRIM(TblName)==lcTableName AND !EMPTY(RmtDefault)
                    IF RmtDefault<>"0" THEN
                        lcSQL = lcSQL + RmtDefault + lcCRLF
                    ENDIF
                    lcSQL = lcSQL + "sp_bindefault " + RTRIM(RDName) + ", '" + RTRIM(TblName) + "." + RTRIM(FldName) +"'" + lcCRLF
                    THIS.StoreSQL(lcSQL,"")
                    lcSQL=""

                ENDSCAN

                *Stored procedures
                lcSQL=SPROC_COMMENT_LOC
                SCAN FOR RTRIM(TblName)==lcTableName AND !EMPTY(RmtRule)
                    lcSQL=lcSQL + RmtDefault
                    THIS.StoreSQL(lcSQL,"")
                    lcSQL=""
                ENDSCAN

                SELECT (lcEnumTablesTbl)
                lnTableCount=lnTableCount+1
                THIS.UpDateTherm(lnTableCount)

            ENDSCAN

        ENDIF

		raiseevent(This, 'CompleteProcess')

    ENDPROC


    PROCEDURE OracleScript
        PARAMETERS lcTableName, lcRmtTblName
        LOCAL lcEnumTablesTbl, lcEnumIndexsTbl, lcEnumFieldsTbl, lcCRLF

        *
        *Called by CreateScript; puts together all the sql for a table
        *including the table, triggers, indexes and defaults
        *

        lcEnumTablesTbl = THIS.EnumTablesTbl
        lcEnumIndexesTbl = THIS.EnumIndexesTbl
        lcEnumFieldsTbl = THIS.EnumFieldsTbl
        lcCRLF = CHR(13)

        * Grab SQL of tables (and their triggers)
        lcSQL = THIS.BuildComment(TABLE_COMMENT_LOC, lcRmtTblName) + lcCRLF + TableSQL
        THIS.StoreSQL(lcSQL, "")

        * Triggers
        lcSQL = ""
        IF !EMPTY(InsertRI) THEN
            lcSQL = InsertRI + lcCRLF
        ENDIF
        IF !EMPTY(UpdateRI) THEN
            lcSQL = lcSQL + UpdateRI + lcCRLF
        ENDIF
        IF !EMPTY(DeleteRI) THEN
            lcSQL = lcSQL + DeleteRI + lcCRLF
        ENDIF
        IF !EMPTY(lcSQL)
            THIS.StoreSQL(lcSQL, TRIGGER_COMMENT_LOC)
        ENDIF

        * Grab index sql
        SELECT (lcEnumIndexesTbl)
        lcSQL = lcCRLF + INDEX_COMMENT_LOC + lcCRLF + lcCRLF
        SCAN FOR RTRIM(IndexName) == lcTableName AND !EMPTY(IndexSQL)
            lcSQL = lcSQL + IndexSQL
            THIS.StoreSQL(lcSQL, "")
            lcSQL = ""
        ENDSCAN

        *Grab default sql
        SELECT (lcEnumFieldsTbl)
        lcSQL = lcCRLF + DEFAULT_COMMENT_LOC + lcCRLF + lcCRLF
        SCAN FOR RTRIM(TblName) == lcTableName AND !EMPTY(RmtDefault)
            lcSQL = lcSQL + RmtDefault
            THIS.StoreSQL(lcSQL,"")
            lcSQL = ""
        ENDSCAN

        SELECT (lcEnumTablesTbl)

    ENDPROC


    FUNCTION UniqueFileName
        PARAMETERS lcFileName, lcExtension
        LOCAL I, lcNewName

        lcNewName=lcFileName
        I=1
        DO WHILE FILE(lcNewName + "." + lcExtension)
            lcNewName=LEFT(lcFileName,MAX_DOSNAME_LEN-LEN(LTRIM(STR(I)))) + LTRIM(STR(I))
            I=I+1
        ENDDO
        RETURN lcNewName

    ENDFUNC


    PROCEDURE JustStem2

        * Return just the stem name from "filname"
        * Unlike JustStem, this returns file name in same case it came in as

        LPARAMETERS m.filname
        IF RAT('\',m.filname) > 0
            m.filname = SUBSTR(m.filname,RAT('\',m.filname)+1,255)
        ENDIF
        IF RAT(':',m.filname) > 0
            m.filname = SUBSTR(m.filname,RAT(':',m.filname)+1,255)
        ENDIF
        IF AT('.',m.filname) > 0
            m.filname = SUBSTR(m.filname,1,AT('.',m.filname)-1)
        ENDIF
        RETURN ALLTRIM(m.filname)

    ENDPROC


    FUNCTION RemotizeName
        PARAMETERS lcLocalName
        LOCAL lcResult, lnLength, I, lcChar, lnOldArea, lnAsc, lcExact, lcServerConstraint

        *all expressions and objects everywhere in the Upsizing Wizard are
        *lower cased, otherwise STRTRAN transformations won't work reliably
        lcExact=SET('EXACT')
        SET EXACT ON
        lcResult = LOWER(ALLTRIM(lcLocalName))
        lnOldArea=SELECT()

        * Check keyword table
        IF !USED("Keywords")
            SELECT 0
            USE Keywords
            SET ORDER TO Keyword
        ELSE
            SELECT Keywords
        ENDIF
*** DH 12/15/2014: use only SQL Server rather than variants of it
***        IF RTRIM(THIS.ServerType)=="SQL Server95" THEN
        IF left(THIS.ServerType, 10) = 'SQL Server'
            lcServerConstraint="SQL Server"
        ELSE
            lcServerConstraint=RTRIM(THIS.ServerType)
        ENDIF

*** DH 2015-09-24: don't filter the KEYWORDS table on server type
***        SET FILTER TO ServerType=lcServerConstraint

        SEEK lcResult

        IF FOUND() THEN
*** DH 2015-09-24: add brackets around name instead of suffixing with "_"
***            lcResult = lcResult + "_"
			lcResult = '[' + lcResult + ']'

        ELSE

            *if it starts with a number, stick a "_" in front of it
            IF LEFT(lcLocalName, 1) >= "0" AND LEFT(lcLocalName, 1) <= "9" THEN
                lcResult= "_" + lcResult
            ENDIF

            lnLength = LEN(lcResult)

            *ISALPHA() will return true but SQL Server will reject when...
            *Codepage 1252 (US): 156, 207
            *Codepage 1250 (E.Eur.): 156, 190, 207
            *Codepage ???? (Russia): 220
            *So these characters are always (on all code pages) turned to underscores
            *Skip for DBCS
            IF lnLength = LENC(lcResult)
                FOR I = 1 TO lnLength
                    lcChar = SUBSTR(lcResult, I, 1)
					lnAsc = ASC(lcChar)
						if (not isalpha(lcChar) and ;
							not between(lcChar, '0', '9') and ;
							lcChar <> ' ') or ;
                            inlist(lnAsc, 156, 190, 207, 220)
                        lcResult=STUFF(lcResult, I, 1, "_")
                    ENDIF
                NEXT I
            ENDIF
        ENDIF

        SET EXACT &lcExact
        SELECT (lnOldArea)
        RETURN lcResult

    ENDFUNC


    FUNCTION UniqueTableName
        PARAMETERS lcStem
        LOCAL lcTest, lcResult, I, lnLength

        lcStem=ALLTRIM(lcStem)
        lcResult=lcStem
        lcTest=lcResult + ".dbf"
        lnLength=LEN(lcStem)
        FOR I=1 TO 10^lnLength-1
            IF FILE(lcTest) THEN
                lcResult=LEFT(lcStem,lnLength-LEN(ALLTRIM(STR(I)))) + ALLTRIM(STR(I))
                lcTest=lcResult + ".dbf"
            ELSE
                EXIT
            ENDIF
        NEXT
        RETURN lcResult

    ENDFUNC


    * 10/30/02 JVF Replaced this function with the new one below. Left intact, but retired, b/c of
    * its legacy nature.
    FUNCTION UniqueCursorName_Retired
        PARAMETERS lcStem
        LOCAL lcResult, I, lnLength

        lcStem=ALLTRIM(lcStem)
        lcResult=lcStem
        lnLength=LEN(lcStem)
        FOR I=1 TO lnLength-1
            IF USED(lcResult) THEN
                lcResult=LEFT(lcStem,lnLength-LEN(ALLTRIM(STR(I)))) + ALLTRIM(STR(I))
            ELSE
                EXIT
            ENDIF
        NEXT
        RETURN lcResult

        * 10/30/02 JVF Issue 16455 Alias in use error with similar tables ending in numeric chars.
        * I see the fatal flaw of original code:
        * For I=1 TO lnLength-1
        * It only loops 6 times b/c the stem length was 7. I guess the USW never ran into situation
        * where their were more cursors open than length of the stem -1.

        * I agree we can loop 999 times and don't have to worry about names longer than 8, so we do not
        * have to keep shortnening the stem. But the following approach is a little better than the
        * recommended b/c we maintain the stem instead of just appending to it.

    FUNCTION UniqueCursorName
        PARAMETERS lcStem

        LOCAL lcResult, I, lnLength

        lcStem=ALLTRIM(lcStem)
        lcResult=lcStem
        lnLength=LEN(lcStem)
        FOR I=1 TO 999
            IF USED(lcResult) THEN
                lcResult=lcStem + PADL(I,3,"0")
            ELSE
                EXIT
            ENDIF
        NEXT
        RETURN lcResult
    ENDFUNC


    PROCEDURE CreateTypeArrays
        *Creates an array for each FoxPro datatype; each array of possible remote datatypes
        *has the same name as the local FoxPro datatype
        PRIVATE aArrays
        LOCAL lcServerConstraint, I

*** DH 09/05/2012: use only SQL Server rather than variants of it
***        IF RTRIM(THIS.ServerType)=="SQL Server95" THEN
		IF left(THIS.ServerType, 10) = 'SQL Server'
            lcServerConstraint="SQL Server"
        ELSE
***            IF THIS.ServerType="SQL Server" THEN
***                lcServerConstraint="SQL Server4x"
***            ELSE
                lcServerConstraint=RTRIM(THIS.ServerType)
***            ENDIF
        ENDIF

        *Find all the local types
        SELECT LocalType,"this." + LocalType FROM TypeMap ;
            WHERE DEFAULT=.T. AND SERVER=lcServerConstraint ;
            INTO ARRAY aArrays

        FOR I=1 TO ALEN(aArrays,1)
            SELECT RemoteType FROM TypeMap WHERE LocalType=aArrays[i,1] AND SERVER=lcServerConstraint ;
                INTO ARRAY &aArrays[i,2]
        NEXT

    ENDPROC


    FUNCTION ValidName
        PARAMETERS lcName
        LOCAL lcNewName

        lcNewName=THIS.RemotizeName(lcName)
        IF LOWER(lcNewName)<>LOWER(lcName) THEN
            *display error message
			if not This.lQuiet
	            MESSAGEBOX(INVALID_NAME_LOC, ICON_EXCLAMATION, TITLE_TEXT_LOC)
			endif not This.lQuiet
            lcName=lcNewName
            RETURN .F.
        ENDIF
        RETURN .T.

    ENDFUNC


    FUNCTION NameObject
        PARAMETERS lcRmtTableName,lcFldName,lcPrefix, lnMaxLength
        LOCAL lnTblNameLength,lnFldNameLength,lnCharsLeft

        lnTblNameLength=LEN(lcRmtTableName)
        lnFldNameLength=LEN(lcFldName)
        lnCharsLeft=(lnMaxLength)-LEN(lcPrefix)-LEN(SEP_CHARACTER)

        *If all the components of the string are too big, clip
        *the table and/or field names

*** DH 2015-09-25: strip brackets from names
		lcRmtTableName = chrtran(lcRmtTableName, '[]', '')
		lcFldName      = chrtran(lcFldName,      '[]', '')
*** DH 2015-09-25: end of new code

        IF lnCharsLeft<lnTblNameLength+lnFldNameLength THEN
            DO CASE
                    *If each name is bigger than half of what's left, clip them both
                CASE lnTblNameLength>(lnCharsLeft/2) AND lnFldNameLength>(lnCharsLeft/2)
                    lcRmtTableName=LEFT(lcRmtTableName,(lnCharsLeft/2))
                    lcFldName=LEFT(lcFldName,(lnCharsLeft/2))

                    *If the field name is super long, clip it
                CASE lnTblNameLength<=(lnCharsLeft/2) AND lnFldNameLength>(lnCharsLeft/2)
                    lcFldName=LEFT(lcFldName,(lnCharsLeft-lnTblNameLength))

                    *If the table name is super long, clip it
                CASE lnTblNameLength>(lnCharsLeft/2) AND lnFldNameLength<=(lnCharsLeft/2)
                    lcTmpTblName=LEFT(lcRmtTableName,(lnCharsLeft-lnFldNameLength))

            ENDCASE
        ENDIF

        lcSprocName=lcPrefix+lcRmtTableName + SEP_CHARACTER + lcFldName

        RETURN lcSprocName

    ENDFUNC


    PROCEDURE MaybeDrop
        PARAMETERS lcObjectName, lcObjectType
        LOCAL llObjectExists, lcSQL, dummy,lcSQT

        *This is called in several places by this.DefaultsAndRules
        *It will drop a sproc or default if it already exists

        *Check to see if the object already exists
        lcSQT=CHR(39)
        lcSQL="select uid from sysobjects where name =" + lcSQT + lcObjectName + lcSQT
        dummy="x"
        llObjectExists=THIS.SingleValueSPT(lcSQL, dummy, "uid")

        IF llObjectExists THEN
            lcSQL="drop " + lcObjectType + " " + lcObjectName
            lnRetVal=THIS.ExecuteTempSPT(lcSQL)
            RETURN lnRetVal
        ELSE
            RETURN .T.
        ENDIF

    ENDPROC


    FUNCTION ExtractFieldNames
        PARAMETERS lcExpression, lcTableName, lnKeyCount, aFieldNames
        LOCAL ii, lcReturnExpression, lcEnum_Fields, lcFieldName, lnRow

        *
        *Takes a FoxPro expression and returns comma separated list of
        *remotized version of all the field names that were in the expression
        *
        *Called by AnalyzeIndexes and BuildRICode
        *

        *Build the array of field names if it wasn't passed
        IF EMPTY(aFieldNames) THEN
            DIMENSION aFieldNames[1]
            lcEnum_Fields=RTRIM(THIS.EnumFieldsTbl)
            *Be sure they come back in order of the longest field names first
            *or the shorter ones which are substrings of longer ones (if any)
            *will mess things up
            SELECT FldName, RmtFldname, 1/LEN(RTRIM(RmtFldname)) AS foo ;
                FROM &lcEnum_Fields ;
                WHERE &lcEnum_Fields..TblName=lcTableName ;
                ORDER BY foo ;
                INTO ARRAY aFieldNames
        ENDIF

        lnKeyCount=0
        lcReturnExpression=""

        *!* jvf: 08/16/99
        * Replaced code that caused indexes to be created out of the intended field sequence.
        * The subsequent code corrects this issue.
        DIMENSION laExp[1]
        THIS.StringToArray(lcExpression, @laExp, "+")
        FOR lnRow = 1 TO ALEN(laExp,1)
            lcFieldName=THIS.StripFunction(laExp[lnRow])
            IF lcReturnExpression=="" THEN
                lcReturnExpression="[" + lcFieldName + "]"
            ELSE
                lcReturnExpression=lcReturnExpression+", ["+lcFieldName+"]"
            ENDIF

            *Keep track of how many fields are in the index expression
            lnKeyCount=lnKeyCount+1
        ENDFOR
        RETURN lcReturnExpression

    ENDFUNC


    PROCEDURE StoreError
        PARAMETERS lnError, lcErrMsg, lcSQL, lcWizErrMsg, lcObjName, lcObjType
        LOCAL lcErrTbl, lnOldArea

        *Stores errors for report

        lnOldArea=SELECT()

        IF THIS.ErrTbl=="" THEN
            THIS.ErrTbl=THIS.CreateWzTable("Errors")
        ENDIF
        lcErrTbl=THIS.ErrTbl
        IF EMPTY(lcWizErrMsg) THEN
            lcWizErrMsg=""
        ENDIF
        INSERT INTO &lcErrTbl (ErrNumber,ErrMsg, WizErr, FailedSQL, ObjName, ObjType) ;
            VALUES (lnError, lcErrMsg, lcWizErrMsg, lcSQL,lcObjName, lcObjType)
        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE StoreSQL
        PARAMETERS lcSQL,lcComment
        LOCAL lcCRLF, lnOldArea, lcScriptTbl

        *Get out of here if user doesn't want a script
        IF !THIS.DoScripts THEN
            RETURN
        ENDIF

        lcCRLF = CHR(13)
        lnOldArea=SELECT()

        IF RTRIM(THIS.ScriptTbl) == "" THEN
            lcScriptTbl = THIS.CreateWzTable("Script")
            THIS.ScriptTbl = lcScriptTbl
        ELSE
            lcScriptTbl = RTRIM(THIS.ScriptTbl)
        ENDIF

        SELECT (lcScriptTbl)

        * There should only be one record in this table
        IF RECCOUNT()=0 THEN
            APPEND BLANK
        ENDIF

        *Add some carriage returns/linefeeds and then stick everything together
        IF !EMPTY(lcComment) THEN
            lcSQL = lcCRLF + lcComment + lcCRLF + lcSQL
        ENDIF
        REPLACE &lcScriptTbl..ScriptSQL WITH lcSQL + lcCRLF ADDITIVE

        * Prevent bulk build-up of memo
        PACK MEMO

        SELECT (lnOldArea)

    ENDPROC


    PROCEDURE SetConnProps
        *Sets connection properties to where the upsizing wizard wants them
        =SQLSETPROP(THIS.MasterConnHand,"Asynchronous",.F.)
        =SQLSETPROP(THIS.MasterConnHand,"Batchmode",.T.)
        =SQLSETPROP(THIS.MasterConnHand,"ConnectTimeOut",45)
        =SQLSETPROP(THIS.MasterConnHand,"DispWarnings",.F.)
        *QueryTimeOut is set to 600 when creating devices and databases
        =SQLSETPROP(THIS.MasterConnHand,"QueryTimeOut",45)
        =SQLSETPROP(THIS.MasterConnHand,"Transactions",1)
        =SQLSETPROP(THIS.MasterConnHand,"WaitTime",100)
        *Never timeout if idle
        =SQLSETPROP(THIS.MasterConnHand,"IdleTimeOut",0)
        *Default wait time of 100 milliseconds

    ENDPROC

    PROCEDURE CreateTS
        PARAMETERS lcTSName, lcTSFName,lcTSFSize
        LOCAL lcSQL, lcSQL1, lcMsg, lnErr, lcErrMsg

        * create new tablespace on Oracle server and an associated data file
        * allocates unlimited space quota for the user on the new tablespace
        lcSQL = "CREATE TABLESPACE " + lcTSName + " DATAFILE '" + lcTSFName + "' SIZE " + ALLTRIM(STR(lcTSFSize)) + " K"
        lcSQL1 = "ALTER USER " + THIS.UserName + " QUOTA UNLIMITED ON " + lcTSName

        *Execute if appropriate
        IF THIS.DoUpsize THEN
            lcMsg=STRTRAN(CREATING_TABLESPACE_LOC,'|1',RTRIM(lcTSName))
            THIS.InitTherm(lcMsg,0,0)
            THIS.UpDateTherm(0,TAKES_AWHILE_LOC)
            =SQLSETPROP(THIS.MasterConnHand,"QueryTimeOut",600)
            THIS.MyError=0

            * create tablespace
            IF !THIS.ExecuteTempSPT(lcSQL, @lnErr,@lcErrMsg) THEN
                IF lnErr = 01543 THEN
                    *User doesn't have CREATE TABLESPACE permissions
*** DH 03/23/2015: DataSourceName may be blank so get it from the connection string.
*** Suggested by Mike Potjer.
					local lcSource
					lcSource = rtrim(This.DataSourceName)
					if empty(lcSource)
						lcSource = evl(This.DataSourceName, strextract(This.ConnectString, 'server=', ';', 1, 3))
					endif empty(lcSource)
***                    lcMsg=STRTRAN(NO_CREATETS_PERM_LOC,'|1',RTRIM(THIS.DataSourceName))
					lcMsg=STRTRAN(NO_CREATETS_PERM_LOC,'|1', lcSource)
                ELSE
                    *Something else went wrong
                    lcMsg=STRTRAN(CREATE_TS_FAILED_LOC,'|1',RTRIM(lcTSName))
                ENDIF
				if This.lQuiet
					This.HadError     = .T.
					This.ErrorMessage = lcMsg
				else
                	MESSAGEBOX(lcMsg, ICON_EXCLAMATION,TITLE_TEXT_LOC)
				endif This.lQuiet
                THIS.Die
            ENDIF

            * allocate unlimited quota on tablespace
            IF !THIS.ExecuteTempSPT(lcSQL1, @lnErr,@lcErrMsg) THEN
                IF lnErr = 01543 THEN
                    *User doesn't have CREATE TABLESPACE permissions
*** DH 03/23/2015: DataSourceName may be blank so get it from the connection string.
*** Suggested by Mike Potjer.
					local lcSource
					lcSource = rtrim(This.DataSourceName)
					if empty(lcSource)
						lcSource = evl(This.DataSourceName, strextract(This.ConnectString, 'server=', ';', 1, 3))
					endif empty(lcSource)
***                    lcMsg=STRTRAN(NO_CREATETS_PERM_LOC,'|1',RTRIM(THIS.DataSourceName))
                    lcMsg=STRTRAN(NO_CREATETS_PERM_LOC,'|1', lcSource)
                ELSE
                    *Something else went wrong
                    lcMsg=STRTRAN(CREATE_TS_FAILED_LOC,'|1',RTRIM(lcTSName))
                ENDIF
				if This.lQuiet
					This.HadError     = .T.
					This.ErrorMessage = lcMsg
				else
                	MESSAGEBOX(lcMsg, ICON_EXCLAMATION,TITLE_TEXT_LOC)
				endif This.lQuiet
                THIS.Die
            ENDIF

            =SQLSETPROP(THIS.MasterConnHand, "QueryTimeOut",30)
        ENDIF

        *Stash sql for script
        THIS.StoreSQL(lcSQL,CREATE_DBSQL_LOC)
        THIS.StoreSQL(lcSQL1,CREATE_DBSQL_LOC)
		raiseevent(This, 'CompleteProcess')

    ENDPROC

    * create new datafile in existing tablespace
    PROCEDURE CreateDataFile
        PARAMETERS lcTSName, lcTSFName,lcTSFSize
        lcSQL = "ALTER TABLESPACE " + lcTSName + " ADD DATAFILE '" + lcTSFName + "' SIZE " + ALLTRIM(STR(lcTSFSize)) + " K"
		llRetVal = This.ExecuteTempSPT(lcSQL)
    ENDPROC

    FUNCTION CreateWzTable
        PARAMETERS lcPassed
        LOCAL lcTableName, lcOldDir

        *
        *All tables used internally (except device table) get created here, indexes where possible
        *

        SELECT 0
        IF THIS.NewDir == "" THEN
            THIS.CreateNewDir
        ENDIF
        lcTableName = THIS.UniqueTableName(lcPassed)

        DO CASE
*** DH 2015-09-29: changed size of RmtTblName, RRuleName, ITrigName, DTrigName, and UTrigName from 30 to 128
            CASE lcPassed="Tables"
                CREATE TABLE &lcTableName FREE;
                    (TblName C (128) NOT NULL, ;
                    TblID I, ;
                    CursName C (128) NOT NULL, ;
                    TblPath M, ;
                    RmtTblName C (128) NOT NULL, ;
                    NewTblName C (128), ;
                    Upsizable L, ;
                    PreOpened L, ;
                    EXPORT L, ;
                    Exported L, ;
                    DataErrs N(9), ;
                    DataErrMsg M, ;
                    ErrTable C (128), ;
                    FldsAnald L, ;
                    CDXAnald L, ;
                    ClustName C(30), ;
                    TableSQL M, ;
                    TStampAdd L, ;
                    IdentAdd L, ;
                    LocalRule M ,;
                    RmtRule M, ;
                    RRuleName C (128), ;
                    RuleExport L, ;
                    RuleError M, ;
                    RuleErrNo N (5), ;
                    ItrigName C (128), ;
                    InsertRI M, ;
                    InsertX L, ;
                    DtrigName C (128), ;
                    DeleteRI M, ;
                    DeleteX L, ;
                    UtrigName C (128), ;
                    UpdateRI M, ;
                    UpdateX L, ;
                    RIError M, ;
                    RIErrNo N (5), ;
                    FKeyCrea L, ;
                    PKeyCrea L, ;
                    PkeyExpr M, ;
                    PKTagName C(10), ;
                    TblError M, ;
                    TblErrNo N (5), ;
                    RmtView C(128), ;
                    ViewErr M, ;
                    Type C(1),;&& Add JEI - RKR - 24.03.2005
                    SQLStr M) && Add JEI - RKR - 24.03.2005

                *DataSent indicates if data was successfully appended to the new table
                *Exported indicates if the table was successfully created

                *"RmtView" contains the name of the "SELECT *" view created if the table was part of a view
                *only some of whose tables were upsized; this field maybe ""

                *RRuleName is the name of remote rule

                *"NewTblName" is name of table after renaming if a remote view was created based on the table
                *Index actually created elsewhere for speed reasons
                *INDEX ON TblName TAG TblName

                *Insert, Delete, and Update contain trigger code of the same name
                *InsertX, DeleteX, and UpdateX show whether the triggers were created successfully

                *FKeyCrea and PKeyCrea are used only in the Oracle or SQL 95 case to indicate
                *whether ALTER TABLE statements to add RI succeeded or not

                * JVF 11/02/02 Added column AutoInNext I, AutoInStep I, to account for VFP 8.0 autoinc attrib.
*** DH 12/06/2012: changed size of RmtFldName from 30 to 128
*** DH 09/09/2013: changed size of RmtType from 13 to 20
*** DH 2015-09-29: changed size of RRuleName from 30 to 128
            CASE lcPassed="Fields"
                CREATE TABLE &lcTableName FREE;
                    (TblName C (128) NOT NULL, ;
                    FldName C (128) NOT NULL, ;
                    DATATYPE C (1) NOT NULL, ;
                    ComboType C (20) NOT NULL, ;
                    FullType C (10) NOT NULL, ;
                    LENGTH N (3) NOT NULL, ;
                    PRECISION N (3) NOT NULL, ;
                    NOCPTRANS L, ;
                    lclnull L, ;
                    RmtFldname C (128), ;
                    RmtType C (20), ;
                    RmtLength N (4), ;
                    RmtPrec N (3),;
                    RmtNull L, ;
                    LocalRule M, ;
                    RmtRule M, ;
                    RRuleName C (128), ;
                    RuleExport L, ;
                    RuleError M, ;
                    RuleErrNo N (5), ;
                    DEFAULT M, ;
                    RmtDefault M, ;
                    RDName C (30), ;
                    DefaExport L, ;
                    DefaBound L, ;
                    DefaError M, ;
                    DefaErrNo N (5), ;
                    InCluster L, ;
                    ClustOrder I, ;
                    TblID I, ;
                    AutoInNext I, ;
                    AutoInStep I)
                    
                   

                *Index actually created elsewhere for speed reasons
                *INDEX ON TblName TAG TblName
                *INDEX ON FldName TAG FldName

*** DH 2015-09-29: changed size of RmtTable from 30 to 128
            CASE lcPassed="Indexes"
                CREATE TABLE &lcTableName FREE;
                    (TblID I, ;
                    IndexName C (128) NOT NULL, ;
                    TagName C (10), ;
                    LclExpr M, ;
                    LclIdxType C (12), ;
                    RmtTable C (128), ;
                    RmtName C (10), ;
                    RmtExpr M, ;
                    RmtType C (20), ;
                    Clustered L, ;
                    Exported L, ;
                    TblUpszd L, ;
                    DontCreate L, ;
                    IdxError M, ;
                    IdxErrNo N (5),;
                    IndexSQL M)

                *Value in IndexName is exactly the same as local tablename
                *INDEX ON RTRIM(IndexName)+TagName TAG TblAndTag (performed elsewhere)

            CASE lcPassed="Views"
                CREATE TABLE &lcTableName FREE;
                    (ViewName C (128) NOT NULL,;
                    NewName C (128), ;
                    ViewSQL M, ;
                    RmtViewSQL M, ;
                    TblsUpszd M, ;
                    NotUpszd  M, ;
                    CONNECTION C (128), ;
                    Remotized L, ;
                    ViewErr M,;
                    ViewErrNo N(5) NULL)

            CASE lcPassed="Script"
                CREATE TABLE &lcTableName FREE;
                    (ScriptSQL M)

            CASE lcPassed="Errors"
                CREATE TABLE &lcTableName FREE;
                    (ObjType C(30), ;
                    ObjName M, ;
                    ErrNumber N(5) NULL,;
                    ErrMsg M,;
                    WizErr M,;
                    FailedSQL M)

            CASE lcPassed=MISC_NAME_LOC
                *This table gets created in the analysis database
                CREATE TABLE &lcTableName NAME MISC_NAME_LOC;
                    (SvType C (20), ;
                    DataSourceName M, ;
                    UserConnection C (128), ;
                    ViewConnection C (128), ;
                    DeviceDBName C (30), ;
                    DeviceDBPName C (12), ;
                    DeviceDBSize N (6), ;
                    DeviceDBNumber N (3), ;
                    DeviceLogName C (30), ;
                    DeviceLogPName C (12), ;
                    DeviceLogSize N (6), ;
                    DeviceLogNumber N (3), ;
                    ServerDBName C (30), ;
                    ServerDBSize N (6), ;
                    ServerLogSize N (6), ;
                    SourceDB M , ;
                    ExportIndexes L, ;
                    ExportValidation L, ;
                    ExportRelations L, ;
                    ExportStructureOnly L, ;
                    ExportDefaults L, ;
                    ExportTimeStamp L, ;
                    ExportTableToView L, ;
                    ExportViewToRmt L, ;
                    ExportDRI L, ;
                    ExportSavePwd L, ;
                    DoUpsize L, ;
                    DoScripts L, ;
                    DoReport L)

            CASE lcPassed="Relation"
                CREATE TABLE &lcTableName FREE;
                    (DD_CHIEXPR M,;
                    DD_CHILD C(128),;
                    Dd_RmtChi C(128),;
                    dd_delete C(1),;
                    dd_insert C(1),;
                    DD_PARENT C(128),;
                    Dd_RmtPar C(128),;
                    DD_PAREXPR M,;
                    dd_update C(1),;
                    ClustName C(30),;
                    ClustType C(8),;
                    ClusterSQL M,;
                    HashKeys N(12),;
                    RIError M, ;
                    RIErrNo N (5), ;
                    ClustErr M,;
                    ClustErrNo N (5),;
                    EXPORT L,;
                    Exported L,;
                    Duplicates N(3))

            CASE lcPassed = "Clusters"
                CREATE TABLE &lcTableName FREE;
                    (ClustName C (30),;
                    ClustType C (5),;
                    HashKeys N (6),;
                    ClustSize N (6),;
                    ClusterSQL M,;
                    ClustErr M,;
                    ClustErrNo N (5),;
                    EXPORT L,;
                    Exported L,;
                    Duplicates N(3))

        ENDCASE

        RETURN lcTableName

    ENDFUNC


    PROCEDURE CreateNewDir
        LOCAL aDirArray, lcDirName
        *Create directory for upsizing files if it doesn't exist already
        IF THIS.NewDir=="" THEN
            THIS.NewDir = NEW_DIRNAME_LOC
            DIMENSION aDirArray[1]
            IF ADIR(aDirArray,THIS.NewDir,'D')=0 THEN
                MD (THIS.NewDir)
                THIS.CreatedNewDir=.T.
            ENDIF
            CD (THIS.NewDir)
            SET DEFAULT TO CURDIR()
            THIS.NewDir=CURDIR()
        ENDIF

    ENDPROC


    PROCEDURE GetFoxDataSize
        LOCAL lcEnumTablesTbl, lcFileName, lcTableStem

        lnTableSize = 0
        lnIndexSize = 0
        THIS.TBFoxTableSize = 0
        THIS.TBFoxIndexSize = 0

        lcEnumTablesTbl = THIS.EnumTablesTbl
        IF EMPTY(lcEnumTablesTbl)
            RETURN
        ENDIF

        SELECT (lcEnumTablesTbl)

        SCAN FOR EXPORT=.T.
            * get table + memo size
            lcFileName = &lcEnumTablesTbl..TblPath
            IF ADIR(laDir, lcFileName) <> 0
                lnTableSize = lnTableSize + laDir[1,2] / 1024.

                IF RAT('.', lcFileName) > 0
                    lcTableStem = LEFT(lcFileName, RAT('.', lcFileName)-1)
                ENDIF

                * lcTableStem = this.JustStem2(&lcEnumTablesTbl..TblPath)
                lcFileName = lcTableStem + ".fpt"
                IF ADIR(laDir, lcFileName) <> 0
                    lnTableSize = lnTableSize + laDir[1,2] / 1024.
                ENDIF

                * get index size
                lcFileName = lcTableStem + ".cdx"
                IF ADIR(laDir, lcFileName) <> 0
                    lnIndexSize = lnIndexSize + laDir[1,2] / 1024.
                ENDIF
            ELSE
                * die
            ENDIF
        ENDSCAN

        THIS.TBFoxTableSize = lnTableSize
        THIS.TBFoxIndexSize = lnIndexSize

    ENDPROC

    PROCEDURE InsArrayRow
        LPARAMETERS aArray, lcElement1, lcElement2, lcElement3, lcElement4, lcElement5, lcElement6F
        LOCAL lnParams, lnRow

        lnParams = PARAMETERS() - 1

        IF ALEN (aArray, 1) = 1 AND EMPTY(aArray[1,1])
            lnRow = 1
        ELSE
            DIMENSION aArray[ALEN(aArray,1)+1, lnParams]
            lnRow = ALEN(aArray, 1)
        ENDIF

        IF lnParams >= 2
            aArray[lnRow,1] = lcElement1
        ENDIF
        IF lnParams >= 2
            aArray[lnRow,2] = lcElement2
        ENDIF
        IF lnParams >= 3
            aArray[lnRow,3] = lcElement3
        ENDIF
        IF lnParams >= 4
            aArray[lnRow,4] = lcElement4
        ENDIF
        IF lnParams >= 5
            aArray[lnRow,5] = lcElement5
        ENDIF
        IF lnParams >= 6
            aArray[lnRow,6] = lcElement6
        ENDIF

    ENDPROC

    PROCEDURE DEBUG
        ACTIVATE WINDOW trace
        ACTIVATE WINDOW DEBUG
        SUSPEND
    ENDPROC

    *set trunc.log on chkpt. option for database on server
    PROCEDURE TruncLogOn
        LOCAL lnOldArea, lcDBName, lnRes

* If we have an extension object and it has a TruncLogOn method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'TruncLogOn', 5) and ;
			not This.oExtension.TruncLogOn(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        lcDBName = ALLTRIM(THIS.ServerDBName)
        lnOldArea = SELECT()
        IF (SQLEXEC(THIS.MasterConnHand, "sp_helpdb ") = 1) AND ;
        	!EMPTY(ALIAS()) and type('STATUS') <> 'U'
            LOCATE FOR NAME = lcDBName
            IF !EOF()
                THIS.TruncLog = IIF(ATC("trunc. log", STATUS) > 0 or ;
                	ATC("trunc. log", strconv(STATUS, 6)) > 0 or ;
                	ATC("Recovery=SIMPLE", strconv(STATUS, 6)) > 0, 1, 0)
            ENDIF
            USE
        ENDIF
        lnRes = SQLEXEC(THIS.MasterConnHand, "sp_dboption " + lcDBName + ", 'trunc. log on chkpt.', true") = 1
        SELECT (lnOldArea)
    ENDPROC

    PROCEDURE TruncLogOff
        LOCAL lnRes

* If we have an extension object and it has a TruncLogOff method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'TruncLogOff', 5) and ;
			not This.oExtension.TruncLogOff(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        * if false or unknown, set to false
        IF THIS.TruncLog <> 1
            lnRes = SQLEXEC(THIS.MasterConnHand, "sp_dboption " + ALLTRIM(THIS.ServerDBName) + ", 'trunc. log on chkpt.', false")
        ENDIF
        THIS.TruncLog = -1

    ENDPROC

    PROCEDURE UpsizeComplete
        LOCAL lcEnumTablesTbl, lcCRLF, myarray[1]

* If we have an extension object and it has a UpsizeComplete method, call it.
* If it returns .F., that means don't continue with the usual processing.

		if vartype(This.oExtension) = 'O' and ;
			pemstatus(This.oExtension, 'UpsizeComplete', 5) and ;
			not This.oExtension.UpsizeComplete(This)
			return
		endif vartype(This.oExtension) = 'O' ...

        THIS.cFinishMsg = ALL_DONE_LOC

        lcCRLF = CHR(10) + CHR(13)
        lcEnumTablesTbl	= THIS.EnumTablesTbl
*** DH 2015-08-18: don't select the Tables table since BuildReport may have closed it.
***		The SQL statement will cause it to be reopened if necessary.
***        SELECT(lcEnumTablesTbl)
		SELECT COUNT(*) FROM (lcEnumTablesTbl) WHERE EXPORT and not Exported INTO ARRAY myarray
		IF (This.DoScripts AND This.DoUpsize) OR This.DoUpsize
            IF !EMPTY(myarray)
                THIS.cFinishMsg = ALL_DONE_LOC + lcCRLF + CANTUPSIZE_TABLE_LOC
                SCAN FOR Exported = .F.
                    THIS.cFinishMsg = THIS.cFinishMsg + RTRIM(LOWER(&lcEnumTablesTbl..TblName)) + ", "
                ENDSCAN
                THIS.cFinishMsg = LEFT(THIS.cFinishMsg, LEN(RTRIM(THIS.cFinishMsg)) - 1)
            ENDIF
        ENDIF
    ENDPROC

    PROCEDURE StripFunction
        LPARAMETER tcString

        LOCAL lnPos, lcString, lcDelim, lnOpenDelimOccurs, lnEndPos

        * Chandsr added code for 40543
        && stripping the fox function names from the expression before sending them over to SQL
        lnOpenDelimOccurs = OCCURS("(", tcString)

        lnPos = AT("(", tcString, IIF (lnOpenDelimOccurs = 0, 1, lnOpenDelimOccurs) )
        * End chandsr added code for 40543

        lcString = tcString

        lcDelim = IIF(AT(",",tcString) > 0, ",", ")" )
        IF lnPos > 0
        	*{ Add and Change JEI RKR 2005.03.30
			lnEndPos = AT(lcDelim, tcString) - 1 - lnPos
			IF lnEndPos > 0
				lcString = SUBSTR(tcString, lnPos + 1, lnEndPos)
			ELSE
				lcString = SUBSTR(tcString, lnPos + 1)
			ENDIF
*			lcString = SUBSTR(tcString, lnPos + 1, AT(lcDelim, tcString) - 1 - lnPos)        	
        	*} Add JEI RKR 2005.03.30
		ELSE
        	*{ Add and Change JEI RKR 2005.03.30
        	lnEndPos = AT(lcDelim, tcString) - 1
        	IF lnEndPos > 0
        		lcString = ALLTRIM(LEFT(tcString,lnEndPos))
        	ENDIF
        	*} Add JEI RKR 2005.03.30		
        ENDIF

        RETURN lcString
    ENDPROC

	&& chandsr added function to make the string ANSI 92 compatible
    PROCEDURE MakeStringANSI92Compatible
        LPARAMETER	cSQLString
        LOCAL cRet, lnBetweenLocation, lnComma, lcBetweenExpression, lcBetweenStart, lcBetweenEnd, lnBetweenBegin, lnBetweenEnd
        LOCAL lcBetweenColumnName

        cRet = STRTRAN (cSQLString, BOOLEAN_FALSE, BOOLEAN_SQL_FALSE,1,1, 1)
        cRet = STRTRAN (cRet, BOOLEAN_TRUE, BOOLEAN_SQL_TRUE, 1, 1, 1)

        lnBetweenLocation = ATC (VFP_BETWEEN, cRet)

        IF (lnBetweenLocation <> 0) THEN
            lcBetweenExpression = SUBSTR (cRet, lnBetweenLocation)
            lnBetweenStart = ATC ("(", lcBetweenExpression)
            lnBetweenEnd = ATC (")", lcBetweenExpression)
            lcBetweenExpression = SUBSTR (cRet, lnBetweenLocation, lnBetweenEnd)

            lnComma = ATC(VFP_COMMA, lcBetweenExpression)
            lcBetweenColumnName = SUBSTR (lcBetweenExpression, lnBetweenStart + 1, lnComma - lnBetweenStart - 1)

            lnBetweenStart = lnComma + 1
            lnComma = ATC(VFP_COMMA, lcBetweenExpression, 2)
            lcBetweenStart = SUBSTR (lcBetweenExpression, lnBetweenStart, lnComma - lnBetweenStart)
            lcBetweenEnd = SUBSTR (lcBetweenExpression, lnComma + 1)
            cRet = STRTRAN (cRet, lcBetweenExpression, lcBetweenColumnName + " BETWEEN " + THIS.ConvertToSQLType (lcBetweenStart) + " AND " + THIS.ConvertToSQLType (lcBetweenEnd), 1, 1, 1)
        ENDIF
        RETURN cRet
    ENDPROC

    PROCEDURE ConvertToSQLType
        LPARAMETERS cData
        LOCAL lcLeadingChar, lcData

        lcData = ALLTRIM (cData)
        lcLeadingChar = SUBSTR(lcData, 1, 1)

        IF (lcLeadingChar = "{") THEN
            lcData = STRTRAN (lcData, "{^", "'")
            lcData = STRTRAN (lcData, "})", "'")
            lcData = STRTRAN (lcData, "}", "'")
        ENDIF

        RETURN lcData
    ENDPROC
	&& chandsr added function to make the string ANSI 92 compatible


	*ADD JEI - RKR - 2005.03.24
	PROCEDURE ReadViews
		LOCAL lcSource, laViews[1,1], lcSQLStr As String
		LOCAL lnOldArea, lnViewCount, lnCurrentView as Integer
		
		
        lcSourceDB=THIS.SourceDB
        lnOldArea=SELECT()
        SET DATABASE TO (lcSourceDB)
		lnViewCount = ADBOBJECTS(laViews,"VIEW")
		FOR lnCurrentView = 1 TO lnViewCount 
			*If this is a remote view, skip it
			IF DBGETPROP(laViews[lnCurrentView],"View","SourceType")=2 THEN
				LOOP
			ENDIF
            lcSQLStr = DBGETPROP(laViews[lnCurrentView],"View","SQL")
            INSERT INTO (This.EnumTablesTbl) (TblName, Type, SQlStr,  EXPORT) ;
            			VALUES (LOWER(laViews[lnCurrentView]), "V", lcSQLStr, .t.)
            
		ENDFOR
		SELECT(This.EnumTablesTbl)
		LOCATE
		SELECT(lnOldArea)
	ENDPROC
	
	PROCEDURE GetChooseView
	
		LOCAL lnOldWorkArea as Integer
		
		DIMENSION This.aChooseViews[1,1]
		
		This.aChooseViews[1,1] = ""
		lnOldWorkArea = SELECT()
		
		SELECT TblName FROM (This.EnumTablesTbl) WHERE Type = "V" AND EXPORT = .T. INTO ARRAY This.aChooseViews
		
		SELECT(This.EnumTablesTbl)
		Replace EXPORT WITH .F. ALL FOR Type == "V"
		DELETE ALL FOR Type == "V"

		LOCATE
		SELECT(lnOldWorkArea)
	ENDPROC
	
	PROCEDURE CheckUpsizeView
		LPARAMETERS tcViewName

		RETURN (ASCAN(This.aChooseViews, tcViewName,1,ALEN(This.aChooseViews),1,15) > 0)
	ENDPROC
	
	&& JEI RKR Add Procedure 
	PROCEDURE CheckForLocalServer
		LOCAL lcTemDBName, lcTempDBFileName, lcCurrentPath, lcSQLStr, laTemp[1,1] as String
		LOCAL lnResult as Integer
		LOCAL llServerIsLocal as Logical
		
		llServerIsLocal = .f.
		
		lcCurrentPath = FULLPATH(".")
		lcTemDBName = SYS(2015)
		lcTempDBFileName = ADDBS(lcCurrentPath) + FORCEEXT(lcTemDBName, "mdf")
		lcSQLStr = "CREATE DATABASE " + lcTemDBName + " ON (NAME = " + lcTemDBName + "_DAT ," + ;
					"FILENAME = '" +lcTempDBFileName + "')"
				
		lnResult = SQLEXEC(This.MasterConnHand, lcSQLStr)
		IF lnResult = 1
			llServerIsLocal = (ADIR(laTemp, lcTempDBFileName) = 1)
			=SQLEXEC(This.MasterConnHand, "DROP DATABASE " + lcTemDBName)
		ENDIF

		This.ServerISLocal = llServerIsLocal

		RETURN llServerIsLocal
	ENDPROC
	
	PROCEDURE CheckForNullValuesInTable
		LPARAMETERS tcTablaName as String
		
		LOCAL lcEnumFieldsTbl, lcTempCursorName, lcNullExistStr, lcEmptyDateExist, lcNullExistOR, lcEmptyDateOR as String
		LOCAL lnOldWorkArea as Integer
		LOCAL llRes as Logical
		LOCAL loException as Exception
		
		lnOldWorkArea = SELECT()
		tcTablaName = UPPER(ALLTRIM(tcTablaName))
		lcEnumFieldsTbl = RTRIM(THIS.EnumFieldsTbl)
		lcTempCursorName = This.UniqueCursorName("CheckForNull")
		lcNullExistStr = "LOCATE FOR "
		lcEmptyDateExist = "LOCATE FOR "
		lcNullExistOR = ""
		lcEmptyDateOR = ""
		llRes = .T.
		
		SELECT * FROM &lcEnumFieldsTbl. WHERE upper(ALLTRIM(TblName)) == tcTablaName INTO CURSOR &lcTempCursorName.
		SELECT(lcTempCursorName)
		SCAN
			lcNullExistStr = lcNullExistStr + lcNullExistOR + " ISNULL(" + ALLTRIM(FldName) + ")" 
			lcNullExistOR = " OR "
			IF ALLTRIM(UPPER(DATATYPE)) $ "DT"
				lcEmptyDateExist = lcEmptyDateExist + lcEmptyDateOR + " EMPTY(" + ALLTRIM(FldName) + ")"
				lcEmptyDateOR = " OR "
			ENDIF
		ENDSCAN
		IF USED(tcTablaName)
			SELECT(tcTablaName)
			TRY
				&lcNullExistStr.
				IF !FOUND()
					IF !EMPTY(lcEmptyDateOR)
						&lcEmptyDateExist.
						IF !FOUND()
							llRes = .F.
						ENDIF
					else
						llRes = .F.
					ENDIF
				ENDIF
			CATCH TO loException
				* Do Nothing
			ENDTRY
		ENDIF
		USE IN (lcTempCursorName)
		SELECT(lnOldWorkArea)
		RETURN llRes
	ENDPROC
	
	PROCEDURE GetColumnNum
		LPARAMETERS tcTableName as String, tcDataTypes as String, tlAllowNull as Logical
		
		LOCAL lnOldWorkArea, lnFieldCount, lnCurrentField as Integer
		LOCAL laDummy[1], lcRes as String
		
		IF VARTYPE(tcDataTypes) <> "C" OR EMPTY(tcDataTypes)
			RETURN ""
		ENDIF
		
		lcRes = ","
		IF USED(tcTableName)
			lnOldWorkArea = SELECT()
			SELECT(tcTableName)
			lnFieldCount = AFIELDS(laDummy)
			FOR lnCurrentField = 1 TO lnFieldCount 
*** DH 07/02/2013: use all DT fields whether they're nullable or not since they may contain
***					blank values
*				IF laDummy[lnCurrentField,2] $ tcDataTypes AND laDummy[lnCurrentField,5] = tlAllowNull
				IF laDummy[lnCurrentField,2] $ tcDataTypes AND ;
					(laDummy[lnCurrentField,5] = tlAllowNull or laDummy[lnCurrentField,2] $ 'DT')
					lcRes = lcRes + TRANSFORM(lnCurrentField) + ","
				ENDIF
			ENDFOR
			SELECT(lnOldWorkArea)
		ENDIF
		
		RETURN lcRes
	ENDPROC
ENDDEFINE