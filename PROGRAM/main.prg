lparameters tcSourceDatabase, ;
	tcTargetDatabase, ;
	tlNewDatabase
loForm = newobject('UpsizingWizardForm', 'lib\upswiz.vcx')
if vartype(loForm) = 'O'
	do case
		case vartype(tcSourceDatabase) = 'C' and ;
			not empty(tcSourceDatabase) and file(forceext(tcSourceDatabase, 'DBC'))
			loForm.SetSourceDatabase(forceext(tcSourceDatabase, 'DBC'))
		case not empty(dbc())
			loForm.SetSourceDatabase(dbc())
	endcase
	do case
		case vartype(tcTargetDatabase) <> 'C' or empty(tcTargetDatabase)
			loForm.nDatabaseType = 2
		case tlNewDatabase
			loForm.cNewTargetDatabase = tcTargetDatabase
			loForm.nDatabaseType      = 2
		otherwise
			loForm.cTargetDatabase = tcTargetDatabase
			loForm.nDatabaseType   = 1
	endcase
	loForm.oEngine.cReturnToProc = 'MAIN'
	loForm.Show()
endif vartype(loForm) = 'O'
