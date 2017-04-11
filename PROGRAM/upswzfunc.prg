*-------------------------------------------------------
* Function....: AODBCDataSources()
* Called by...:
*
* Abstract....:
*
* Returns.....:
*
* Parameters..:
*
* Notes.......:
*-------------------------------------------------------
FUNCTION AODBCDataSources
	LPARAMETERS taDSN

	LOCAL lnDataSourcesNum, lnRetVal, lnODBCEnv, dsn, dsndesc, mdsn, mdesc, lcServerType

	DECLARE short SQLDataSources IN odbc32 ;
					Integer henv,Integer fDir,String @ DSN,;
					Integer DSNMax, Integer @pcbDSN,String @Description,;
					Integer DescMax,Integer @desclen
					
	*Do until all the data source names have been retrieved
	lnRetVal = 0
	lnODBCEnv = VAL(SYS(3053))
	lnDataSourcesNum = 0
	DIMENSION taDSN[1,2]

	taDSN[1,2] = .Null.

	DO WHILE lnRetVal = 0 && SUCCESS
		dsn = space(100)
		dsndesc = space(100)
		mdsn = 0
		mdesc = 0
		lcServerType = ""
		
		lnRetVal = sqldatasources(lnODBCEnv, 1 ,; &&SQL_FETCH_NEXT
							 @dsn, 100, @mdsn, @dsndesc, 100, @mdesc)

		IF lnRetVal = 0 THEN &&if no error occurred
			lnDataSourcesNum = lnDataSourcesNum + 1
			DIMENSION taDSN[lnDataSourcesNum ,2]
			taDSN[lnDataSourcesNum ,1] = LEFT(dsn, AT(CHR(0), dsn)-1)
			taDSN[lnDataSourcesNum ,2] = LEFT(dsndesc, AT(CHR(0), dsndesc)-1)
		ENDIF
	ENDDO
	
	RETURN lnDataSourcesNum
ENDFUNC