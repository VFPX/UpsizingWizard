*==============================================================================
* Function:			BulkXMLLoad
* Purpose:			Performs a SQL Server bulk XML load
* Author:			Doug Hennig
* Last revision:	02/01/2019
* Parameters:		tcAlias      - the alias of the cursor to export
*					tcTable      - the name of the table to import into
*					ttBlank      - the value to use for blank DateTime values
*						(optional: if it isn't specified, 01/01/1900 is used)
*					tcDatabase   - the database the table belongs to
*					tcServer     - the SQL Server name
*					tcUserName   - the user name for the connection (optional:
*						if it isn't specified, Windows Integrated Security is
*						used)
*					tcPassword   - the password for the connection (optional:
*						if it isn't specified, Windows Integrated Security is
*						used)
*					tcTempFolder - a temporary file path. It must be a shared
*						location accessible to the service account of the
*						target instance of SQL Server and to the account
*						running the bulk load application. Unless you are bulk
*						loading on a local server, the temporary file path must
*						be a UNC path (such as \\servername\sharename). If
*						specified, transactional processing will be enabled for
*						the XML bulk load
* Returns:			an empty string if the bulk load succeeded or the text of
*						an error message if it failed
* Environment in:	the alias specified in tcAlias must be open
*					the specified table and database must exist
*					the specified server must be accessible
*					there must be enough disk space for the XML files
* Environment out:	if an empty string is returned, the data was imported into
*						the specified table
* Notes:			These are comments from Mike Potjer about changes made
*						02/22/2017:
*					BulkXMLLoad couldn't be used for most of our larger tables,
*					because they have character fields that often contain blank
*					values. I found that, by default, SQLXMLBulkLoad wants to
*					store a blank character value as NULL, but if the SQL Server
*					table does not allow NULL values in those fields, the XML
*					bulk load will fail. In researching this, I found out that
*					setting the Transaction property of SQLXMLBulkLoad to true
*					would allow the bulk insert to succeed with blank character
*					fields. I never found an explanation for why that works, but
*					apparently when Transaction is set, a temporary text file is
*					created and imported via a T-SQL BULK INSERT. Some of the
*					information I found also recommended setting the
*					SQLXMLBulkLoad TempFilePath property when Transaction is
*					set.
*
*					There is also a block of code that strips out characters
*					that are illegal according to the XML 1.0 specification.
*					The reason I added that is because when I was testing, the
*					XML bulk load choked on a table where a CHR(2) had somehow
*					gotten into a field of user-entered data.
*
*					Change made 2019-01-24: to handle large tables, XML output
*					is now done in batches of up to 500 MB of records at a
*					time.
*==============================================================================

lparameters tcAlias, ;
	tcTable, ;
	ttBlank, ;
	tcDatabase, ;
	tcServer, ;
	tcUserName, ;
	tcPassword, ;
	tcTempFolder
local lnSelect, ;
	lcAlias, ;
	laFields[1], ;
	lnFields, ;
	lnPos, ;
	ltBlank, ;
	lnI, ;
	lcField, ;
	llClose, ;
	lcType, ;
	lcReplaceCommand, ;
	lcSchema, ;
	lcData, ;
	lnRecords, ;
	lnRecsProcessed, ;
	lcReturn, ;
	loException as Exception, ;
	lcXSD, ;
	loBulkLoad
#include ..\Include\AllDefs.H

* These characters are illegal according to the XML 1.0 specification.

#define XML_ILLEGAL_CHARS	'chr(0)+chr(1)+chr(2)+chr(3)+chr(4)+chr(5)+' + ;
							'chr(6)+chr(7)+chr(8)+chr(11)+chr(12)+chr(14)+' + ;
							'chr(15)+chr(16)+chr(17)+chr(18)+chr(19)+' + ;
							'chr(20)+chr(21)+chr(22)+chr(23)+chr(24)+' + ;
							'chr(25)+chr(26)+chr(27)+chr(28)+chr(29)+' + ;
							'chr(30)+chr(31)'

* If there are any date fields in the selected cursor and we're not supposed to
* upsize blank dates as NULL, create a cursor from it with the appropriate
* value for blank dates.

lnSelect = select()
lcAlias  = tcAlias
select (tcAlias)
lnFields = afields(laFields)
if not isnull(ttBlank)
	lnPos = ascan(laFields, 'D', -1, -1, 2, 15)
	if lnPos = 0
		lnPos = ascan(laFields, 'T', -1, -1, 2, 15)
	endif lnPos = 0
	if lnPos > 0
		lcAlias = sys(2015)
		select * from (tcAlias) into cursor (lcAlias) readwrite
		ltBlank = iif(vartype(ttBlank) $ 'TD', ttBlank, SQL_SERVER_EMPTY_DATE_Y2K)
		for lnI = 1 to lnFields
			if laFields[lnI, 2] $ 'TD'
				lcField = laFields[lnI, 1]
				replace (lcField) with ltBlank for empty(&lcField)
			endif laFields[lnI, 2] $ 'TD'
		next lnI
		llClose = .T.
	endif lnPos > 0
endif not isnull(ttBlank)

* If the current table contains any string fields, strip out illegal XML
* characters.

if ascan(laFields, 'C', -1, -1, 2, 7) > 0 or ;
	ascan(laFields, 'M', -1, -1, 2, 7) > 0 or ;
	ascan(laFields, 'V', -1, -1, 2, 7) > 0

* If a read-write cursor wasn't created earlier for blank date processing, do
* it now.

	if not llClose
		lcAlias = sys(2015)
		select * from (tcAlias) into cursor (lcAlias) readwrite
		llClose = .T.
	endif not llClose

* Loop through the non-NOCPTRANS string fields and strip out any illegal
* characters. This assumes that those characters were not entered
* intentionally and that we do not need to substitute something else for the
* characters being removed. Note: the reason "(lcField)" is used is to prevent
* an error if the field name is a reserved word like "FROM".

	for lnI = 1 to lnFields
		lcField = laFields[lnI, 1]
		lcType  = laFields[lnI, 2]
		if inlist(lcType, 'C', 'M', 'V') and not laFields[lnI, 6]
			lcReplaceCommand = 'REPLACE (lcField) WITH CHRTRAN(' + ;
				lcField + ', ' + XML_ILLEGAL_CHARS + ;
				", '') FOR NOT EMPTY(" + ;
				iif(lcType = 'M', 'RTRIM(' + lcField + ')', lcField) + ')'
try
			&lcReplaceCommand
catch to loException
messagebox(tcAlias + chr(13) + lcReplaceCommand)
set step on 
endtry
		endif inlist(lcType, 'C', 'M', 'V') ...
	next lnI
endif ascan(laFields ...

* We don't want the XML files to exceed 500 MB (the files can be up to 2 GB but
* XML is verbose so we have to account for that) so we may have to process the
* table in batches.

lnRecords       = ceiling(500000000/recsize())
lnRecsProcessed = 0

* Create the XML data and schema files.

lcSchema = forceext(tcTable, 'xsd')
lcData   = forceext(tcTable, 'xml')
lcReturn = ''
go top
do while lnRecsProcessed < reccount()
	try
		cursortoxml(alias(), lcData, 1, 512 + 8, lnRecords, lcSchema)
	catch to loException
		lcReturn = loException.Message
	endtry

* Convert the XSD into a format acceptable by SQL Server. Add the SQL
* namespace, convert the <xsd:choice> start and end tags to <xsd:sequence>,
* use the sql:datatype attribute for DateTime fields, and specify the table
* imported into with the sql:relation attribute. Note: although the XSD content
* doesn't change, the file has to be recreated on each pass or the bulk XML
* load won't import any records on the second and subsequent passes.

	if empty(lcReturn)
		lcXSD = filetostr(lcSchema)
		lcXSD = strtran(lcXSD, ':xml-msdata">', ;
			':xml-msdata" xmlns:sql="urn:schemas-microsoft-com:mapping-schema">')
		lcXSD = strtran(lcXSD, 'IsDataSet="true">', ;
			'IsDataSet="true" sql:is-constant="1">')
		lcXSD = strtran(lcXSD, '<xsd:choice maxOccurs="unbounded">', ;
			'<xsd:sequence>')
		lcXSD = strtran(lcXSD, '</xsd:choice>', ;
			'</xsd:sequence>')
		lcXSD = strtran(lcXSD, 'type="xsd:dateTime"', ;
			'type="xsd:dateTime" sql:datatype="dateTime"')
		lcXSD = strtran(lcXSD, 'minOccurs="0"', ;
			'sql:relation="' + lower(tcTable) + '" minOccurs="0"')
		strtofile(lcXSD, lcSchema)

* Instantiate the SQLXMLBulkLoad object and set its ConnectionString and other
* properties. Note: we can set the ErrorLogFile property to the name of a file
* to write import errors to; that isn't done here.

		try
			loBulkLoad   = createobject('SQLXMLBulkLoad.SQLXMLBulkload.4.0')
			lcConnString = 'Provider=SQLOLEDB.1;Initial Catalog=' + tcDatabase + ;
				';Data Source=' + tcServer + ';Persist Security Info=False;'
			if empty(tcUserName)
				lcConnString = lcConnString + 'Integrated Security=SSPI'
			else
				lcConnString = lcConnString + 'User ID=' + tcUserName + ;
					';Password=' + tcPassword
			endif empty(tcUserName)
			loBulkLoad.ConnectionString = lcConnString
			loBulkLoad.KeepNulls        = .T.
			loBulkLoad.ForceTableLock   = .T.

* Turn on transaction processing for the bulk load. This allows the XML bulk
* load to succeed when you have blank string fields and do not allow NULL
* values. I have not been able to find any explanation WHY transaction
* processing allows blank strings when non-transaction processing forces NULL
* values to be saved. If a temporary folder was specified, use it; otherwise
* the location specified in the TEMP environment variable is used.

			loBulkLoad.Transaction = .T.
			if not empty(tcTempFolder)
				loBulkLoad.TempFilePath = tcTempFolder
			endif not empty(tcTempFolder)

* Call Execute to perform the bulk import.

			loBulkLoad.Execute(lcSchema, lcData)
		catch to loException
			lcReturn = loException.Message
messagebox(lcReturn)
set step on 
		endtry
	endif empty(lcReturn)
	lnRecsProcessed = lnRecsProcessed + lnRecords
	if not eof()
		skip
	endif not eof()
	if not empty(lcReturn)
		exit
	endif not empty(lcReturn)
enddo while lnRecsProcessed < reccount()

* Clean up. If we created a cursor, close it.

if llClose
	use
endif llClose
erase (lcSchema)
erase (lcData)
select (lnSelect)
return lcReturn
