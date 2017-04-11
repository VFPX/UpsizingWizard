# Upsizing Wizard

**Project Manager**: [Doug Hennig](mailto:dhennig@stonefield.com)

This is an update to the Visual FoxPro 9.0 SP2 Upsizing Wizard. To launch the wizard run UpsizingWizard.app.

This Sedna update release includes:

*  Updated, cleaner look and feel

*  Streamlined, simpler steps

*  Support for bulk insert to improve performance.

*  Allows you to specify the connection as a DBC, a DSN, one of the existing connections or a new connection string.

*  Fields using names with reserved SQL keywords are now delimited.

*  If lQuiet is set to true when calling the wizard, no UI is displayed. It uses RAISEEVENT() during the progress of the upsizing so the caller can show progress.

*  Performance improvement when upsizing to Microsoft SQL Server 2005 and later.

*  Trims all Character fields being upsized to Varchar. 

*  BlankDateValue property available. It specifies that blank dates should be upsized as nulls. (Old behavior was to set them to 01/01/1900).

*  Support for an extension object. This allows developers to hook into every step of the upsizing process and change the behavior. Another way is to subclass the engine.

*  Support for table names with spaces.

*  UpsizingWizard.APP can be started with default settings (via params) for source name and path, target db, and a Boolean indicating if the target database is to be created.

The Upsizing Wizard is part of Sedna, a collection of libraries, samples and add-ons to Visual FoxPro 9.0 SP2.

**2017.02.22 Release**  
This update implements better support for BulkXMLLoad by Mike Potjer.

**2015.12.01 Release**  
This update implements a couple of pathing issue fixes by Thierry Nivelet.

**2015.09.28 Release**  
This update fixes several issues found by Matt Slay and Jim Nelson: the Upsizing Wizard didn't properly handle remote table and field names delimited with square brackets, nor did it automatically delimit names using reserved words such as "order". It also didn't handle the rare case where TypeMap.dbf was missing.

**2015.08.20 Release**  
This update fixes another issue found by Jim Nelson: the settings of CURSOSETPROP("FetchMemo"), SET PROCEDURE, and SET PATH weren't restored after running the Upsizing Wizard.

**2015.08.18 Release**  
This update fixes a few bugs found by Jim Nelson: an error when BlankDateValue is set to NULL and NULLDISPLAY is blank, an error when SourceDB is a relative rather than full path, and a warning message when copying files for reporting purposes.

**2015.01.13 Release**  
This update fixes a bug found by Matt Slay that caused the precision (decimals) of numeric fields to be incorrect.

**2013.11.25 Release**  
This update removes the old license and readme files, which are no longer applicable.

**2013.11.20 Release**  
This update fixes a bug, discovered and fixed by Jon Love, that caused a "variable not found" error in the BulkInsert method.

**2013.07.24 Release**  
This update has the following bug fixes (thanks to Mike Potjer for finding and even fixing some of these issues):

* Upsizing a logical field to a bit field no longer causes an error when the table has a lot of records.

* You no longer get an error upsizing tables that have field rules, table rules, or triggers.

* Quotes in the content of fields are preserved.

* You no longer get a warning that 6.5 compatibility cannot be used.

**2013.07.02 Release**  
This update has the following bug fixes (thanks to Mike Potjer for helping with these issues):

* Under some conditions, an error occurred while trying to clean up (specifically removing a working directory the wizard created, which may not be empty) after the wizard is done. This code is now wrapped in a TRY to avoid the error.

* A temporary DBF file wasn't deleted but its FPT was, causing an error if you tried to open it. The file is now deleted.

* A temporary file used for bulk upload wasn't erased when the wizard is done with it. It is now.

* One of the export mechanisms ("FastExport") used null dates rather than the defined date for blank VFP date fields. This was fixed.

* You are no longer told that the bulk insert mechanism failed when it in fact succeeds. This also resolves a problem with duplicate records being created.

* Under some conditions, the last few records in a VFP table weren't imported into the SQL Server table. This was fixed.

**2012.12.06 Release**  
This update has the following bug fixes:

* Handles converting Memo to Varchar(Max)

* You can now change the date used in SQL Server for empty VFP dates in one place: SQL_SERVER_EMPTY_DATE_Y2K in AllDefs.H

* One of the export mechanisms ("JimExport") used null dates rather than the defined date for blank VFP date fields. This was fixed.

* The progress meter now correctly shows the progress of the sending data for each table.

* A bug that sometimes caused an error sending the last bit of data for a table to SQL Server was fixed.

* You no longer get a "string too long" error when using bulk insert with a record with more than 30 MB of data.

* You can now use field names up to 128 characters; the former limit was 30.

Review article on CODE Magazine: [http://www.code-magazine.com/Article.aspx?quickid=0703052](http://www.code-magazine.com/Article.aspx?quickid=0703052)
