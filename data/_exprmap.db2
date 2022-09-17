*--------------------------------------------------------------------------------------------------------------------------------------------------------
* (ES) AUTOGENERADO - ��ATENCI�N!! - ��NO PENSADO PARA EJECUTAR!! USAR SOLAMENTE PARA INTEGRAR CAMBIOS Y ALMACENAR CON HERRAMIENTAS SCM!!
* (EN) AUTOGENERATED - ATTENTION!! - NOT INTENDED FOR EXECUTION!! USE ONLY FOR MERGING CHANGES AND STORING WITH SCM TOOLS!!
*--------------------------------------------------------------------------------------------------------------------------------------------------------
*< FOXBIN2PRG: Version="1.19" SourceFile="_exprmap.dbf" /> (Solo para binarios VFP 9 / Only for VFP 9 binaries)
*


<TABLE>
	<MemoFile></MemoFile>
	<CodePage>1252</CodePage>
	<LastUpdate></LastUpdate>
	<Database></Database>
	<FileType>0x00000003</FileType>
	<FileType_Descrip>FoxBASE+ / FoxPro /dBase III PLUS / dBase IV, no memo</FileType_Descrip>

	<FIELDS>
		<FIELD>
			<Name>FOXEXPR</Name>
			<Type>C</Type>
			<Width>10</Width>
			<Decimals>0</Decimals>
			<Null>.F.</Null>
			<NoCPTran>.F.</NoCPTran>
			<Field_Valid_Exp></Field_Valid_Exp>
			<Field_Valid_Text></Field_Valid_Text>
			<Field_Default_Value></Field_Default_Value>
			<Table_Valid_Exp></Table_Valid_Exp>
			<Table_Valid_Text></Table_Valid_Text>
			<LongTableName></LongTableName>
			<Ins_Trig_Exp></Ins_Trig_Exp>
			<Upd_Trig_Exp></Upd_Trig_Exp>
			<Del_Trig_Exp></Del_Trig_Exp>
			<TableComment></TableComment>
			<Autoinc_Nextval>0</Autoinc_Nextval>
			<Autoinc_Step>0</Autoinc_Step>
		</FIELD>
		<FIELD>
			<Name>SQLSERVER</Name>
			<Type>C</Type>
			<Width>20</Width>
			<Decimals>0</Decimals>
			<Null>.F.</Null>
			<NoCPTran>.F.</NoCPTran>
			<Field_Valid_Exp></Field_Valid_Exp>
			<Field_Valid_Text></Field_Valid_Text>
			<Field_Default_Value></Field_Default_Value>
			<Table_Valid_Exp></Table_Valid_Exp>
			<Table_Valid_Text></Table_Valid_Text>
			<LongTableName></LongTableName>
			<Ins_Trig_Exp></Ins_Trig_Exp>
			<Upd_Trig_Exp></Upd_Trig_Exp>
			<Del_Trig_Exp></Del_Trig_Exp>
			<TableComment></TableComment>
			<Autoinc_Nextval>0</Autoinc_Nextval>
			<Autoinc_Step>0</Autoinc_Step>
		</FIELD>
		<FIELD>
			<Name>ORACLE</Name>
			<Type>C</Type>
			<Width>20</Width>
			<Decimals>0</Decimals>
			<Null>.F.</Null>
			<NoCPTran>.F.</NoCPTran>
			<Field_Valid_Exp></Field_Valid_Exp>
			<Field_Valid_Text></Field_Valid_Text>
			<Field_Default_Value></Field_Default_Value>
			<Table_Valid_Exp></Table_Valid_Exp>
			<Table_Valid_Text></Table_Valid_Text>
			<LongTableName></LongTableName>
			<Ins_Trig_Exp></Ins_Trig_Exp>
			<Upd_Trig_Exp></Upd_Trig_Exp>
			<Del_Trig_Exp></Del_Trig_Exp>
			<TableComment></TableComment>
			<Autoinc_Nextval>0</Autoinc_Nextval>
			<Autoinc_Step>0</Autoinc_Step>
		</FIELD>
		<FIELD>
			<Name>PAD</Name>
			<Type>L</Type>
			<Width>1</Width>
			<Decimals>0</Decimals>
			<Null>.F.</Null>
			<NoCPTran>.F.</NoCPTran>
			<Field_Valid_Exp></Field_Valid_Exp>
			<Field_Valid_Text></Field_Valid_Text>
			<Field_Default_Value></Field_Default_Value>
			<Table_Valid_Exp></Table_Valid_Exp>
			<Table_Valid_Text></Table_Valid_Text>
			<LongTableName></LongTableName>
			<Ins_Trig_Exp></Ins_Trig_Exp>
			<Upd_Trig_Exp></Upd_Trig_Exp>
			<Del_Trig_Exp></Del_Trig_Exp>
			<TableComment></TableComment>
			<Autoinc_Nextval>0</Autoinc_Nextval>
			<Autoinc_Step>0</Autoinc_Step>
		</FIELD>
	</FIELDS>

	<IndexFiles>

		<IndexFile Type="Structural" >

			<INDEXES>
				<INDEX>
					<TagName>ORACLE</TagName>
					<TagType>REGULAR</TagType>
					<Key>ORACLE</Key>
					<Filter></Filter>
					<Order>ASCENDING</Order>
					<Collate>MACHINE</Collate>
				</INDEX>
				<INDEX>
					<TagName>SQLSERVER</TagName>
					<TagType>REGULAR</TagType>
					<Key>SQLSERVER</Key>
					<Filter></Filter>
					<Order>ASCENDING</Order>
					<Collate>MACHINE</Collate>
				</INDEX>
			</INDEXES>
		</IndexFile>

	</IndexFiles>


	<RECORDS>

		<RECORD>
			<FOXEXPR>.F.</FOXEXPR>
			<SQLSERVER>0</SQLSERVER>
			<ORACLE>0</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>.T.</FOXEXPR>
			<SQLSERVER>1</SQLSERVER>
			<ORACLE>1</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>#</FOXEXPR>
			<SQLSERVER>&lt;&gt;</SQLSERVER>
			<ORACLE>&lt;&gt;</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>ASC(</FOXEXPR>
			<SQLSERVER>ASCII(</SQLSERVER>
			<ORACLE>ASCII(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>AT(</FOXEXPR>
			<SQLSERVER>CHARINDEX(</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>CDOW(</FOXEXPR>
			<SQLSERVER>DATENAME(dw,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>CHR(</FOXEXPR>
			<SQLSERVER>CHAR(</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>CMONTH(</FOXEXPR>
			<SQLSERVER>DATENAME(mm,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>CTOD(</FOXEXPR>
			<SQLSERVER>CONVERT(datetime,</SQLSERVER>
			<ORACLE>TO_DATE(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>CTOT(</FOXEXPR>
			<SQLSERVER>CONVERT(datetime,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>DATE()</FOXEXPR>
			<SQLSERVER>GETDATE()</SQLSERVER>
			<ORACLE>SYSDATE</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>DAY(</FOXEXPR>
			<SQLSERVER>DATEPART(dd,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>DOW(</FOXEXPR>
			<SQLSERVER>DATEPART(dw,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>DTOC(</FOXEXPR>
			<SQLSERVER>CONVERT(varchar,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>DTOR(</FOXEXPR>
			<SQLSERVER>RADIANS(</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>DTOT(</FOXEXPR>
			<SQLSERVER>CONVERT(datetime,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>HOUR(</FOXEXPR>
			<SQLSERVER>DATEPART(hh,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>LIKE(</FOXEXPR>
			<SQLSERVER>PATINDEX(</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>MINUTE(</FOXEXPR>
			<SQLSERVER>DATEPART(mi,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>MONTH(</FOXEXPR>
			<SQLSERVER>DATEPART(mm,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>MTON(</FOXEXPR>
			<SQLSERVER>CONVERT(money,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>NTOM(</FOXEXPR>
			<SQLSERVER>CONVERT(float,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>RTOD(</FOXEXPR>
			<SQLSERVER>DEGREES(</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>SUBSTR(</FOXEXPR>
			<SQLSERVER>SUBSTRING(</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>TTOC(</FOXEXPR>
			<SQLSERVER>CONVERT(char,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>TTOD(</FOXEXPR>
			<SQLSERVER>CONVERT(datetime,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>YEAR(</FOXEXPR>
			<SQLSERVER>DATEPART(yy,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>.AND.</FOXEXPR>
			<SQLSERVER> AND</SQLSERVER>
			<ORACLE>AND</ORACLE>
			<PAD>.T.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>.OR.</FOXEXPR>
			<SQLSERVER>OR</SQLSERVER>
			<ORACLE>OR</ORACLE>
			<PAD>.T.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>.NULL.</FOXEXPR>
			<SQLSERVER>NULL</SQLSERVER>
			<ORACLE>NULL</ORACLE>
			<PAD>.T.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>.NOT.</FOXEXPR>
			<SQLSERVER>NOT</SQLSERVER>
			<ORACLE>NOT</ORACLE>
			<PAD>.T.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>&amp;</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>||</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>CEILING(</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>CEIL(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>LOG(</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>LN(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>LEN(</FOXEXPR>
			<SQLSERVER>DATALENGTH(</SQLSERVER>
			<ORACLE>LENGTH(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>MAX(</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>GREATEST(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>MIN(</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>LEAST(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>PADL(</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>LPAD(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>PADR(</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>RPAD(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>PROPER(</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>INITCAP(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>STR(</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>TO_CHAR(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>SUBS(</FOXEXPR>
			<SQLSERVER>SUBSTRING(</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>PROP(</FOXEXPR>
			<SQLSERVER></SQLSERVER>
			<ORACLE>INITCAP(</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>MONT(</FOXEXPR>
			<SQLSERVER>DATEPART(mm,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>MINU(</FOXEXPR>
			<SQLSERVER>DATEPART(mi,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>CMON(</FOXEXPR>
			<SQLSERVER>DATENAME(mm,</SQLSERVER>
			<ORACLE></ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>DATETIME()</FOXEXPR>
			<SQLSERVER>GETDATE()</SQLSERVER>
			<ORACLE>SYSDATE</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>=&lt;</FOXEXPR>
			<SQLSERVER>&lt;=</SQLSERVER>
			<ORACLE>&lt;=</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

		<RECORD>
			<FOXEXPR>=&gt;</FOXEXPR>
			<SQLSERVER>&gt;=</SQLSERVER>
			<ORACLE>&gt;=</ORACLE>
			<PAD>.F.</PAD>
		</RECORD>

	</RECORDS>


</TABLE>

