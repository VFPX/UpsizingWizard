#define PARSE_ERROR	(-1)

function IsSeparator
parameters cStr
return at(left(cStr, 1), chr(10)+chr(13)+chr(9)+chr(32)) > 0


function IsDelimiter
parameters cStr
local lcChar
lcChar = substr(cStr, 1, 1)
return len(lcChar) == 0 or IsSeparator(lcChar) or lcChar == ','


function IsFirstEscapeChar
parameters cStr
return at(left(cStr, 1), ["'([]) > 0


function IsQualifier
parameters cStr
return at(left(cStr, 1), ".!") > 0


function IsSqlKeyword
parameters cStr, pStr, lClause
local lcEnd

if len(substr(cStr, pStr)) = 0
	return ""
endif

lpEnd = SkipToken(cStr, pStr)
cToken = upper(substr(cStr, pStr, lpEnd - pStr))
do case
case cToken == "SELECT"
	return "SELECT"
case cToken == "FROM"
	return "FROM"
case cToken == "WHERE"
	return "WHERE"
case cToken == "GROUP"
	return "GROUP"
case cToken == "HAVING"
	return "HAVING"
case cToken == "UNION"
	return "UNION"
case cToken == "ORDER"
	return "ORDER"
case cToken == "INTO"
	return "INTO"
case cToken == "TO"
	return "TO"
case cToken == "PREFERENCES"
	return "PREFERENCES"
case cToken == "NOCONSOLE"
	return "NOCONSOLE"
case cToken == "PLAIN"
	return "PLAIN"
case cToken == "NOWAIT"
	return "NOWAIT"
endcase

if !lClause
	do case
	case cToken == "LEFT"
		return "LEFT"
	case cToken == "RIGHT"
		return "RIGHT"
	case cToken == "OUTER"
		return "OUTER"
	case cToken == "INNER"
		return "INNER"
	case cToken == "FULL"
		return "FULL"
	case cToken == "JOIN"
		return "JOIN"
	case cToken == "IN"
		return "IN"
	endcase
endif
return ""
endfun


* skip token delimited by spaces or comma
function SkipToken
parameters cStr, pStr, fComma
local lcChar
lcChar = substr(cStr, pStr, 1)
do while len(lcChar) > 0 and !IsSeparator(lcChar) and (!fComma or lcChar != ',')
	if IsFirstEscapeChar(lcChar)
		pStr = SkipProtectedString(cStr, pStr)
	else
		pStr = pStr + 1
	endif
	lcChar = substr(cStr, pStr, 1)
enddo
return pStr
endfun


* spaces, tabs, cr, lf
function SkipSpaces
parameters cStr, pStr
local lcChar
lcChar = substr(cStr, pStr, 1)
do while len(lcChar) > 0 and IsSeparator(lcChar)
	pStr = pStr + 1
	lcChar = substr(cStr, pStr, 1)
enddo
return pStr
endfunc


* go to next sql clause or eos
function NextClause
parameters cStr, pStr, cClause
pStr = SkipSpaces(cStr, pStr)
do while len(substr(cStr, pStr)) > 0 and empty(IsSqlKeyword(cStr, pStr, .T.))
	pStr = SkipToken(cStr, pStr, .F.)
	pStr = SkipSpaces(cStr, pStr)
enddo
cClause = IsSqlKeyword(cStr, pStr, .T.)
return pStr
endfunc


* returns the length of the protected string
function SkipProtectedString
parameters cStr, pStr
local lcChar, lcFirst, lcLast, lnCount

* we have at least a first escapr char
lcFirst = substr(cStr, pStr, 1)
lnCount = 1

do case
case lcFirst = '['
	lcLast = ']'
case lcFirst = '('
	lcLast = ')'
otherwise
	lcLast = lcFirst
endcase

pStr = pStr + 1
lcChar = substr(cStr, pStr, 1)
do while len(lcChar) > 0
	pStr = pStr + 1
	lnCount = lnCount + iif(lcChar = lcLast, - 1, iif(lcChar = lcFirst, 1, 0))
	if lnCount = 0
		exit
	endif
	lcChar = substr(cStr, pStr, 1)
enddo
* doesn't error at eos to simplify the code.
return pStr
endfunc


function SkipSymbol
parameters cStr, pStr
local lcChar

lcChar = substr(cStr, pStr, 1)
if IsFirstEscapeChar(lcChar)
	pStr = SkipProtectedString(cStr, pStr)
else
	do while isalpha(lcChar) or lcChar = '_'
		pStr = pStr + 1
		lcChar = substr(cStr, pStr, 1)
	enddo
endif
return pStr
endfunc


function ParseFromTable
parameters cStr, pStr, pNext, aTables

* parse <table> ::= <symbol> [{.|!} <symbol>...]
pStr = SkipSpaces(cStr, pStr)
pStart = pStr
pStr = SkipSymbol(cStr, pStr)
do while pStart < pStr and IsQualifier(substr(cStr, pStr, 1))
	pStr = SkipSymbol(cStr, pStr + 1)
enddo

* save table name into aTables
if pStart <= pStr and isDelimiter(substr(cStr, pStr, 1))
	if empty(aTables[1])
		nLen = 1
	else
		nLen = alen(aTables, 1) + 1
		dimension aTables[nLen]
	endif
	aTables[nLen] = substrc(cStr, pStart, pStr - pStart)
	
	* parse [<alias>]
	pStr = SkipSpaces(cStr, pStr)
	if pStr < pNext and substrc(cStr, pStr, 1) != ',' and empty(IsSqlKeyword(cStr, pStr, .F.))
		pStr = SkipSymbol(cStr, pStr)
	endif
else
	* if we got no table name, bail
	return PARSE_ERROR
endif
	
pStr = SkipSpaces(cStr, pStr)
return iif(pStr < pNext, pStr, pNext)
endfun


function ParseJoinList
parameters cStr, pStr, pNext, aTables, lLeftOuter
local lLeft

* parse JOIN | INNER JOIN | { LEFT | RIGHT | FULL } [OUTER] JOIN
pStr = SkipSpaces(cStr, pStr)
pStart = pStr
cKey = IsSqlKeyword(cStr, pStr, .F.)

if cKey == "INNER"
	pStr = SkipToken(cStr, pStr)
	pStr = SkipSpaces(cStr, pStr)
	cKey = IsSqlKeyword(cStr, pStr, .F.)
else
	if cKey == "LEFT" or cKey == "RIGHT" or cKey == "FULL"
		lLeft = (cKey = "LEFT")
		pStr = SkipToken(cStr, pStr)
		pStr = SkipSpaces(cStr, pStr)
		cKey = IsSqlKeyword(cStr, pStr, .F.)
		if cKey == "OUTER"
			lLeftOuter = lLeft
			pStr = SkipToken(cStr, pStr)
			pStr = SkipSpaces(cStr, pStr)
			cKey = IsSqlKeyword(cStr, pStr, .F.)
		endif
	endif
endif

if cKey == "JOIN"
	pStr = SkipToken(cStr, pStr)
	pStr = SkipSpaces(cStr, pStr)

	if substrc(cStr, pStr, 1) == '('
		pStr = pStr + 1
	endif
	
	* parse <table> [<alias>]
	pStr = ParseFromTable(cStr, pStr, pNext, @aTables)

	* parse { <join_op [(] <table> [<alias>]... }
	pStr = ParseJoinList(cStr, pStr, pNext, @aTables, @lLeft)
	
	* jump to next ',' or pNext
	do while pStr < pNext and substr(cStr, pStr, 1) != ','
		pStr = pStr + 1
	enddo
else
	* error on partial parsing
	if pStart < pStr
		return PARSE_ERROR
	endif
endif

return iif(pStr < pNext, pStr, pNext)
endfunc


function ParseFromClause
parameters cStr, aTables
local cTalk, pFrom, pNext, pStart, pEnd, cClause, lLeftOuter

pStr = 1
pStr = NextClause(cStr, pStr, @cClause)
if cClause <> "SELECT"
	return .F.
endif

pStr = SkipToken(cStr, pStr)
pStr = NextClause(cStr, pStr, @cClause)
if cClause <> "FROM"
	return .F.
endif

pFrom = SkipToken(cStr, pStr)
pNext = NextClause(cStr, pFrom, @cClause)
pStr = pFrom

do while pStr < pNext
	* save start pointer
	pStr = SkipSpaces(cStr, pStr)
	pStart = pStr
	lLeftOuter = .F.
	
	* parse <table> [<alias>]
	pStr = ParseFromTable(cStr, pStr, pNext, @aTables)
	if pStr = PARSE_ERROR
		return .F.
	endif

	* parse { <join_op [(] <table> [<alias>]... }
	pStr = ParseJoinList(cStr, pStr, pNext, @aTables, @lLeftOuter)
	if pStr = PARSE_ERROR
		return .F.
	endif
	
	* convert to odbc escape sequence {oj <local loj> }
	if lLeftOuter
		pEnd = pStr - 1
		do while IsSeparator(substr(cStr, pEnd, 1))
			pEnd = pEnd - 1
		enddo
		cStr = left(cStr, pEnd) + "}" + substr(cStr, pEnd + 1)
		cStr = left(cStr, pStart - 1) + "{oj " + substr(cStr, pStart)
		pStr = pStr + 5
		pNext = pNext + 5
	endif

	* if parsing is out of sync, cancel
	if pStr < pNext
		if substrc(cStr, pStr, 1) == ','
			pStr = pStr + 1
		else
			return .F.
		endif
	endif
enddo

return .T.
endfunc


