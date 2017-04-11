#define C_DEBUG .F.

#define C_VERS3B_LOC		"The Wizards require Visual FoxPro 3.0b or higher."
#define C_BUILD3B			690       && minimum 3.0b build to check for

#define DEFAULTTITLE_LOC	"Upsizing Wizard"

#define NODISABLE_LOC		"The current step cannot be disabled."

#define NOPROCESS_LOC		"No process defined."

#define MESSAGE_LOC			"Upsizing Wizard"

#define ERRORTITLE_LOC		"Upsizing Wizard"

#define ERRORMESSAGE_LOC ;
	"Error #" + alltrim(str(m.nError)) + " in " + m.cMethod + ;
	" (" + alltrim(str(m.nLine)) + "): " + m.cMessage

* The result of the above message will look like this:
*
*		Error #1 in WIZTEMPLATE.INIT (14): File does not exist.

#define MB_ICONEXCLAMATION		48
#define MB_ABORTRETRYIGNORE		2
#define MB_OK					0
#DEFINE MB_YESNO                4       && Yes and No buttons
#DEFINE IDYES           		6       && Yes button pressed

#define THERMCOMPLETE_LOC	"Complete."

#DEFINE NUM_AFIELDS  18   				&& number of columns in AFIELDS array
#DEFINE DT_MEMO  	"M"
#DEFINE DT_GENERAL  "G"
#DEFINE DT_BLOB		"W"

* The following string is used to form a message when
* _LOGFILE cannot be created or opened. The message will
* look like one of these two examples:
*
*	FCREATE("test1.log") failed. Event logging disabled.
*	FOPEN("test1.log", 2) failed. Event logging disabled.

#define LOGFILEERROR_LOC	" failed. Event logging disabled."

#DEFINE C_MAXFIELDS_LOC "The maximum number of fields to sort by is "

#DEFINE C_FREETABLE_LOC	"Free Tables"

#DEFINE BMP_LOCAL		"dblview.bmp"
#DEFINE BMP_REMOTE		"dbrview.bmp"
#DEFINE BMP_TABLE		"dbtable.bmp"

#DEFINE C_READONLY_LOC		"File is read-only and not allowed by this wizard/builder. Please select another."
#DEFINE C_READ2_LOC			"File is used exclusively by another."

#DEFINE E_BADDBCTABLE_LOC	"The table selected does not have a valid backlink to its DBC. "+;
							"You can fix this with the VALIDATE DATABASE RECOVER command."

#DEFINE C_TPROMPT_LOC		"Select file to open:"
#DEFINE C_READ3_LOC			"File is in use by Wizard. Select another."
#DEFINE C_READ4_LOC			"The DBF is part of a DBC. Select table from DBC container."

#define C_BADALIAS_LOC 'An alias in the list "THIS, THISFORM, THISFORMSET, OWIZARD, OENGINE" is ' + ;
	'in use. These aliases conflict with the wizards. Close the file(s) and continue?'

#define ALERTTITLE_LOC		"Upsizing Wizard"

#define C_WINONLY_LOC     "The Wizards require Visual FoxPro for Windows or Visual FoxPro for Macintosh."

* This message is displayed if the user enters the name of a file 
* that does not exist in WizEngine.WizLocFile()
#define E_FILENOTFOUND_LOC "File not found."


* These are the countries and regions to enable DBCS:  Japan, Korea, PRC, Taiwan
#DEFINE DBCS_LOC "81 82 86 88"

*- these are the problem characters, Mac-style
#DEFINE C_BADCHARS_MAC CHR(129) + CHR(130) + CHR(131) + CHR(132) + CHR(133) + CHR(134) + CHR(135) + CHR(136) + ;
	CHR(137) + CHR(138) + CHR(139) + CHR(140) + CHR(142) + CHR(143) + CHR(144) + CHR(145) + ;
	CHR(146) + CHR(147) + CHR(148) + CHR(149) + CHR(150) + CHR(151) + CHR(152) + CHR(153) + ;
	CHR(154) + CHR(160) + CHR(161) + CHR(162) + CHR(163) + CHR(164) + CHR(165) + CHR(47) + ;
	CHR(92) + CHR(44) + CHR(45) + CHR(61) + CHR(58) + CHR(59) + CHR(123) + CHR(125) + ;
	CHR(91) + CHR(93) + CHR(33) + CHR(64) + CHR(35) + CHR(36) + CHR(37) + CHR(94) + ;
	CHR(38) + CHR(42) + CHR(46) + CHR(60) + CHR(62) + CHR(40) + CHR(41) + CHR(63) + ;
	CHR(43) + CHR(124) + CHR(128) + CHR(155) + CHR(156) + CHR(157) + CHR(158) + CHR(159) + ;
	CHR(166) + CHR(167) + CHR(168) + CHR(169) + CHR(170) + CHR(171) + CHR(172) + CHR(173) + ;
	CHR(174) + CHR(175) + CHR(176) + CHR(177) + CHR(178) + CHR(179) + CHR(180) + CHR(181) + ;
	CHR(182) + CHR(183) + CHR(184) + CHR(185) + CHR(186) + CHR(187) + CHR(188) + CHR(189) + ;
	CHR(190) + CHR(191) + CHR(192) + CHR(193) + CHR(194) + CHR(195) + CHR(196) + CHR(197) + ;
	CHR(198) + CHR(199) + CHR(200) + CHR(201) + CHR(202) + CHR(203) + CHR(204) + CHR(205) + ;
	CHR(206) + CHR(207) + CHR(208) + CHR(209) + CHR(210) + CHR(211) + CHR(212) + CHR(213) + ;
	CHR(214) + CHR(215) + CHR(216) + CHR(217) + CHR(218) + CHR(219) + CHR(220) + CHR(221) + ;
	CHR(222) + CHR(223) + CHR(224) + CHR(225) + CHR(226) + CHR(227) + CHR(228) + CHR(229) + ;
	CHR(230) + CHR(231) + CHR(232) + CHR(233) + CHR(234) + CHR(235) + CHR(236) + CHR(237) + ;
	CHR(238) + CHR(239) + CHR(240) + CHR(241) + CHR(242) + CHR(243) + CHR(244) + CHR(245) + ;
	CHR(246) + CHR(247) + CHR(248) + CHR(249) + CHR(250) + CHR(251) + CHR(252) + CHR(253) + ;
	CHR(254) + CHR(34) + CHR(39) + CHR(32)


*- for setting FoxTools in Wizards
#DEFINE EN_FOXTOOLS_LOC		"The Setup Wizard requires "
#DEFINE EN_FXTVER_LOC		"This Wizard requires version 3.00 or higher of "
#DEFINE C_WINLIBRARY		"FT3.DLL"							&& latest version of foxtools.fll
#DEFINE C_MACPPCLIBRARY		"FXTOOL30.CFM"						&& latest version of foxtools (code fragment version)
#DEFINE C_MAC68KLIBRARY		"FXTOOL30.ASLM"						&& latest version of foxtools (shared library manager version)

#DEFINE C_MACPPC_TAG_LOC	"POWER MAC"							&& in the VERS(1) return string

#DEFINE C_NOTAG_LOC 		"You cannot combine index tags and fields."
#DEFINE TAGDELIM	 " *"

* Registry strings for looking up scrollbox color
#DEFINE HKEY_CURRENT_USER           -2147483647  && BITSET(0,31)+1
#DEFINE CONTROL_KEY					"Control Panel\Colors"
#DEFINE SCROLLBAR_KEY				"Scrollbar"
#DEFINE SHADOW_KEY					"ButtonShadow"
