'// publics.bas
'//---------------------------publics----------------------------------------------
'//		6/6/20.	wmk.
'// publics - module-wide vars for Module1 code in accounts sheets
'// module-wide constants. (used in processing bank download sheets)
'// (mirrored in file publics)

'// Modification History.
'// ---------------------
'//	5/??/20.	wmk.	Original code
'//	5/20/20.	wmk.	Released as Production into Alpha Clean Accounting.ods
'//	5/23/20.	wmk.	ASTSHEET, LIASHEET, INCSHEET, corrected to match
'//						Production spreadsheet; RJUST corrected
'//	5/27/20.	wmk.	DEVDEBUG constant added
'//	5/28/20.	wmk.	ERRUNKNOWN universal error code value added
'//	6/6/20.		wmk.	FMTDATETIME, DATEROW constants added

'//
'// Note. The Module1 functions D and E are referenced in cells for executing test
'// cases against the General Ledger sheet. autocalc MUST BE OFF if these two
'// functions are active (DEVDEBUG=true), since both functions fire on the current
'// user selection. They really go haywire with recalc or autocalc since the current
'// user row/column range may have nothing to do with their intended reference.
'//
'// debugging control.
public const DEVDEBUG=true						'// true if functions D(), E() activated
'public const DEVDEBUG=false					'// false if functions D(), E() de-activated

'// universal error code(s).
public const ERRUNKNOWN=-9999					'// universal unknown error code

'// production sheet names
public const GLSHEET = "General Ledger"			'// General Ledger Sheet.Name
public const ASTSHEET = "Asset Accounts"		'// (Production) Assets Sheet.Name
public const LIASHEET = "Liability Accounts"	'// (Production) Liabilities Sheet.Name
public const INCSHEET = "Income Accounts"		'// (Production) Income Sheet.Name
public const EXPSHEET = "Expense Accounts"		'// Expense Accounts Sheet.Name
'// development sheet names
'public const ASTSHEET = "Assets"				'// Assets Sheet.Name
'Public const LIASHEET = "Liabilities"			'// Liabilities Sheet.Name
'public const INCSHEET = "Income"				'// Income Sheet.Name
'public const EXPSHEET = "Expense Accounts_2"						'// mod051720
public const MON1="January"
public const MON2="February"
public const MON3="March"
public const MON4="April"
public const MON5="May"
public const MON6="June"
public const MON7="July"
public const MON8="August"
public const MON9="September"
public const MON10="October"
public const MON11="November"
public const MON12="December"
public const YELLOW=16776960		'// decimal value of YELLOW color
public const LTGREEN=10092390		'// decimal value of LTGREEN color

'// column index values and string lengths for column data			
public const	INSBIAS=3		'// insert line count bias
public const	COALEN=4		'// length of COA field strings
public const COLDATE=0		'// Date - column A
public const COLTRANS=1	'// Transaction - column B
public const COLDEBIT=2	'// Debit - column C
public const COLCREDIT=3	'// Credit - column D
public const COLACCT=4		'// COA Account - column E
public const COLREF=5		'// Reference - column F
public const DATEROW=1		'// Sheet date row index					'// mod060620

'// cell formatting constants.
public const DEC2=123		'// number format for (x,xxx.yy)			'// mod052020
public const MMDDYY=37		'// date format mm/dd/y						'// mod052020
public const FMTDATETIME=50	'// date/time format mm/dd/yyy hh:mm:ss 	'// mod060620
public const LJUST=1		'// left-justify HoriJustify				'// mod052020
public const CJUST=2		'// center HoriJustify						'// mod052020
public const RJUST=3		'// right-justify HoriJustify				'// mod052320
public const MAXTRANSL=50	'// maximum transaction text length			'// mod052020
'//----------------------end publics---------------------------------------------
'/**/
