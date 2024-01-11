'// Module1Hdr.bas
REM  *****  BASIC  *****
'// Module1Hdr - BadAss Library Module1 header.
'// Module1 (Beta Release) - Local macros for Chase Activity Download Spreadsheet and
'//	transaction processing in Production Alpha Clean Accounting spreadsheet.
'// 8/22/22.	wmk.	'/**/ and // <module-name>.bas bracketing lines added; application
'//				 extended to general banking activity in Accounting subsystem.
'// Legacy Mods.
'//	5/14/20.	wmk.	original code with mods
'//	5/16/20.	wmk.	new functions and subs to span sheets for
'//						transaction mapping to accounts sheets
'//	5/17/20.	wmk.	correct test constant EXPSHEET to match this
'//						document 'Expense Accounts_2'
'//	5/18/20.	wmk.	row data processing constants added.
'//	5/19/20.	wmk.	sheet name constants changed to match production
'//	5/20/20.	wmk.	added OPTION EXPLICIT for code control
'//	5/22/20.	wmk.	added PlaceSplitTrans function.
'//	5/23/20.	wmk.	added BeanField function; bug fix with ASTSHEET
'//	5/25/20.	wmk.	added errhandling.bas and associated subs
'//	5/26/20.	wmk.	errhandling.bas update resinserted
'//	6/28/20.	wmk.	Beta Release to BadAss Library

'// Code protection.
OPTION EXPLICIT				'// every variable must be declared before use

'/**/

'// errhandling.bas
'//---------------------------------------------------------------
'// errhandling.bas - Definitions for development error handling.
'//		wmk. 5/26/20.	17:30
'//---------------------------------------------------------------
'//
'// Header Description.
'// -------------------
'//	errhandling.bas contains definitions that may be inserted in any
'// OOoBasic Module for handing error conditions. This is intended for
'// use with development modules, and should be removed for Production
'// spreadsheets. This will ensure clean code released into the
'// Production environment.
'//
'// error handling interface methods.
'// ErrLogSetup - set up error logging (gbErrLogging, goCellRangeAddress, gsErrModule)
'// ErrLogDisable() - disable error logging by setting gbErrLogging = false
'// ErrLogGetCellInfo(lColumn, lRow) - get error focus cell column, row
'// ErrLogSetCellInfo(lColumn, lRow) - set error focus cell column, row
'//	ErrLogGetCellName() - get error focus cell alphanumeric name
'//	ErrLogGetDisplay() - get error message box enabled status
'//	ErrLogSetDisplay(bDisplayOn) - set error message box enabled status
'//	ErrLogGetModule() - get current module name gsErrModule
'// ErrLogSetModule(sName) - set current module name gsErrModule
'// ErrLogGetRecording - Get error recording status from error log globals
'// ErrLogSetRecording - Set error recording status in error log globals
'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
'// LogError - make error log entry([goCellRangeAddress], psErrName, psErrMsg)

'// global constants.

'// public constants.
public const ERRLOGSHEET="Error Log"			'// error logging sheet name
public const ERRLOGINSROW=5						'// insertion row for log messages
'// global variables.
'// global variables.
global gbErrDisplay	As Boolean					'// msgBox on/off flag
global gbErrLogging as Boolean					'// logging on/off flag; LogSetup sets/clears
global gbErrRecording As Boolean				'// log sheet recording on/off flag
global gsErrSheet As String						'// name of sheet error thrown on
global gsErrModule As String					'// sub/function throwing error
global gsErrName As String						'// name of error code
global gsErrMsg As String						'// error message
global goErrRangeAddress As Object				'// cell range address

'// public variables.
'//

'// error handling interface methods.
'// ErrLogSetup - set up error logging (gbErrLogging, goCellRangeAddress, gsErrModule)
'// ErrLogDisable() - disable error logging by setting gbErrLogging = false
'// ErrLogGetCellInfo(lColumn, lRow) - get error focus cell column, row
'// ErrLogSetCellInfo(lColumn, lRow) - set error focus cell column, row
'//	ErrLogGetCellName() - get error focus cell alphanumeric name
'//	ErrLogGetDisplay() - get error message box enabled status
'//	ErrLogSetDisplay(bDisplayOn) - set error message box enabled status
'//	ErrLogGetModule() - get current module name gsErrModule
'// ErrLogSetModule(sName) - set current module name gsErrModule
'// ErrLogGetRecording - Get error recording status from error log globals
'// ErrLogSetRecording - Set error recording status in error log globals
'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
'// LogError - make error log entry([goCellRangeAddress], psErrName, psErrMsg)
'//-----------------end errhandling.bas header--------------------------
'/**/

'// publics.bas
'//---------------------------publics.bas----------------------------------------------
'//		6/6/20.	wmk.
'// publics.bas - module-wide vars for Module1 code in accounts sheets
'// module-wide constants. (used in processing bank download sheets)
'// (mirrored in file publics.bas)

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
'//----------------------end publics.bas---------------------------------------------
'/**/

'// Main.bas
'//---------------------------------------------------------------
'// Main - generic Main sub for module.
'//		wmk. 5/15/20.
'//---------------------------------------------------------------

Sub Main

End Sub
'/**/

'// CheckDoubleEntry.bas
'//---------------------------------------------------------------
'// CheckDoubleEntry - Check double entry transaction fields.
'//		7/3/20.	wmk.
'//---------------------------------------------------------------

public function CheckDoubleEntry(poTransRange As Object, rpsMonth As String,_
					rpsAcct1 As String, rpsAcct2 As String) As Integer

'//	Usage.	iVal = CheckDoubleEntry(oTransRange, rpsMonth, rpsAcct1,
'//						rpsAcct2)
'//
'//		oTransRange = RangeAddress of double-entry transaction
'//		rpsMonth = Month name string (e.g. "January")
'//		rpsAcct1 = (returned)
'//		rpsAcct2 = (returned)
'//
'// Entry.	<entry conditions>
'//
'//	Exit.	iVal = 0 if no error,
'//				   ERRCONTINUE (-1) if error; skip to next transaction
'//				   ERRSTOP (-2) if error; end processing user selction
'//			rpsAcct1 = COA from first line of transaction
'//			rpsAcct2 = COA from 2nd line of transaction
'//
'// Calls.	GetMonthName
'//
'//	Modification history.
'//	---------------------
'//	6/1/20.	wmk.	original code
'//	6/2/20.	wmk.	error handling setup added
'//	6/3/20.	wmk.	csERRBADDATE constant added
'// 6/5/20.	wmk.	bug fix; was not detecting 2nd line Debit and Credit empty
'//	7/3/20.	wmk.	sERRCOANOTFOUND corrected to csERRCOANOTFOUND
'//	Notes.
'//

const csERRCOANOTFOUND="ERRCOANOTFOUND"
const csERROUTOFROOM="ERROUTOFROOM"
const csERRCOACHANGED="ERRCOACHANGED"

'//----------in Module error handling setup-------------------------------
'// LogError setup snippet.
const csERRUNKNOWN="ERRUNKNOWN"
const sMyName="CheckDoubleEntry"
'// add additional error code strings here...
Dim sErrName as String		'// error name for LogError
Dim sErrMsg As String		'// error message for LogError
'//*---------error handling setup---------------------------
'//	ErrLogGetModule() - get current module name gsErrModule
'// ErrLogSetModule(sName) - set current module name gsErrModule
'// ErrLogGetCellInfo(lColumn, lRow) - get error focus cell column, row
'// ErrLogSetCellInfo(lColumn, lRow) - set error focus cell column, row
'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
dim iErrSheetIx as Integer
dim lErrColumn as Long
dim lErrRow as Long
dim sErrCurrentMod As String
'//----------end in Module error handling setup---------------------------

'//	constants.
'// process control constants.
const ERRCONTINUE=-1	'// error, but continue with next transaction
const ERRSTOP=-2		'// error, stop processing transactions
const csERRBADDATE="ERRBADDATE"		'// bad date error name

'//	local variables.
dim iRetValue As Integer
dim iBrkPt As Integer		'// easy breakpoint
dim iStatus As Integer		'// general status var


'//------------ CheckDoubleEntry code begins here---------------------------
'//	Usage.	iVal = CheckDoubleEntry(oTransRange, rpsMonth, rpsAcct1,
'//						rpsAcct2)
		dim Doc As Object
		dim oSheets as Object
		dim oGLSheet as Object		'// ledger sheet
		dim iSheetIx As Integer		'// sheet index of ledger
		dim bTransValid As Boolean	'// transaction is valid flag
  		dim bADEmpty As Boolean		'// 1st Debit/2nd Credit empty flag
  		dim bBCEmpty As Boolean		'// 1st Credit/2nd Debit empty flag
		
		'// Cell access objects for first transaction line.
		dim oCellDate as Object		'// Date cell
		dim oCellTrans As Object	'// Transaction cell
		dim oCellDebit As Object	'// Debit cell
		dim oCellCredit As Object	'// Credit cell
		dim oCellAcct As Object		'// COA cell
		dim oCellRef As Object		'// Reference cell
		dim sDate as String			'// Date string
		dim sTrans as String		'// Transaction string
		dim sStyle as String		'// Style property
		dim lDateFormat As Long		'// date format code
		dim sMonth As String		'// month name
		dim lGLCurrRow As Long		'// first row of double entry
		dim lGLNextRow As Long		'// 2nd row of double entry
		dim lGLEndRow As Long		'// end of ledger selection by user
		dim sDebit As String		'// Debit field
		dim sCredit As String		'// Credit field
		dim sAcct As String			'// COA field
		
		'// Cell access objects for 2nd transaction line.
		dim oCellDateB as Object	'// Date cell
		dim oCellTransB As Object	'// Transaction cell
		dim oCellDebitB As Object	'// Debit cell 
		dim oCellCreditB As Object	'// Credit cell
		dim oCellAcctB As Object	'// COA cell
		dim oCellRefB As Object		'// Reference cell
		dim sDateB As String
		dim sTransB As String
		dim sStyleB As String
		dim lDateBFormat As Long
		dim sDebitB As String
		dim sCreditB As String
		dim sAcctB As String
		
	'// code.

'//*------------error handling code initialization---------------
	'// preserve entry error settings.
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	sErrCurrentMod = ErrLogGetModule()
'//*	
	'// set local error settings.
	ErrLogSetModule(sMyName)
	ErrLogSetSheet(poTransRange.Sheet)
	ErrLogSetCellInfo(COLDATE, poTransRange.StartRow)
	ErrLogSetRecording(true)		'// enable log entries
'//*----------end error handling code initialization-------------

	iRetValue = 0
	iStatus = 0
'		poTransRange.Sheet = iSheetIx of ledger spreadsheet
'		poTransRange.StartColumn = COLDATE
'		poTransRange.StartRow = lGLCurrRow
'		poTransRange.EndColumn = COLREF
'		poTransRange.EndRow = lGLCurrRow + 1
		'// Access ledger sheet from passed information
	iSheetIx = poTransRange.Sheet
	Doc = ThisComponent
	oSheets = Doc.getSheets()
	oGLSheet = oSheets.GetByIndex(iSheetIx)
	lGLCurrRow = poTransRange.StartRow
	lGLEndRow = poTransRange.EndRow
			
		'// load 1st transaction line row from GL sheet
		oCellDate = oGLSheet.getCellByPosition( COLDATE, lGLCurrRow)
		oCellTrans = oGLSheet.getCellByPosition( COLTRANS, lGLCurrRow)
		oCellDebit = oGLSheet.getCellByPosition( COLDEBIT, lGLCurrRow)
		oCellCredit = oGLSheet.getCellByPosition( COLCREDIT, lGLCurrRow)
		oCellAcct = oGLSheet.getCellByPosition( COLACCT, lGLCurrRow)
		oCellRef = oGLSheet.getCellByPosition( COLREF, lGLCurrRow)

		'// extract date & transaction information from 1st row
		sDate = oCellDate.String
		sTrans = oCellTrans.String
	'//	Ensure nonempty date field before GetMonthName call
	if Len(Trim(sDate)) = 0 then
		sErrName = "ERRBADDATE"
		sErrMsg = "1st Date field empty - skipping"
		iStatus = ERRSTOP
		GoTo DateError
	endif
		

iBrkpt=1
		sStyle = oCellDate.Text.CellStyle
		lDateFormat = oCellDate.Text.NumberFormat
		sMonth = GetMonthName(sDate)		'// get month search name
		
iBrkpt = 1	'// breakpoint line
		
		'// load 2nd transaction line row from GL sheet
		lGLNextRow = lGLCurrRow + 1
		
		'// check that 2nd row is not beyond selection
		if lGLNextRow > lGLEndRow then
'			iStatus = msgBox("Selection ends in mid-transaction!"_
'					+CHR(13)+CHR(10)+ "Check last row selection of transaction")	
			oCellDate.Text.CellBackColor = YELLOW
			sErrName = "ERRODDROWS"
			sErrMsg = "Selection ends in mid-transaction!"_
				+CHR(13)+CHR(10)+ "Check last row selection of transaction"
			Call LogError(sErrName, sErrMsg)
			iRetValue = ERRSTOP
			
			GoTo AdvanceTrans
			'exit For
		endif	'// end would advance past last row conditional

		'// load 2nd line transaction values
		oCellDateB = oGLSheet.getCellByPosition( COLDATE, lGLNextRow)
		oCellTransB = oGLSheet.getCellByPosition( COLTRANS, lGLNextRow)
		oCellDebitB = oGLSheet.getCellByPosition( COLDEBIT, lGLNextRow)
		oCellCreditB = oGLSheet.getCellByPosition( COLCREDIT, lGLNextRow)
		oCellAcctB = oGLSheet.getCellByPosition( COLACCT, lGLNextRow)
		oCellRefB = oGLSheet.getCellByPosition( COLREF, lGLNextRow)
'XRay oCellDateB

		'// extract date information from 2nd row of transaction
		sDateB = oCellDateB.Text.String
		
		'//	Ensure nonempty date field before GetMonthName call
		if Len(Trim(sDateB)) = 0 then
			sErrName = "ERRBADDATE"
			sErrMsg = "2nd Date field empty - skipping"
			GoTo DateError
		endif

		sStyleB = oCellDateB.Text.CellStyle
		lDateBFormat = oCellDateB.Text.NumberFormat
'		sMonth = GetMonthName(sDate)		'// use month from 1st line

iBrkpt = 1	'// breakpoint line
		
		'// make sure dates, descriptions, and amounts agree
			'// if not, issue message box and bail out
		iStatus = StrComp(sDate, sDateB)
		if iStatus <> 0 then
			sErrName = csERRBADDATE
			sErrMsg = "Transaction date mismatch!"+CHR(13)+CHR(10)_
						+ "Check 1st row selection of transaction"
			GoTo DateError
		endif

		GoTo CheckDesc
		
'//----------Date Error handler-------------------------------
iBrkpt = 1
'//	Date error(s) are terminating errors; iRetVal = -2
DateError:
		if iStatus <> 0 then
			'// log error and set Date Cell YELLOW
			oCellDate.Text.CellBackColor = YELLOW
			Call LogError(sErrName, sErrMsg)
			iRetValue = ERRSTOP		'// bail out on return
			GoTo AdvanceTrans		
		endif		'// end any date error conditional

'//-------end Date Error Handler-----------------------------

CheckDesc:
		
iBrkpt = 1	'// breakpoint line

		'// extract transaction information from both rows
		'// make sure that descriptions agree
		
		sTrans = oCellTrans.String
		sTransB = oCellTransB.String
		iStatus = strComp(sTrans, sTransB)
		
		if iStatus <> 0 then
'			iStatus = msgBox("Transaction description mismatch!"+CHR(13)+CHR(10)_
'						+ "Check transaction")
			sErrName = "ERRDESCNOMATCH"
			sErrMsg = 	"Transaction description mismatch!"+CHR(13)+CHR(10)_
						+ "Check transaction"
			Call LogError(sErrName, sErrMsg)
			oCellDate.Text.CellBackColor = YELLOW
			oCellDateB.Text.CellBackColor = YELLOW
			iRetValue = ERRSTOP			'// force stop on return
			GoTo AdvanceTrans
		endif		'// end descriptions not same

		'// extract debit and credit amounts for both halves of transaction
		'// ensure both halves have same amount, if not, bail out
		
		sDebit = oCellDebit.Text.String
		sCredit = oCellCredit.Text.String

		sDebitB = oCellDebitB.Text.String
		sCreditB = oCellCreditB.Text.String

'//-------------------------------------------------------------------------
		'// ensure that is actual double entry; either debit on 1st line and
		'// Credit on second line are same amount, or Credit on 1st line
		'// and Debit on second line are same amount etc.

'//	Debit Credit
'//   A     B			sDebit and sCredit
'//   C     D			sDebitB and sCreditB		
'//   A=D and B=C		correct
'//  if A and D empty, B and C must be nonempty
'//  if B and C empty, A and D must be nonempty
'//-------------------------------------------------------------------------
iBrkpt = 1 
 
' dim bTransValid as boolean
' dim bADEmpty as boolean
' dim bBCEmpty as boolean
  		bTransValid = (StrComp(sDebit,sCreditB) = 0) AND  (StrComp(sCredit,sDebitB)=0)
  		bTransValid = bTransValid AND ((Len(sDebit) > 0) OR (Len(sCredit) > 0))_
  								  AND  ((Len(sCreditB) > 0) OR (Len(sDebitB) > 0))			'// mod060520
  		bADEmpty = (Len(sDebit)=0) AND (Len(sCreditB)=0)
  		bBCEmpty = (Len(sCredit)=0) AND (Len(sDebitB)=0)
  		bTransValid = bTransValid AND ((bADEmpty AND (Not bBCEmpty))_
  									 OR (bBCEmpty AND (Not bADEmpty)))						'// mod052020
  		if Not bTransValid then
  		'// message, flag both lines, skip
'  			iStatus = msgBox("Invalid transaction...check Debit and Credit values!"+CHR(13)+CHR(10)_
'  							+"Transaction skipped")
			sErrName = "ERRDEB<>CRED"
			sErrMsg = "Invalid transaction...check Debit and Credit values!"+CHR(13)+CHR(10)_
  							+"Transaction skipped"
  			Call LogError(sErrName, sErrMsg)
  			oCellDate.Text.CellBackColor = YELLOW
			oCellDateB.Text.CellBackColor = YELLOW
			iRetValue = ERRCONTINUE
			GoTo AdvanceTrans							
  		endif	'// end invalid transaction
  		    
		'// extract COA account numbers for both halves of transaction
		sAcct = oCellAcct.String
		sAcctB = oCellAcctB.String
iBrkpt = 1
		'// ensure both halves have COA account number; if not, just skip
		if Len( sAcct ) <> 4 OR Len(sAcctB) <> 4 then
			iStatus = msgBox("Transaction contains invalid COA number(s)!"_
					+CHR(13)+CHR(10)+ "Check 1st row selection of transaction")	
			sErrName = csERRCOANOTFOUND
			sErrMsg = "Transaction contains invalid COA number(s)!"_
				+CHR(13)+CHR(10)+ "Check 1st row selection of transaction"
			Call LogError(sErrName, sErrMsg)
			oCellDate.Text.CellBackColor = YELLOW
			oCellDateB.Text.CellBackColor = YELLOW
			'// Note. have to do the following rather than use if-then-else to loop
			'// again, since this Basic does not support "Continue For"
			iRetValue = ERRCONTINUE 		'// force Advance to next on return
			GoTo AdvanceTrans
		endif	'// end length not == 4
		
		'// set returned month and COA accounts
		rpsMonth = sMonth
		rpsAcct1 = sAcct
		rpsAcct2 = sAcctB
		
AdvanceTrans:
'//*-----------------restore error handling for caller-----------
	'// restore entry error settings.
	ErrLogSetModule(sErrCurrentMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
'//*----------end error handling setup---------------------------

	'// iRetValue = 0, no error
	'//				ERRCONTINUE (-1) - error; continue with next transaction
	'//				ERRSTOP (-2) - error; stop processing transactions
	CheckDoubleEntry = iRetValue

end function 	'// end CheckDoubleEntry	7/3/20
'/**/

'// CheckInserts.bas
'//------------------------------------------------------------------
'// CheckInserts - Check target COA sheet(s) for insertion available.
'//		7/3/20.	wmk.
'//------------------------------------------------------------------

public function CheckInserts(psMonth As String, psAcct1 As String,_
		psAcct2 As String, rpoCat1Range As Object, rpoCat2Range As Object) as Integer

'//	Usage.	iVal = CheckInserts(sMonth, sAcct1, sAcct2, oCat1Range, oCat2Range)
'//
'//		sMonth = Month name string (e.g. "January")
'//		sAcct1 = COA from 1st line of transaction
'//		sAcct2 = COA from 2nd line of transaction
'//		rpoCat1Range = (returned)
'//		rpoCat2Range = (returned)
'//
'// Entry.	Spreadsheet has category sheets for all possible COA #s
'//
'//	Exit.	rplIns1Ptr = row for insertion of line 1 transaction
'//			rplIns2Ptr = row for insertion of line 2 transaction
'//
'// Calls.	GetInsRow.
'//
'//	Modification history.
'//	---------------------
'//	6/1/20.		wmk.	original code; stub
'//	6/2/20.		wmk.	code transported from within PlaceTransM; error handling
'//						set up
'//	6/8/20.		wmk.	bug fix where sCOA1Msg and sCOA2Msg not declared	
'// 7/3/20.		wmk.	bug fix in error handling using lGLCurrRow, needs lErrRow
'//
'//	Notes. Passed paramters are COA #s; they will be used to access the
'// correct account catgory sheets, for the month of the transaction
'//

'//	constants.
'// insert error codes.
const ERRCOANOTFOUND=-1
const ERROUTOFROOM=-2
const ERRNOCOAMONTH=-3
const sERRCOANOTFOUND="ERRCOANOTFOUND"
const sERROUTOFROOM="ERROUTOFROOM"
const sERRNOCOAMONTH="ERRNOCOAMONTH"
const sERRUNKNOWN="ERRUNKNOWN"

'//----------in Module error handling setup-------------------------------
'// LogError setup snippet.
const csERRUNKNOWN="ERRUNKNOWN"
const sMyName="CheckInserts"
'// add additional error code strings here...
Dim sErrName as String		'// error name for LogError
Dim sErrMsg As String		'// error message for LogError
'//*---------error handling setup---------------------------
'//	ErrLogGetModule() - get current module name gsErrModule
'// ErrLogSetModule(sName) - set current module name gsErrModule
'// ErrLogGetCellInfo(lColumn, lRow) - get error focus cell column, row
'// ErrLogSetCellInfo(lColumn, lRow) - set error focus cell column, row
'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
dim iErrSheetIx as Integer
dim lErrColumn as Long
dim lErrRow as Long
dim sErrCurrentMod As String
'//----------end in Module error handling setup---------------------------

'// process control constants.
const ERRCONTINUE=-1	'// error, but continue with next transaction
const ERRSTOP=-2		'// error, stop processing transactions

'//	local variables.
dim iRetValue As Integer	'// function returned value
dim sMonth As String		'// local month

'//	Object setup variables.
dim oDoc As Object			'// ThisComponent
dim oSheets As Object		'// Doc.getSheets()

'// COA Sheet processing variables.
dim sAcct As String			'// first line COA#
dim sAcctB As String		'// second line COA#
dim sAcctCat1 as String 	'// accounting category of 1st line of transaction
dim sAcctCat2 as String		'// account category of 2nd line of transaction
dim bSameAcctCat as boolean	'// account categories the same flag		
dim oCat1Sheet As Object	'// COA sheet for 1st transaction line
dim oCat2Sheet As Object	'// COA sheet for 2nd transaction line
dim bSameCOA As Boolean		'// COAs match flag
dim lCat1InsRow As Long		'// COA sheet insertion row, line 1
dim lCat2InsRow As Long		'// COA sheet insertion row, line 2
dim sBadCOA As String		'// first COA bad string
dim sBadCOA2 As String		'// second COA bad string
dim iStatus As Integer		'// processing status flag
dim sCOA1Msg As String		'// COA1 error message string
dim sCOA2Msg As String		'// COA2 error message string

	'// code.
	
	iRetValue = 0
	iStatus = 0				'// clear general status
'	rplIns1Ptr = -1			'// set bad pointers for return
'	rplIns2Ptr = -1
	sMonth = psMonth		'// set local copy of passed month
		
'//*------------error handling code initialization---------------
	'// preserve entry error settings.
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	sErrCurrentMod = ErrLogGetModule()
'//*	
	'// set local error settings.
	ErrLogSetModule(sMyName)
	'// keep entry sheet/cell selctions, since errors will be in GL
'	ErrLogSetSheet(poTransRange.Sheet)
'	ErrLogSetCellInfo(COLDATE, poTransRange.StartRow)
	ErrLogSetRecording(true)		'// enable log entries
'//*----------end error handling code initialization-------------

	
	CheckInserts = iRetValue
'	if true then
'		Exit Function
'	endif
'//---------------end stub code--------------------	
'//-----------------CheckInserts code starts here----------------------------------------
'//
'//	Usage.	iVal = CheckInserts(psMonth, psAcct1, psAcct2, rplIns1Ptr, rplIns2Ptr)
'//
'//		psMonth = Month name string (e.g. "January")
'//		psAcct1 = COA from 1st line of transaction
'//		psAcct2 = COA from 2nd line of transaction
'//		rplIns1Ptr = (returned)
'//		rplIns2Ptr = (returned)
'//

		'// get appropriate sheet names from COA account numbers
		iStatus = 0			'// clear processing exceptions
		sAcct = psAcct1		'// copy COAs to local vars
		sAcctB = psAcct2
		sAcctCat1 = GetTransSheetName(sAcct)
		sAcctCat2 = GetTransSheetName(sAcctB)
		bSameAcctCat = (StrComp(sAcctCat1, sAcctCat2) = 0)
dim iBrkpt As Integer
iBrkpt = 1

		'// set up oCat1Sheet as sheet object for first half of transaction
		'//			oCat2Sheet as sheet object for second half of transaction
'static oCat1Sheet as Object	'// account sheet object of 1st category in transaction
'static oCat2Sheet as Object	'// account sheet object of 2nd category in transaction
'dim lCat1InsRow as long		'// insert index for 1st category in transaction
'dim lCat2InsRow as long		'// insert index for 2nd cateory in transaction

'XRay Doc
		oDoc = ThisComponent
		oSheets = oDoc.getSheets()

'XRay oSheets
		'// check if same sheet, and just copy instance						'// mod051920
		oCat1Sheet = oSheets.getByName(sAcctCat1)
		if bSameAcctCat then
			oCat2Sheet = oCat1Sheet
		else
			oCat2Sheet = oSheets.getByName(sAcctCat2)
		endif
		
		rpoCat1Range.Sheet = oCat1Sheet.RangeAddress.Sheet
		rpoCat2Range.Sheet = oCat2Sheet.RangeAddress.Sheet
		
'XRay oCat1Sheet		
iBrkpt=1

		'// check here if COA#s different; if not, will bail out later		'// mod051920
		bSameCOA = StrComp(sAcct, sAcctB) = 0								'// mod051920
		lCat1InsRow = GetInsRow(oCat1Sheet, sAcct, sMonth)					'// mod051920

		if bSameCOA then													'// mod051920
			'// Note: might get away with setting same insertion row, since will
			'// insert 2 lines... for now, skipped
			lCat2InsRow = lCat1InsRow										'// mod051920
		else																'// mod051920
			lCat2InsRow = GetInsRow(oCat2Sheet, sAcctB, sMonth)				'// mod051920
		endif	'// end same COA conditional								'// mod051920
				
		'// add code to verify have a row in each sheet before committing to change
'dim lBadRow1 	as long		'// 1st acct insert row return value
'dim lBadRow2	as lont		'// 2nd acct insert row return value

		if lCat1InsRow < 0 OR lCat2InsRow < 0 then
			'// issue appropriate message and don't mark rows as processed
			sBadCOA = ""
			sBadCOA2 = ""
			'// handle transaction line 1 error.
			if lCat1InsRow < 0 then
				sBadCOA = sAcct
				sCOA1Msg = ""			'// clear line 1 message
				ErrLogSetSheet(oCat1Sheet.RangeAddress.Sheet)
				ErrLogSetCellInfo(COLDATE, lErrRow)
				Select Case lCat1InsRow
'				Case -1
				Case ERRCOANOTFOUND
					sCOA1Msg = "Account "+sBadCOA+" not found"
					sErrName = sERRCOANOTFOUND
'				Case -2
				Case ERROUTOFROOM
					sCOA1Msg = "Account "+sBadCOA+" month "+sMonth+" not enough rows"_
								+" to insert"
					sErrName = sERROUTOFROOM
'				Case -3
				Case ERRNOCOAMONTH
					sCOA1Msg = "Account "+sBadCOA+" month "+sMonth+" not found"
					sErrName = sERRNOCOAMONTH
				Case else
					sCOA1Msg = "Account "+sBadCOA+"Undocumented error"
					sErrName = sERRUNKNOWN
				end Select
				
				sErrMsg = sCOA1Msg
				Call LogError(sErrName, sErrMsg)
				
			endif	'// end error in 1st line conditional
			
			if lCat2InsRow < 0 then
				sBadCOA2 = sAcctB
				sCOA2Msg = ""			'// clear line 2 message
				Select Case lCat2InsRow
'				Case -1
				Case ERRCOANOTFOUND
					sCOA2Msg = "Account "+sBadCOA2+" not found"
					sErrName = sERRCOANOTFOUND
'				Case -2
				Case ERROUTOFROOM
					sCOA2Msg = "Account "+sBadCOA2+" month "+sMonth+" not enough rows"_
								+" to insert"
					sErrName = sERROUTOFROOM
'				Case -3
				Case ERRNOCOAMONTH
					sCOA2Msg = "Account "+sBadCOA2+" month "+sMonth+" not found"
					sErrName = sERRNOCOAMONTH
	
					sCOA2Msg = "Account "+sBadCOA2+"Undocumented error"
					sErrName = sERRUNKNOWN
				end Select
				
				sErrMsg = sCOA2Msg
				Call LogError(sErrName, sErrMsg)	'// log error and continue
				
			endif		'// end bad line 2 insert row conditional
			
			iStatus = ERRCONTINUE			'// flag error to continue with next transaction
'			msgBox(sCOA1Msg+CHR(13)+CHR(10) + sCOA2Msg)
			'// advance to next transaction
			GoTo AdvanceTrans
			
 		endif	'// end problem with either COA insert conditional
 		
 		'// no problems, set returned insertion rows
' 		rplIns1Ptr = lCat1InsRow
'		rplIns2Ptr = lCat2InsRow
 		rpoCat1Range.StartRow = lCat1InsRow
 		rpoCat2Range.StartRow = lCat2InsRow
 		rpoCat1Range.EndRow = lCat1InsRow
 		rpoCat2Range.EndRow = lCat2InsRow
 		rpoCat1Range.StartColumn = COLDATE
 		rpoCat2Range.StartColumn = COLDATE
 		rpoCat1Range.EndColumn = COLREF
 		rpoCat2Range.EndColumn = COLREF

AdvanceTrans:
'//*-----------------restore error handling for caller-----------
	'// restore entry error settings.
	ErrLogSetModule(sErrCurrentMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
'//*----------end error handling setup---------------------------

	'// iStatus = 0, no error
	'//				ERRCONTINUE (-1) - error; continue with next transaction
	'//				ERRSTOP (-2) - error; stop processing transactions
	iRetValue = iStatus				'// set returned status
	CheckInserts = iRetValue
	
end function 	'// end CheckInserts	7/3/20
'/**/

'// CheckSplitTrans.bas
'//---------------------------------------------------------------
'// CheckSplitTrans - Check split transaction for validity.
'//		7/4/20. 	wmk.	10:45
'/---------------------------------------------------------------

function CheckSplitTrans(poGLSheet as Object, plSplit1stRow as Long )_
		as Long

'//	Usage.	lSplitEndRow = CheckSplitTrans(oGLSheet, lSplit1stRow)
'//			call PlaceSplitTrans( oGLSheet, lSplit1stRow )
'//
'//		oGLSheet = GL sheet object
'//		lSplit1stRow = row index of first row of split transaction
'//
'// Entry.	GL sheet processing has found row with "split" as the
'//			.Text.String value in COLSPLIT. The split entry contains
'//			multiple rows until another row is found with "split" as the
'//			.Text.String value in COLSPLIT. The second row of the split
'//			transaction MUST contain the COA target and the total amount
'//			of the split, regardless of whether it is a Debit or Credit.
'//			Subsequent rows contain the amounts split into the opposing
'//			COAs, and their value sums MUST equal the COA target amount.
'//
'//	Exit.	lSplitEndRow = last row of split, "split" in COLSPLIT
'//						< 0 - error in split transaction
'//						-1 - (ERRNOTSPLIT) first row not "split" in COLSPLIT
'//						-2 - (ERRBADTALLY) split values don't add to total
'//						-3 - (ERRROWSEXCEEDED) too many/few rows before second "split"
'//						-4 - (ERRBADVALS) bad Debit/Credit values
'//						-5 - (ERRBADDATE) Date field too short
'//						-6 - (ERRDATENOTSAME) Date field doesn't match others
'//						-7 - (ERRSPLITDESC) "split" descriptions don't match
'//
'// Calls.	ErrLogGetModule, ErrLogGetCellInfo, ErrLogGetSheet,
'//			ErrLogSetSheet, ErrLogSetModule, ErrLogSetCellInfo,
'//			LogError
'//
'//	Modification history.
'//	---------------------
'//	5/22/20		wmk.	original code
'//	5/23/20		wmk.	bug fix where not tallying properly; fix bug where
'//						empty Date field within loop range checking rows
'//	5/24/20.	wmk.	documentation added on .Type cell property; modify
'//						bad row Date cell with own content to update color
'//	5/27/20.	wmk.	include new error handling capability
'//	5/28/20.	wmk.	error handling improvements and documentation
'//	6/5/20.		wmk.	added check to ensure that transaction desc
'//						matches in both "split" lines; added check to
'//						ensure that all dates are the same in all
'//						transaction lines; cell addresses updated for
'//						all error reporting conditions
'//	7/4/20.		wmk.	double check tally value mismatch with strings to
'//						correct issue where 9.62 did not match the sum
'//						of 4.81 and 4.81
'//
'//	Notes.
'// ensure current row has "split" in COLSPLIT field; return -1 if not
'// determine how many rows are in the "split", searching for next
'//  row with "split"; set some sort of maximum so don't search forever
'//  if maximum exceeded return -3
'// ensure that total values form split rows = total in 1st row
'//  if not, return -2
'// set up cat1 sheet to point to total's category/account sheet
'// loop processing split rows, setting up cat2 sheet to point to
'//   each category/account sheet
'// on successful exit, point to last row with "split" in COLSPLIT field
'// Programming notes. oCell.Type returns the cell type {EMPTY,VALUE,TEXT
'// FORMULA). The Value, String, and Formula properties are used to
'// set the values of a cell.

'//	constants.
'// local contants. (common with PlaceSplitTrans)
const MAXSPLITROWS=10	'// maximum rows in "split" transaction
const MINROWLIMIT=1		'// 0-based index minimum total row index
const MINDATELEN=6		'// minimum Date string length
const ERRNOTSPLIT=-1	'// first row not "split" in COLSPLIT
const ERRBADTALLY=-2	'// split values don't add to total
const ERRROWSEXCEEDED=-3	'// too many/few rows before second "split"
const ERRBADVALS=-4		'// bad Debit/Credit values
const ERRBADDATE=-5		'// Date field too short
const ERRDATENOTSAME=-6	'// Date field doesn't match others			'// mod060520
const ERRSPLITDESC=-7	'// "split" descriptions don't match		'// mod060520
const COLSPLIT=5		'// "split" column
const csSplit="split"	'// split row identifier string


'//	local variables.
dim oCellSplit	as Object	'// cell from COLSPLIT column
dim oGLSheet	as Object	'// local ptr to poGLSheet
dim oCellCredit as Object		'// credit field
dim oCellDebit 	as Object		'// debit field		

'// other local vars.
dim iBrkPt as integer		'// easy breakpoint reference
dim lRetValue as long		'// returned value
dim bBadSplit as Boolean	'// bad split flag
dim sSplit As String		'// COLSPLIT .Text.String
dim lSplitCount as long		'// split row count
dim i as integer			'// loop counter
dim lGLCurrRow as Long			'// GL current row
dim bFoundSplit as boolean	'// found "split" 2nd time
dim lTrans1stRow as long	'// transaction 1st row
dim dTransTotal as Double		'// split entered total
dim dSplitTally as Double		'// split tallied total
dim dDebitVal as Double		'// Debit value
dim dCreditVal as Double	'// Credit value
dim bDebitSplits as boolean		'// first line is Debit flag
dim sDebit as String		'// debit value string
dim sCredit as String		'// credit value string
dim lGLStartRow as Long		'// start row
dim dAddAmt	as Double		'// amount to add to tally
dim sSourceAcct as String	'// source account being split
dim oCellDate as Object		'// cell Date from transaction row
dim sDate As String			'// cell Date string from transaction	'// mod060520
dim sDateSplit As String	'// Date field from first "split" line	'// mod060520
dim oCellTrans As Object	'// cell Transaction desc from row		'// mod060520
dim sTrans As String		'// transaction from end "split" line	'// mod060520
dim sTransSplit As String	'// Transaction from 1st "split" line	'// mod060520
dim bDateMatched As Boolean	'// Date field(s) all match flag		'// mod060520
dim sErrMsg as String		'// error message
dim sErrCode As String		'// error code string

	'// code.

	lRetValue = ERRNOTSPLIT	'// set error return

'//*---------end error handling setup---------------------------
'//	ErrLogGetModule() - get current module name gsErrModule
'// ErrLogSetModule(sName) - set current module name gsErrModule
'// ErrLogGetCellInfo(lColumn, lRow) - get error focus cell column, row
'// ErrLogSetCellInfo(lColumn, lRow) - set error focus cell column, row
'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
dim iErrSheetIx as Integer
dim lErrColumn as Long
dim lErrRow as Long
dim sErrCurrentMod As String
'//*
	'// preserve entry error settings.
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	sErrCurrentMod = ErrLogGetModule()
'//*	
	'// set local error settings.
	ErrLogSetModule("CheckSplitTrans")
	ErrLogSetSheet(poGLSheet.RangeAddress.Sheet)
	ErrLogSetCellInfo(COLSPLIT, plSplit1stRow)
'//*----------end error handling setup---------------------------

	'// ensure current row has "split" in COLSPLIT field;
	'//	 return ERRNOTSPLIT if not
	lGLCurrRow = plSplit1stRow
	lGLStartRow = plSplit1stRow	
	lTrans1stRow = lGLCurrRow + 1	'// transaction first data row
'	oGLSheet = poGLSheet
	oCellSplit = poGLSheet.getCellByPosition(COLSPLIT, lGLCurrRow)
	sSplit = oCellSplit.Text.String
	bBadSplit = (StrComp(sSplit, csSplit) <> 0)
	lRetValue = ERRNOTSPLIT			'// set initial error	
iBrkPt=1

	if bBadSplit then
		lRetValue = ERRNOTSPLIT
		GoTo Bailout
	endif	'// end not "split"

	'// preserve Date field from "split" 1st row					'// mod060520
	oCellDate = poGLSheet.getCellByPosition(COLDATE, lGLCurrRow)	'// mod060520
	sDateSplit = oCellDate.String									'// mod060520
	'// date must be at least #/#/## long							'// mod060520
	if Len(sDateSplit) < MINDATELEN then							'// mod060520
ErrLogSetCellInfo(COLDATE, lGLCurrRow)									'// mod060520
		lRetValue = ERRBADDATE										'// mod060520
		GoTo BailOut												'// mod060520
	endif	'// end date too short conditional						'// mod060520
	
	bDateMatched = true		'// for AND in loop to test failure		'// mod060520

	'// preserve Transaction field from "split" 1st row				'// mod060520
	oCellTrans = poGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)	'// mod060520
	sTransSplit = oCellTrans.String	'// Transaction desc			'// mod060520

	'// check Date in 1st transaction row
	oCellDate = poGLSheet.getCellByPosition(COLDATE, lTrans1stRow)	'//	mod060520
	bDateMatched = bDateMatched AND _
					  (StrComp(oCellDate.String, sDateSplit)=0)	'// mod060520
	if NOT bDateMatched then										'// mod060520
ErrLogSetCellInfo(COLDATE, lTrans1stRow)								'// mod060520
		lRetValue = ERRDATENOTSAME									'// mod060520
		GoTo BailOut												'// mod060520
	endif
	
	'// load debit, and credit, cells from 1st transaction line
ErrLogSetCellInfo(COLDEBIT, lTrans1stRow)
	oCellDebit = poGLSheet.getCellByPosition( COLDEBIT, lTrans1stRow)
	oCellCredit = poGLSheet.getCellByPosition( COLCREDIT, lTrans1stRow)
		
	'// declare tally vars as double dTransTotal, dSplitTally
	sDebit = oCellDebit.Text.String
	sCredit = oCellCredit.Text.String

	'// verify that one or the other is empty and set flag
	bDebitSplits = (Len(sCredit)=0) 		'// debit line is split
	bBadSplit = (Len(sDebit)=0) AND (Len(sCredit)=0)	'// one MUST be nonempty
		
	'// check for entries in both Debit and Credit columns
	if bDebitSplits then
		bBadSplit = bBadSplit OR Len(sCredit)>0		
	else	
		bBadSplit = bBadSplit OR Len(sDebit)>0
	endif	'// end debit splits conditional
		
	if bBadSplit then
		lRetValue = ERRBADVALS
		GoTo BailOut
	endif

	if bDebitSplits then
		dDebitVal = oCellDebit.getValue()		'// look at debit and Credit
		dTransTotal = dDebitVal
	else
		dCreditVal = oCellCredit.getValue()
		dTransTotal = dCreditVal
	endif
		
	dSplitTally = 0.			'// clear split tally

	'// loop on split transaction rows, searching for 2nd "split"
	'// and recording amount from 1st row as total,
	'// then tallying amounts from subsequent rows and
	'// comparing tally with 1st row value
	'// credit amount, and COA account number cells
	'//	when encounter "split", set lRetVal to row index
	
'//	loop on transaction rows looking for "split"

	lGLCurrRow = lTrans1stRow + 1	'// skip over 1st row - already processed
	'// loop beyond MAXSPLITROWS to determine if "split" out of bounds
	for i = 0 to MAXSPLITROWS
ErrLogSetCellInfo(COLSPLIT, lGLCurrRow)

		oCellSplit = poGLSheet.getCellByPosition(COLSPLIT, lGLCurrRow)
		sSplit = oCellSplit.Text.String
		bFoundSplit = (StrComp(sSplit, csSplit) = 0)
		
		if bFoundSplit then

			'// check that last date matches the rest
			oCellDate = poGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
			sDate = oCellDate.String
			bDateMatched = bDateMatched AND _
					 (StrComp(oCellDate.String, sDateSplit)=0)		'// mod060520
			if NOT bDateMatched then								'// mod060520
ErrLogSetCellInfo(COLDATE, lGLCurrRow)									'// mod060520
				lRetValue = ERRDATENOTSAME							'// mod060520
				GoTo BailOut										'// mod060520
			endif

			'// check that the last "split" description matches		'// mod060520
			oCellTrans = poGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
			sTrans = oCellTrans.String								'// mod060520
			if StrComp(sTrans, sTransSplit) <> 0 then				'// mod060520
ErrLogSetCellInfo(COLTRANS, lGLCurrRow)								'// mod060520
				lRetValue = ERRSPLITDESC							'// mod060520
				GoTo BailOut										'// mod060520
			endif
			
			'// check not enough rows to be a valid split			
			if i < MINROWLIMIT then
ErrLogSetCellInfo(COLDATE, lTrans1stRow)								'// mod060520
				lRetValue = ERRROWSEXCEEDED
				GoTo BailOut
			endif
if true then
GoTo ExitFor
endif

			'// check tally against total
			
			if dSplitTally <> dTransTotal then
ErrLogSetCellInfo(COLREF, lGLCurrRow)								'// mod060520
				lRetValue = ERRBADVALS
				GoTo BailOut
			endif
ExitFor:			
			lRetValue = lGLCurrRow
			exit For		'// leave For loop

			
		else	'// not found	
			'// not found yet, check max exceeded			
			if i >= MAXSPLITROWS then
				lRetValue = ERRROWSEXCEEDED
				GoTo BailOut
			endif
			
			'// if empty Date field, is error
			oCellDate = poGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
		ErrLogSetCellInfo(COLDATE, lGLCurrRow)

			sDate = oCellDate.String		'// get date			'// mod060520
			if Len(sDate) < MINDATELEN then
				lRetValue = ERRBADDATE
				GoTo BailOut
			endif

			bDateMatched = bDateMatched AND _
					 (StrComp(sDate, sDateSplit)=0)				'// mod060520
			if NOT bDateMatched then								'// mod060520
				lRetValue = ERRBADDATE								'// mod060520
				GoTo BailOut										'// mod060520
			endif													'// mod060520
	

			'// check Debit and Credit columns; one MUST have a nonempty string
						
			'// increment tally value from Debit or Credit
			if bDebitSplits then	'// Debit splits - tally Credits
	ErrLogSetCellInfo(COLCREDIT, lGLCurrRow)
				oCellCredit = poGLSheet.getCellByPosition( COLCREDIT, lGLCurrRow)
				bBadSplit = (Len(oCellCredit.String) = 0)
				dAddAmt = oCellCredit.getValue()
			else	''// Credit splits; tally Debits
	ErrLogSetCellInfo(COLDEBIT, lGLCurrRow)
				oCellDebit = poGLSheet.getCellByPosition( COLDEBIT, lGLCurrRow)
				dAddAmt = oCellDebit.getValue()
				bBadSplit = (Len(oCellDebit.String) = 0)
			endif
			
			'// bad split if opposing column of transaction total is empty
			if bBadSplit then
				lRetValue = ERRBADVALS
				GoTo BailOut
			endif
						
			dSplitTally = dSplitTally +	dAddAmt		'// add value to tally
			
		endif	'// end split found conditional

		lRetValue = lGLCurrRow		'// point to current row
		lGLCurrRow = lGLCurrRow + 1	'// advance to next entry
		
	next i
	
	'// normal return if tally matches transaction total
	if dSplitTally = dTransTotal then
		GoTo ExitNormal		'// lRetVal = lGLCurrRow
	else		'// tally mismatch on values, but check strings
		if StrComp(Str(dSplitTally), Str(dTransTotal))=0 then
			GoTo ExitNormal
		endif
	endif 	'// end valid tally conditional
	
	lRetValue = ERRBADVALS
	
'//----------- Error handler--------------------	
BailOut:
	if lRetValue >= 0 then
		GoTo ExitNormal
	endif

	'// lGLCurrRow points to row in which error occurred
	oCellDate = poGLSheet.getCellByPosition( COLDEBIT, lGLCurrRow)
	oCellDate.String = oCellDate.String		'// modify cell with own content
	oCellDate.CellBackColor = YELLOW

	'// ErrLog has corrupted cell column, row; change to YELLOW
dim lCurrErrCol as Long
dim lCurrErrRow As Long
ErrLogGetCellInfo(lCurrErrCol, lCurrErrRow)
	oCellDate = poGLSheet.getCellByPosition(lCurrErrCol, lCurrErrRow)
	oCellDate.String = oCellDate.String
	oCellDate.CellBackColor = YELLOW
	
	Select Case lRetValue
		Case ERRNOTSPLIT
			sErrMsg = "Transaction not 'split'..."_
					+ CHR(13)+CHR(10)+"stopped."
			sErrCode = "ERRNOTSPLIT"
			
		Case ERRBADVALS
			sErrMsg = "Bad Debit/Credit field(s); check values"_
					+ CHR(13)+CHR(10)+"stopped."
			sErrCode = "ERRBADVALS"
					
		Case ERRBADTALLY
			sErrMsg = "Debit and Credit values do not balance..."_
					+ CHR(13)+CHR(10)+"stopped."
			sErrCode = "ERRBADTALLY"
			
		Case ERRROWSEXCEEDED
			sErrMsg = "Too few or too many rows in 'split'..."_
					+ CHR(13)+CHR(10)+"stopped."
			sErrCode = "ERRROWSEXCEEDED"

		Case ERRBADDATE
			sErrMsg = "Date field too short or empty..."_
					+ CHR(13)+CHR(10)+"stopped."			
			sErrCode = "ERRBADDATE"

		Case ERRDATENOTSAME	'// Date field doesn't match others			'// mod060520
const csDATENOTSAME="ERRDATESNOTSAME"
			sErrCode = csDATENOTSAME
			sErrMsg = "Split - one or more dates don't match"
		
		Case ERRSPLITDESC	'// "split" descriptions don't match		'// mod060520
const csSPLITDESC="ERRSPLITDESC"
			sErrCode = csSPLITDESC
			sErrMsg = "Split - 1st and last split descriptions don't match"

		Case else
			sErrMsg = "Split transaction funky; check it..."_
					+ CHR(13)+CHR(10)="stopped."
			sErrCode = "ERRUNKNOWN"

		End Select
		
	Call LogError(sErrCode, sErrMsg)
'	msgBox(sErrMsg)		'// issue error message
'XRay oCellSplit

ExitNormal:

'//*-----------------restore error handling for caller-----------
	'// restore entry error settings.
	ErrLogSetModule(sErrCurrentMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
'//*----------end error handling setup---------------------------
			

	CheckSplitTrans = lRetValue
	
end function 	'// end CheckSplitTrans	7/4/20
'/**/

'// CreateReverseTrans.bas
'//---------------------------------------------------------------
'// CreateReverseTrans - Create reversal transaction.
'//		6/28/20.	wmk.	23:15
'//---------------------------------------------------------------

public function CreateReverseTrans(poGLSheet As Object, _
					plStartRow As Long, plRowCount As Long) as Integer

'// Usage.	iVal = CreateReverseTrans(oGLSheet, lStartRow, lRowCount)
'//
'//		oGLSheet = ledger sheet object
'//		lStartRow = first row of transaction; if "split", 1st "split" row
'//		lRowCount = last row of transaction; if "split", last "split" row
'//
'// Entry.	lStartRow points to the first row of either a "split" or
'//			double entry transaction.
'//			lRowCount = the total row count of the transaction
'//
'//	Exit.	iVal = 0 - normal exit
'//				 = -1 - reverse transaction issue
'//			if normal exit, original transaction is duplicated, with
'//			the reversal transaction below it in the ledger sheet
'//
'// Calls.	none.
'//
'//	Modification history.
'//	---------------------
'//	6/12/20.	wmk.	original code
'//	6/28/20.	wmk.	ErrLogSetModule call corrected
'//
'//	Notes. 
'//

'//	constants.
const COLSPLIT=5			'// column index looking for "split"
const ERRREVERSE=-1			'// error in reverse transaction
const sERRREVERSE="ERRREVERSE"

'//	local variables.
dim iRetValue As Integer	'// returned value
dim iSheetIx As Integer		'// GL sheet index

'//	CellRange content variables for GL sheet.
dim oGLRange as Object		'// RangeAddress object of two rows processing
dim lGLStartRow as long
dim lGLEndRow as long
dim lGLCurrRow as long		'// current row in GL
dim lGLOrigRow As Long		'// original transaction row index
dim oDateCell As Object		'// Date field cell
dim lRowCount As Long		'// local row count

'// error handling setup------------------------------------------------
const sMyName="CreateReverseTrans"
dim sErrCode as String	'// error code string for user code
Dim sErrMsg As String	'// error message for LogError
dim lLogRow as long		'// cell row working on
dim lLogCol as Long		'// cell column working on
dim iLogSheetIx as Integer	'// sheet index module working on

dim iErrSheetIx As Integer	'// entry error sheet index
dim lErrColumn As Long		'// entry error column ptr
dim lErrRow As Long			'// entry error row
dim sErrCallerMod As String	'// caller module name
'// end error handling setup--------------------------------------------


	'// code.
	
	iRetValue = -1			'// set abnormal return
	ON ERROR GOTO ErrorHandler
'//*
	'// preserve entry error settings.
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	sErrCallerMod = ErrLogGetModule()
'//*	
	'// set local error settings.
	ErrLogSetModule(sMyName)
	ErrLogSetSheet(poGLSheet.RangeAddress.Sheet)
	ErrLogSetCellInfo(COLSPLIT, plStartRow)
'//*----------end error handling setup---------------------------
	
'//---------------CreateReverseTrans starts here -----------------------
'// CreateReverseTrans starts here
'// Usage.	iVal = CreateReverseTrans(poGLSheet, lStartRow, lRowCount)
'//
'//		poGLSheet = ledger sheet object
'//		lStartRow = first row of transaction; if "split", 1st "split" row
'//		lRowCount = last row of transaction; if "split", last "split" row

	'// insert lRowCount new rows above 1st transaction row in ledger
	iSheetIx = poGLSheet.RangeAddress.Sheet
	lGLStartRow = plStartRow
	lRowCount = plRowCount
	poGLSheet.Rows.insertByIndex(lGLStartRow, lRowCount)	'// insert new transaction rows
	lGLOrigRow = lGLStartRow + lRowCount				'// old rows moved down
	
	'// copy old transaction to new rows above.
dim oTargetCellAddress As new com.sun.star.table.CellAddress
	oTargetCellAddress.Sheet = iSheetIx
	oTargetCellAddress.Column = COLDATE
	oTargetCellAddress.Row = lGLStartRow
		
dim oGLCellRange As new com.sun.star.table.CellRangeAddress
	oGLCellRange.Sheet = iSheetIx
	oGLCellRange.StartColumn = COLDATE
	oGLCellRange.EndColumn = COLREF
	oGLCellRange.StartRow = lGLOrigRow
	oGLCellRange.EndRow = lGLOrigRow + lRowCount - 1

	poGLSheet.copyRange(oTargetCellAddress, oGLCellRange)

	'// now swap Debit and Credit columns in 2nd double-entry
	'// to emulate transaction reversal
	'// copy Debit Column to Credit column below first
	oGLCellRange.StartColumn = COLDEBIT
	oGLCellRange.StartRow = lGLStartRow
	oGLCellRange.EndColumn = COLDEBIT
	oGLCellRange.EndRow = lGLStartRow + lRowCount - 1
	
	oTargetCellAddress.Column = COLCREDIT
	oTargetCellAddress.Row = lGLOrigRow
	poGLSheet.copyRange(oTargetCellAddress, oGLCellRange)
	
	'// copy Credit column to Debit column below
	
	oGLCellRange.StartColumn = COLCREDIT
	oGLCellRange.StartRow = lGLStartRow
	oGLCellRange.EndColumn = COLCREDIT
	oGLCellRange.EndRow = lGLStartRow + lRowCount - 1
	
	oTargetCellAddress.Column = COLDEBIT
	oTargetCellAddress.Row = lGLOrigRow
	poGLSheet.copyRange(oTargetCellAddress, oGLCellRange)

'//end CreateReverseTrans-----------------------------------------
	iRetValue = 0		'// set normal return
	
NormalExit:
	'// restore caller error handling setup
	'// set local error settings.
	ErrLogSetModule(sErrCallerMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
	
	CreateReverseTrans = iRetValue
	exit function
	
ErrorHandler:
	'// log error and highlight date field in 1st row
	sErrCode = sERRREVERSE
	sErrMsg = "Error in reverse transaction..." _
			+ CHR(13)+CHR(10) + "Reverse transaction failed."
	Call LogError(sErrCode, sErrMsg)
	oDateCell = poGLSheet.getCellByPosition(COLDATE, lGLStartRow)
	oDateCell.String = oDateCell.String
	oDateCell.CellBackColor = YELLOW
	GoTo NormalExit
	
end function 	'// end CreateReverseTrans	6/28/20
'/**/

'// D.bas
'//------------------------------------------------------------------
'// D - stub function to enable calls to development subs/functions.
'//		5/27/20.	wmk.
'//------------------------------------------------------------------

public function D() as void

'//	Usage.	=D()
'//
'// Entry.	public const DEVDEBUG set = true/false
'//			most tests coded here depend upon the =D() reference being
'//			in the same row as the data to be tested, since the called
'//			subs/functions will look at the current range selection.
'//			When the user presses <Enter> after completing the formula
'//			referencing D(), the range selection is the cell in which
'//			the formula occurs.
'//
'//	Exit.	varies according to Calls.
'//			if DEVDEBUG=false, exits without doing anything
'//
'//	Calls.	varies according to test needed.
'//			PlaceSplitTrans
'/
'//	Modification history.
'//	---------------------
'//	5/13/20		wmk.	original code
'//	5/16/20.	wmk.	documentation; test calls to GetTransSheetName
'//	5/22/20.	wmk.	test call to PlaceSplitTrans
'//	5/27/20.	wmk.	DEVDEBUG constant used to activate/deactivate
'//
'//	Notes. If DEVDEBUG=true, D acts to test whatever functions/subs
'// its code is set up for. Usually the <Insert notes here>
'//

'// local variables. (defined as needed)
dim Doc as object		'// current component
dim oSel as object		'// current selection
dim GLSheet as object	'// current sheet selected
dim oRange as object	'// current selection within sheet
dim oSheets As Object		'// sheets array of objects
Dim oGLSheet As Object		'// GL sheet object
dim iSheetIx as Integer		'// sheet index
dim sGLSheet as String		'// GL sheet name 	

dim	lGLStartRow	as long
dim	lGLEndRow	as long
dim	lGLRowCount	as long
dim	lGLNewRow	as long
dim	lGLCurrRow	as long
dim	lNRowsSelected	as long
dim	lTransCount	as long
dim	bOddRowCount	as boolean

dim iBrkpt as integer
dim lSplitEnd	as long		'// end row of split transaction
dim oCellDate	as Object

	'// code.
	'// NEVER REMOVE THIS CONDITIONAL.
	if NOT DEVDEBUG then
		Exit Function
	endif
	
	Doc = ThisComponent
	oSel = Doc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet				'// get sheet index value
	oGLSheet = Doc.Sheets(iSheetIx)		'// set up sheet object
'XRay oGLSheet		'// xray to get .sun. reference
	sGLSheet = oGLSheet.CodeName			'// get this sheet name

iBrkpt=1	

	'// set row bounds from selection information
	lGLStartRow = oRange.StartRow
	lGLEndRow = oRange.EndRow
	lGLRowCount = lGLEndRow-lGLStartRow
	lGLNewRow = lGLStartRow - 2			'// kludge new row for loop
	lGLCurrRow = lGLStartRow - 2		'// current GL processing row preset for increment
	lNRowsSelected = lGLEndRow+1 - lGLStartRow
	lTransCount = lNRowsSelected/2	'// get transaction count (2 rows per)
	bOddRowCount = lTransCount*2 <> lNRowsSelected

	lSplitEnd = PlaceSplitTrans(oGLSheet, lGLStartRow)
	
	if lSplitEnd < 0 then
		oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLStartRow)
		oCellDate.Text.CellBackColor = YELLOW
		msgBox("Split transaction - bad transaction; check data")
	endif
	
'	lSplitEnd = PlaceSplitTrans(oGLSheet, lGLStartRow)	'// check static var
	
iBrkPt=1
				
end function	'// end D
'/**/

'// DayName.bas
'// DayName - Return text name of day from week day number
public Function DayName(nDay as integer) as string

'//	Modification history.
'//	---------------------
'//	12/8/15.	wmk.	Added "??" return value if nDay type is unkown.

'DayName = "Monday"
dim strDay as string

'//	Code.

select case nDay

case 1
	strDay = "Sunday"
case 2
	strDay ="Monday"
case 3
	strDay = "Tuesday"
case 4
	strDay = "Wednesday"
case 5
	strDay = "Thursday"
case 6
	strDay = "Friday"
case 7
	strDay = "Saturday"
	
case else
	strDay = "??"
	
end select


DayName = strDay
end function	'// end Dayname
'/**/

'// E.bas
'//------------------------------------------------------------------
'// E - stub function to test error handling capabilities.
'//		5/27/20.	wmk.
'//------------------------------------------------------------------

public function E() as void

'//	Usage.	=E()
'//
'// Entry.	public const DEVDEBUG set = true/false
'//			most tests coded here depend upon the =D() reference being
'//			in the same row as the data to be tested, since the called
'//			subs/functions will look at the current range selection.
'//			When the user presses <Enter> after completing the formula
'//			referencing D(), the range selection is the cell in which
'//			the formula occurs.
'//
'//	Exit.	varies according to Calls.
'//			if DEVDEBUG=false, exits without doing anything
'//
'//	Calls.	varies according to test needed.
'//			ErrLogSetup, LogError, ErrLogDisable
'/
'//	Modification history.
'//	---------------------
'//	5/25/20.	wmk.	original code
'//	5/27/20.	wmk.	DEVDEBUG constant used to activate/deactivate
'//
'//	Notes. If DEVDEBUG=true, D acts to test whatever functions/subs
'// its code is set up for. Usually the <Insert notes here>

'// local variables. (defined as needed)
dim Doc As Object			'// this component
dim oSel As Object			'// current selection
dim oRange As Object		'// CellRangeAddress object

	'// code.
	if NOT DEVDEBUG then
		Exit Function
	endif
	
	Doc = ThisComponent
	oSel = Doc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	Call ErrLogSetup(oRange, "Test sub E")
	Call LogError("TESTERROR", "Test sub E - error")
	Call ErrLogDisable
					
end function	'// end E
'/**/

'// ErrLogDisable.bas
'//---------------------------------------------------------------
'// ErrLogDisable - Disable error logging.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public Function ErrLogDisable() As Void

'//	Usage.	macro call or
'//			call ErrLogDisable
'//
'// Entry.	<entry conditions>
'//
'//	Exit.	gbErrLogging = false
'//			goErrRangeAddress object released
'//			gsErrMsg = ""
'//			gsErrName = ""
'//			gsErrModule = ""
'//			gsErrSheet = ""
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	5/25/20.	wmk.	original code
'// 5/26/20.	wmk.	change from sub to function

'//	constants.

'//	local variables.

	'// code.
	'// if not Basic, add code to release goErrRangeAddress conditionally
	
	gbErrLogging = false
	gsErrMsg = ""
	gsErrName = ""
	gsErrModule = ""
	gsErrSheet = ""
	
end Function		'// end ErrLogDisable
'/**/

'// ErrLogGetCellInfo.bas
'//-----------------------------------------------------------------
'// ErrLogGetCellInfo - Get cell information from error log globals.
'//		wmk. 5/26/20.
'//-----------------------------------------------------------------

public Function ErrLogGetCellInfo(rplColumn As Long, rplRow As Long) As Void

'//	Usage.	ErrLogGetCellInfo(lColumn, lRow)
'//
'// Entry.	goErrLogRange = current cell tracking
'//					.StartColumn = starting column index
'//					.StartRow = starting row index
'//
'//	Exit.	lColumn = goErrLogRange.StartColumn
'//			lRow = goErrLogRange.StartRow
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. Complemented by ErrLogSetInfo(lColumn, lRow)
'//

'//	local variables

	'// code.
	rplColumn = goErrRangeAddress.StartColumn
	rplRow = goErrRangeAddress.StartRow
	
end Function		'// end ErrLogGetCellnfo
'/**/

'// ErrLogGetCellName.bas
'//---------------------------------------------------------------
'// ErrLogGetCellName - Get cell name from error log globals.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public Function ErrLogGetCellName() As String

'//	Usage.	sCellName = ErrLogGetCellName()
'//
'// Entry.	goLogRange = current cell tracking
'//
'//	Exit.	sCellName = "XYY" where
'//				X = column letter converted from goLogRange.StartColumn
'//				Y = row number converted from goLogRange.StartRow
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code; stub
'//
'//	Notes. Complemented by ErrLogGetModule()
'//

'//	local variables
dim sRetStr As String		'// returned cell string

	'// code.
	sRetStr = ""
	ErrLogGetCellName = sRetStr	'// return current name
	
end Function		'// end ErrLogGetCellName
'/**/

'// ErrLogGetDisplay.bas
'//---------------------------------------------------------------
'// ErrLogGetDisplay - Get msgBox error messaging on/off status.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public function ErrLogGetDisplay() As Boolean

'//	Usage.	bMsgDisplay = ErrLogGetDisplay()
'//
'// Entry.	gbErrDisplay = current msgBox error messaging status
'//
'//	Exit.	bMsgDisplay = gbErrDisplay
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'// Notes. gbErrDisplay flag is used by LogError to determine if
'//	error messages are displayed in a msgBox onscreen

	'// code.
	bRetVal = gbErrDisplay
	ErrLogGetDisplay = bRetVal

end function 	'// end ErrLogGetDisplay
'/**/

'// ErrLogGetModule.bas
'//---------------------------------------------------------------
'// ErrLogGetModule - Get module name in error log globals.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public Function ErrLogGetModule() As String

'//	Usage.	sCurrModule = ErrLogGetModule()
'//
'// Entry.	gsErrModule = current module name global
'//
'//	Exit.	sCurrModule = gsErrModule
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. Complemented by ErrLogGetModule()
'//

'//	local variables.

	'// code.
	ErrLogGetModule = gsErrModule		'// return current name
	
end Function		'// end ErrLogGetModule
'/**/

'// ErrLogGetRecording.bas
'//-----------------------------------------------------------------------
'// ErrLogGetRecording - Get error recording status from error log globals
'//		wmk. 5/26/20.
'//-----------------------------------------------------------------------

public function ErrLogGetRecording() As Boolean

'//	Usage.	bLogRecording = ErrLogGetRecording()
'//
'// Entry.	gbErrRecording = current log recording status
'//
'//	Exit.	bLogRecording = gbErrRecording
'//							true if log recording on
'//							false if log recording off
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. LogError uses the gbErrRecording flag to determine if Error Log
'// sheet (ERRLOGSHEET) entries are to be made.

'//	local variables.
dim bRetValue	as Boolean

	'// code.
	
	bRetValue = gbErrRecording
	ErrLogGetRecording = bRetValue

end function 	'// end ErrLogGetRecording
'/**/

'// ErrLogGetSheet.bas
'//---------------------------------------------------------------
'// ErrLogGetSheet - Get sheet index from error log globals.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public Function ErrLogGetSheet() As Integer

'//	Usage.	iSheetIx = ErrLogGetSheet()
'//
'//			iSheet = new sheet index to set in goErrRangeAddress.Sheet global
'//
'// Entry.	goErrRangeAddress.Sheet = current error focus sheet index
'//
'//	Exit.	iSheetIx = goErrRangeAddress.Sheet = new error focus sheet index
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. Complemented by ErrLogSetSheet(iSheet)
'//

'//	local variables.

	'// code.
	
	ErrLogGetSheet = goErrRangeAddress.Sheet	'// return global sheet index
	
end Function		'// end ErrLogGetSheet
'/**/

'// ErrLogMakeEntry.bas
'//---------------------------------------------------------------
'// ErrLogMakeEntry - record entry in ERRLOGSHEET spreadsheet.
'//		wmk. 6/6/20. 09:45
'//---------------------------------------------------------------

public function ErrLogMakeEntry() As Integer

'//	Usage.	iStatus = ErrLogMakeEntry()
'//
'// Entry.	error handling global values set
'//
'//			ERRLOGSHEET = (string) name of error log sheet
'//
'//	Exit.	iStatus = 0 if log entry made successfully
'//					< 0 if failed
'//
'// Calls.	SetSheetDate(oLogRange, FMTDATETIME)
'//
'//	Modification history.
'//	---------------------
'//	5/26/20.	wmk.	original code
'//	5/28/20.	wmk.	time stamp column added
'//	6/3/20.		wmk.	add date/time stamp to error log sheet heading.
'// 6/6/20.		wmk.	update SetSheetDate call with format spec
'//
'//	Notes. <Insert notes here>
'//	Sheet	Column	Row	Module	ERRCODE	Message

'//	constants.
const COLTIME=0			'// Time Stamp column 
const COLSHEET=1		'// Sheet column
const COLCOLUMN=2		'// Column column
const COLROW=3			'// Row column
const COLMODULE=4		'// Module column
const COLERRCODE=5		'// ERRCODE column
const COLERRMSG=6		'// Message column
const ERRINSROW=5		'// insertion row index
'const FMTDATETIME=50	'// numeric value of mm/dd/yyy hh:mm:ss spec

'//	local variables.
dim iBrkPt as Integer		'// easy breakpoint spotter
dim iRetValue As Integer	'// return value
dim iStatus As Integer		'// general status
dim oDoc as Object			'// ThisComponent
dim oRange as Object		'// range address of this document
dim oLogSheet as Object		'// sheet ERRLOGSHEET
dim oSheets as Object		'// Doc.Sheets()
dim iSheetIx as Integer		'// sheet index of ERRLOGSHEET
dim oSel as Object
dim sSheet As String		'// sheet name for ERRLOGSHEET
dim oCellSheet As Object	'// Sheet cell
dim oCellTime As Object		'// Time stamp cell
dim oCellColumn As Object	'// Column cell
dim oCellRow As Object		'// Row cell
dim oCellModule As Object	'// Module cell
dim oCellErrCode As Object	'// Error Code cell
dim oCellErrMsg As Object	'// Error Message cell
dim oRangeAdd As Object		'// range address object
dim oErrSheet As Object		'// sheet error occurred in

	'// code.
	
	iRetValue = -1	'// set initial failure
	ON ERROR GoTo BailOut
	oDoc = ThisComponent
	oSheets = oDoc.getSheets()			'// get sheet collection
	oSel = oDoc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet				'// get sheet index value
'	oGLSheet = Doc.Sheets(iSheetIx)		'// set up sheet object
'XRay oGLSheet		'// xray to get .sun. reference
'	sGLSheet = oGLSheet.CodeName			'// get this sheet name

'XRay oSheets	
	sSheet = ERRLOGSHEET
	oLogSheet = oSheets.getByName(sSheet)
	
	'// see if ERRLOGSHEET exists and set up insertion pointers for
	'// new row
	oRangeAdd = new com.sun.star.table.CellRangeAddress
	oRangeAdd.Sheet = oLogSheet.RangeAddress.Sheet
	oRangeAdd.StartColumn = COLTIME
	oRangeAdd.EndColumn = COLERRMSG
	oRangeAdd.StartRow = ERRINSROW
	oRangeAdd.EndRow = ERRINSROW
	
	'// insert new row into log sheet
	oLogSheet.insertCells(oRangeAdd,_
						com.sun.star.sheet.CellInsertMode.ROWS)
	
	'// set values in new row fields from global error variables
iBrkPt=1
	oCellTime = oLogSheet.getCellByPosition(COLTIME, ERRINSROW)		'// Time stamp cell
	oCellSheet = oLogSheet.getCellByPosition(COLSHEET, ERRINSROW)	'// Sheet cell
	oCellColumn = oLogSheet.getCellByPosition(COLCOLUMN, ERRINSROW)	'// Column cell
	oCellRow = oLogSheet.getCellByPosition(COLROW, ERRINSROW)	'// Row cell
	oCellModule = oLogSheet.getCellByPosition(COLMODULE, ERRINSROW)	'// Module cell
	oCellErrCode = oLogSheet.getCellByPosition(COLERRCODE, ERRINSROW)'// Error Code cell
	oCellErrMsg = oLogSheet.getCellByPosition(COLERRMSG, ERRINSROW)	'// Error Message cell
	oErrSheet = oSheets.getByIndex(goErrRangeAddress.Sheet)

	oCellTime.setValue(Now())					'// time stamp
	oCellTime.NumberFormat = FMTDATETIME		'// set number format
	oCellSheet.String = oErrSheet.CodeName		'// set sheet name
	oCellColumn.String = Str(goErrRangeAddress.StartColumn)
	oCellRow.String = Str(goErrRangeAddress.StartRow)
	oCellModule.String = gsErrModule
	oCellErrCode.String = gsErrName
	oCellErrMsg.String = gsErrMsg
	
	'// set sheet modification date in log sheet header
	iStatus = SetSheetDate(oRangeAdd, FMTDATETIME)
	
	'// set iRetValue = 0 for success

	iRetValue = 0		'// if make it here, success...

BailOut:

	if iRetValue < 0 then	'// if problem accessing ERRLOGSHEET
		msgBox("Log Sheet '"+ERRLOGSHEET+"' not found!")
	endif
	
	ErrLogMakeEntry = iRetValue

end function 	'// end ErrLogMakeEntry	6/3/20
'/**/

'// ErrLogSetCellInfo.bas
'//-----------------------------------------------------------------
'// ErrLogSetCellInfo - Set cell information in error log globals.
'//		wmk. 5/26/20.	16:00
'//-----------------------------------------------------------------

public Function ErrLogSetCellInfo(plColumn As Long, plRow As Long) As Void

'//	Usage.	ErrLogSetCellInfo(lColumn, lRow)
'//
'// Entry.	goLogRange = current cell tracking
'//					.StartColumn = starting column index
'//					.StartRow = starting row index
'//
'//	Exit.	lColumn = goLogRange.StartColumn
'//			lRow = goLogRange.StartRow
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. Complemented by ErrLogGetInfo(lColumn, lRow)
'//

'//	local variables

	'// code.
	goErrRangeAddress.StartColumn = plColumn
	goErrRangeAddress.StartRow = plRow
	
end Function		'// end ErrLogSetCellnfo
'/**/

'// ErrLogSetDisplay.bas
'//---------------------------------------------------------------
'// ErrLogSetDisplay - Set msgBox error messaging on/off.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public function ErrLogSetDisplay(pbDisplayOn As Boolean) As Void

'//	Usage.	ErrLogSetDisplay(bDisplayOn)
'//
'//		bDisplayOn = true to turn on msgBpx error messaging
'//					 false to turn off msgBox error messaging
'//
'// Entry.	gbErrDisplay = current msgBox error messaging status
'//
'//	Exit.	gbErrDisplay  = bDisplayOn
'//
'//	Modification history.
'//	---------------------
'//	5/26/20.	wmk.	original code
'//
'// Notes. gbErrDisplay flag is used by LogError to determine if
'//	error messages are displayed in a msgBox onscreen

	'// code.
	gbErrDisplay = pbDisplayOn

end function 	'// end ErrLogSetDisplay
'/**/

'// ErrLogSetModule.bas
'//---------------------------------------------------------------
'// ErrLogSetModule - Set module name in error log globals.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public Function ErrLogSetModule(psName as String) As Void

'//	Usage.	ErrLogSetModule(sName)
'//
'//			sName = new name to set in error module caller global
'//
'// Entry.	gsErrModule = current module name global
'//
'//	Exit.	gsErrModule = new error module caller name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. Complemented by ErrLogGetModule()
'//

'//	local variables.

	'// code.
	gsErrModule = psName	'// reset global module name
	
end Function		'// end ErrLogSetModule
'/**/

'// ErrLogSetRecording.bas
'//---------------------------------------------------------------
'// ErrLogSetRecording - Set error recording status in error log globals
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public function ErrLogSetRecording(pbRecordingOn as Boolean) As Void

'//	Usage.	ErrLogSetRecording(bRecordingOn)
'//
'//		bRecordingOn = true to turn on log recording in ERRLOGSHEET
'//					   false to turn off log recording in ERRLOGSHEET
'//
'// Entry.	gbErrRecording = current ERRLOGSHEET recording status
'//
'//	Exit.	gbErrRecording = bRecordingOn
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. LogError uses the gbErrRecording flag to determine if Error Log
'// sheet (ERRLOGSHEET) entries are to be made.
'//	Eventually, this should issue a message to the Error Log
'// sheet reporting the recording on/off event.

	'// code.
	gbErrRecording = pbRecordingOn

end function 	'// end ErrLogSetRecording
'/**/

'// ErrLogSetSheet.bas
'//---------------------------------------------------------------
'// ErrLogSetSheet - Set sheet index in error log globals.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public Function ErrLogSetSheet(piSheet as Integer) As Void

'//	Usage.	ErrLogSetModule(iSheet)
'//
'//			iSheet = new sheet index to set in goErrRangeAddress.Sheet global
'//
'// Entry.	goErrRangeAddress.Sheet = current error focus sheet index
'//
'//	Exit.	goErrRangeAddress.Sheet = new error focus sheet index
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. Complemented by ErrLogGetSheet()
'//

'//	local variables.

	'// code.
	goErrRangeAddress.Sheet = piSheet	'// reset global sheet index
	
end Function		'// end ErrLogSetSheet
'/**/

'// ErrLogSetup.bas
'//---------------------------------------------------------------
'// ErrLogSetup - set up error logging to ERRLOGSHEET.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public sub ErrLogSetup(poCellRange as Object, psModule as String)

'//	Usage.	macro call or
'//			call ErrLogSetup( oCellRange, sModule)
'//
'//			oCellRange = .CellRangeAddress object with following
'//							.Sheet = sheet index
'//							.StartColumn = starting column of range
'//							.EndColumn = ending column of range
'//							.StartRow = starting row of range
'//							.EndRow = ending row of range
'//			sModule = (string) name of sub/function for logging
'//
'// Entry.	none.
'//
'//	Exit.	global variables set for ErrLogError
'//			gbErrLogging = true
'//			gsErrModule = sModule
'//			gsErrMsg = ""
'//			gsErrName = ""
'//			gsErrSheet = ""
'//			goErrRangeAddress = oCellRange object values
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. This sub should be the first call from any sub/function
'//	desiring debug/error logging. The gbErrLogging var should be
'// tested before calling the LogError sub to ensure that ErrLogSetup
'//	has been called and errors within the error logging will be avoided.
'// 

'//	constants.

'//	local variables.
Static bBeenSetup as boolean	'// very 1st time flag
dim oDoc	As Object		'// this component
dim oSheets As Object		'// Sheets() array
dim oCurrSheet as Object	'// current sheet

	'// code.
	if IsNull(poCellRange) then
		msgBox("ErrLogSetup - null pointer passed for cell range.")
		GoTo BailOut
	endif

	if IsNull(psModule) then
		msgBox("ErrLogSetup - null string passed for module name.")
		GoTo BailOut
	endif

	if NOT bBeenSetup then
		bBeenSetup = true
		goErrRangeAddress = new com.sun.star.table.CellRangeAddress
	endif
		
	'// new module creates fresh instance
	if StrComp(psModule, gsErrModule) <> 0 then
'XRay goErrRangeAddress	'// no .Release method
		goErrRangeAddress = new com.sun.star.table.CellRangeAddress
	endif 		'// end new module conditional

	oDoc = ThisComponent
	oCurrSheet = oDoc.Sheets(poCellRange.Sheet)		'// set up sheet object
	gsErrSheet = oCurrSheet.CodeName				'// set up sheet name
	goErrRangeAddress.Sheet = oCurrSheet.RangeAddress.Sheet
	goErrRangeAddress.StartColumn = poCellRange.StartColumn
	goErrRangeAddress.EndColumn = poCellRange.EndColumn
	goErrRangeAddress.StartRow = poCellRange.StartRow
	goErrRangeAddress.EndRow = poCellRange.EndRow
	gbErrDisplay = true							'// display loggin on
	gbErrRecording = false						'// log sheet recording on/off flag
	gbErrLogging = true							'// enable error logging
	gsErrMsg = ""
	gsErrName = ""
	gsErrModule = psModule						'// initialize module name
	
Bailout:

end sub		'// end ErrLogSetup
'/**/

'// GetInsRow.bas
'//------------------------------------------------------------------
'// GetInsRow - Get insertion row given sheet, COA account, month.
'//		wmk. 6/8/20. 15:00
'//------------------------------------------------------------------

function GetInsRow( poSheet, psAcctCat, psMonth) as long

'//	Usage.	lInsRow = GetInsRow( oSheet, sAcctCat, sMonth)
'//
'//		oSheet = Sheet object from Doc.Sheets[]
'//		sAcctCat = accounting COA to search for
'//		sMonth = month to search for within category
'//
'//	Exit.	lInsRow = >= 0; row index at which to insert new transaction
'//							returned index is where COA and month found + INSBIAS
'//					= -1 COA not found at all
'//					= -2 COA and month found, but not enough room to insert
'//					= -3 COA found, but month not found before COA changed
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	5/17/20		wmk.	original code; stub.
'//	5/18/20.	wmk.	code fleshed out
'//	5/20/20.	wmk.	vars compliant with EXPLICIT
'// 5/23/20.	wmk.	clarify calling sequence documentation
'// 6/8/20.		wmk.	mod to stop checking if have encountered a COA
'//						that is > COA looking for and flag not found

'//
'//	Notes. Does row-by-row search for line containing the COA account number
'// in Column A, and the month in ColumnB
'//

'//	constants. (see module-wide constants)

'//	local variables.
dim il			as Long		'// loop index
dim iBrkPt		as Integer	'// easy breakpoint marker
dim	j			as integer	'// short loop index
dim lRowIndex	as Long		'// returned row index
dim sAcctCat	as String	'// localized reference to psAcctCat
dim sMonth		as String	'// localized reference to psMonth
dim bCOAMatched	as Boolean	'// found COA flag
dim bNoMatch	as Boolean	'// this one not COA looking for flag

'// sheet information variables.
dim lStartRow	as Long		'// sheet starting row (should be 0)
dim lStartColumn	as Long		'// sheet starting column
dim lEndRow		as Long		'// sheet end row
dim lEndColumn		as Long		'// sheet end column
dim oRange		as Object	'// RangeAddress object
Static oCellDate	as Object
Static oCellTrans	as Object
dim sRowDate	as String	'// string from Date column
dim sRowTrans	as String	'// string from Transaction column
dim sScanDate	as String	'// Date field from current scan row
'dim oRow		as Object	'// row cells from search row index

	'// code.

	lRowIndex = -1		'// set not found
	sAcctCat = psAcctCat
	sMonth = psMonth
	oRange = poSheet.RangeAddress
	lStartRow = oRange.StartRow
	lStartColumn = oRange.StartColumn
	lEndRow = oRange.EndRow
	lEndColumn = oRange.EndColumn
	bCOAMatched = false		'// set COA not found yet	

iBrkpt = 1
'XRay poSheet
	
	'// for each row 0 to sheet last row
	for il = lStartRow to lEndRow

		'// search for month first to ease debugging, since there are so many
		'// less month matches...		

		'// check Transaction field for month match
		
		oCellTrans = poSheet.getCellByPosition(COLTRANS, il)		
		sRowTrans = oCellTrans.String

'// debugging
if Len(sRowTRans)<>Len(sMonth) then
iBrkPt=0
else
iBrkPt=1
endif
iBrkPt = StrComp(sRowTrans, sMonth)
'//---end debugging

		if StrComp(sRowTrans, sMonth) <> 0 then		'// if not month match, skip out
			GoTo AdvanceRow
		endif

iBrkPt=1	

		'// check Date field for COA account match
		oCellDate = poSheet.getCellByPosition(COLDATE, il)
		sRowDate = oCellDate.String
		if Len(sRowDate) <> COALEN then 	'// if not correct COA account length, skip out
			GoTo AdvanceRow
		endif
	
iBrkPt=1	
		bNoMatch = StrComp(sRowDate, sAcctCat) <> 0	
		if bNoMatch then	'// if not a match
'			if bCOAMatched then	'// but matched before, COA changed
'				lInsertRow = -3		'// no month match for account
'				exit For
'			else
'				GoTo AdvanceRow		'// keep looking for matching COA
'			endif	'// end found COA before condtional
			
			'// check to see if past COA looking for; if so, stop searching
			'// and flag COANOTOFUND
			if StrComp(sRowDate, sAcctCat) > 0 then					'// mod060820
				exit for											'// mod060820
			endif													'// mod060820
			
			
			GoTo AdvanceRow
		endif	'// end no match conditional

		bCOAMatched = true
		lRowIndex = il
		
		'// Check next rows to be within INSBIAS on this account number
		for j = 1 to INSBIAS
			oCellDate = poSheet.getCellByPosition(COLDATE, il+j)
			sScanDate = oCellDate.String
			if Len(sScanDate) = COALEN then		'// found another account
				lRowIndex = -2		'// regardless of account, not enough room for insert
				exit For
			endif
		next j		'// end look-ahead loop

		'// check if enough room to insert
		if lRowIndex >= 0 then
			lRowIndex = il + INSBIAS			'// have insertion row index
		endif
		exit For
		'//	if match exit loop with lRowIndex = index row + INSBIAS(=3)
		
AdvanceRow:	
	next il		'// end loop on sheet rows

	'// set return value
	GetInsRow = lRowIndex		

iBrkpt = 1	

end function 	'// end GetInsRow	6/8/20
'/**/

'// GetMonthName.bas
'//---------------------------------------------------------------
'// GetMonthName - return month name given date number.
'//		wmk. 5/20/20.
'//---------------------------------------------------------------

function GetMonthName(psDate as String) as String

'//	Usage.	sMonth = GetMonthName( sDate )
'//
'//		sDate = date string "mm/dd/yy" or compatible format
'//
'//	Exit.	sMonth = appropriate month from list
'//			"January", "February", "March", "April", "May", "June"
'//			"July", "August", "September", "October", "November", "December"
'//
'//	Modification history.
'//	---------------------
'//	5/15/20		wmk.	original code
'//	5/20/20.	wmk.	vars compliant with EXPLICIT
'//
'//	Notes. <Insert notes here>
'//

'//	constants. (see module-wide constants)

'//	local variables.

dim nMonth as integer	'// number of month
dim sName as String		'// returned name of month

	'// code.

	sName = ""	'// default to empty string
	nMonth = Month(DateValue(psDate))
	
	Select Case nMonth
	Case 1
		sName = MON1
	Case 2
		sName = MON2
	Case 3
		sName = MON3
	Case 4
		sName = MON4
	Case 5
		sName = MON5
	Case 6
		sName = MON6
	Case 7
		sName = MON7
	Case 8
		sName = MON8
	Case 9
		sName = MON9
	Case 10
		sName = MON10
	Case 11
		sName = MON11
	Case 12
		sName = MON12

	end Select	'// end month number case
	
	GetMonthName = sName	'// set return value
	
'//--------------Debugging
'msgBox( "In GetMonthName, returning month: "+sName)
'//-----------end Debugging
	
end function 	'// end GetMonthName 
'/**/

'// GetTransSheetName.bas
'//---------------------------------------------------------------
'// GetTransSheetName - Get sheet name where transaction belongs.
'//		wmk. 5/20/20.
'//---------------------------------------------------------------

function GetTransSheetName( sAcct as String ) as String

'//	Usage.	sSheetName = GetTransSheetName( sAcct )
'//
'//		sAcct = account number for transaction
'//				1xxx - Asset account number
'//				2xxx - Liability account number
'//				3xxx - Income account number
'//				4xxx - Expenses account number
'//				5xxx - Expenses account number	(non-deductible)
'//				9xxx - Expenses account number	(uncategorized)
'//
'//	Exit.	sSheetName = <target> = <result>
'//
'//	Modification history.
'//	---------------------
'//	5/15/20		wmk.	original code
'//	5/20/20.	wmk.	var dims to comply with EXPLICIT
'//
'//	Notes. <Insert notes here>
'//

'//	constants. (see module-wide constants)

'//	local variables.

dim sName	as String	'// returned string
dim sDigit 	as String	'// extracted 1st digit

	'// code.
	
	sName = ""					'// default is empty string
	sDigit = Left(sAcct, 1)		'// trigger on leftmost digit
	Select Case sDigit
	
		Case "1"
			sName = ASTSHEET
		Case "2"
			sName = LIASHEET
		Case "3"
			sName = INCSHEET
		Case "4"
			sName = EXPSHEET
		Case "5"
			sName = EXPSHEET
		Case "9"
			sName = EXPSHEET
			
	end Select	'// end 1st digit case statement

'// Note: the Switch statement must account for all values of sDigit or else
'// a runtime error occurs... when "6" was passed, function failed.
'	sName = Switch( sDigit="1",ASTSHEET, sDigit="2",LIASHEET, sDigit="3",INCSHEET,_
'				sDigit="4",EXPSHEET, sDigit="5",EXPSHEET, sDigit="9",EXPSHEET)		

	GetTransSheetName = sName					'// return name string

'//--------------Debugging
'msgBox( "In GetTransSheetName, returning sheet name: "+sName)
'//-----------end Debugging

end function 	'// end GetTransSheetName
'/**/

'// LogError.bas
'//---------------------------------------------------------------
'// LogError - Make error log entry to ERRLOGSHEET.
'//		wmk. 6/2/20.
'//---------------------------------------------------------------

public sub LogError(psErrName, psErrMsg)

'//	Usage.	macro call or
'//			call LogError(sErrName, sErrMsg)
'//
'//		sErrName = error code name (e.g. "ERRBADVALS")
'//		sErrMsg = error message string
'//
'// Entry.	gbErrDisplay = true if msgBox to be displayed
'//						   false otherwise
'//			gbErrRecording = true if to write ERRLOGSHEET entry
'//							 false otherwise
'//
'//	Exit.	gsErrName = psErrName
'//			gsErrMsg = psErrMsg
'//
'// Calls.	ErrLogGetCellInfo, ErrLogGetModule, ErrLogMakeEntry
'//
'//	Modification history.
'//	---------------------
'//	5/26/20.	wmk.	original code; issue message in box; set
'//						up call to ErrMakeEntry
'//	6/2/20.		wmk.	added caller module name to msgBox	
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim lErrCol As Long			'// error column index
dim lErrRow As Long			'// error row index
dim lErrSheet as Integer	'// error sheet index
dim sErrCol As String	'// string conversions
dim sErrRow As String
dim sErrSheet As String
dim iStatus As Integer		'// recording status

	'// code.
	gsErrName = psErrName
	gsErrMsg =  psErrMsg
	ErrLogGetCellInfo(lErrCol, lErrRow)		'// extract column and row information
	lErrSheet = ErrLogGetSheet()
	sErrSheet = Str(lErrSheet)
	sErrCol = Str(lErrCol)
	sErrRow = Str(lErrRow)	

	if gbErrDisplay then
		msgBox("In LogError, Called by '" + ErrLogGetModule() + CHR(13)+CHR(10)_
			+ "Error Name: "+psErrName+CHR(13)+CHR(10)_
			+ "Sheet index: "+sErrSheet+CHR(13)+CHR(10)_
			+ "At cell index: "+sErrCol+":"+sErrRow+CHR(13)+CHR(10)_
			+psErrMsg)
	endif	'// end display message conditional
	
	if gbErrRecording then
		iStatus = ErrLogMakeEntry()
	endif	'// end log recording active conditional
	
			
end sub		'// end LogError	6/2/20
'/**/

'// PlaceSplitTotal.bas
'//------------------------------------------------------------------
'// PlaceSplitTotal - Itemize split transaction in total COA account.
'//		7/8/20.	wmk.	15:00
'//------------------------------------------------------------------

public function PlaceSplitTotal(poGLRange As Object, psAcct As String) As Integer

'//	Usage.	iVal = PlaceSplitTotal(oGLRange, sAcct)
'//
'//		oGLRange = RangeAddress information for GL sheet
'//		sAcct = Total COA account from which items split
'//
'// Entry.	oGLRange.Sheet = index of ledger sheet
'//					.StartColumn = index of starting column
'//					.StartRow = index of first "split" row
'//					.EndColumn = index of ending column
'//					.EndRow = index of last "split" row
'//
'//	Exit.	iVal = 0 - normal return
'//				   ERRCONTINUE - nonfatal error; continue processing
'//				   ERRSTOP - fatal error; stop processing transactions
'//
'// Calls.	GetMonthName, GetTransSheetName, GetInsRow, SetSumFormula
'//
'//	Modification history.
'//	---------------------
'//	6/4/20.		wmk.	original code
'//	6/6/20.		wmk.
'//	6/27/20.	wmk.	superfluous "if" removed
'//	7/8/20.		wmk.	code to remove extra row in source COA transaction	
'//
'//	Notes.
'// This function adds lines to the splitting account category
'// then copies lines from the original "split" transaction
'// except for the Total transaction line; then moves the
'// nonempty Debit or Credit range into the opposing column
'// and ensures that the SUM formula in the last "split" row
'// is corrected to sum the moved data.
'//

'//	constants.
const FORMULA=3		'// FORMULA type on cell

'//	local variables.
dim iRetValue As Integer		'// returned value
dim oDoc As Object				'// ThisComponent
dim oSheets As Object			'// Doc.getSheets()
dim iSheetIx As Integer			'// sheet index of ledger
dim oGLSheet As Object			'// ledger sheet object
dim sSplitAcct As String		'// COA of account splitting from
dim iBrkPt As Integer			'// easy breakpoint

'//	GL transaction variables.
dim oGLCellDate	As Object		'// Date cell from GL
dim sStyle	As String			'// cell style
dim sDate	As String			'// date field string
dim sMonth	As String			'// month name of date (e.g. "January")
dim l1stRow	As Long				'// first "split" row
dim lLastRow As Long			'// last "split" row
dim lRowCount	As Long			'// transaction row count between "split"s

'//	COA category sheet variables.
dim sAcctCat2 As String			'// COA account category total transaction
dim oCat2Sheet As Object		'// COA account sheet for total
dim lCat2InsRow	As Long			'// COA sheet insertion row
dim oDebitCell as Object
dim bMoveDebits as Boolean
dim oFormulaCell As Object
dim sFormula As String
dim sFormStart as String
dim lCat2CurrRow as Long	'// insertion sheet current row
dim li as Long				'// long loop counter
dim oDateCell as Object		'// Date cell object

	'// code.
	sSplitAcct = psAcct			'// copy COA account string
	
	iRetValue = 0
	
	PlaceSplitTotal = iRetValue
	
'	if true then
'		Exit Function
'	endif

'//---------------begin PlaceSplitTotal here------------------------
'//	poGLSheet is the ledger sheet with the whole transaction
'// sSplitAcct is the account that has the total COA
		'//==========================================================
'//-----------------error handling setup-----------------------------
dim sErrCallerMod 	As String	'// caller routine name
dim iErrSheetIx		As Integer	'// caller sheet index
dim lErrColumn		As Long		'// caller error focus column
dim lErrRow			As Long		'// caller error focus row
dim sMyName			As String	'// this routine name
	'//*-----------------preserve error handling for caller-----------
	'// preserve caller error settings.
	sErrCallerMod = ErrLogGetModule()
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	'// set up new error settings.
	sMyName = "PlaceSplitTotal"
	ErrLogSetModule(sMyName)
	ErrLogSetSheet(poGLRange.Sheet)
	ErrLogSetCellInfo(poGLRange.StartColumn, poGLRange.StartRow)
'//*----------end error handling setup-------------------------------

		'//	insertCells...
'// entry.	sSplitAcct is the COA from the total line entry
'//         oGLCellDate.Text.String is the Date field from the totals
'//			 line entry
'//			oSheets object is the Sheets[] array this component

	oDoc = ThisComponent			'// set component and access objects
	oSheets = oDoc.getSheets()
	iSheetIx = poGLRange.Sheet
	oGLSheet = oSheets.getByIndex(iSheetIx)
	
	'// extract date information from ledger entry
	l1stRow = poGLRange.StartRow	'// first row of "split"
	lLastRow = poGLRange.EndRow		'// last row of "split"
	lRowCount = poGLRange.EndRow - l1stRow - 1			'// transaction lines with COAs
	oGLCellDate = oGLSheet.getCellByPosition(COLDATE, l1stRow)

	'// set up parameters for GetTransSheetName and GetInsRow
	sStyle = oGLCellDate.Text.String
	sDate = oGLCellDate.Text.String
	sMonth = GetMonthName(sDate)		'// get month search name
	sAcctCat2 = GetTransSheetName(sSplitAcct)	'// get sheet name matching COA

	'// access target COA sheet and set up insertion row
	oCat2Sheet = oSheets.getByName(sAcctCat2)
iBrkpt=1		'// oCat2Sheet.RangeAddress.Sheet = sheet index
ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)	'// set sheet for errors
	lCat2InsRow = GetInsRow(oCat2Sheet, sSplitAcct, sMonth)

	if lCat2InsRow < 0 then
		sErrName = "ERRCOAINSERT"
		sErrMsg = "Unable to insert Total rows in split account"
		GoTo ErrHandler
	endif
		
	
Dim oCOARange As New com.sun.star.table.CellRangeAddress

	'// set up oCOARange object
	'// copy entire transaction just like if doing it manually
	'// move filled values to other column (Debit > Credit) or
	'//		(Credit to Debit)
	'// delete total row from transaction
	'// set sum formula in COA field of last "split" row
	
	'// at this point l1stRow = 1st row with "split"
	'//               lLastRow = last row with "split
	'//				  set lStartColumn = COLDATE
	'//				  set lEndColumn = COLREF
	'//				  set sheet number (integer) = index of GL sheet
	'//				  set target start row at lCat2InsRow
	'//				  set target start column to COLDATE
	'//				  set number of rows to lRowCount + 1
	'// Doc = ThisComponent
	'// Sheet = Doc.Sheets(0)
	'//	oSel = Doc.getCurrentSelection()
	'//	oRange = oSel.RangeAddress			'// extract range information
iBrkPt=1		
	oCOARange.Sheet = oCat2Sheet.RangeAddress.Sheet
	oCOARange.StartColumn = COLDATE
	oCOARange.EndColumn = COLREF
	oCOARange.StartRow = lCat2InsRow	'// we are inserting into the 2nd sheet
	oCOARange.EndRow = lCat2InsRow + lRowCount + 2	'// start+transaction rowcount+1
ErrLogSetCellInfo(COLDATE, lCat2InsRow)
	oCat2Sheet.insertCells(oCOARange,_
							com.sun.star.sheet.CellInsertMode.ROWS)
dim lSplitEnd As Long		'// save split end in target sheet
	lSplitEnd = oCOARange.EndRow
		
	'// now copy original transaction split range address to same range
	'// overwriting blank rows;
dim oTargetCellAddress As new com.sun.star.table.CellAddress
	oTargetCellAddress.Sheet = oCOARange.Sheet
	oTargetCellAddress.Column = COLDATE
	oTargetCellAddress.Row = lCat2InsRow
		
dim oGLCellRange As new com.sun.star.table.CellRangeAddress
	oGLCellRange.Sheet = poGLRange.Sheet
	oGLCellRange.StartColumn = COLDATE
	oGLCellRange.EndColumn = COLREF
	oGLCellRange.StartRow = l1stRow
	oGLCellRange.EndRow = lLastRow

'// ErrLogGetSheet - Get sheet index from error log globals.
ErrLogSetSheet(oTargetCellAddress.Sheet)
ErrLogSetCellInfo(oTargetCellAddress.Column, oTargetCellAddress.Row)

	'// copy the entire split transaction to the COA sheet
	'// then delete the total row;
	oGLSheet.copyRange(oTargetCellAddress, oGLCellRange)
		
	'// use removeCells to remove Total row
	'// uses same parameters as insertCells; only change Row range
	oCOARange.Sheet = oTargetCellAddress.Sheet	
	oCOARange.StartRow = lCat2InsRow + 1	'// total row
	oCOARange.EndRow = oCOARange.StartRow
	
ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)
ErrLogSetCellInfo(COLDATE, oCOARange.EndRow)
	
	oCat2Sheet.removeRange(oCOARange,_
						com.sun.star.sheet.CellDeleteMode.ROWS)
	lSplitEnd = lSplitEnd - 1		'// adjust split end for removed row
		
iBrkPt=1
							
	'// then move the Debit/Credit values to the other column
	'// change background color of colored cells to LTGREEN
	'// test which column is .Type empty cells
	'// then move the Debit/Credit values to the empty column
	oDebitCell = oCat2Sheet.getCellByPosition(COLDEBIT,_
					oCOARange.StartRow)
	oTargetCellAddress.Row = oDebitCell.CellAddress.Row
	oCOARange.EndRow = lSplitEnd - 2	'// 2 back; no Total row and not end "split"

	bMoveDebits = (oDebitCell.Type <> EMPTY)
		'	EMPTY,VALUE,TEXT,FORMULA

	'// if there is something in the Debits column; the Debits move to Credits		
	if bMoveDebits then	'// set range to debits column
		oTargetCellAddress.Column = COLCREDIT
		oCOARange.StartColumn = COLDEBIT
		oCOARange.EndColumn = COLDEBIT
	else				'// otherwise, the Credits move to Debits
		oTargetCellAddress.Column = COLDEBIT
		oCOARange.StartColumn = COLCREDIT
		oCOARange.EndColumn = COLCREDIT
	endif	'// end move Debits conditional

'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
ErrLogSetCellInfo(oTargetCellAddress.Column, oTargetCellAddress.Row)

	'// move the column containing values to opposing column
	oCat2Sheet.moveRange(oTargetCellAddress, oCOARange)

		
	'// set .formula on COLACCT of last line to
	'//	 "=SUM()" previous cells above in either Debit or Credit

ErrLogSetCellInfo(COLACCT, lSplitEnd-1)

	'// row is now lSplitEnd-1 since Total row deleted	
	oFormulaCell = oCat2Sheet.getCellByPosition(COLACCT, lSplitEnd-1)
		
	'// Always force SUM formula into last row COLACCT field
'	if oFormulaCell.Type = FORMULA then

	'// set up RangeAddress object for SetSumFormula
dim oNewSplitRange as new com.sun.star.table.CellRangeAddress
	oNewSplitRange.Sheet = oCat2Sheet.RangeAddress.Sheet
	oNewSplitRange.StartRow = lCat2InsRow		'// first "split" row
	oNewSplitRange.EndRow = lSplitEnd-1			'// last "split" row (minus 1, row deleted)
	oNewSplitRange.StartColumn = COLDATE		'// Date column
	oNewSplitRange.EndColumn = COLREF			'// Reference column
	sFormula = oFormulaCell.String			'// just for kicks

ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)
ErrLogSetCellInfo(COLACCT, oNewSplitRange.EndRow)
			
	sFormula = SetSumFormula(oNewSplitRange, bMoveDebits)		
	oFormulaCell.String = sFormula			'// set formula in cell
	oFormulaCell.setFormula(sFormula)
	'// loop changing Date fields' color to LTGREEN
ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)
ErrLogSetCellInfo(COLDATE, oNewSplitRange.StartRow)
	lRowCount = oNewSplitRange.EndRow - oNewSplitRange.StartRow + 1
	lCat2CurrRow = oNewSplitRange.StartRow
	for li = 0 to lRowCount-1
		oDateCell = oCat2Sheet.getCellByPosition(COLDATE, lCat2CurrRow)
		oDateCell.CellBackColor = LTGREEN
		lCat2CurrRow = lCat2CurrRow + 1	
	next li					

'// add code to remove superfluous row at bottom
	'// row to remove is at lSplitEnd								'// mod070820
	'// use removeCells to remove extraneous						'// mod070820
	'// uses same parameters as insertCells; only change Row range	'// mod070820
	oCOARange.Sheet = oTargetCellAddress.Sheet						'// mod070820
	oCOARange.StartRow = lSplitEnd	'// blank row					'// mod070820
	oCOARange.EndRow = lSplitEnd									'// mod070820
	
ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)						'// mod070820
ErrLogSetCellInfo(COLDATE, lSplitEnd)								'// mod070820	
	oCat2Sheet.removeRange(oCOARange,_
						com.sun.star.sheet.CellDeleteMode.ROWS) 		'// mod070820

		'// update sheet modified date cell							'// mod060620
		oCOARange.Sheet = oCat2Sheet.RangeAddress.Sheet				'// mod060620
		Call SetSheetDate(oCOARange, MMDDYY)						'// mod060620
		oDateCell = oCat2Sheet.getCellByPosition(COLDEBIT, DATEROW)	'// mod060620
		oDateCell.CellBackColor = LTGREEN							'// mod060620

	GoTo NormalExit

ErrHandler:
	'// here because insertion row returned <0; unable to insert rows
	'// iStatus is bad value returned
	iRetValue = iStatus
	Select Case iStatus
	Case ERRCOANOTFOUND
		sErrCode = csERRCOANOTFOUND
		sErrMsg = "Account "+sBadCOA2+" not found"
	Case ERROUTOFROOM
		sErrCode = csERROUTOFROOM
		sErrMsg = "Account "+sBadCOA2+" month "+sMonth+" not enough rows"_
					+" to insert"
	Case ERRCOACHANGED
		sErrCode = csCOACHANGED
		sErrMsg = "Account "+sBadCOA2+" month "+sMonth+" not found"
	Case else
		sErrCode = csERRUNKNOWN
		sErrMsg = "Account "+sBadCOA2+"Undocumented error"
	end Select

	Call LogError(sErrCode, sErrMsg)			
	
'//-----------------end PlaceSplitTotal here----------------------------
NormalExit:
	'//*-----------------restore error handling for caller-----------
	'// restore entry error settings.
	ErrLogSetModule(sErrCallerMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
	'//*----------end error handling restore----------------------------

	'// iStatus = 0, no error
	'//				ERRCONTINUE (-1) - error; continue with next transaction
	'//				ERRSTOP (-2) - error; stop processing transactions

	PlaceSplitTotal = iRetValue

end Function		'// end PlaceSplitTotal		7/8/20
'/**/

'// PlaceSplitTrans.bas
'//------------------------------------------------------------------
'// PlaceSplitTrans - Place split transaction into multiple accounts.
'//		7/8/20.	wmk.	07:00
'//-------------------------------------------------------------------

function PlaceSplitTrans( poGLSheet as Object, plSplit1stRow as Long)_
	as long

'//	Usage.	lSplitEndRow = PlaceSplitTrans(oGLSheet, lSplit1stRow)
'//			call PlaceSplitTrans( oGLSheet, lSplit1stRow )
'//
'//		oGLSheet = GL sheet object
'//		lSplit1stRow = row index of first row of split transaction
'//
'// Entry.	GL sheet processing has found row with "split" as the
'//			.Text.String value in COLSPLIT. The split entry contains
'//			multiple rows until another row is found with "split" as the
'//			.Text.String value in COLSPLIT. The second row of the split
'//			transaction MUST contain the COA target and the total amount
'//			of the split, regardless of whether it is a Debit or Credit.
'//			Subsequent rows contain the amounts split into the opposing
'//			COAs, and their value sums MUST equal the COA target amount.
'//
'//	Exit.	lSplitEndRow >=0 - row index of GL last row of split
'//						 <0 - error, caller should terminate processing`
'//			ErrLogDisable() called, so caller will have to re-enable
'//			caller error logging
'//
'//	Calls.	CheckSplitTrans, GetTransSheetName, GetInsRow, GetMonthName
'//			ErrLogSetup, LogError, ErrLogDisable
'//
'//	Modification history.
'//	---------------------
'//	5/22/20		wmk.	original code
'//	5/23/20		wmk.	code added to put lines in split accounts; bug
'//						fixes; fix splits loop count; changes inserted
'//						rows Date field color to LTGREEN
'//	5/24/20.	wmk.	start adding code to copy rows into COA total;
'//						add insert error code checking and ErrorHandler
'//	5/26/20.	wmk.	error handling modified; error logging code added
'// 5/27/20.	wmk.	bug fixed where COA error in GL sheet being reported
'//						as in accounting category sheet; error code string
'//						ERRCOANOTFOUND corrected in constants fix bug where
'//						last split row with "split" not being copied into
'//						account where other accounts split from
'//	5/28/20.	wmk.	bug fix where debit/credit column swap copying 1
'//						too many rows; SUM formula insertion fix for COA
'//						category sheet entry of split details; feature implemented
'//						where split block inserted in sheet gets LTGREEN Date
'//						cell background color
'//	6/3/20.		wmk.	add code to change source transaction REF fields
'//						cell background color to LTGREEN between "split"
'//						rows; dead code block removed
'//	6/4/20.		wmk.	modularize code; add PlaceSplitTotal function;
'//						dead code and unused vars removed
'//	6/5/20.		wmk.	adjust code for SplitOutCOAs call; dead code
'//						removed; additional error codes from SplitTrans
'//	7/8/20.		wmk.	debugging superfluous msgbox removed
'//
'//	Notes.
'//	Any return value < 0 should be treated as a termination error,
'// and the caller should stop processing rows, as likely there will
'// only be errors in the subsequent rows until the user corrects
'// the GL spreadsheet entries.
'// Feature; 5/17/20; transfer "split" transaction to appropriate
'// multiple accounts (generalized add code to process split
'// transactions; caveat: all split transactions should have total in
'// first line with COA that will be used in all other line items as
'// moved to ease cross-referencing, regardless of whether total is a
'// debit or credit.
'//	Method.
'//	Stolen code from PlaceTransM. Only use oCat2Sheet, since only 
'// dealing with one target account sheet at a time. oGLSheet contains
'// the original "split" transaction lines. oCat2Sheet is the
'// category account sheet that gets each inserted line.

'//	constants. (see module-wide definitions, and add these)

'// local contants.

'// split controls and error codes.
const MAXSPLITROWS=13	'// maximum rows in "split" to "split"; 10 subaccounts
const ERRNOTSPLIT=-1	'// first fow not "split" in COLSPLIT
const ERRBADTALLY=-2	'// split values don't add to total
const ERRROWSEXCEEDED=-3	'// too many rows bedore second "split" found
const ERRBADVALS=-4		'// bad values, either missing or don't add
const ERRBADDATE=-5		'// Date field too short
const ERRDATENOTSAME=-6	'// Date field doesn't match others			'// mod060520
const ERRSPLITDESC=-7	'// "split" descriptions don't match		'// mod060520

'// insert error codes.
const ERRCOANOTFOUND=-1
const ERROUTOFROOM=-2
const ERRCOACHANGED=-3

'//------------------begin error handling block---------------
'// LogError setup snippet.

'// split error code strings.
const csERRUNKNOWN="ERRUNKNOWN"
const csERRNOTSPLIT="ERRNOTSPLIT"
const csERRBADTALLY="ERRBADTALLY"
const csERRROWSEXCEEDED="ERRROWSEXCEEDED"
const csERRBADVALS="ERRBADVALS"
const csERRBADDATE="ERRBADDATE"

'// insert error codes.
const csERRCOANOTFOUND="ERRCOANOTFOUND"
const csERROUTOFROOM="ERROUTOFROOM"
const csERRCOACHANGED="ERRCOACHANGED"

'//------------------end error handling block---------------

'//	local variables.

Dim Doc As Object			'// This component
Dim oSheets as Object		'// Sheets() collection

'//	GL XCell information objects (original line)
dim oGLCellDate as Object
dim oGLCellTrans as Object
dim oGLCellDebit as Object
dim oGLCellCredit as Object
dim oGLCellAcct as Object
dim oGLCellRef as Object

'//	oCat2Sheet XCell information objects.
dim oCat2Date as object		'// date of transaction
dim oCat2Trans as object	'// transaction description
dim oCat2Debit as object	'// debit amount
dim oCat2Credit as object	'// credit amount
dim oCat2Acct as object		'// COA account number
dim oCat2Ref as object		'// reference info

'// GL field data extraction objects (.Text or .Value) and strings
dim sStyle as String		'// Text.Style string

'// Date field variables for both GL objects.
dim sDate as String
dim lDateFormat as long		'// date format mm/dd/yy
dim sMonth as String

'//	CellRange content variables for GL sheet.
dim lGLStartRow as long
dim lGLEndCol as long
dim lGLCurrRow as long		'// current row in GL

'// more local vars.
dim iStatus as integer		'// general status return var
dim sAcctCat2 as String		'// account category of 2nd line of transaction
dim oCat2Sheet as Object	'// account sheet object of 2nd category in transaction
dim lCat2InsRow as long		'// insert index for 2nd cateory in transaction

'// row processing variables; use to process selected area
dim i as long				'// loop counter
dim li as Long				'// long loop counter
dim sBadCOA2 as String		'// bad acct # in 2nd line
dim sCOA2Msg	as String	'// bad 2nd line message

'// other local vars.
dim iBrkPt as integer		'// easy breakpoint reference
dim lRetValue as long		'// returned value
dim l1stRow as Long			'// 1st row defining "split"
dim sSplitAcct as String	'// account splitting into
dim lLastRow as Long		'// last line of "split" block
dim sAcct2 as String		'// COA target account
dim lTrans1stRow as Long	'// first row after "split"
dim lRowCount as Long		'// transaction row count

'// error handling setup------------------------------------------------
const sMyName="PlaceSplitTrans"
dim sErrCode as String	'// error code string for user code
Dim sErrMsg As String	'// error message for LogError
dim lLogRow as long		'// cell row working on
dim lLogCol as Long		'// cell column working on
dim iLogSheetIx as Integer	'// sheet index module working on
dim oLogRange as new com.sun.star.table.CellRangeAddress

	'// code.

'// error handling initialization---------------------------------------
	sErrCode = ""
	iLogSheetIx = poGLSheet.RangeAddress.Sheet
	lLogCol = poGLSheet.RangeAddress.StartColumn
	lLogRow = poGLSheet.RangeAddress.StartRow
	oLogRange.Sheet = iLogSheetIx
	oLogRange.StartColumn = lLogCol
	oLogRange.StartRow = lLogRow
	oLogRange.EndRow = lLogRow
	oLogRange.EndColumn = lLogCol
	Call ErrLogSetup(oLogRange, sMyName)
	ErrLogSetRecording(true)		'// turn on log sheet
'//--------------end error handling setup-------------------------------

	lRetValue = ERRNOTSPLIT	'// set bad return
	Doc = ThisComponent				'// set up for sheet access within document
	oSheets = Doc.getSheets()			'// get sheet collection

	lGLStartRow = plSplit1stRow	'// one of these is superfluous?
	l1stRow = plSplit1stRow		'// set local ptr to 1st "split" row
	lTrans1stRow = l1stRow + 1	'// first actual transaction data row

	'// CheckSplitTrans validates and returns last "split" row
	lRetValue = CheckSplitTrans(poGlSheet, l1stRow)
	
	if lRetValue < 0 then		'// if invalid split, get out
		GoTo BailOut			'// CheckSplitTrans handled error
	else
		lLastRow = lRetValue	'// set last row of split block
	endif

	'// l1stRow = first row of "split"
	'// lLastRow = last row of "split"
	'// get "total" account COA and preserve as split "from" account
	lGLCurrRow = l1stRow + 1			'// move back to total line to split lines
	oGLCellAcct = poGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
ErrLogSetCellInfo(COLACCT, lGLCurrRow)
	sSplitAcct = oGLCellAcct.Text.String			'// set total COA splitting

'// set up Range object for SplitOutCOAs and PlaceSplitTotal calls.
dim oGLRange As new com.sun.star.table.CellRangeAddress

'	if true then			'// jump around new code 1st pass
'		GoTo SplitTrans
'	endif
	
'// set RangeAddress object pointing to the "split" transaction
'// in the ledger sheet
'//	iStatus = SplitOutCOAs(oGLRange, sSplitAcct)
'//		oGLRange specifies sheet, and row indexes of the
'//         entire split transaction
'//		sSplitAcct = COA of Total account for xrefs
	oGLRange.Sheet = poGLSheet.RangeAddress.Sheet
	oGLRange.StartColumn = COLDATE
	oGLRange.EndColumn = COLREF
	oGLRange.StartRow = l1stRow
	oGLRange.EndRow = lLastRow
	lRetValue = SplitOutCOAs(oGLRange, sSplitAcct)

	if lRetValue < 0 then
		lCat2InsRow = lRetValue
		GoTo ErrHandler
	endif
GoTo SplitTotal
	
SplitTrans:

SplitTotal:
iBrkPt=1

'//	setup code for PlaceSplitTotal
'// oGLRange = RangeAddress information from poGLSheet
'// sSplitAcct = COA for transaction total
		oGLRange.Sheet = poGLSheet.RangeAddress.Sheet
		oGLRange.StartColumn = COLDATE
		oGLRange.EndColumn = COLREF
		oGLRange.StartRow = l1stRow
		oGLRange.EndRow = lLastRow
		iStatus = PlaceSplitTotal(oGLRange, sSplitAcct)

		if iStatus < 0 then
			'// iStatus should be set to negative insert row
			lCat2InsRow = iStatus
			GoTo ErrHandler
		endif	'// end error splitting total conditional


'//	change Ref field background to LTGREEN for all split
'//	rows between "split"s in source transaction
	lGLCurrRow = l1stRow + 1			'// start at Total row
	lRowCount = lLastRow - l1stRow -1	'// row count of lines between "split"
	for li = 0 to lRowCount-1			'// minus 1, since indexed on 0
		oGLCellRef = poGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		oGLCellRef.CellBackColor = LTGREEN
		lGLCurrRow = lGLCurrRow + 1		'// advance to next transaction row
	next li			'// next transaction line
	
	lRetValue = lLastRow		'// ensure return last "split" row	
	GoTo ExitNormal		'// jump around error handler
			

ErrHandler:
	'// issue appropriate message and don't mark rows as processed
	sBadCOA2 = sAcct2
	sCOA2Msg = ""			'// clear line 2 message
	Select Case lCat2InsRow
	Case ERRCOANOTFOUND
		sErrCode = csERRCOANOTFOUND
		sErrMsg = "Account "+sBadCOA2+" not found"
	Case ERROUTOFROOM
		sErrCode = csERROUTOFROOM
		sErrMsg = "Account "+sBadCOA2+" month "+sMonth+" not enough rows"_
					+" to insert"
	Case ERRCOACHANGED
		sErrCode = csCOACHANGED
		sErrMsg = "Account "+sBadCOA2+" month "+sMonth+" not found"
		
	Case ERRDATENOTSAME	'// Date field doesn't match others			'// mod060520
const csDATENOTSAME="ERRDATESNOTSAME"
		sErrCode = csDATENOTSAME
		sErrMsg = "Split - one or more dates don't match"
		
	Case ERRSPLITDESC	'// "split" descriptions don't match		'// mod060520
const csSPLITDESC="ERRSPLITDESC"
		sErrCode = csSPLITDESC
		sErrMsg = "Split - 1st and last split descriptions don't match"
		
	Case else
		sErrCode = csERRUNKNOWN
		sErrMsg = "Account "+sBadCOA2+"Undocumented error"
	end Select

	Call LogError(sErrCode, sErrMsg)			
	GoTo BailOut

	'// end up here if insert row < 0
ExitNormal:	
BailOut:
	ErrLogDisable()
	PlaceSplitTrans = lRetValue		'// set return value

iBrkPt = 1
	
end function		'// end PlaceSplitTrans	7/8/20
'/**/

'// PlaceTransM.bas
'//-------------------------------------------------------------------
'// PlaceTransM - place multiple transactions into appropriate sheets.
'//		wmk. 6/12/20.	18:00
'//-------------------------------------------------------------------

sub PlaceTransM()

'//	Usage.	macro call or
'//			call PlaceTransM
'//
'// Entry.	user has selected one or more transaction rows in ledger sheet
'//			from which transactions are to be placed in appropriate
'//			accounts sheet
'//
'//	Exit.	transaction (2 rows) is copied with each row of the transaction
'//			copied to the appropriate sheet by account number, and the
'//			account numbers transposed in the new entries to cross-reference
'//			the corresponding double-entry account for the transaction
'//
'//	Calls.	GetTransSheetName, GetInsRow, GetMonthName, SetSheetDate,
'//			PlaceSplitTrans
'//
'//	Modification history.
'//	---------------------
'//	5/17/20.	wmk.	original code
'//	5/18/20.	wmk.	finish original code and get functional; fix show-
'//						stopping bugs
'// 5/19/20.	wmk.	fixes where inserting transactions into different
'//						COAs within same sheet
'//	5/20/20.	wmk.	bug fixes in COA validation (5/18 bugs); var
'//						declarations to match EXPLICIT option
'//	5/30/20.	wmk.	bug fix with error handling setup code; bug fixes to comply
'//						with OPTION EXPLICIT
'//	6/1/20.		wmk.	add error code ERRBADSEL for when user selects
'//						odd number of rows or selection splits across
'//						a transaction; improved error checking on date
'//						field
'//	6/1/20.		wmk.	add process controls ERRCONTINUE and ERRSTOP;
'//						modularize code to processes CheckDoubleEntry,
'//						CheckInserts, StoreTrans, and SetProcessed;
'//						dead code cleanup
'//	6/2/20.		wmk.	CheckDoubleEntry code completed; old inline code
'//						removed; LogError modified to display calling
'//						module name; CheckInserts code completed; StoreTrans
'//						code completed
'//	6/3/20.		wmk.	Modularized code removed; error handling on last
'//						line improved; eliminate extraneous leftover
'//						variable declarations
'//	6/6/20.		wmk.	SetSheetDate call added to set modification date
'//						in any changed sheets; background LTGREEN
'// 6/6/20.		wmk.	mod to add capability to process "split"
'//						transactions by calling PlaceSplitTrans
'// 6/12/20.	wmk.	code rearranged to check for "split" first in loop
'//					 	to allow user to select only 1st row if just
'//						wants to process 1 split transaction
'//
'//	Notes. This emulates the process of taking general ledger
'// transactions (double entries) and creating new entries in
'// the appropriate account sheets for each line of the transaction. 
'// Each line of the transaction uses the following columns
'// Column A=date, B=transaction, C=debit D=credit  E=COA account
'//			 F=reference
'// Method.
'// (maybe write a utility just to verify transactions)
'// user should "select" area in GL sheet with ledger entries to clone
'//  to other sheets;
'// const INSBIAS = 3 to set row bias from start of Month/Acct#
'// convert month field from date data to capitalized month name for
'// use CurrentSelection function to get General Ledger rows to push into
'// appropriate category account sheet;
'// function GetTransSheetName (<sAcct>) as String - get sheet name, given
'// account number string
'// function GetInsRow( <oExSheet>, <sAcct>, <sMonth> ) as long returning row #
'//   if -1 not found; -2 if won't fit before next month or account;
'//  -3 COA found, but month not found before COA changed, 
'// insert new row at row returned from GetInsRow and copy row information
'// from General Ledger sheet to Expenses sheet; change color of date
'// field in both General Ledger and target account category sheet to a
'// light green. This will help the user to avoid the mistake of
'// double-updating General Ledger entries into account category sheets
'// use getSheetByName() call on Sheets object to accesss category
'// sheets
'// search in both category sheets looks for <month-name> in the second
'// field, and <account-number> in the first field, then move forward
'// in rows to check if enough space for insert without affecting
'// either month tallies; if not, transaction is skipped.
'// new rows are marked with light green background in Date field to
'// make it easy for the user to spot "automated" entries.
'// processed transaction rows in GL sheet are marked with light green
'//	background in Reference field to make it easy for user to spot
'// entries that have been disseminated with "automated" entries in
'// the account category sheets; if a transaction has been skipped
'// by the automated dissemination, its Reference field is left in
'// its original background color state
'//----------------------------------------------------------------------
'***************BUGS - * = still outstanding
'// 5/18/20.	failed to flag as error and skip transaction with at least
'//				one line with values in both debit and credit fields
'// 5/18/20.	failed to flag as error and skip transaction with at least
'//				one l line with no values in both debit and credit fields
'//	5/19/20.	transaction on accounts 1010/1020 ATM withdrawal
'//				skipped with "duplicate accounts" error
'//	5/19/20.	"duplicate accounts" trasaction error still colored
'//				reference field green
'//	5/19/20.	insert into same sheet with two different COAs causes
'//				insert of second COA line only 2 rows in from COA, not 3

'//	constants. (see module-wide constants)
'const COLDATE=0		'// Date - column A
'const COLTRANS=1	'// Transaction - column B
'const COLDEBIT=2	'// Debit - column C
'const COLCREDIT=3	'// Credit - column D
'const COLACCT=4		'// COA Account - column E
'const COLREF=5		'// Reference - column F


'// column index values for accessing information.

'// cell formatting constants.

'const DEC2=123		'// number format for (x,xxx.yy)
'const MMDDYY=37		'// date format mm/dd/y
'const LJUST=1		'// left-justify HoriJustify
'const CJUST=2		'// center HoriJustify
'const RJUST=0		'// right-justify HoriJustify
'const MAXTRANSL=50	'// maximum transaction text length

'// process control constants.
const ERRCONTINUE=-1	'// error, but continue with next transaction
const ERRSTOP=-2		'// error, stop processing transactions

'//	local variables.	6/3/20
Dim Doc As Object
Dim oGLSheet As Object		'// GL (source) sheet object
dim oSheets As Object		'// .Sheets() current Doc
dim oSel as Object
dim iSheetIx as integer
dim oRange as Object

'//	GL XCell information objects (1st transaction line)	6/3/20
dim oCellRef as object		'// reference info
dim sRef As String			'// reference field

'//	GL XCell information objects (line 2 of transaction)
dim oCellDateB as object	'// date of transaction
dim oCellRefB as object		'// reference info

'// GL field data extraction objects (.Text or .Value) and strings
dim sAcct as String			'// COA account field .String

'// Date field variables for both GL objects.
dim sMonth as String

'//	CellRange content variables for GL sheet.
dim oGLRange as Object		'// RangeAddress object of two rows processing
dim lGLStartRow as long
dim lGLEndRow as long
dim lGLCurrRow as long			'// current row in GL
dim lTransCount	as long			'// number of transactions to process

'// more local vars.
dim iBrkPt as integer			'// dummy var for setting breakpoint lines
dim iStatus as integer				'// general status return var
dim lStatus as Long				'// long status for PlaceSplitTrans

'// row processing variables; use to process selected area
dim i as long		'// loop counter
dim lNRowsSelected as long
dim bOddRowCount as boolean
dim sAcctB		as String		'// COA field from 2nd line
dim oCat1Range As Object		'// COA information for line 1
dim oCat2Range As Object		'// COA information for line 2		

'//----------in Module error handling setup-------------------------------
Dim sErrName as String		'// error name for LogError
Dim sErrMsg As String		'// error message for LogError
'*
'//-----------error handling setup-----------------------
'// Error Handling setup snippet.
	const sMyName="PlaceTransM"
	dim lLogRow as long		'// cell row working on
	dim lLogCol as Long		'// cell column working on
	dim iLogSheetIx as Integer	'// sheet index module working on
	dim oLogRange as new com.sun.star.table.CellRangeAddress
'//-----------end in Module error handling setup-----------------------


	'// code.
	
	Doc = ThisComponent			'// set up for sheet access within document
	oSel = Doc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet				'// get sheet index value

'//------------error handling setup code.--------------------
	iLogSheetIx = iSheetIx
	oLogRange.Sheet = iLogSheetIx
	lLogCol = oRange.StartColumn
	lLogRow = oRange.StartRow
	oLogRange.StartColumn = lLogCol
	oLogRange.StartRow = lLogRow
	oLogRange.EndRow = lLogRow
	oLogRange.EndColumn = lLogCol
	Call ErrLogSetup(oLogRange, sMyName)
'//-----------end error handling setup-----------------------

	oGLSheet = Doc.Sheets(iSheetIx)		'// set up sheet object

iBrkpt=1	

	'// set row bounds from selection information
	oGLRange = new com.sun.star.table.CellRangeAddress				'// mod060120
	oGLRange.Sheet = iSheetIx										'// mod060120
	lGLStartRow = oRange.StartRow
	lGLEndRow = oRange.EndRow
	lGLCurrRow = lGLStartRow - 2		'// current GL processing row preset for increment
	lNRowsSelected = lGLEndRow+1 - lGLStartRow
	lTransCount = lNRowsSelected/2	'// get transaction count (2 rows per)
'	bOddRowCount = lTransCount*2 <> lNRowsSelected
	
	for i = 1 to lTransCount

		lGLCurrRow = lGLCurrRow + 2	'// advance to next entry

		'// Note. By checking for "split" first, allows user to select
		'// first row only if wants to process 1 split transaction and no others
		'// Check for "split"; branch into and check
		oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		sRef = oCellRef.String

		if StrComp(sRef, "split")=0 then
			lStatus = PlaceSplitTrans(oGLSheet, lGLCurrRow)
			if lStatus < 0 then
				iStatus = ERRSTOP
				exit for
			else
				lGLCurrRow = lStatus - 1	'// minus 1 since advances by 2
				'// restore error handling setup
	iLogSheetIx = iSheetIx
	oLogRange.Sheet = iLogSheetIx
	lLogCol = oRange.StartColumn
	lLogRow = oRange.StartRow
	oLogRange.StartColumn = lLogCol
	oLogRange.StartRow = lLogRow
	oLogRange.EndRow = lLogRow
	oLogRange.EndColumn = lLogCol
	Call ErrLogSetup(oLogRange, sMyName)
	
				if lGLCurrRow < lGLEndRow-1 then 
					GoTo AdvanceTrans
				else
					exit for
				endif
				
			endif	'// bad return from PlaceSplitTrans conditional
			
		endif	'// end "split" conditional
		
		'// not a "split", check for end/odd row count
		bOddRowCount = (lGLCurrRow = lGLEndRow)
		if bOddRowCount then
			GoTo CheckLastRow
		endif
		
		'// account for case where a "split" was handled and
		'// adding 2 advances past end of selection by user
		if lGLCurrRow > lGLEndRow then
			exit for
		endif
				
		'// set RangeAddress object fields to point to 1st row
		'// and last row of transaction
		'// set up the range of the current transaction
		oGLRange.StartColumn = COLDATE
		oGLRange.StartRow = lGLCurrRow
		oGLRange.EndColumn = COLREF
		oGLRange.EndRow = lGLEndRow			'// set end of user selection

	
		'// Check the Date, Transaction, Debit/Credit and COA fields
		'//	for validity
		sMonth = ""
		sAcct = ""
		sAcctB = ""
'//	Usage.	iVal = CheckDoubleEntry(oTransRange, rpsMonth, rpsAcct1,
'//						rpsAcct2)
		iStatus = CheckDoubleEntry(oGLRange, sMonth, sAcct, sAcctB)
		if iStatus < 0 then
			Select Case iStatus
			Case ERRCONTINUE
				GoTo AdvanceTrans
			Case ERRSTOP
				exit for
			End Select
		endif		'// end error on Double Entry conditional
		
		'// Check target sheets/accounts for ability to insert
		'// row at correct accounting line
'//	Usage.	iVal = CheckInserts(sMonth, sAcct1, sAcct2, oCat1Range, oCat2Range)

		oCat1Range = new com.sun.star.table.CellRangeAddress
		oCat2Range = new com.sun.star.table.CellRangeAddress
		oCat1Range.Sheet = 0
		oCat1Range.StartRow = 0
		oCat1Range.StartColumn = 0
		oCat1Range.EndRow = 0
		oCat1Range.EndColumn = 0
		oCat2Range.Sheet =  0
		oCat2Range.StartRow = 0
		oCat2Range.StartColumn = 0
		oCat2Range.EndRow = 0
		oCat2Range.EndColumn = 0
		iStatus = CheckInserts(sMonth, sAcct, sAcctB, oCat1Range, oCat2Range)
		if iStatus < 0 then
			Select Case iStatus
			Case ERRCONTINUE
				GoTo AdvanceTrans
			Case ERRSTOP
				exit for
			End Select
		endif		'// end error on CheckInserts conditional
		
		'// Store transaction in new rows in account sheet(s)
		iStatus = StoreTrans(oGLRange, oCat1Range, oCat2Range, sAcct, sAcctB)
		if iStatus < 0 then
			Select Case iStatus
			Case ERRCONTINUE
				GoTo AdvanceTrans
			Case ERRSTOP
				exit for
			End Select
		endif		'// end error on StoreTrans conditional
		
		'// Update processed status in current transaction
		'//	by setting COLREF cells background to LTGREEN
'//	Usage.	iVar = SetProcessed(oTransRange)
		iStatus = SetProcessed(oGLRange)
		if iStatus < 0 then
			Select Case iStatus
			Case ERRCONTINUE
				GoTo AdvanceTrans
			Case ERRSTOP
				exit for
			End Select
		endif		'// end error on SetProcessed conditional

SetProcessed:
			'// Change color in GL lines to indicate processed.					'// mod051920
			oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
			oCellRefB = oGLSheet.getCellByPosition(COLREF, lGLCurrRow+1)	
			oCellRef.Text.CellBackColor = LTGREEN								'// mod051920
			oCellRefB.Text.CellBackColor = LTGREEN								'// mod051920

iBrkpt=1
	
AdvanceTrans:						'// use since this Basic does not support Continue For
			
	next i		'// end loop on GL rows to process

CheckLastRow:
	'// check for last row orphaned.
	if bOddRowCount then
		oCellDateB = oGLSheet.getCellByPosition( COLDATE, lGLEndRow)	'// load last row date field
		oCellDateB.Text.CellBackColor = YELLOW
		sErrName = "ERRODDROWS"
		sErrMsg = "Selection ends in mid-transaction!"_
				+CHR(13)+CHR(10)+ "Check last row selection of transaction"
		Call LogError(sErrName, sErrMsg)
	endif	'// end odd row count conditional

DontCheck:
	'// Update Sheet modified date stamp to indicate changes		'// mod060620
	'// Note: COA sheets' dates updated in StoreTrans
	Call SetSheetDate(oGLRange, MMDDYY)									'// mod060620

end sub		'// end PlaceTransM		6/12/20
'/**/

'// ReverseDETrans.bas
'//---------------------------------------------------------------
'// ReverseDETrans - Reverse double-entry transaction.
'//		6/12/20.	wmk.	17:30
'//---------------------------------------------------------------

public function ReverseDETrans(poGLRange As Object) As Integer

'//	Usage.	iVal = ReverseDETrans(oGLRange)
'//
'//		oGLRange = GL RangeAddress of double-entry transaction
'//
'// Entry.	oGLRange.Sheet = sheet index of ledger
'//					.StartColumn = COLDATE
'//					.StartRow = 1st row of transaction
'//					.EndColumn = COLREF
'//					.EndRow = last row of transaction
'//
'//	Exit.	iVal = 0 - normal return
'//				  -1 - abnormal return
'//
'// Calls.	CheckDoubleEntry, CreateReverseTrans
'//			ErrLogSetup, ErrLogSetRecording
'//
'//	Modification history.
'//	---------------------
'//	6/11/20.	wmk.	original code
'//	6/12/20.	wmk.	documentaton updated; code updated to call
'//						CheckDoubleEntry and CreateReverseTrans
'//
'//	Notes. Method: For double-entry transaction,insert 2 lines at first
'// row position with "ROWS" option; copy original transaction to 2 new
'// rows; swap Debit and Credit columns in old transaction by copying
'// back from original transaction into opposite columns.
'// Once the reversing transaction has been put in place in the GL,
'// the user runs PlaceTransM or PlaceSplitTrans to store the 
'// new transaction in the appropriate COA sheets
'// as though the user entered them for processing.

'//	constants.
const DEROWCOUNT=2			'// double entry row count

'//	local variables.
dim iRetValue As Integer
Dim Doc As Object
Dim oGLSheet As Object		'// GL (source) sheet object
dim oSheets As Object		'// .Sheets() current Doc
dim oSel as Object
dim oRange as Object
dim iSheetIx As Integer		'// GL sheet index
dim iStatus As Integer

'//	CellRange content variables for GL sheet.
dim oGLRange as Object		'// RangeAddress object of two rows processing
dim lGLStartRow as long
dim lGLEndRow as long
dim lGLCurrRow as long			'// current row in GL
dim lTransCount	as long			'// number of transactions to process
dim sMonth As String			'// month name
dim sAcct1 As String			'// line 1 COA field
dim sAcct2 As String			'// line 2 COA field
dim oCellDate As Object		'// ledger date field
dim oCellRef As Object		'// ledger reference field
dim lGLOrigRow As Long		'// original transaction row after insert

'// error handling setup------------------------------------------------
const sMyName="ReverseSplitTrans"
dim sErrCode as String	'// error code string for user code
Dim sErrMsg As String	'// error message for LogError
dim lLogRow as long		'// cell row working on
dim lLogCol as Long		'// cell column working on
dim iLogSheetIx as Integer	'// sheet index module working on
dim oLogRange as new com.sun.star.table.CellRangeAddress

	'// code.

	Doc = ThisComponent			'// set up for sheet access within document
	oSel = Doc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = poGLRange.Sheet			'// get sheet index value
	lGLStartRow = poGLRange.StartRow	'// set local start row for new
	
'// error handling initialization---------------------------------------
	sErrCode = ""
	iLogSheetIx = poGLRange.Sheet
	lLogCol = poGLRange.StartColumn
	lLogRow = poGLRange.StartRow
	oLogRange.Sheet = iLogSheetIx
	oLogRange.StartColumn = lLogCol
	oLogRange.StartRow = lLogRow
	oLogRange.EndRow = lLogRow
	oLogRange.EndColumn = lLogCol
	Call ErrLogSetup(oLogRange, sMyName)
	ErrLogSetRecording(true)		'// turn on log sheet
'//--------------end error handling setup-------------------------------

	
	iRetValue = -1		'// abnormal return
	oGLSheet = Doc.Sheets(iSheetIx)		'// set up sheet object

	'// set up call to CheckDouble entry
	'// verify transaction before reversing
	iStatus = CheckDoubleEntry(poGLRange, sMonth, sAcct1, sAcct2)
	if iStatus < 0 then
		iRetValue = iStatus
		GoTo ErrorHandler
	endif
	
	iStatus = CreateReverseTrans(oGLSheet, lGLStartRow, DEROWCOUNT)
	if iStatus < 0 then
		iRetValue = iStatus
		GoTo ErrorHandler
	endif	'// end CreateReversTrans error conditional
	
	lGLOrigRow = lGLStartRow + DEROWCOUNT

	'// highlight reversal transaction date fields
	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLOrigRow)
	oCellDate.String = oCellDate.String
	oCellDate.CellBackColor = LTGREEN
	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLOrigRow+1)
	oCellDate.String = oCellDate.String
	oCellDate.CellBackColor = LTGREEN

	'// document "reversal" in REF column
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLOrigRow)
	oCellRef.String = "reversal"
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLOrigRow+1)
	oCellRef.String = "reversal"
	
	'// set normal return
	iRetValue = 0

NormalExit:	
	ReverseDETrans = iRetValue
	exit function
	
ErrorHandler:
	ReverseDETrans = iRetValue
	
end function 	'// end ReverseDETrans	6/12/20
'/**/

'// ReverseSplitTrans.bas
'//---------------------------------------------------------------
'// ReverseSplitTrans - Reverse split transaction.
'//		7/9/20.	wmk.	07:45
'//---------------------------------------------------------------

public function ReverseSplitTrans(poGLRange As Object) As Integer

'//	Usage.	iVal = ReverseSplitTrans(oGLRange)
'//
'//		oGLRange = GL RangeAddress of double-entry transaction
'//
'// Entry.	oGLRange.Sheet = sheet index of ledger
'//					.StartColumn = COLDATE
'//					.StartRow = 1st row of transaction
'//					.EndColumn = COLREF
'//					.EndRow = last row of transaction
'//
'//	Exit.	iVal = 0 - normal return
'//				  -1 - abnormal return
'//
'// Calls.	CreateReverseTrans, SetSumFormula
'//
'//	Modification history.
'//	---------------------
'//	6/11/20.	wmk.	original code; stub
'//	6/12/20.	wmk.	original complete
'//	7/9/20.		wmk.	bug fix where "reversal" field LTGREEN
'//
'//	Notes. Method. For "split" entry transaction, verify with
'// CheckSplitTrans which will return the last "split" row of the
'// transaction; insert lRowCount rows at first row of transaction with
'// "ROWS" option; copy original transaction rows into new rows; swap
'// Debit and Credit columns in old.
'// The original transaction is placed first in the
'// GL, and the new reversing transaction immediately following.
'// Once the reversing transaction has been put in place in the GL,
'// the user runs PlaceTransM or PlaceSplitTrans to store the 
'// new transaction in the appropriate COA sheets
'// as though the user entered them for processing.

'//

'//	constants.

'//	local variables.
dim iRetValue As Integer
Dim Doc As Object
Dim oGLSheet As Object		'// GL (source) sheet object
dim oSheets As Object		'// .Sheets() current Doc
dim oSel as Object
dim iSheetIx as integer
dim oRange as Object
dim iStatus As Integer
dim i As Integer			'// short loop counter

'// GL Row processing variables.
dim lSplit1stRow As Long
dim lSplitEndRow As Long
dim lGLCurrRow As Long		'// current GL row being processed
dim lRowCount As Integer	'// count of rows in transaction
dim oCellDate As Object		'// Date field cell
dim oCellCredit As Object	'// Credit field cell
dim oCellRef As Object		'// Reference field cell
dim oCellAcct As Object		'// COA field cell
dim bDebitsMoving As Boolean	'// debits moving to credits flag
dim sSumStr As String		'// SUM formula string

'// error handling setup------------------------------------------------
const sMyName="ReverseSplitTrans"
dim sErrCode as String	'// error code string for user code
Dim sErrMsg As String	'// error message for LogError
dim lLogRow as long		'// cell row working on
dim lLogCol as Long		'// cell column working on
dim iLogSheetIx as Integer	'// sheet index module working on
dim oLogRange as new com.sun.star.table.CellRangeAddress

	'// code.

'// error handling initialization---------------------------------------
	sErrCode = ""
	iLogSheetIx = poGLRange.Sheet
	lLogCol = poGLRange.StartColumn
	lLogRow = poGLRange.StartRow
	oLogRange.Sheet = iLogSheetIx
	oLogRange.StartColumn = lLogCol
	oLogRange.StartRow = lLogRow
	oLogRange.EndRow = lLogRow
	oLogRange.EndColumn = lLogCol
	Call ErrLogSetup(oLogRange, sMyName)
	ErrLogSetRecording(true)		'// turn on log sheet
'//--------------end error handling setup-------------------------------

	iRetValue = -1
	Doc = ThisComponent			'// set up for sheet access within document
	oSel = Doc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet				'// get sheet index value

	'// set up sheet object for CheckSplitTrans
	'// verify original split transaction
	oSheets = Doc.getSheets()
	oGLSheet = Doc.Sheets(iSheetIx)
	lSplit1stRow = poGLRange.StartRow
	lSplitEndRow = CheckSplitTrans(oGLSheet, lSplit1stRow)
	if lSplitEndRow < 0 then
		GoTo ErrorHandler
	endif	'// end problem with split conditional

	lRowCount = lSplitEndRow - lSplit1stRow + 1
	iStatus = CreateReverseTrans(oGLSheet, lSplit1stRow, lRowCount)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
	
	'// highlight reversal transaction rows DATE fields with LTGREEN
	'// and set REF field to "reversal"
	lGLCurrRow = lSplit1stRow + lRowCount + 1	'// start at total row
	for i = 0 to lRowCount-3
		oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
		oCellDate.String = oCellDate.String
		oCellDate.CellBackColor = LTGREEN
		oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		oCellRef.String = "reversal"
		oCellRef.CellBackColor = NOFILL
		lGLCurrRow = lGLCurrRow + 1
	next i		'// end loop on transaction rows

	'// get transaction sum formula, store and highlight as LTGREEN
	oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lSplit1stRow+1)
	bDebitsMoving = (len(oCellCredit.String)>0)
dim	oGLRange As new com.sun.star.table.CellRangeAddress
	oGLRange.Sheet = iSheetIx
	oGLRange.StartColumn = COLDATE
	oGLRange.StartRow = poGLRange.StartRow
	oGLRange.EndColumn = COLREF
	oGLRange.EndRow = oGLRange.StartRow + lRowCount -1
	sSumStr = SetSumFormula(oGLRange, bDebitsMoving)
	if len(sSumStr) =  0 then
		GoTo ErrorHandler
	endif	'// end SetSumFormula error conditional
	
	oCellAcct = oGLSheet.getCellByPosition(COLACCT, lSplitEndRow + lRowCount)
	oCellAcct.String = sSumStr		'// set sum formula
	oCellAcct.setFormula(sSumStr)
	oCellAcct.Text.HoriJustify = CJUST
	oCellAcct.CellBackColor = LTGREEN
	oCellAcct.NumberFormat = DEC2
	iRetValue = 0			'// set normal return
	
NormalExit:
	ReverseSplitTrans = iRetValue
	exit function
	
ErrorHandler:	
	ReverseSplitTrans = iRetValue

end function 	'// end ReverseSplitTrans 7/9/20
'/**/

'// ReverseTrans.bas
'//---------------------------------------------------------------
'// ReverseTrans - Reverse transaction.
'//		6/12/20.	wmk.	17:45
'//---------------------------------------------------------------

public sub ReverseTrans()

'//	Usage.	macro call or
'//			call ReverseTrans()
'//
'// Entry.	User has first row of transaction selected;
'//			may either be a double entry or "split"
'//			
'//	Exit.	Reversing transaction entered below original
'//			Reversing transaction processed by appropriate
'//			double-entry or "split" transaction
'//
'// Calls.	ReverseSplit, ReverseDETrans, SetSheetDate
'//
'//	Modification history.
'//	---------------------
'//	6/11/20.		wmk.	original code
'//	6/12/20.		wmk.	add call to SetSheetDate
'//
'//	Notes. Method: For double-entry transaction,insert 2 lines at first
'// row position with "ROWS" option; copy original transaction to 2 new
'// rows; swap Debit and Credit columns in old transaction by copying
'// back from original transaction into opposite columns.
'// For "split" entry transaction, verify with CheckSplitTrans which will
'// return the last "split" row of the transaction; insert lRowCount rows
'// at first row of transaction with "ROWS" option; copy original
'// transaction rows into new rows; swap Debit and Credit columns in old
'// Both these methods will place the original transaction first in the
'// GL, and the new reversing transaction immediately following.
'// Once the reversing transaction has bee put in place in the GL,
'// the new transaction should be stored in the appropriate COA sheets
'// as though the user entered them for processing.


'//	constants.

'//	local variables.
Dim Doc As Object
Dim oGLSheet As Object		'// GL (source) sheet object
dim oSheets As Object		'// .Sheets() current Doc
dim oSel as Object
dim iSheetIx as integer
dim oRange as Object
dim iStatus As Integer

'// GL entry processing variables.
dim lGLStartRow As Long
dim lGLEndRow As Long
dim lGLCurrRow As Long
dim oCellRef As Object
dim sRef As String

'//----------in Module error handling setup-------------------------------
Dim sErrName as String		'// error name for LogError
Dim sErrMsg As String		'// error message for LogError
'*
'//-----------error handling setup-----------------------
'// Error Handling setup snippet.
	const sMyName="ReverseTrans"
	dim lLogRow as long		'// cell row working on
	dim lLogCol as Long		'// cell column working on
	dim iLogSheetIx as Integer	'// sheet index module working on
	dim oLogRange as new com.sun.star.table.CellRangeAddress
'//-----------end in Module error handling setup-----------------------

	'// code.
	Doc = ThisComponent			'// set up for sheet access within document
	oSel = Doc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet				'// get sheet index value

'//------------error handling setup code.--------------------
	iLogSheetIx = iSheetIx
	oLogRange.Sheet = iLogSheetIx
	lLogCol = oRange.StartColumn
	lLogRow = oRange.StartRow
	oLogRange.StartColumn = lLogCol
	oLogRange.StartRow = lLogRow
	oLogRange.EndRow = lLogRow
	oLogRange.EndColumn = lLogCol
	Call ErrLogSetup(oLogRange, sMyName)
'//-----------end error handling setup-----------------------

	oGLSheet = Doc.Sheets(iSheetIx)		'// set up sheet object

	'// set row bounds from selection information
	'// Note. ReverseTrans will only process one transaction, regardless
	'// of how many rows the user has selected.
dim	oGLRange As new com.sun.star.table.CellRangeAddress				'// mod060120
	oGLRange.Sheet = iSheetIx										'// mod060120
	lGLStartRow = oRange.StartRow
	lGLEndRow = oRange.EndRow
	lGLCurrRow = lGLStartRow 		'// current GL processing row
'	lNRowsSelected = lGLEndRow+1 - lGLStartRow

	'// set RangeAddress object fields to point to 1st row
	'// and last row of transaction
	'// set up the range of the current transaction
	oGLRange.StartColumn = COLDATE
	oGLRange.StartRow = lGLCurrRow
	oGLRange.EndColumn = COLREF
	oGLRange.EndRow = lGLEndRow			'// set end of user selection

	'// Check for "split"; branch into and check
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
	sRef = oCellRef.String

	if StrComp(sRef, "split")=0 then
		iStatus = ReverseSplitTrans(oGLRange)
	else		'// is double entry transaction
		iStatus = ReverseDETrans(oGLRange)
	endif

	if iStatus < 0 then
		GoTo ErrorHandler
	endif

	'// Update Sheet modified date stamp to indicate changes			'// mod060620
	'// Note: COA sheets' dates updated in StoreTrans
	Call SetSheetDate(oGLRange, MMDDYY)									'// mod060620
	
NormalExit:
	Exit Sub

ErrorHandler:


end sub		'// end ReverseTrans	6/12/20
'/**/

'// SetProcessed.bas
'//---------------------------------------------------------------
'// SetProcessed - Set double-entry transaction processed.
'//		6/1/20.	wmk.
'//---------------------------------------------------------------

public function SetProcessed(oTransRange as Object) as Integer

'//	Usage.	iVar = SetProcessed(oTransRange)
'//
'//		oTransRange = RangeAddress of double-entry transaction
'//
'// Entry.	<entry conditions>
'//
'//	Exit.	iVar = 0 if no error; -1 otherwise
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/1/20.	wmk.	original code; stub
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim iRetValue As Integer

	'// code.
	
	iRetValue = 0		'// set normal return
	
	SetProcessed = iRetValue

end function 	'// end SetProcessed  6/1/20
'/**/

'// SetSheetDate.bas
'//---------------------------------------------------------------
'// SetSheetDate - Set date/time stamp in common sheet date field.
'//		6/8/20.	wmk.
'//---------------------------------------------------------------

public function SetSheetDate(poRange As Object, piFmt As Integer) As Integer

'//	Usage.	iVar = SetSheetDate(oRange, iFmt)
'//
'//		oRange = RangeAddress of sheet to set date/time stamp in
'//		iFmt = NumberFormat for date/time stamp {FMTDATETIME, MMDDYY}
'//
'// Entry.	system clock running
'//
'//	Exit.	Sheet at Sheet index within oRange has Now() date/time
'//			stamp set in Cell column=COLDATEMOD, row=ROWDATEMOD
'//			iVar = 0 - no error
'//				   -1 - untrapped error occurred; date/time stamp
'//						setting probably failed
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/3/20.		wmk.	original code
'//	6/6/20.		wmk.	date/time format parameter added	
'//	6/8/20.		wmk.	add LTGRN background color to new date
'//	Notes.


'//	constants.
const COLDATEMOD=2		'// corresponds to "C"
const ROWDATEMOD=1		'// corresponds to "2" for cell C2
'const FMTDATETIME=50	'// numeric value of mm/dd/yyy hh:mm:ss spec

'//	local variables.
dim iRetValue As Integer
dim oDoc As Object		'// ThisComponent
dim oSheets As Object	'// Doc.getSheets()
dim oSheet As Object	'// target sheet
dim iSheetIx As Integer	'// target sheet index
dim oCellTime as Object	'// date/time modified cell
dim iDateFormat As Integer	'// date/time format to set

	'// code.
	
	iRetValue = 0
	ON ERROR GoTo Abnormal
	iSheetIx = poRange.Sheet	'// extract sheet index from passed range
	oDoc = ThisComponent		'// access target sheet using range information
	oSheets = oDoc.getSheets()
	oSheet = oSheets.getByIndex(iSheetIx)
	oCellTime = oSheet.getCellByPosition(COLDATEMOD, ROWDATEMOD)
	
	oCellTime.setValue(Now())					'// time stamp
	
	Select Case piFmt
	CASE FMTDATETIME
		iDateFormat = FMTDATETIME
		
	CASE MMDDYY
		iDateFormat = MMDDYY
		
	CASE else
		oDateFormat = MMDDYY
		
	End Select
	
	oCellTime.NumberFormat = iDateFormat		'// set number format
	oCellTime.CellBackColor = LTGREEN			'// set changed color
	GoTo NormalExit
	
Abnormal:
	iRetValue = -1		'// set abnormal return code
	
NormalExit:
	SetSheetDate = iRetValue

end function 	'// end SetSheetDate	6/8/20
'/**/

'// SetSumFormula.bas
'//---------------------------------------------------------------
'// SetSumFormula - Create SUM formula for last split account cell.
'//		wmk. 5/28/20.	23:15
'//---------------------------------------------------------------

public function SetSumFormula(poRange as Object, pbDebitsMoving as Boolean) As String

'//	Usage.	sNewFormula = SetSumFormula(oRange, bDebitsMoving)
'//
'//			oRange = RangeAddress object with .StartRow = first row of split
'//						.EndRow = last row of split
'//			bDebitsMoving = true if moving Debits to Credits; formula should be SUM(Credits)
'//					   false if moving Credits to Debits; formula should be SUM(Debits)
'//
'//	Exit.	sNewFormula = sum formula; either "=SUM(D..:D)" for SUM(Credits)
'//						 or "=SUM(C..:C..)" for SUM(Debits)
'//
'// Calls.	ErrLogGetModule, ErrLogSetModule
'//
'//	Modification history.
'//	---------------------
'//	5/25/20.	wmk.	original code
'//	5/26/20.	wmk.	error handling upgrade
'// 5/28/20.	wmk.	bug fix; changed calling sequence for readability;
'//						changed logic to properly regenerate formula; code clarified
'//						cell range index strings trimmed; convert row indices to
'//						row numbers by adding 1
'//
'//	Notes. <Insert notes here>

'//	constants.
const CHECKFORMCHARS=6		'// number of leading characters to check in formula
const DEBITCHARS="=SUM(C"	'// Debit summing formula
const CREDCHARS="=SUM(D"	'// Credit summing formula
const RANGEDEBIT=":C"
const RANGECREDIT=":D"
const ERRBADRANGE=-1		'// bad range error code; missing ':'

'//------------------begin error handling block---------------
'// LogError setup snippet.

'// error code strings.
const csERRUNKNOWN="ERRUNKNOWN"
const csERRBADRANGE="ERRBADRANGE"

Dim sErrName as String		'// error name for LogError
Dim sErrMsg As String		'// error message for LogError
'//------------------end error handling block---------------

'//	local variables.
dim iMidPtr As Integer		'// mid-string char pos pointer
dim sRetString As String	'// returned string
dim sStartRow As String		'// starting row
dim sEndRow As String		'// ending row
dim iScanPos As Integer		'// string scanning position
dim sScanChar As String		'// current scan char
dim iFormulaLen As Integer	'// formula length
dim sWhatsLeft As String	'// remaining formula to scan
dim iErr as Integer			'// error flag
dim iColonPos as Integer	'// pos of ":" in formula string
dim sMidRange As String		'// midrange string ":C" or ":D"

	'// code.
'//-----------error handling setup-----------------------
'// Error Handling setup snippet.
	const sMyName="SetSumFormula"
'	dim lLogRow as long		'// cell row working on
'	dim lLogCol as Long		'// cell column working on
'	dim iLogSheetIx as Integer	'// sheet index module working on
'	dim oLogRange as new com.sun.star.table.CellRangeAddress
'//	dim oDoc as Object		'// ThisComponent
'//	dim oSheets as Object	'// oDoc.Sheets
'//	code to put sheet index into oLogRange.Sheet
'	oLogRange.Sheet = iLogSheetIx
'	oLogRange.StartColumn = lLogCol
'	oLogRange.StartRow = lLogRow
'	oLogRange.EndRow = lLogRow
'	oLogRange.EndColumn = lLogCol
'	Call ErrLogSetup(oLogRange, sMyName)
'//OR
	dim sCallerModule As String
	sCallerModule = ErrLogGetModule()
	ErrLogSetModule(sMyName)
	'// on exit...
'	ErrLogSetModule(sCallerModule)
'//----------------end error handling setup-----------------------

	iErr = ERRUNKNOWN			'// set unknown error code
	ON ERROR GOTO ErrHandler
	sRetString = "test string"
	sStartRow = Ltrim(Str(poRange.StartRow+1))	'// add 1 to Row indices to get numbers
	sEndRow = Ltrim(Str(poRange.EndRow+1))

		
'// build returned formula as DEBITCHARS+<startrow>+RANGEDEBIT+
'//							  +<endrow>+")"
'// or CREDCHARS+<startrow>+RANGECREDIT+
'//							  +<endrow>+")"
	if pbDebitsMoving then
 		sRetString = CREDCHARS	'// + "XX" + RANGECREDIT + "YY" + ")"
 		sMidRange = RANGECREDIT
	else
		sRetString = DEBITCHARS '// + "XX" + RANGEDEBIT + "YY" + ")"
		sMidRange = RANGEDEBIT
	endif
			
	'// append start row number to returned string
	sRetString = sRetString + sStartRow	+ sMidRange
	sRetString = sRetString + sEndRow + ")"	
	iErr = 0
	GoTo NormalExit
	
ErrHandler:

	if iErr = 0 then
		GoTo NormalExit
	endif
	
	Select Case ierr

		Case ERRBADRANGE
			sErrName = csERRBADRANGE
			sErrMsg = "Expected ':' in split SUM formula..."
				
		Case else
			sErrName = csERRUNKNOWN
			sErrMsg = "Unrecognized error generating split SUM formula range..."
			
	End Select
	
	Call LogError(sErrName, sErrMsg)
	
NormalExit:
	ErrLogSetModule(sCallerModule)	'// restore caller module name
	SetSumFormula = sRetString
	
end function 	'// end SetSumFormula
'/**/

'// SplitOutCOAs.bas
'//-----------------------------------------------------------------
'// SplitOutCOAs - Split out COA transaction lines from split entry.
'//		8/26/22.	wmk.	11:14
'//-----------------------------------------------------------------

public function SplitOutCOAs(poGLRange As Object, psSplitAcct As String) As Integer

'//	Usage.	iVal = SplitOutCOAs(oGLRange, sSplitAcct)
'//
'//		oGLRange = RangeAddress of the entire "split" transaction
'//		sSplitAcct = COA account splitting total from
'//
'// Entry.	ThisComponent = spreadsheet document
'//
'//	Exit.	iVal = 0 - no error encountered
'//				 < 0 - error encountered during split operation
'//
'// Calls.	GetMonthName, GetTransSheetName
'//			ErrLogGetModule, ErrLogGetSheet, ErrLogGetCellInfo,
'//			ErrLogSetModule, ErrLogSetSheet, ErrLogSetCellInfo
'//
'//	Modification history.
'//	---------------------
'//	6/5/20.		wmk.	original code
'// 6/6/20.		wmk.	add code to update sheet cell C2 with date
'//						modified for COAs
'//	6/7/20.		wmk.	add code to set SUM in last split GL row
'//	6/10/20.	wmk.	expande SUM range to go "split" to "split"
'//	7/6/20.		wmk.	bug fix; eliminate code overwriting sSplitAcct
'// 8/22/22.	wmk.	bug fix; text values being stored in Debit/Credit
'//				 target fields, changed to .value values.
'// 8/26/22.	wmk.	bug fix; empty string debit/credit setting value 0.
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim iRetValue As Integer		'// returned value
dim i As Integer				'// loop counter
dim iBrkPt As Integer			'// easy breakpoint
dim iSheetIx As integer			'// sheet index in oSheets()
dim oDoc As Object				'// ThisComponent
dim oSheets As Object			'// oDoc.getSheets()
dim oGLSheet As Object			'// ledger sheet

'// GL sheet processing variables.
dim lRowCount as Long
dim lLastRow As Long
dim l1stRow As Long
dim lGLCurrRow as Long
dim oGLCellDate as Object
dim lDateFormat As Long
dim sStyle As String
dim sDate As String
dim sMonth As String
dim sSplitAcct As String
dim oGLCellAcct As Object
dim sAcct2 As String

'//	target COA sheet processing variables.
dim lCat2InsRow As Long			'// COA insert row
dim sAcctCat2 As String			'// COA category and month
dim oCat2Sheet As Object		'// COA sheet inserting into
'dim oGLCellDate As Object		'// Date cell
dim	oGLCellTrans As Object		'// Transaction description cell
dim	oGLCellDebit As Object		'// Debit cell
dim	oGLCellCredit As Object		'// Credit cell
dim	oGLCellRef As Object		'// Reference cell


'// target COA sheet column values.
dim	oCat2Date As Object		'// Date field
dim	oCat2Trans As Object	'// Transaction description
dim	oCat2Debit As Object	'// Debit field
dim	oCat2Credit As Object 	'// Credit field
dim	oCat2Acct As Object		'// COA field
dim	oCat2Ref As Object		'// Reference field
		

	'// code.
	
	iRetValue = 0

	SplitOutCOAs = iRetValue

'	if true then
'		Exit function
'	endif

'// SplitOutCOAs code begins here...------------------------------------------
'//-----------------error handling setup-----------------------------
dim sErrCallerMod 	As String	'// caller routine name
dim iErrSheetIx		As Integer	'// caller sheet index
dim lErrColumn		As Long		'// caller error focus column
dim lErrRow			As Long		'// caller error focus row
dim sMyName			As String	'// this routine name
	'//*-----------------preserve error handling for caller-----------
	'// preserve caller error settings.
	sErrCallerMod = ErrLogGetModule()
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	'// set up new error settings.
	sMyName = "SplitOutCOAs"
	ErrLogSetModule(sMyName)
	ErrLogSetSheet(poGLRange.Sheet)
	ErrLogSetCellInfo(poGLRange.StartColumn, poGLRange.StartRow)
'//*----------end error handling setup-------------------------------

	oDoc = ThisComponent			'// document object
	oSheets = oDoc.getSheets()		'// sheets array
	iSheetIx = poGLRange.Sheet
	oGLSheet = oSheets.getByIndex(iSheetIx)
	l1stRow = poGLRange.StartRow	'// first row of "split"
	lLastRow = poGLRange.EndRow		'// last row of "split"
	sSplitAcct = psSplitAcct		'// localize split account
	
	oGLCellDate = oGLSheet.getCellByPosition(COLDATE, l1stRow)

	'// set up parameters for GetTransSheetName and GetInsRow
'	sStyle = oGLCellDate.Text.String
	sDate = oGLCellDate.Text.String
	sMonth = GetMonthName(sDate)		'// get month search name
	sAcctCat2 = GetTransSheetName(sSplitAcct)	'// get sheet name matching COA

	'// access target COA sheet and set up insertion row
	oCat2Sheet = oSheets.getByName(sAcctCat2)
iBrkpt=1		'// oCat2Sheet.RangeAddress.Sheet = sheet index
ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)	'// set sheet for errors
	lCat2InsRow = GetInsRow(oCat2Sheet, sSplitAcct, sMonth)

	if lCat2InsRow < 0 then
		sErrName = "ERRCOAINSERT"
		sErrMsg = "Unable to insert Total rows in split account"
		GoTo ErrHandler
	endif
		
	'// set up oCat2Range for SetSheetDate calls
dim oCat2Range as new com.sun.star.table.CellRangeAddress			'// mod060620
	oCat2Range.Sheet = oCat2Sheet.RangeAddress.Sheet				'// mod060620
	oCat2Range.StartColumn = COLDEBIT								'// mod060620
	oCat2Range.StartRow = DATEROW									'// mod060620
	oCat2Range.EndColumn = COLDEBIT									'// mod060620
	oCat2Range.EndRow = DATEROW										'// mod060620
		
	lGLCurrRow = l1stRow + 1			'// move to total line to split lines
	lRowCount = lLastRow - l1stRow -1	'// transactions row count

iBrkPt=1	

	'// loop on lines between "split"s past total line
	for i=0 to lRowCount-2
		lGLCurrRow = lGLCurrRow + 1		'// advance GL row
		
		'// extract line date from GL line
		'// get month from date field.
		'// get category sheet name
		'// set up parameters for GetTransSheetName and GetInsRow
		oGLCellDate = oGLSheet.getCellByPosition(COLDATE,lGLCurrRow)

	ErrLogSetCellInfo(COLDATE, lGLCurrRow)
		lDateFormat = oGLCellDate.Text.NumberFormat
		sStyle = oGLCellDate.Text.String
		sDate = oGLCellDate.Text.String
		sMonth = GetMonthName(sDate)		'// get month search name
	ErrLogSetCellInfo(COLACCT, lGLCurrRow)
	ErrLogSetSheet(oGLSheet.RangeAddress.Sheet)
		oGLCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
		sAcct2 = oGLCellAcct.Text.String		'// COA field

		'// check for short or empty account field
		if Len(sAcct2) < 4 then
			lCat2InsRow = ERRCOANOTFOUND

			GoTo ErrHandler
		endif
		
		sAcctCat2 = GetTransSheetName(sAcct2)	'// get sheet name matching COA

		oCat2Sheet = oSheets.getByName(sAcctCat2)
iBrkpt=1
		lCat2InsRow = GetInsRow(oCat2Sheet, sAcct2, sMonth)

		'// if any problem with insert, bail out - this is all or nothing
		if lCat2InsRow < 0 then
			GoTo ErrHandler
		endif	'// end insert row error conditional


		'// load category/account splitting transactions from
		'// at this point, just add the lines to the appropriate accounts,
	
		'// extract line data; treat every line like a 2nd line
		oGLCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
		oGLCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
		oGLCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
		oGLCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
		oGLCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)

iBrkPt=1
		'// lCat2InsRow set above...
		'// Insert new line in appropriate category sheet.
	ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)
	ErrLogSetCellInfo(COLDATE, lCat2InsRow)
		oCat2Sheet.Rows.insertByIndex(lCat2InsRow, 1)	'// insert new category 2 row


		'// set up target cell objects.
		oCat2Date = oCat2Sheet.getCellByPosition(COLDATE, lCat2InsRow)
		oCat2Trans = oCat2Sheet.getCellByPosition(COLTRANS, lCat2InsRow)
		oCat2Debit = oCat2Sheet.getCellByPosition(COLDEBIT, lCat2InsRow)
		oCat2Credit = oCat2Sheet.getCellByPosition(COLCREDIT, lCat2InsRow)
		oCat2Acct = oCat2Sheet.getCellByPosition(COLACCT, lCat2InsRow)
		oCat2Ref = oCat2Sheet.getCellByPosition(COLREF, lCat2InsRow)

		'// copy date, transaction, debit, credit, and reference fields from GL entry.
		'// set account field to splitting account COA
		oCat2Date.String = oGLCellDate.String
		oCat2Date.Text.HoriJustify = RJUST
		oCat2Date.CellBackColor = LTGREEN
		oCat2Date.HoriJustify = RJUST
		oCat2Trans.Text.String = oGLCellTrans.Text.String
		oCat2Trans.Text.HoriJustify = LJUST
		oCat2Debit.Text.String = oGLCellDebit.Text.String
		if len(oGLCellDebit.String) > 0 then
		 oCat2Debit.setvalue(oGLCellDebit.getvalue())
		 oCat2Debit.NumberFormat = DEC2
		endif
 		oCat2Debit.Text.HoriJustify = RJUST
		oCat2Credit.Text.String = oGLCellCredit.Text.String
		if len(oGLCellCredit.String) > 0 then
 		 oCat2Credit.setvalue(oGLCellCredit.getvalue())
		 oCat2Credit.NumberFormat = DEC2
		endif
		oCat2Credit.Text.HoriJustify = RJUST
		oCat2Ref.Text.String = oGLCellRef.Text.String
		oCat2Ref.Text.HoriJustify = CJUST
		
		'// replace account field with splitting account field
		oCat2Acct.Text.String = sSplitAcct
		oCat2Acct.Text.HoriJustify = CJUST

		'// update sheet modified date cell							'// mod060620
		oCat2Range.Sheet = oCat2Sheet.RangeAddress.Sheet			'// mod060620
		Call SetSheetDate(oCat2Range, MMDDYY)						'// mod060620
		oCat2Date = oCat2Sheet.getCellByPosition(COLDEBIT, DATEROW)	'// mod060620
		oCat2Date.CellBackColor = LTGREEN							'// mod060620

	next i		'// advance to next split transaction line

	'// set up RangeAddress object for SetSumFormula
	'// poGLRange has the entire GL range
dim oSumSplitRange as new com.sun.star.table.CellRangeAddress
dim lSumRow As Long			'// row for SUM formula in GL
dim sFormula as String		'// target for SUM in GL
dim bOpposite As Boolean	'// flag for SetSumFormula
	oSumSplitRange.Sheet = poGLRange.Sheet
	oSumSplitRange.StartRow = poGLRange.StartRow + 2	'// first COA row
	oSumSplitRange.EndRow = poGLRange.EndRow - 1		'// last COA row
	oSumSplitRange.StartColumn = poGLRange.StartColumn	'// Date column
	oSumSplitRange.EndColumn = poGLRange.EndColumn		'// Reference column
	lSumRow = poGLRange.EndRow
	oGLCellAcct = oGLSheet.getCellByPosition(COLACCT, lSumRow)
	sFormula = oGLCellAcct.String			'// just for kicks
	
ErrLogSetSheet(oGLSheet.RangeAddress.Sheet)
ErrLogSetCellInfo(COLACCT, lSumRow)
	'// oGLCellDebit contains last Debit value; non-empty 
	'// bMoveDebits would be true, so pass the opposite
	'// to SetSumFormula..
	oSumSplitRange.StartRow = poGLRange.StartRow	'// 1st "split"	'// mod061020
	oSumSplitRange.EndRow = poGLRange.EndRow		'// 2nd "split"	'// mod061020
	bOpposite = (Len(oGLCellDebit.String)=0)
	sFormula = SetSumFormula(oSumSplitRange, bOpposite)		
	oGLCellAcct.String = sFormula			'// set formula in cell
	oGLCellAcct.setFormula(sFormula)
	oGLCellAcct.CellBackColor = LTGREEN
	oGLCellAcct.NumberFormat = DEC2	
	Call SetSheetDate(oSumSplitRange, MMDDYY)
	GoTo NormalExit
	
ErrHandler:
	'// lCat2InsRow < 0 sets error code.
	'//*-----------------restore error handling for caller-----------
	'// restore entry error settings.
	ErrLogSetModule(sErrCallerMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
	'//*----------end error handling restore----------------------------

	iRetValue = lCat2InsRow
	
NormalExit:
	'// restore caller error handling.
	
	SplitOutCOAs = iRetValue


end function 	'// end SplitOutCOAs	8/22/22.
'/**/

'// StoreTrans.bas
'//---------------------------------------------------------------
'// StoreTrans - Store double-entry transaction.
'//		6/11/20.	wmk.	21:30
'//---------------------------------------------------------------

public function StoreTrans(poGLRange as Object, poCat1Range as Object,_
			poCat2Range as Object, psAcct1 As String,_
					psAcct2 As String) as Integer

'//	Usage.	iVal = StoreTrans(oGLRange, oCat1Range, oCat2Range,
'//							sAcct1, sAcct2)
'//
'//		oGLRange = RangeAddress of double-entry transaction
'//		oCat1Range = RangeAddress of COA sheet for 1st line
'//		oCat2Range = RangeAddress of COA sheet for 2nd line
'//		sAcct1 = COA account number for line 1 of transaction
'//		sAcct2 = COA account number for line 2 of transaction
'//
'// Entry.	oGLRange.Sheet = sheet index for GL 2-line transaction
'//					.StartRow = row index of 1st line of transaction
'//			oCat1Range.Sheet = sheet index for COA of line 1
'//					  .StartRow = insertion row index for new line
'//			oCat2Range.Sheet = sheet index for COA of line 2
'//					  .StartRow = insertion row index for new line
'//			sAcct1 may equal sAcct2 if double entry is for only one
'//				COA account
'//
'//	Exit.	iVal = 0 - no error occurred during insertion of new lines
'//					into both COA account sheets
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/2/20.		wmk.	original code; stub
'//	6/2/20.		wmk.	code transported in from PlaceTransM and
'//						modified
'//	6/6/20.		wmk.	runtime bugs fixed; SetSheetDate called for changes
'//						to COA sheets
'//	6/11/20.	wmk.	bug fix where stray dValue used instead of
'//						dSetValue when debit/credit fields blank
'//                     in swapped rows
'//	Notes.


'//	constants.

'//	local variables.
dim iRetValue As Integer
dim oDoc As Object				'// This Component
dim oSheets As Object			'// Doc.getSheets()
dim oGLSheet As Object			'// ledger sheet
dim oCat1Sheet As Object		'// COA sheet for line 1
dim oCat2Sheet As Object		'// COA sheet for line 2
dim iSheetIx As Integer			'// ledger sheet index
dim iCat1Ix As Integer			'//	COA line 1 sheet index
dim iCat2Ix As Integer			'// COA line 2 sheet index
dim lGLRow As Long				'// ledger row

'//	ledger access variables.
dim oCellDate As Object		'// ledger date field
dim oCellTrans As Object	'// transaction field
dim oCellDebit As Object	'// Debit field	
dim oCellCredit As Object	'// Credit field
dim oCellAcct As Object		'// COA field
dim oCellRef As Object		'// reference field
dim sRef as String			'// reference string
dim sTrans As String		'// transaction text string
dim dSetValue As Double		'// transaction value to set
dim iNumberFormat As Integer	'// transaction number format
dim sAcct As String			'// line 1 COA
dim sAcctB As String		'// line 2 COA

'//	COA transaction 1st line access variables.
dim oCat1Date As Object		'// ledger date field
dim oCat1Trans As Object	'// transaction field
dim oCat1Debit As Object	'// Debit field	
dim oCat1Credit As Object	'// Credit field
dim oCat1Acct As Object		'// COA field
dim oCat1Ref As Object		'// reference field

'// COA transaction 2nd line access variables.
dim oCat2Date As Object		'// ledger date field
dim oCat2Trans As Object	'// transaction field
dim oCat2Debit As Object	'// Debit field	
dim oCat2Credit As Object	'// Credit field
dim oCat2Acct As Object		'// COA field
dim oCat2Ref As Object		'// reference field


	'// code.
	
	iRetValue = 0
	
	StoreTrans = iRetValue

'	if true then
'		Exit Function
'	endif
	
'//-----------------StoreTrans code begins here
'dim oCat1Date as Object		'// date cell from accounting category 1 sheet
'dim oCat2Date as Object		'// date cell from accounting category 2 sheet
'const LTGREEN=10092390			'// light green background color
'// StoreTrans - Store transactions in inserted rows
'// parameters: oGLRange, oCat1Range, oCat2Range, Acct1, sAcct2
'// Entry conditions.
'// Entry.	oGLRange.Sheet = sheet index for GL 2-line transaction
'//					.StartRow = row index of 1st line of transaction
'//			oCat1Range.Sheet = sheet index for COA of line 1
'//					  .StartRow = insertion row index for new line
'//			oCat2Range.Sheet = sheet index for COA of line 2
'//					  .StartRow = insertion row index for new line
'//			sAcct1 may equal sAcct2 if double entry is for only one
'//				COA account
'//'//	lCat1Sheet = COA sheet for sAcct1
'// lCat2Sheet = COA sheet for sAcct2
'//	lCat1InsRow = insert fow for lCat1Sheet
'// lCat2InsRow = insert row for lCat2Sheet
'//	oCellDate = Date cell information from 1st ledger liner
'// sTrans = transaction description field from ledger
'//	oCellDebit = Debit cell information from 1st ledger line
'//	oCellDebitB = Debit cell information from 2nd ledger line
'// oCellCredit = Credit cell information from 1st ledger line
'// oCellCreditB = Credit cell inromation from 2nd ledger line
'// oCellAcct = COA # from 1st ledger line
'// oCellAcctB = COA # from 2nd ledger line
'// oCellRef = Reference field from 1st ledger line
'// oCellRefB = Reference field from 2nd ledger line

'// Notes: sAcct, sAcctB were set above 
'//	the transaction dates had to match, so are the same
'// the transaction descriptions had to match, so are the same
'// the Debit and Credit amounts had to match, so are the same
'// the Reference fields can be assumed to be the same
'// the COA#s are different, but have been preserved in sAcct, sAcctB
'// passed parameter is the poGLRange which has all the information
'// necessary to acces the first transaction row, from which
'// all date except the COAs can be extracted and duplicated in both
'// COA sheet entries
'// setup code can be this..
'	oDoc = ThisComponent
'	oSheets = oDoc.getSheets()
'	iSheetIx = oGLRange.Sheet
'	lGLRow = oGLRange.StartRow
'	oGLSheet = oSheets.getByPosition(iSheetIx)
'	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLRow)
'	oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLRow)
'	oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLRow)
'	'// Debit and Credit are same values; balanced transaction
'	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLRow)

	'// copy account strings to local vars
	sAcct = psAcct1			'// 1st COA
	sAcctB = psAcct2		'// 2nd COA
	
	'// set up access to all relevant sheets
	oDoc = ThisComponent
	oSheets = oDoc.getSheets()
	iSheetIx = poGLRange.Sheet
	iCat1Ix = poCat1Range.Sheet
	iCat2Ix = poCat2Range.Sheet
	oGLSheet = oSheets.getByIndex(iSheetIx)
	oCat1Sheet = oSheets.getByIndex(iCat1Ix)
	oCat2Sheet = oSheets.getByIndex(iCat2Ix)
	
	'// set up get/insert information for each COA sheet
	lGLRow = poGLRange.StartRow
dim lCat1InsRow As Long
dim lCat2InsRow As Long
	lCat1InsRow = poCat1Range.StartRow
	lCat2InsRow = poCat2Range.StartRow
	
	if iCat1Ix = iCat2Ix then
	
		'// if same sheet, adjust the insert row of the highest COA by 1
		'// only if the lower COA is in the first line of the transaction
		'// since the row number of the second line of the transaction
		'// will change by one when the lower COA inserts its line

		if StrComp(sAcct, sAcctB) <= 0 then									'// mod051920
			lCat2InsRow = lCat2InsRow + 1									'// mod051920
		endif																'// mod051920
		
	endif	'// end same sheet for both COAs conditional	

	'// get row index of 1st account sheet for insertion
	'// insert new line in account sheet with 2nd COA number
	'// set date field color in inserted rows to LTGREEN
	oCat1Sheet.Rows.insertByIndex(lCat1InsRow, 1)	'// insert new category 1 row
	oCat2Sheet.Rows.insertByIndex(lCat2InsRow, 1)	'// insert new category 2 row
	oCat1Date = oCat1Sheet.getCellByPosition(COLDATE, lCat1InsRow)
	oCat1Date.Text.CellBackColor = LTGREEN
	oCat2Date = oCat2Sheet.getCellByPosition(COLDATE, lCat2InsRow)
	oCat2Date.Text.CellBackColor = LTGREEN
		
dim iBrkPt as Integer
iBrkpt=1
		
	'// extract key ledger fields for setting in new entries
	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLRow)
	oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLRow)
	oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLRow)
	oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLRow)
	oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLRow)
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLRow)

	'// set new date fields to line 1 date
	oCat1Date.setValue(oCellDate.getValue())
	oCat2Date.setValue(oCellDate.getValue())
	oCat1Date.Text.NumberFormat = oCellDate.Text.NumberFormat
	oCat2Date.Text.NumberFormat = oCellDate.Text.NumberFormat
	oCat1Date.Text.HoriJustify = RJUST		'// reset alignment since row above might be diff
	oCat2Date.Text.HoriJustify = RJUST		'// reset alignment since row above might be diff

iBrkpt=1

	'// duplicate transaction field in both category sheets/accounts
	sTrans = oCellTrans.String
	oCat1Trans = oCat1Sheet.getCellByPosition(COLTRANS, lCat1InsRow)
	oCat2Trans = oCat2Sheet.getCellByPosition(COLTRANS, lCat2InsRow)
	oCat1Trans.String = sTrans
	oCat2Trans.String = sTrans
	oCat1Trans.Text.HoriJustify = LJUST		'// left-justify since row above might be diff		'// mod051620
	oCat2Trans.Text.HoriJustify = LJUST
		
iBrkpt=1

	'// copy old credit and debit fields to new sheets
	oCat1Debit = oCat1Sheet.getCellByPosition(COLDEBIT, lCat1InsRow)
	oCat2Debit = oCat2Sheet.getCellByPosition(COLDEBIT, lCat2InsRow)
	oCat1Credit = oCat1Sheet.getCellByPosition(COLCREDIT, lCat1InsRow)
	oCat2Credit = oCat2Sheet.getCellByPosition(COLCREDIT, lCat2InsRow)
	
	'// Test for empty strings and set accordingly to avoid setting 0 values.
	if Len(oCellDebit.Text.String) <> 0 then
		'// line 1 debit is non-empty, so line cat1 Debit = line 1 debit
		'//   and cat1 Credit is empty
		dSetValue = oCellDebit.getValue()
		iNumberFormat = oCellDebit.NumberFormat
		oCat1Debit.setValue(dSetValue)
		oCat1Debit.Text.NumberFormat = iNumberFormat
		oCat1Debit.Text.HoriJustify = RJUST
		oCat1Credit.String = ""
		'// AND line 2 credit = line 1 debit; line 2 Debit is empty
		oCat2Credit.setValue(dSetValue)
		oCat2Credit.Text.NumberFormat = iNumberFormat
		oCat2Credit.HoriJustify = RJUST
		oCat2Debit.String = ""
	else	'// line 1 debit is empty, so line2 credit should be empty
		dSetValue = oCellCredit.getValue()
		iNumberFormat = oCellCredit.NumberFormat
		oCat1Credit.setValue(dSetValue)
		oCat1Credit.Text.NumberFormat = iNumberFormat
		oCat1Credit.Text.HoriJustify = RJUST
		oCat1Debit.String = ""
		'// and line 2 debit = line 1 credit; line 2 Credit is empty
		oCat2Debit.setValue(dSetValue)
		oCat2Debit.Text.NumberFormat = iNumberFormat
		oCat2Debit.HoriJustify = RJUST
		oCat2Credit.String = ""
	endif

iBrkpt=1

	'// swap COA #s in transactions to xref
	oCat1Acct = oCat1Sheet.getCellByPosition(COLACCT, lCat1InsRow)
	oCat2Acct = oCat2Sheet.getCellByPosition(COLACCT, lCat2InsRow)
	
'	oCellAcctB.setValue(oCellAcct.getValue())
'	oCellAcctB.HoriJustify = CJUST
'	oCellAcctB.NumberFormat = oAcct.NumberFormat
'	oCat1Acct.SetValue(oCellAcctB.getValue())
	iNumberFormat = oCellAcct.NumberFormat
	oCat1Acct.String = sAcctB		'// 2nd line account
	oCat1Acct.Text.HoriJustify = CJUST
	oCat1Acct.NumberFormat = iNumberFormat
	oCat2Acct.String = sAcct		'// 1st line account
	oCat2Acct.Text.HoriJustify = CJUST
	oCat2Acct.NumberFormat = iNumberFormat
		
iBrkpt=1

'// duplicate reference field and center it
	oCat1Ref = oCat1Sheet.getCellByPosition(COLREF, lCat1InsRow)
	oCat2Ref = oCat2Sheet.getCellByPosition(COLREF, lCat2InsRow)
	sRef = oCellRef.String
	oCat1Ref.String = sRef
	oCat2Ref.String = sRef
	oCat1Ref.Text.HoriJustify = CJUST
	oCat2Ref.Text.HoriJustify = CJUST

	'// Update Sheets modified date stamp to indicate changes		'// mod060620
	'// Note: COA sheets' dates updated in StoreTrans
	Call SetSheetDate(poCat1Range, MMDDYY)							'// mod060620
'const DATEROW=1
	oCellDate = oCat1Sheet.getCellByPosition(COLDEBIT, DATEROW)		'// mod060620
	oCellDate.CellBackColor = LTGREEN
	Call SetSheetDate(poCat2Range, MMDDYY)							'// mod060620
	oCellDate = oCat2Sheet.getCellByPosition(COLDEBIT, DATEROW)		'// mod060620
	oCellDate.CellBackColor = LTGREEN

	StoreTrans = iRetValue

end function 	'// end StoreTrans	6/11/20
'/**/

