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

