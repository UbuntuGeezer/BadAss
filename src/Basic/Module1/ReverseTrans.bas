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
