'// ReverseTrans.bas
'//---------------------------------------------------------------
'// ReverseTrans - Reverse transaction.
'//		4/20/25.	wmk.
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
'// 4/20/25.	wmk.	superfluous function definition removed from beginning.
'//	6/12/20.	wmk.	add call to SetSheetDate.
'//	6/11/20.	wmk.	original code.
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


end sub		'// end ReverseTrans	4/20/25.
'/**/
