'// ReverseDETrans.bas
'//---------------------------------------------------------------
'// ReverseDETrans - Reverse double-entry transaction.
'//		11/14/22.	wmk.	08:04
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
'// 11/14/22.	wmk.	clear Ref cell background color
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
	oCellRef.CellBackColor = 0			'// no background
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLOrigRow+1)
	oCellRef.String = "reversal"
	oCellRef.CellBackColor = 0			'// no background
	
	'// set normal return
	iRetValue = 0

NormalExit:	
	ReverseDETrans = iRetValue
	exit function
	
ErrorHandler:
	ReverseDETrans = iRetValue
	
end function 	'// end ReverseDETrans	11/14/22
'/**/
