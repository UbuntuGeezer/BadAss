'// ReverseSplitTrans.bas
'//---------------------------------------------------------------
'// ReverseSplitTrans - Reverse split transaction.
'//		11/14/22.	wmk.	08:16
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
'// 11/14/22.	wmk.	change Ref field background color to 0 (black).
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
		oCellRef.CellBackColor = 0
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
