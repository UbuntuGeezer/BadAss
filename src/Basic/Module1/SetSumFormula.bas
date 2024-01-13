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
