'// LoadCOAs.bas
'//---------------------------------------------------------------
'// LoadCOAs - Load COA account#s for list boxes.
'//		10/22/25.	wmk.	17:59
'//---------------------------------------------------------------

public function LoadCOAs() as Integer

'//	Usage.	iVal = LoadCOAs()
'//
'// Entry.	pusCOA_Assets = array allocated for Asset COAs
'//			pusCOA_Liabilities = array for Liability COAs
'//			pusCOA_Income = array for Income COAs
'//			pusCOA_Expenses = array for Expense COAs
'//			pusCOAList = array for COA list for listbox
'//			COA_SHEET = sheet name of COA worksheet
'//			COA_COLSTART = starting column of COA loadable list
'//			COA_ROWSTART = starting row of COA loadable list
'//			sCOA_1STROWID = string to match verifying 1st row COA list
'//			sCOA_LASTROWID = string to match verifying last row COA list
'//			COA_NASSETS = Assets COA entry count
'//
'//	Exit.	iVal = 0 - no errors encountered
'//				  -1 - bad load
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'// 10/22/25.	wmk.	(automated) Modification History sorted.
'// 7/6/20.		wmk.	modified to get COA entry count from 2 columns to 
'// 7/6/20.		 right of "BEGIN COA" (COA_COLROWS). 
'// 6/30/20.	wmk.	comments updated. 
'// 6/14/20.	wmk.	original. 
'//
'//	Notes. Initial version knows how many rows to expect by public constants
'//	set internally for each accounting category. Future version can be
'// generalized to load "n" rows where "n" is specified at the head of
'// the COA table in the Chart of Accounts.


'//	constants.
const ERRNOTCOATBL=-2		'// COA table not where expected

'//	local variables.
dim iRetValue As Integer
dim iAsset As Integer		'// asset COA counter
dim iLiability As Integer	'// liability COA counter
dim iIncome As Integer		'// income COA counter
dim iExpense As Integer		'// expense COA counter
dim oDoc As Object			'// component
dim oSheets As Object		'// sheets array
dim oCOASheet As Object		'// COA sheet
dim oCOACell As Object		'// COA cell general object
dim oCOACntCell	As Object	'// cell with COA count
dim bIsCOATable As Boolean	'// is COA table flag
dim sCOAList(0) As String	'// COA list array

dim iRowsExpected As Integer		'// expected row count
dim i	As Integer					'// loop count
dim oAcctCell As Object				'// account name cell
dim bTableEndFound As Boolean		'// end of table found
dim lCurrRow As Long				'// current row index
dim iBrkPt As Integer				'// breakpoint finder


'//----------in sub/function error handling setup-------------------------------
'// Error Handling setup snippet.
	const sMyName="LoadCOAs"
	dim lLogRow as long		'// cell row working on
	dim lLogCol as Long		'// cell column working on
	dim iLogSheetIx as Integer	'// sheet index module working on
	dim oLogRange as new com.sun.star.table.CellRangeAddress

'// LogError setup snippet.
	const csERRUNKNOWN="ERRUNKNOWN"
'// add additional error code strings here...
	Dim sErrName as String		'// error name for LogError
	Dim sErrMsg As String		'// error message for LogError

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sErrName = "ERRUNKNOWN"
	sErrMsg = "In LoadCOAs - unprocessed error."

	'// set up access to COA sheet
	oDoc = ThisComponent		'// get into document
	oSheets = oDoc.getSheets()
	oCOASheet = oSheets.getByName(COA_SHEET)
	
'//-----------in sub/function error handling initialization-------------
'//	dim oDoc as Object		'// ThisComponent
'//	dim oSheets as Object	'// oDoc.Sheets
'//	code to put sheet index into oLogRange.Sheet
	oLogRange.Sheet = oCOASheet.RangeAddress.Sheet
	lLogCol = COA_COLSTART
	lLogRow = COA_ROWSTART
	oLogRange.StartColumn = lLogCol
	oLogRange.StartRow = lLogRow
	oLogRange.EndRow = lLogRow
	oLogRange.EndColumn = lLogCol
	Call ErrLogSetup(oLogRange, sMyName)
	ErrLogSetRecording(true)
'//-----------end in sub/function error handling initialization---------

	
	iAsset = 0			'// clear counters
	iLiability = 0
	iIncome = 0
	iExpense = 0
	
	'// get COA table 1st row and verify
	oCOACell = oCOASheet.getCellByPosition(COA_COLSTART, COA_ROWSTART)
	bIsCOATable = (StrComp(Trim(oCOACell.String), sCOA_1STROWID)=0)
	if NOT bIsCOATable then
		iRetValue = ERRNOTCOATBL
		sErrName = "ERRNOTCOATBL"
		sErrMsg = "LoadCOAs/Invalid COA table"
		GoTo ErrorHandler
	endif

	oCOACntCell = oCOASheet.getCellByPosition(COA_COLROWS, COA_ROWSTART)

'	iRowsExpected = COA_NASSETS + COA_NLIABILITIES + COA_NINCOME _
'					+ COA_NEXPENSES
	iRowsExpected = oCOACntCell.Value
	if iRowsExpected < 1 then
		sErrName = "ERRCOACOUNT"
		sErrMsg = "LoadCOAs/Invalid COA count in table"
		GOTO ErrorHandler
	endif
redim pusCOAList(iRowsExpected-1)		'// minus 1, since 0-based

	lCurrRow = COA_ROWSTART + 1
	'// loop on rows until termination string found "****"
	for i = 0 to iRowsExpected
		oAcctCell = oCOASheet.getCellByPosition(COA_COLSTART, lCurrRow)
		bTableEndFound = (StrComp(oAcctCell.String, sCOA_LASTROWID)=0)
		if bTableEndFound then
			if i <> iRowsExpected then
				iRetValue = ERRCOACOUNT
				sErrName = "ERRCOACOUNT"
				sErrMsg = "LoadCOAs/Table entry count mismatch"
				GoTo ErrorHandler
			else
				exit for
			endif	'// end wrong row count conditional			
		endif	'// end table end found conditional
	
		'// combine and store COA strings
		oCOACell = oCOASheet.getCellByPosition(COA_COLSTART+1, lCurrRow)
		pusCOAList(i) = oCOACell.String +" " + oAcctCell.String
		lCurrRow = lCurrRow + 1
	next i

	'// check for premature table end
	if i < iRowsExpected then
		iRetValue = ERRCOACOUNT
		sErrName = "ERRCOACOUNT"
		sErrMsg = "LoadCOAs/Table entry count mismatch"
		GoTo ErrorHandler
	endif
	
iBrkPt = 1		'// no errors		
					
	iRetValue = 0		'// set normal return
	
NormalExit:
	call ErrLogDisable()		'// disable error logging
	iRetValue = iRowsExpected	'// return row count
	LoadCOAs = iRetValue
	exit function
	
ErrorHandler:
	Call LogError(sErrName, sErrMsg)
	GoTo NormalExit
	
end function 	'// end LoadCOAs	10/23/25.
'/**/
