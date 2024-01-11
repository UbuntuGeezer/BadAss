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

