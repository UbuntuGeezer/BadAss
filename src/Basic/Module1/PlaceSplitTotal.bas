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

