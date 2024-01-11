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

