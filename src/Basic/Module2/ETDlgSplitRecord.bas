'// ETDlgSplitRecord.bas
'//---------------------------------------------------------------
'// ETDlgSplitRecord - Record split transaction from ET Dialog.
'//		7/7/20.	wmk.	12:00	
'//---------------------------------------------------------------

public function ETDlgSplitRecord() As Integer

'//	Usage.	iVar = ETDlgSplitRecord()
'//
'// Entry.	ET Dialog public vars have entered values
'//			gbETDebitIsTotal = true if credits are split
'//							   false if debits are split
'//
'//	Exit.	Double entry split transaction recorded in GL sheet at
'//			user selected position
'//
'// Calls.	SetSumFormula
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; adapted from ETDialogRecord
'//	6/26/20.	wmk.	bug fix where both debit and credit splits being
'//						recorded in Credit column
'//	6/27/20.	wmk.	bug fix where transaction rows being set to
'//						color of row above where split inserted; need
'//						to be set to NOFILL
'// 7/7/10.		wmk.	bug fix where total line background color not being
'//						set to NOFILL; bug fix where description centered if
'//						line above was centered
'//
'//	Notes. When the ET dialog completes, either the DebitField or CreditField
'// may have been set to "split". If that is the case, ETDlgSplitRecord will
'// be called to record the split transaction.
'//

'//	constants.

'//----------------constants borrowed from header.bas-------------------
const LTGREEN=10092390		'// decimal value of LTGREEN color

'// column index values and string lengths for column data			
'public const	INSBIAS=3		'// insert line count bias
'public const	COALEN=4		'// length of COA field strings
const COLDATE=0		'// Date - column A
const COLTRANS=1	'// Transaction - column B
const COLDEBIT=2	'// Debit - column C
const COLCREDIT=3	'// Credit - column D
const COLACCT=4		'// COA Account - column E
const COLREF=5		'// Reference - column F

'// cell formatting constants.
const DEC2=123		'// number format for (x,xxx.yy)			'// mod052020
const MMDDYY=37		'// date format mm/dd/y						'// mod052020
const FMTDATETIME=50	'// date/time format mm/dd/yyy hh:mm:ss 	'// mod060620
const LJUST=1		'// left-justify HoriJustify				'// mod052020
const CJUST=2		'// center HoriJustify						'// mod052020
const RJUST=3		'// right-justify HoriJustify				'// mod052320
const MAXTRANSL=50	'// maximum transaction text length			'// mod052020

'//----------------end constants borrowed from publics.bas--------------

'//	local variables
dim iRetValue As Integer
dim i		As Integer
dim iStatus	As Integer			'// general status
dim sSumFormula	As String		'// split sum formula

'// GL Sheet access variables for inserting transaction
dim oDoc		As Object		'// ThisComponent
dim oSel		As Object		'// current user selection
dim oRange		As Object		'// cell range selection info
dim iSheetIx	As Integer		'// sheet index
dim oGLSheet	As Object		'// GL sheet object
dim lGLRow		As Long			'// GL 1st line row
dim lGLCurrRow	As Long			'// current new transaction row
dim lGLBottomRow	As Long		'// bottom row of split transaction
dim iInsertCnt	As Integer		'// insertion row count
dim iSplitLim	As Integer		'// splits limit 0-based

'// GL Transaction field cells
dim	oCellDate 	As Object		'// Date field
dim	oCellTrans 	As Object		'// Description field
dim	oCellDebit 	As Object		'// Debit field
dim	oCellCredit	As Object		'// Credit field
dim	oCellAcct	As Object		'// COA field
dim	oCellRef	As Object		'// Reference field
dim oDate		As Object		'// oCellDate.Text

	'// code.
	
	iRetValue = 0
	ON ERROR GOTO ErrorHandler

	'// set up access to GL on user selection
	oDoc = ThisComponent				'// set up for sheet access within document
	oSel = oDoc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet	'// get sheet index value
	oGLSheet = oDoc.Sheets.getByIndex(iSheetIx)	'// get GL sheet
	lGLRow = oRange.StartRow			'// set row pointer to 1st row
	lGLCurrRow = lGLRow
	
	'// rows to insert = upper bound of amounts + 1, +2 for upper and
	'// lower "split" divider rows +1 for total row
	iSplitLim = uBound(gdSTAmts)
	iInsertCnt = iSplitLim + 4
	lGLBottomRow = lGLCurrRow + iInsertCnt - 1	
	'// set up and insert iInsertCnt rows at first row pointed to by user
	oGLSheet.Rows.insertByIndex(lGLRow, iInsertCnt)	'// insert transaction rows
	oRange.EndRow = lGLBottomRow			'// set for SetSumFormula call
		
	'// update sheet date modified
	iStatus = SetSheetDate(oRange, MMDDYY)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif	'// end abnormal return from SetSheetDate

'/*--------------------------------------------------------------------------
	'// create top row "split" divider
	'// extract key ledger fields for setting in new entries
	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
	oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
	oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
	oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
	oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		
	'// set Date field in transaction line		
	oCellDate.setValue(DateValue(gsETDate))		'// set date field to value
	oCellDate.NumberFormat = MMDDYY				'// set mm/dd/yy format
	oCellDate.HoriJustify = RJUST				'// right-justify date
	oCellDate.CellBackColor = LTGREEN			'// set light green background

	'// set description field to ET description
	'// gsETDescription = entered description
	oCellTrans.String = gsETDescription
	oCellTrans.CellBackColor = LTGREEN			'// color entire 1st row
	oCellTrans.HoriJustify = LJUST
	
	'//	set debit and credit amount fields to ""
	oCellDebit.String = ""
	oCellDebit.CellBackColor = LTGREEN
	oCellCredit.String = ""
	oCellCredit.CellBackColor = LTGREEN
	
	'// set account and ref fields to "-" and "split"
	oCellAcct.String = "-"
	oCellAcct.HoriJustify = CJUST
	oCellAcct.CellBackColor = LTGREEN
	oCellRef.String = "split"
	oCellRef.HoriJustify = CJUST
	oCellRef.CellBackColor = LTGREEN
'*/----------------------------------------------------------------------------

	lGLCurrRow = lGLCurrRow + 1		'// advance to total row
'/*------------------enter total row-------------------------------------------
	'// extract key ledger fields for setting in new entries
	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
	oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
	oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
	oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
	oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		
	'// set Date field in transaction line		
	oCellDate.setValue(DateValue(gsETDate))		'// set date field to value
	oCellDate.NumberFormat = MMDDYY				'// set mm/dd/yy format
	oCellDate.HoriJustify = RJUST				'// right-justify date
	oCellDate.CellBackColor = LTGREEN			'// set light green background

	'// set description field to ET description
	'// gsETDescription = entered description
	oCellTrans.String = gsETDescription
	oCellTrans.CellBackColor = NOFILL
	oCellTrans.HoriJustify = LJUST

	'// set Total in debit or credit and appropriate COA
	if gbETDebitIsTotal then		'// credits split
		oCellDebit.Value = gdETAmount
		oCellDebit.NumberFormat = DEC2
		oCellCredit.String = ""
		oCellAcct.String = gsETAcct1
	else		'// debits split
		oCellCredit.Value = gdETAmount
		oCellCredit.NumberFormat = DEC2
		oCellDebit.String = ""
		oCellAcct.String = gsETAcct2
	endif		'// end debit is total conditional
	oCellDebit.HoriJustify = RJUST
	oCellDebit.CellBackColor = NOFILL
	oCellCredit.HoriJustify = RJUST
	oCellCredit.CellBackColor = NOFILL
	oCellAcct.HoriJustify = CJUST
	oCellAcct.CellBackColor = NOFILL
		
	'// set ref field
	oCellRef.String = gsETRef
	oCellRef.HoriJustify = CJUST
	oCellRef.CellBackColor = NOFILL
'*/---------end Total row------------------------------------------------------
 if false then
 GoTo JumpAhead
 endif
 
'/*----------------------------------------------------------------------------
	'// for each row, copy transaction information, placing debit amt
	'// in Debit column of 1st row, credit amt in Credit column of
	'// 2nd row, all other columns in same position
	lGLCurrRow = lGLCurrRow + 1		'// advance past total row
	for i = 0 to iSplitLim
		'// extract key ledger fields for setting in new entries
		oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
		oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
		oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
		oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
		oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
		oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		
		'// set Date field in transaction line		
		oCellDate.setValue(DateValue(gsETDate))		'// set date field to value
		oCellDate.NumberFormat = MMDDYY				'// set mm/dd/yy format
		oCellDate.HoriJustify = RJUST				'// right-justify date
		oCellDate.CellBackColor = LTGREEN			'// set light green background
		
		'// set Description field
		oCellTrans.String = gsSTDescs(i)
		oCellTrans.HoriJustify = LJUST
		oCellTrans.CellBackColor = NOFILL
		
		'// set amount in proper column
		if gbETDebitIsTotal then		'// credits split
			oCellDebit.String = ""
			oCellCredit.String = Str(gdSTAmts(i))
			oCellCredit.setValue(gdSTAmts(i))					
			oCellCredit.NumberFormat = DEC2
			oCellAcct.String = gsSTCOAs(i)
		else		'// debits split
			oCellDebit.String = Str(gdSTAmts(i))
			oCellCredit.String = ""
			oCellDebit.setValue(gdSTAmts(i))					
			oCellDebit.NumberFormat = DEC2
			oCellAcct.String = gsSTCOAs(i)
		endif		'// end debit is total conditional
	
		oCellDebit.HoriJustify = RJUST
		oCellDebit.CellBackColor = NOFILL
		oCellCredit.HoriJustify = RJUST
		oCellCredit.CellBackColor = NOFILL
		oCellAcct.HoriJustify = CJUST
		OCellAcct.CellBackColor = NOFILL
		
		oCellRef.String = gsSTRefs(i) 
		oCellRef.HoriJustify = CJUST
		oCellRef.CellBackColor = NOFILL
		
		lGLCurrRow = lGLCurrRow + 1	'// advance to next transaction row
		
	next i		'// advance split transaction row
'*/-----------------------------------------------------------------------------
 if false then
 GoTo JumpAhead
 endif

'/*--------------------------------------------------------------------------
	'// create bottom row "split" divider with total
	'// extract key ledger fields for setting in new entries
	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
	oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
	oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
	oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
	oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		
	'// set Date field in transaction line		
	oCellDate.setValue(DateValue(gsETDate))		'// set date field to value
	oCellDate.NumberFormat = MMDDYY				'// set mm/dd/yy format
	oCellDate.HoriJustify = RJUST				'// right-justify date
	oCellDate.CellBackColor = LTGREEN			'// set light green background

	'// set description field to ET description
	'// gsETDescription = entered description
	oCellTrans.String = gsETDescription
	oCellTrans.CellBackColor = LTGREEN			'// color entire last row
	oCellTrans.HoriJustify = LJUST
	
	'//	set debit and credit amount fields to ""
	oCellDebit.String = ""
	oCellDebit.CellBackColor = LTGREEN
	oCellCredit.String = ""
	oCellCredit.CellBackColor = LTGREEN
	
	'// set account field to appropriate SUM formula
	sSumFormula = SetSumFormula(oRange, gbETDebitIsTotal)
	oCellAcct.String = sSumFormula
	oCellAcct.SetFormula(sSumFormula)
	oCellAcct.NumberFormat = DEC2
	oCellAcct.HoriJustify = CJUST
	oCellAcct.CellBackColor = LTGREEN
	
	'// set ref field to "split"
	oCellRef.String = "split"
	oCellRef.HoriJustify = CJUST
	oCellRef.CellBackColor = LTGREEN
'*/----------------------------------------------------------------------------
JumpAhead:
	iRetValue = 0
	
NormalExit:
	ETDlgSplitRecord = iRetValue
	exit function
	
ErrorHandler:
	iRetValue = -1
	GoTo NormalExit
	

end function 	'// end ETDlgSplitRecord	7/7/20
'/**/
