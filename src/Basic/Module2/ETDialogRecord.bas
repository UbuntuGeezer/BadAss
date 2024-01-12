'// ETDialogRecord.bas
'//---------------------------------------------------------------
'// ETDialogRecord - Record entered fields from ET Dialog.
'//		6/27/20.	wmk.	21:30
'//---------------------------------------------------------------

public function ETDialogRecord() As Integer

'//	Usage.	iVar = ETDialogRecord()
'//
'// Entry.	ET Dialog public vars have entered values
'//
'//	Exit.	Double entry transaction recorded in GL sheet at
'//			user selected position
'//			gbETSplitTrans = true if split transaction
'//							 false if regular transaction
'//			gbETDebitIsTotal = true if credits split;
'//							   false if debits split
'//
'// Calls.	ETDlgSplitRecord
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code; stub
'//	6/17/20.	wmk.	code fleshed out gradually
'//	6/24/20.	wmk.	change to record numeric debit and credit amount
'//						from gdETAmount
'//	6/25/20.	wmk.	bug fix eliminating gsETAmount which was replaced by
'//						gdETAmount 6/24
'// 6/27/20.	wmk.	ensure background in all but date fields is NOFILL
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

'// GL Sheet access variables for inserting transaction
dim oDoc		As Object		'// ThisComponent
dim oSel		As Object		'// current user selection
dim oRange		As Object		'// cell range selection info
dim iSheetIx	As Integer		'// sheet index
dim oGLSheet	As Object		'// GL sheet object
dim lGLRow		As Long			'// GL 1st line row
dim lGLCurrRow	As Long			'// current new transaction row

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

	'// determine if split transaction and handle
	gbETDebitIsTotal = (StrComp(gsETAcct2,"split")=0)
	gbETSplitTrans = (gbETDebitIsTotal OR (StrComp(gsETAcct1,"split")=0))
	if gbETSplitTrans then
		iStatus = ETDlgSplitRecord()
		if iStatus < 0 then
			GoTo ErrorHandler
		else
			GoTo NormalExit
		endif
	endif	'// end split transaction conditional
	
	'// set up access to GL on user selection
	oDoc = ThisComponent				'// set up for sheet access within document
	oSel = oDoc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet	'// get sheet index value
'	oSheets = oDoc.getSheets()
	oGLSheet = oDoc.Sheets.getByIndex(iSheetIx)	'// get GL sheet
	lGLRow = oRange.StartRow			'// set row pointer to 1st row
	
	'// set up and insert 2 rows at first row pointed to by user
	oGLSheet.Rows.insertByIndex(lGLRow, 2)	'// insert transaction rows
	'// add code here to update sheet date modified
	iStatus = SetSheetDate(oRange, MMDDYY)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif	'// end abnormal return from SetSheetDate
	
	'// for each row, copy transaction information, placing debit amt
	'// in Debit column of 1st row, credit amt in Credit column of
	'// 2nd row, all other columns in same position
	lGLCurrRow = lGLRow
	for i = 0 to 1
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
		
		'// add code here to set remaining transaction cells
		oCellTrans.String = gsETDescription
		oCellTrans.HoriJustify = LJUST
		oCellTrans.CellBackColor = NOFILL
		
		'// Debit amt is in first line; credit in second
		if i = 0 then
			oCellDebit.String = Str(gdETAmount)
			oCellDebit.setValue(gdETAmount)							'//	mod062420
			oCellDebit.NumberFormat = DEC2
			oCellCredit.String = ""
			oCellAcct.String = gsETAcct1
		else
			oCellDebit.String = ""
			oCellCredit.String = Str(gdETAmount)
			oCellCredit.setValue(Val(gdETAmount))					'//	mod062420
			oCellCredit.NumberFormat = DEC2
			oCellAcct.String = gsETAcct2
		endif
		oCellDebit.HoriJustify = RJUST
		oCellDebit.CellBackColor = NOFILL
		oCellCredit.HoriJustify = RJUST
		oCellCredit.CellBackColor = NOFILL
		oCellAcct.HoriJustify = CJUST
		oCellAcct.CellBackColor = NOFILL
				
		oCellRef.String = gsETRef 
		oCellRef.HoriJustify = CJUST
		oCellRef.CellBackColor = NOFILL
		
		lGLCurrRow = lGLCurrRow + 1	'// advance to next transaction row
		
	next i		'// advance transaction row

	GoTo NormalExit
	
NormalExit:
	ETDialogRecord = iRetValue
	exit function
	
ErrorHandler:
	iRetValue = -1
	GoTo NormalExit
	

end function 	'// end ETDialogRecord	6/25/20
'/**/
