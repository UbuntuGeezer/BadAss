'// CheckDoubleEntry.bas
'//---------------------------------------------------------------
'// CheckDoubleEntry - Check double entry transaction fields.
'//		7/3/20.	wmk.
'//---------------------------------------------------------------

public function CheckDoubleEntry(poTransRange As Object, rpsMonth As String,_
					rpsAcct1 As String, rpsAcct2 As String) As Integer

'//	Usage.	iVal = CheckDoubleEntry(oTransRange, rpsMonth, rpsAcct1,
'//						rpsAcct2)
'//
'//		oTransRange = RangeAddress of double-entry transaction
'//		rpsMonth = Month name string (e.g. "January")
'//		rpsAcct1 = (returned)
'//		rpsAcct2 = (returned)
'//
'// Entry.	<entry conditions>
'//
'//	Exit.	iVal = 0 if no error,
'//				   ERRCONTINUE (-1) if error; skip to next transaction
'//				   ERRSTOP (-2) if error; end processing user selction
'//			rpsAcct1 = COA from first line of transaction
'//			rpsAcct2 = COA from 2nd line of transaction
'//
'// Calls.	GetMonthName
'//
'//	Modification history.
'//	---------------------
'//	6/1/20.	wmk.	original code
'//	6/2/20.	wmk.	error handling setup added
'//	6/3/20.	wmk.	csERRBADDATE constant added
'// 6/5/20.	wmk.	bug fix; was not detecting 2nd line Debit and Credit empty
'//	7/3/20.	wmk.	sERRCOANOTFOUND corrected to csERRCOANOTFOUND
'//	Notes.
'//

const csERRCOANOTFOUND="ERRCOANOTFOUND"
const csERROUTOFROOM="ERROUTOFROOM"
const csERRCOACHANGED="ERRCOACHANGED"

'//----------in Module error handling setup-------------------------------
'// LogError setup snippet.
const csERRUNKNOWN="ERRUNKNOWN"
const sMyName="CheckDoubleEntry"
'// add additional error code strings here...
Dim sErrName as String		'// error name for LogError
Dim sErrMsg As String		'// error message for LogError
'//*---------error handling setup---------------------------
'//	ErrLogGetModule() - get current module name gsErrModule
'// ErrLogSetModule(sName) - set current module name gsErrModule
'// ErrLogGetCellInfo(lColumn, lRow) - get error focus cell column, row
'// ErrLogSetCellInfo(lColumn, lRow) - set error focus cell column, row
'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
dim iErrSheetIx as Integer
dim lErrColumn as Long
dim lErrRow as Long
dim sErrCurrentMod As String
'//----------end in Module error handling setup---------------------------

'//	constants.
'// process control constants.
const ERRCONTINUE=-1	'// error, but continue with next transaction
const ERRSTOP=-2		'// error, stop processing transactions
const csERRBADDATE="ERRBADDATE"		'// bad date error name

'//	local variables.
dim iRetValue As Integer
dim iBrkPt As Integer		'// easy breakpoint
dim iStatus As Integer		'// general status var


'//------------ CheckDoubleEntry code begins here---------------------------
'//	Usage.	iVal = CheckDoubleEntry(oTransRange, rpsMonth, rpsAcct1,
'//						rpsAcct2)
		dim Doc As Object
		dim oSheets as Object
		dim oGLSheet as Object		'// ledger sheet
		dim iSheetIx As Integer		'// sheet index of ledger
		dim bTransValid As Boolean	'// transaction is valid flag
  		dim bADEmpty As Boolean		'// 1st Debit/2nd Credit empty flag
  		dim bBCEmpty As Boolean		'// 1st Credit/2nd Debit empty flag
		
		'// Cell access objects for first transaction line.
		dim oCellDate as Object		'// Date cell
		dim oCellTrans As Object	'// Transaction cell
		dim oCellDebit As Object	'// Debit cell
		dim oCellCredit As Object	'// Credit cell
		dim oCellAcct As Object		'// COA cell
		dim oCellRef As Object		'// Reference cell
		dim sDate as String			'// Date string
		dim sTrans as String		'// Transaction string
		dim sStyle as String		'// Style property
		dim lDateFormat As Long		'// date format code
		dim sMonth As String		'// month name
		dim lGLCurrRow As Long		'// first row of double entry
		dim lGLNextRow As Long		'// 2nd row of double entry
		dim lGLEndRow As Long		'// end of ledger selection by user
		dim sDebit As String		'// Debit field
		dim sCredit As String		'// Credit field
		dim sAcct As String			'// COA field
		
		'// Cell access objects for 2nd transaction line.
		dim oCellDateB as Object	'// Date cell
		dim oCellTransB As Object	'// Transaction cell
		dim oCellDebitB As Object	'// Debit cell 
		dim oCellCreditB As Object	'// Credit cell
		dim oCellAcctB As Object	'// COA cell
		dim oCellRefB As Object		'// Reference cell
		dim sDateB As String
		dim sTransB As String
		dim sStyleB As String
		dim lDateBFormat As Long
		dim sDebitB As String
		dim sCreditB As String
		dim sAcctB As String
		
	'// code.

'//*------------error handling code initialization---------------
	'// preserve entry error settings.
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	sErrCurrentMod = ErrLogGetModule()
'//*	
	'// set local error settings.
	ErrLogSetModule(sMyName)
	ErrLogSetSheet(poTransRange.Sheet)
	ErrLogSetCellInfo(COLDATE, poTransRange.StartRow)
	ErrLogSetRecording(true)		'// enable log entries
'//*----------end error handling code initialization-------------

	iRetValue = 0
	iStatus = 0
'		poTransRange.Sheet = iSheetIx of ledger spreadsheet
'		poTransRange.StartColumn = COLDATE
'		poTransRange.StartRow = lGLCurrRow
'		poTransRange.EndColumn = COLREF
'		poTransRange.EndRow = lGLCurrRow + 1
		'// Access ledger sheet from passed information
	iSheetIx = poTransRange.Sheet
	Doc = ThisComponent
	oSheets = Doc.getSheets()
	oGLSheet = oSheets.GetByIndex(iSheetIx)
	lGLCurrRow = poTransRange.StartRow
	lGLEndRow = poTransRange.EndRow
			
		'// load 1st transaction line row from GL sheet
		oCellDate = oGLSheet.getCellByPosition( COLDATE, lGLCurrRow)
		oCellTrans = oGLSheet.getCellByPosition( COLTRANS, lGLCurrRow)
		oCellDebit = oGLSheet.getCellByPosition( COLDEBIT, lGLCurrRow)
		oCellCredit = oGLSheet.getCellByPosition( COLCREDIT, lGLCurrRow)
		oCellAcct = oGLSheet.getCellByPosition( COLACCT, lGLCurrRow)
		oCellRef = oGLSheet.getCellByPosition( COLREF, lGLCurrRow)

		'// extract date &' transaction information from 1st row
		sDate = oCellDate.String
		sTrans = oCellTrans.String
	'//	Ensure nonempty date field before GetMonthName call
	if Len(Trim(sDate)) = 0 then
		sErrName = "ERRBADDATE"
		sErrMsg = "1st Date field empty - skipping"
		iStatus = ERRSTOP
		GoTo DateError
	endif
		

iBrkpt=1
		sStyle = oCellDate.Text.CellStyle
		lDateFormat = oCellDate.Text.NumberFormat
		sMonth = GetMonthName(sDate)		'// get month search name
		
iBrkpt = 1	'// breakpoint line
		
		'// load 2nd transaction line row from GL sheet
		lGLNextRow = lGLCurrRow + 1
		
		'// check that 2nd row is not beyond selection
		if lGLNextRow > lGLEndRow then
'			iStatus = msgBox("Selection ends in mid-transaction!"_
'					+CHR(13)+CHR(10)+ "Check last row selection of transaction")	
			oCellDate.Text.CellBackColor = YELLOW
			sErrName = "ERRODDROWS"
			sErrMsg = "Selection ends in mid-transaction!"_
				+CHR(13)+CHR(10)+ "Check last row selection of transaction"
			Call LogError(sErrName, sErrMsg)
			iRetValue = ERRSTOP
			
			GoTo AdvanceTrans
			'exit For
		endif	'// end would advance past last row conditional

		'// load 2nd line transaction values
		oCellDateB = oGLSheet.getCellByPosition( COLDATE, lGLNextRow)
		oCellTransB = oGLSheet.getCellByPosition( COLTRANS, lGLNextRow)
		oCellDebitB = oGLSheet.getCellByPosition( COLDEBIT, lGLNextRow)
		oCellCreditB = oGLSheet.getCellByPosition( COLCREDIT, lGLNextRow)
		oCellAcctB = oGLSheet.getCellByPosition( COLACCT, lGLNextRow)
		oCellRefB = oGLSheet.getCellByPosition( COLREF, lGLNextRow)
'XRay oCellDateB

		'// extract date information from 2nd row of transaction
		sDateB = oCellDateB.Text.String
		
		'//	Ensure nonempty date field before GetMonthName call
		if Len(Trim(sDateB)) = 0 then
			sErrName = "ERRBADDATE"
			sErrMsg = "2nd Date field empty - skipping"
			GoTo DateError
		endif

		sStyleB = oCellDateB.Text.CellStyle
		lDateBFormat = oCellDateB.Text.NumberFormat
'		sMonth = GetMonthName(sDate)		'// use month from 1st line

iBrkpt = 1	'// breakpoint line
		
		'// make sure dates, descriptions, and amounts agree
			'// if not, issue message box and bail out
		iStatus = StrComp(sDate, sDateB)
		if iStatus <> 0 then
			sErrName = csERRBADDATE
			sErrMsg = "Transaction date mismatch!"+CHR(13)+CHR(10)_
						+ "Check 1st row selection of transaction"
			GoTo DateError
		endif

		GoTo CheckDesc
		
'//----------Date Error handler-------------------------------
iBrkpt = 1
'//	Date error(s) are terminating errors; iRetVal = -2
DateError:
		if iStatus <> 0 then
			'// log error and set Date Cell YELLOW
			oCellDate.Text.CellBackColor = YELLOW
			Call LogError(sErrName, sErrMsg)
			iRetValue = ERRSTOP		'// bail out on return
			GoTo AdvanceTrans		
		endif		'// end any date error conditional

'//-------end Date Error Handler-----------------------------

CheckDesc:
		
iBrkpt = 1	'// breakpoint line

		'// extract transaction information from both rows
		'// make sure that descriptions agree
		
		sTrans = oCellTrans.String
		sTransB = oCellTransB.String
		iStatus = strComp(sTrans, sTransB)
		
		if iStatus <> 0 then
'			iStatus = msgBox("Transaction description mismatch!"+CHR(13)+CHR(10)_
'						+ "Check transaction")
			sErrName = "ERRDESCNOMATCH"
			sErrMsg = 	"Transaction description mismatch!"+CHR(13)+CHR(10)_
						+ "Check transaction"
			Call LogError(sErrName, sErrMsg)
			oCellDate.Text.CellBackColor = YELLOW
			oCellDateB.Text.CellBackColor = YELLOW
			iRetValue = ERRSTOP			'// force stop on return
			GoTo AdvanceTrans
		endif		'// end descriptions not same

		'// extract debit and credit amounts for both halves of transaction
		'// ensure both halves have same amount, if not, bail out
		
		sDebit = oCellDebit.Text.String
		sCredit = oCellCredit.Text.String

		sDebitB = oCellDebitB.Text.String
		sCreditB = oCellCreditB.Text.String

'//-------------------------------------------------------------------------
		'// ensure that is actual double entry; either debit on 1st line and
		'// Credit on second line are same amount, or Credit on 1st line
		'// and Debit on second line are same amount etc.

'//	Debit Credit
'//   A     B			sDebit and sCredit
'//   C     D			sDebitB and sCreditB		
'//   A=D and B=C		correct
'//  if A and D empty, B and C must be nonempty
'//  if B and C empty, A and D must be nonempty
'//-------------------------------------------------------------------------
iBrkpt = 1 
 
' dim bTransValid as boolean
' dim bADEmpty as boolean
' dim bBCEmpty as boolean
  		bTransValid = (StrComp(sDebit,sCreditB) = 0) AND  (StrComp(sCredit,sDebitB)=0)
  		bTransValid = bTransValid AND ((Len(sDebit) > 0) OR (Len(sCredit) > 0))_
  								  AND  ((Len(sCreditB) > 0) OR (Len(sDebitB) > 0))			'// mod060520
  		bADEmpty = (Len(sDebit)=0) AND (Len(sCreditB)=0)
  		bBCEmpty = (Len(sCredit)=0) AND (Len(sDebitB)=0)
  		bTransValid = bTransValid AND ((bADEmpty AND (Not bBCEmpty))_
  									 OR (bBCEmpty AND (Not bADEmpty)))						'// mod052020
  		if Not bTransValid then
  		'// message, flag both lines, skip
'  			iStatus = msgBox("Invalid transaction...check Debit and Credit values!"+CHR(13)+CHR(10)_
'  							+"Transaction skipped")
			sErrName = "ERRDEB<>CRED"
			sErrMsg = "Invalid transaction...check Debit and Credit values!"+CHR(13)+CHR(10)_
  							+"Transaction skipped"
  			Call LogError(sErrName, sErrMsg)
  			oCellDate.Text.CellBackColor = YELLOW
			oCellDateB.Text.CellBackColor = YELLOW
			iRetValue = ERRCONTINUE
			GoTo AdvanceTrans							
  		endif	'// end invalid transaction
  		    
		'// extract COA account numbers for both halves of transaction
		sAcct = oCellAcct.String
		sAcctB = oCellAcctB.String
iBrkpt = 1
		'// ensure both halves have COA account number; if not, just skip
		if Len( sAcct ) <> 4 OR Len(sAcctB) <> 4 then
			iStatus = msgBox("Transaction contains invalid COA number(s)!"_
					+CHR(13)+CHR(10)+ "Check 1st row selection of transaction")	
			sErrName = csERRCOANOTFOUND
			sErrMsg = "Transaction contains invalid COA number(s)!"_
				+CHR(13)+CHR(10)+ "Check 1st row selection of transaction"
			Call LogError(sErrName, sErrMsg)
			oCellDate.Text.CellBackColor = YELLOW
			oCellDateB.Text.CellBackColor = YELLOW
			'// Note. have to do the following rather than use if-then-else to loop
			'// again, since this Basic does not support "Continue For"
			iRetValue = ERRCONTINUE 		'// force Advance to next on return
			GoTo AdvanceTrans
		endif	'// end length not == 4
		
		'// set returned month and COA accounts
		rpsMonth = sMonth
		rpsAcct1 = sAcct
		rpsAcct2 = sAcctB
		
AdvanceTrans:
'//*-----------------restore error handling for caller-----------
	'// restore entry error settings.
	ErrLogSetModule(sErrCurrentMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
'//*----------end error handling setup---------------------------

	'// iRetValue = 0, no error
	'//				ERRCONTINUE (-1) - error; continue with next transaction
	'//				ERRSTOP (-2) - error; stop processing transactions
	CheckDoubleEntry = iRetValue

end function 	'// end CheckDoubleEntry	7/3/20
'/**/
