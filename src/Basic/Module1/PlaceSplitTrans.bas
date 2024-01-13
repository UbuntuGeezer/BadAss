'// PlaceSplitTrans.bas
'//------------------------------------------------------------------
'// PlaceSplitTrans - Place split transaction into multiple accounts.
'//		7/8/20.	wmk.	07:00
'//-------------------------------------------------------------------

function PlaceSplitTrans( poGLSheet as Object, plSplit1stRow as Long)_
	as long

'//	Usage.	lSplitEndRow = PlaceSplitTrans(oGLSheet, lSplit1stRow)
'//			call PlaceSplitTrans( oGLSheet, lSplit1stRow )
'//
'//		oGLSheet = GL sheet object
'//		lSplit1stRow = row index of first row of split transaction
'//
'// Entry.	GL sheet processing has found row with "split" as the
'//			.Text.String value in COLSPLIT. The split entry contains
'//			multiple rows until another row is found with "split" as the
'//			.Text.String value in COLSPLIT. The second row of the split
'//			transaction MUST contain the COA target and the total amount
'//			of the split, regardless of whether it is a Debit or Credit.
'//			Subsequent rows contain the amounts split into the opposing
'//			COAs, and their value sums MUST equal the COA target amount.
'//
'//	Exit.	lSplitEndRow >=0 - row index of GL last row of split
'//						 <0 - error, caller should terminate processing`
'//			ErrLogDisable() called, so caller will have to re-enable
'//			caller error logging
'//
'//	Calls.	CheckSplitTrans, GetTransSheetName, GetInsRow, GetMonthName
'//			ErrLogSetup, LogError, ErrLogDisable
'//
'//	Modification history.
'//	---------------------
'//	5/22/20		wmk.	original code
'//	5/23/20		wmk.	code added to put lines in split accounts; bug
'//						fixes; fix splits loop count; changes inserted
'//						rows Date field color to LTGREEN
'//	5/24/20.	wmk.	start adding code to copy rows into COA total;
'//						add insert error code checking and ErrorHandler
'//	5/26/20.	wmk.	error handling modified; error logging code added
'// 5/27/20.	wmk.	bug fixed where COA error in GL sheet being reported
'//						as in accounting category sheet; error code string
'//						ERRCOANOTFOUND corrected in constants fix bug where
'//						last split row with "split" not being copied into
'//						account where other accounts split from
'//	5/28/20.	wmk.	bug fix where debit/credit column swap copying 1
'//						too many rows; SUM formula insertion fix for COA
'//						category sheet entry of split details; feature implemented
'//						where split block inserted in sheet gets LTGREEN Date
'//						cell background color
'//	6/3/20.		wmk.	add code to change source transaction REF fields
'//						cell background color to LTGREEN between "split"
'//						rows; dead code block removed
'//	6/4/20.		wmk.	modularize code; add PlaceSplitTotal function;
'//						dead code and unused vars removed
'//	6/5/20.		wmk.	adjust code for SplitOutCOAs call; dead code
'//						removed; additional error codes from SplitTrans
'//	7/8/20.		wmk.	debugging superfluous msgbox removed
'//
'//	Notes.
'//	Any return value < 0 should be treated as a termination error,
'// and the caller should stop processing rows, as likely there will
'// only be errors in the subsequent rows until the user corrects
'// the GL spreadsheet entries.
'// Feature; 5/17/20; transfer "split" transaction to appropriate
'// multiple accounts (generalized add code to process split
'// transactions; caveat: all split transactions should have total in
'// first line with COA that will be used in all other line items as
'// moved to ease cross-referencing, regardless of whether total is a
'// debit or credit.
'//	Method.
'//	Stolen code from PlaceTransM. Only use oCat2Sheet, since only 
'// dealing with one target account sheet at a time. oGLSheet contains
'// the original "split" transaction lines. oCat2Sheet is the
'// category account sheet that gets each inserted line.

'//	constants. (see module-wide definitions, and add these)

'// local contants.

'// split controls and error codes.
const MAXSPLITROWS=13	'// maximum rows in "split" to "split"; 10 subaccounts
const ERRNOTSPLIT=-1	'// first fow not "split" in COLSPLIT
const ERRBADTALLY=-2	'// split values don't add to total
const ERRROWSEXCEEDED=-3	'// too many rows bedore second "split" found
const ERRBADVALS=-4		'// bad values, either missing or don't add
const ERRBADDATE=-5		'// Date field too short
const ERRDATENOTSAME=-6	'// Date field doesn't match others			'// mod060520
const ERRSPLITDESC=-7	'// "split" descriptions don't match		'// mod060520

'// insert error codes.
const ERRCOANOTFOUND=-1
const ERROUTOFROOM=-2
const ERRCOACHANGED=-3

'//------------------begin error handling block---------------
'// LogError setup snippet.

'// split error code strings.
const csERRUNKNOWN="ERRUNKNOWN"
const csERRNOTSPLIT="ERRNOTSPLIT"
const csERRBADTALLY="ERRBADTALLY"
const csERRROWSEXCEEDED="ERRROWSEXCEEDED"
const csERRBADVALS="ERRBADVALS"
const csERRBADDATE="ERRBADDATE"

'// insert error codes.
const csERRCOANOTFOUND="ERRCOANOTFOUND"
const csERROUTOFROOM="ERROUTOFROOM"
const csERRCOACHANGED="ERRCOACHANGED"

'//------------------end error handling block---------------

'//	local variables.

Dim Doc As Object			'// This component
Dim oSheets as Object		'// Sheets() collection

'//	GL XCell information objects (original line)
dim oGLCellDate as Object
dim oGLCellTrans as Object
dim oGLCellDebit as Object
dim oGLCellCredit as Object
dim oGLCellAcct as Object
dim oGLCellRef as Object

'//	oCat2Sheet XCell information objects.
dim oCat2Date as object		'// date of transaction
dim oCat2Trans as object	'// transaction description
dim oCat2Debit as object	'// debit amount
dim oCat2Credit as object	'// credit amount
dim oCat2Acct as object		'// COA account number
dim oCat2Ref as object		'// reference info

'// GL field data extraction objects (.Text or .Value) and strings
dim sStyle as String		'// Text.Style string

'// Date field variables for both GL objects.
dim sDate as String
dim lDateFormat as long		'// date format mm/dd/yy
dim sMonth as String

'//	CellRange content variables for GL sheet.
dim lGLStartRow as long
dim lGLEndCol as long
dim lGLCurrRow as long		'// current row in GL

'// more local vars.
dim iStatus as integer		'// general status return var
dim sAcctCat2 as String		'// account category of 2nd line of transaction
dim oCat2Sheet as Object	'// account sheet object of 2nd category in transaction
dim lCat2InsRow as long		'// insert index for 2nd cateory in transaction

'// row processing variables; use to process selected area
dim i as long				'// loop counter
dim li as Long				'// long loop counter
dim sBadCOA2 as String		'// bad acct # in 2nd line
dim sCOA2Msg	as String	'// bad 2nd line message

'// other local vars.
dim iBrkPt as integer		'// easy breakpoint reference
dim lRetValue as long		'// returned value
dim l1stRow as Long			'// 1st row defining "split"
dim sSplitAcct as String	'// account splitting into
dim lLastRow as Long		'// last line of "split" block
dim sAcct2 as String		'// COA target account
dim lTrans1stRow as Long	'// first row after "split"
dim lRowCount as Long		'// transaction row count

'// error handling setup------------------------------------------------
const sMyName="PlaceSplitTrans"
dim sErrCode as String	'// error code string for user code
Dim sErrMsg As String	'// error message for LogError
dim lLogRow as long		'// cell row working on
dim lLogCol as Long		'// cell column working on
dim iLogSheetIx as Integer	'// sheet index module working on
dim oLogRange as new com.sun.star.table.CellRangeAddress

	'// code.

'// error handling initialization---------------------------------------
	sErrCode = ""
	iLogSheetIx = poGLSheet.RangeAddress.Sheet
	lLogCol = poGLSheet.RangeAddress.StartColumn
	lLogRow = poGLSheet.RangeAddress.StartRow
	oLogRange.Sheet = iLogSheetIx
	oLogRange.StartColumn = lLogCol
	oLogRange.StartRow = lLogRow
	oLogRange.EndRow = lLogRow
	oLogRange.EndColumn = lLogCol
	Call ErrLogSetup(oLogRange, sMyName)
	ErrLogSetRecording(true)		'// turn on log sheet
'//--------------end error handling setup-------------------------------

	lRetValue = ERRNOTSPLIT	'// set bad return
	Doc = ThisComponent				'// set up for sheet access within document
	oSheets = Doc.getSheets()			'// get sheet collection

	lGLStartRow = plSplit1stRow	'// one of these is superfluous?
	l1stRow = plSplit1stRow		'// set local ptr to 1st "split" row
	lTrans1stRow = l1stRow + 1	'// first actual transaction data row

	'// CheckSplitTrans validates and returns last "split" row
	lRetValue = CheckSplitTrans(poGlSheet, l1stRow)
	
	if lRetValue < 0 then		'// if invalid split, get out
		GoTo BailOut			'// CheckSplitTrans handled error
	else
		lLastRow = lRetValue	'// set last row of split block
	endif

	'// l1stRow = first row of "split"
	'// lLastRow = last row of "split"
	'// get "total" account COA and preserve as split "from" account
	lGLCurrRow = l1stRow + 1			'// move back to total line to split lines
	oGLCellAcct = poGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
ErrLogSetCellInfo(COLACCT, lGLCurrRow)
	sSplitAcct = oGLCellAcct.Text.String			'// set total COA splitting

'// set up Range object for SplitOutCOAs and PlaceSplitTotal calls.
dim oGLRange As new com.sun.star.table.CellRangeAddress

'	if true then			'// jump around new code 1st pass
'		GoTo SplitTrans
'	endif
	
'// set RangeAddress object pointing to the "split" transaction
'// in the ledger sheet
'//	iStatus = SplitOutCOAs(oGLRange, sSplitAcct)
'//		oGLRange specifies sheet, and row indexes of the
'//         entire split transaction
'//		sSplitAcct = COA of Total account for xrefs
	oGLRange.Sheet = poGLSheet.RangeAddress.Sheet
	oGLRange.StartColumn = COLDATE
	oGLRange.EndColumn = COLREF
	oGLRange.StartRow = l1stRow
	oGLRange.EndRow = lLastRow
	lRetValue = SplitOutCOAs(oGLRange, sSplitAcct)

	if lRetValue < 0 then
		lCat2InsRow = lRetValue
		GoTo ErrHandler
	endif
GoTo SplitTotal
	
SplitTrans:

SplitTotal:
iBrkPt=1

'//	setup code for PlaceSplitTotal
'// oGLRange = RangeAddress information from poGLSheet
'// sSplitAcct = COA for transaction total
		oGLRange.Sheet = poGLSheet.RangeAddress.Sheet
		oGLRange.StartColumn = COLDATE
		oGLRange.EndColumn = COLREF
		oGLRange.StartRow = l1stRow
		oGLRange.EndRow = lLastRow
		iStatus = PlaceSplitTotal(oGLRange, sSplitAcct)

		if iStatus < 0 then
			'// iStatus should be set to negative insert row
			lCat2InsRow = iStatus
			GoTo ErrHandler
		endif	'// end error splitting total conditional


'//	change Ref field background to LTGREEN for all split
'//	rows between "split"s in source transaction
	lGLCurrRow = l1stRow + 1			'// start at Total row
	lRowCount = lLastRow - l1stRow -1	'// row count of lines between "split"
	for li = 0 to lRowCount-1			'// minus 1, since indexed on 0
		oGLCellRef = poGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		oGLCellRef.CellBackColor = LTGREEN
		lGLCurrRow = lGLCurrRow + 1		'// advance to next transaction row
	next li			'// next transaction line
	
	lRetValue = lLastRow		'// ensure return last "split" row	
	GoTo ExitNormal		'// jump around error handler
			

ErrHandler:
	'// issue appropriate message and don't mark rows as processed
	sBadCOA2 = sAcct2
	sCOA2Msg = ""			'// clear line 2 message
	Select Case lCat2InsRow
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
		
	Case ERRDATENOTSAME	'// Date field doesn't match others			'// mod060520
const csDATENOTSAME="ERRDATESNOTSAME"
		sErrCode = csDATENOTSAME
		sErrMsg = "Split - one or more dates don't match"
		
	Case ERRSPLITDESC	'// "split" descriptions don't match		'// mod060520
const csSPLITDESC="ERRSPLITDESC"
		sErrCode = csSPLITDESC
		sErrMsg = "Split - 1st and last split descriptions don't match"
		
	Case else
		sErrCode = csERRUNKNOWN
		sErrMsg = "Account "+sBadCOA2+"Undocumented error"
	end Select

	Call LogError(sErrCode, sErrMsg)			
	GoTo BailOut

	'// end up here if insert row < 0
ExitNormal:	
BailOut:
	ErrLogDisable()
	PlaceSplitTrans = lRetValue		'// set return value

iBrkPt = 1
	
end function		'// end PlaceSplitTrans	7/8/20
'/**/
