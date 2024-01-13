'// CheckSplitTrans.bas
'//---------------------------------------------------------------
'// CheckSplitTrans - Check split transaction for validity.
'//		7/4/20. 	wmk.	10:45
'/---------------------------------------------------------------

function CheckSplitTrans(poGLSheet as Object, plSplit1stRow as Long )_
		as Long

'//	Usage.	lSplitEndRow = CheckSplitTrans(oGLSheet, lSplit1stRow)
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
'//	Exit.	lSplitEndRow = last row of split, "split" in COLSPLIT
'//						< 0 - error in split transaction
'//						-1 - (ERRNOTSPLIT) first row not "split" in COLSPLIT
'//						-2 - (ERRBADTALLY) split values don't add to total
'//						-3 - (ERRROWSEXCEEDED) too many/few rows before second "split"
'//						-4 - (ERRBADVALS) bad Debit/Credit values
'//						-5 - (ERRBADDATE) Date field too short
'//						-6 - (ERRDATENOTSAME) Date field doesn't match others
'//						-7 - (ERRSPLITDESC) "split" descriptions don't match
'//
'// Calls.	ErrLogGetModule, ErrLogGetCellInfo, ErrLogGetSheet,
'//			ErrLogSetSheet, ErrLogSetModule, ErrLogSetCellInfo,
'//			LogError
'//
'//	Modification history.
'//	---------------------
'//	5/22/20		wmk.	original code
'//	5/23/20		wmk.	bug fix where not tallying properly; fix bug where
'//						empty Date field within loop range checking rows
'//	5/24/20.	wmk.	documentation added on .Type cell property; modify
'//						bad row Date cell with own content to update color
'//	5/27/20.	wmk.	include new error handling capability
'//	5/28/20.	wmk.	error handling improvements and documentation
'//	6/5/20.		wmk.	added check to ensure that transaction desc
'//						matches in both "split" lines; added check to
'//						ensure that all dates are the same in all
'//						transaction lines; cell addresses updated for
'//						all error reporting conditions
'//	7/4/20.		wmk.	double check tally value mismatch with strings to
'//						correct issue where 9.62 did not match the sum
'//						of 4.81 and 4.81
'//
'//	Notes.
'// ensure current row has "split" in COLSPLIT field; return -1 if not
'// determine how many rows are in the "split", searching for next
'//  row with "split"; set some sort of maximum so don't search forever
'//  if maximum exceeded return -3
'// ensure that total values form split rows = total in 1st row
'//  if not, return -2
'// set up cat1 sheet to point to total's category/account sheet
'// loop processing split rows, setting up cat2 sheet to point to
'//   each category/account sheet
'// on successful exit, point to last row with "split" in COLSPLIT field
'// Programming notes. oCell.Type returns the cell type {EMPTY,VALUE,TEXT
'// FORMULA). The Value, String, and Formula properties are used to
'// set the values of a cell.

'//	constants.
'// local contants. (common with PlaceSplitTrans)
const MAXSPLITROWS=10	'// maximum rows in "split" transaction
const MINROWLIMIT=1		'// 0-based index minimum total row index
const MINDATELEN=6		'// minimum Date string length
const ERRNOTSPLIT=-1	'// first row not "split" in COLSPLIT
const ERRBADTALLY=-2	'// split values don't add to total
const ERRROWSEXCEEDED=-3	'// too many/few rows before second "split"
const ERRBADVALS=-4		'// bad Debit/Credit values
const ERRBADDATE=-5		'// Date field too short
const ERRDATENOTSAME=-6	'// Date field doesn't match others			'// mod060520
const ERRSPLITDESC=-7	'// "split" descriptions don't match		'// mod060520
const COLSPLIT=5		'// "split" column
const csSplit="split"	'// split row identifier string


'//	local variables.
dim oCellSplit	as Object	'// cell from COLSPLIT column
dim oGLSheet	as Object	'// local ptr to poGLSheet
dim oCellCredit as Object		'// credit field
dim oCellDebit 	as Object		'// debit field		

'// other local vars.
dim iBrkPt as integer		'// easy breakpoint reference
dim lRetValue as long		'// returned value
dim bBadSplit as Boolean	'// bad split flag
dim sSplit As String		'// COLSPLIT .Text.String
dim lSplitCount as long		'// split row count
dim i as integer			'// loop counter
dim lGLCurrRow as Long			'// GL current row
dim bFoundSplit as boolean	'// found "split" 2nd time
dim lTrans1stRow as long	'// transaction 1st row
dim dTransTotal as Double		'// split entered total
dim dSplitTally as Double		'// split tallied total
dim dDebitVal as Double		'// Debit value
dim dCreditVal as Double	'// Credit value
dim bDebitSplits as boolean		'// first line is Debit flag
dim sDebit as String		'// debit value string
dim sCredit as String		'// credit value string
dim lGLStartRow as Long		'// start row
dim dAddAmt	as Double		'// amount to add to tally
dim sSourceAcct as String	'// source account being split
dim oCellDate as Object		'// cell Date from transaction row
dim sDate As String			'// cell Date string from transaction	'// mod060520
dim sDateSplit As String	'// Date field from first "split" line	'// mod060520
dim oCellTrans As Object	'// cell Transaction desc from row		'// mod060520
dim sTrans As String		'// transaction from end "split" line	'// mod060520
dim sTransSplit As String	'// Transaction from 1st "split" line	'// mod060520
dim bDateMatched As Boolean	'// Date field(s) all match flag		'// mod060520
dim sErrMsg as String		'// error message
dim sErrCode As String		'// error code string

	'// code.

	lRetValue = ERRNOTSPLIT	'// set error return

'//*---------end error handling setup---------------------------
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
'//*
	'// preserve entry error settings.
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	sErrCurrentMod = ErrLogGetModule()
'//*	
	'// set local error settings.
	ErrLogSetModule("CheckSplitTrans")
	ErrLogSetSheet(poGLSheet.RangeAddress.Sheet)
	ErrLogSetCellInfo(COLSPLIT, plSplit1stRow)
'//*----------end error handling setup---------------------------

	'// ensure current row has "split" in COLSPLIT field;
	'//	 return ERRNOTSPLIT if not
	lGLCurrRow = plSplit1stRow
	lGLStartRow = plSplit1stRow	
	lTrans1stRow = lGLCurrRow + 1	'// transaction first data row
'	oGLSheet = poGLSheet
	oCellSplit = poGLSheet.getCellByPosition(COLSPLIT, lGLCurrRow)
	sSplit = oCellSplit.Text.String
	bBadSplit = (StrComp(sSplit, csSplit) <> 0)
	lRetValue = ERRNOTSPLIT			'// set initial error	
iBrkPt=1

	if bBadSplit then
		lRetValue = ERRNOTSPLIT
		GoTo Bailout
	endif	'// end not "split"

	'// preserve Date field from "split" 1st row					'// mod060520
	oCellDate = poGLSheet.getCellByPosition(COLDATE, lGLCurrRow)	'// mod060520
	sDateSplit = oCellDate.String									'// mod060520
	'// date must be at least #/#/## long							'// mod060520
	if Len(sDateSplit) < MINDATELEN then							'// mod060520
ErrLogSetCellInfo(COLDATE, lGLCurrRow)									'// mod060520
		lRetValue = ERRBADDATE										'// mod060520
		GoTo BailOut												'// mod060520
	endif	'// end date too short conditional						'// mod060520
	
	bDateMatched = true		'// for AND in loop to test failure		'// mod060520

	'// preserve Transaction field from "split" 1st row				'// mod060520
	oCellTrans = poGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)	'// mod060520
	sTransSplit = oCellTrans.String	'// Transaction desc			'// mod060520

	'// check Date in 1st transaction row
	oCellDate = poGLSheet.getCellByPosition(COLDATE, lTrans1stRow)	'//	mod060520
	bDateMatched = bDateMatched AND _
					  (StrComp(oCellDate.String, sDateSplit)=0)	'// mod060520
	if NOT bDateMatched then										'// mod060520
ErrLogSetCellInfo(COLDATE, lTrans1stRow)								'// mod060520
		lRetValue = ERRDATENOTSAME									'// mod060520
		GoTo BailOut												'// mod060520
	endif
	
	'// load debit, and credit, cells from 1st transaction line
ErrLogSetCellInfo(COLDEBIT, lTrans1stRow)
	oCellDebit = poGLSheet.getCellByPosition( COLDEBIT, lTrans1stRow)
	oCellCredit = poGLSheet.getCellByPosition( COLCREDIT, lTrans1stRow)
		
	'// declare tally vars as double dTransTotal, dSplitTally
	sDebit = oCellDebit.Text.String
	sCredit = oCellCredit.Text.String

	'// verify that one or the other is empty and set flag
	bDebitSplits = (Len(sCredit)=0) 		'// debit line is split
	bBadSplit = (Len(sDebit)=0) AND (Len(sCredit)=0)	'// one MUST be nonempty
		
	'// check for entries in both Debit and Credit columns
	if bDebitSplits then
		bBadSplit = bBadSplit OR Len(sCredit)>0		
	else	
		bBadSplit = bBadSplit OR Len(sDebit)>0
	endif	'// end debit splits conditional
		
	if bBadSplit then
		lRetValue = ERRBADVALS
		GoTo BailOut
	endif

	if bDebitSplits then
		dDebitVal = oCellDebit.getValue()		'// look at debit and Credit
		dTransTotal = dDebitVal
	else
		dCreditVal = oCellCredit.getValue()
		dTransTotal = dCreditVal
	endif
		
	dSplitTally = 0.			'// clear split tally

	'// loop on split transaction rows, searching for 2nd "split"
	'// and recording amount from 1st row as total,
	'// then tallying amounts from subsequent rows and
	'// comparing tally with 1st row value
	'// credit amount, and COA account number cells
	'//	when encounter "split", set lRetVal to row index
	
'//	loop on transaction rows looking for "split"

	lGLCurrRow = lTrans1stRow + 1	'// skip over 1st row - already processed
	'// loop beyond MAXSPLITROWS to determine if "split" out of bounds
	for i = 0 to MAXSPLITROWS
ErrLogSetCellInfo(COLSPLIT, lGLCurrRow)

		oCellSplit = poGLSheet.getCellByPosition(COLSPLIT, lGLCurrRow)
		sSplit = oCellSplit.Text.String
		bFoundSplit = (StrComp(sSplit, csSplit) = 0)
		
		if bFoundSplit then

			'// check that last date matches the rest
			oCellDate = poGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
			sDate = oCellDate.String
			bDateMatched = bDateMatched AND _
					 (StrComp(oCellDate.String, sDateSplit)=0)		'// mod060520
			if NOT bDateMatched then								'// mod060520
ErrLogSetCellInfo(COLDATE, lGLCurrRow)									'// mod060520
				lRetValue = ERRDATENOTSAME							'// mod060520
				GoTo BailOut										'// mod060520
			endif

			'// check that the last "split" description matches		'// mod060520
			oCellTrans = poGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
			sTrans = oCellTrans.String								'// mod060520
			if StrComp(sTrans, sTransSplit) <> 0 then				'// mod060520
ErrLogSetCellInfo(COLTRANS, lGLCurrRow)								'// mod060520
				lRetValue = ERRSPLITDESC							'// mod060520
				GoTo BailOut										'// mod060520
			endif
			
			'// check not enough rows to be a valid split			
			if i < MINROWLIMIT then
ErrLogSetCellInfo(COLDATE, lTrans1stRow)								'// mod060520
				lRetValue = ERRROWSEXCEEDED
				GoTo BailOut
			endif
if true then
GoTo ExitFor
endif

			'// check tally against total
			
			if dSplitTally <> dTransTotal then
ErrLogSetCellInfo(COLREF, lGLCurrRow)								'// mod060520
				lRetValue = ERRBADVALS
				GoTo BailOut
			endif
ExitFor:			
			lRetValue = lGLCurrRow
			exit For		'// leave For loop

			
		else	'// not found	
			'// not found yet, check max exceeded			
			if i >= MAXSPLITROWS then
				lRetValue = ERRROWSEXCEEDED
				GoTo BailOut
			endif
			
			'// if empty Date field, is error
			oCellDate = poGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
		ErrLogSetCellInfo(COLDATE, lGLCurrRow)

			sDate = oCellDate.String		'// get date			'// mod060520
			if Len(sDate) < MINDATELEN then
				lRetValue = ERRBADDATE
				GoTo BailOut
			endif

			bDateMatched = bDateMatched AND _
					 (StrComp(sDate, sDateSplit)=0)				'// mod060520
			if NOT bDateMatched then								'// mod060520
				lRetValue = ERRBADDATE								'// mod060520
				GoTo BailOut										'// mod060520
			endif													'// mod060520
	

			'// check Debit and Credit columns; one MUST have a nonempty string
						
			'// increment tally value from Debit or Credit
			if bDebitSplits then	'// Debit splits - tally Credits
	ErrLogSetCellInfo(COLCREDIT, lGLCurrRow)
				oCellCredit = poGLSheet.getCellByPosition( COLCREDIT, lGLCurrRow)
				bBadSplit = (Len(oCellCredit.String) = 0)
				dAddAmt = oCellCredit.getValue()
			else	''// Credit splits; tally Debits
	ErrLogSetCellInfo(COLDEBIT, lGLCurrRow)
				oCellDebit = poGLSheet.getCellByPosition( COLDEBIT, lGLCurrRow)
				dAddAmt = oCellDebit.getValue()
				bBadSplit = (Len(oCellDebit.String) = 0)
			endif
			
			'// bad split if opposing column of transaction total is empty
			if bBadSplit then
				lRetValue = ERRBADVALS
				GoTo BailOut
			endif
						
			dSplitTally = dSplitTally +	dAddAmt		'// add value to tally
			
		endif	'// end split found conditional

		lRetValue = lGLCurrRow		'// point to current row
		lGLCurrRow = lGLCurrRow + 1	'// advance to next entry
		
	next i
	
	'// normal return if tally matches transaction total
	if dSplitTally = dTransTotal then
		GoTo ExitNormal		'// lRetVal = lGLCurrRow
	else		'// tally mismatch on values, but check strings
		if StrComp(Str(dSplitTally), Str(dTransTotal))=0 then
			GoTo ExitNormal
		endif
	endif 	'// end valid tally conditional
	
	lRetValue = ERRBADVALS
	
'//----------- Error handler--------------------	
BailOut:
	if lRetValue >= 0 then
		GoTo ExitNormal
	endif

	'// lGLCurrRow points to row in which error occurred
	oCellDate = poGLSheet.getCellByPosition( COLDEBIT, lGLCurrRow)
	oCellDate.String = oCellDate.String		'// modify cell with own content
	oCellDate.CellBackColor = YELLOW

	'// ErrLog has corrupted cell column, row; change to YELLOW
dim lCurrErrCol as Long
dim lCurrErrRow As Long
ErrLogGetCellInfo(lCurrErrCol, lCurrErrRow)
	oCellDate = poGLSheet.getCellByPosition(lCurrErrCol, lCurrErrRow)
	oCellDate.String = oCellDate.String
	oCellDate.CellBackColor = YELLOW
	
	Select Case lRetValue
		Case ERRNOTSPLIT
			sErrMsg = "Transaction not 'split'..."_
					+ CHR(13)+CHR(10)+"stopped."
			sErrCode = "ERRNOTSPLIT"
			
		Case ERRBADVALS
			sErrMsg = "Bad Debit/Credit field(s); check values"_
					+ CHR(13)+CHR(10)+"stopped."
			sErrCode = "ERRBADVALS"
					
		Case ERRBADTALLY
			sErrMsg = "Debit and Credit values do not balance..."_
					+ CHR(13)+CHR(10)+"stopped."
			sErrCode = "ERRBADTALLY"
			
		Case ERRROWSEXCEEDED
			sErrMsg = "Too few or too many rows in 'split'..."_
					+ CHR(13)+CHR(10)+"stopped."
			sErrCode = "ERRROWSEXCEEDED"

		Case ERRBADDATE
			sErrMsg = "Date field too short or empty..."_
					+ CHR(13)+CHR(10)+"stopped."			
			sErrCode = "ERRBADDATE"

		Case ERRDATENOTSAME	'// Date field doesn't match others			'// mod060520
const csDATENOTSAME="ERRDATESNOTSAME"
			sErrCode = csDATENOTSAME
			sErrMsg = "Split - one or more dates don't match"
		
		Case ERRSPLITDESC	'// "split" descriptions don't match		'// mod060520
const csSPLITDESC="ERRSPLITDESC"
			sErrCode = csSPLITDESC
			sErrMsg = "Split - 1st and last split descriptions don't match"

		Case else
			sErrMsg = "Split transaction funky; check it..."_
					+ CHR(13)+CHR(10)="stopped."
			sErrCode = "ERRUNKNOWN"

		End Select
		
	Call LogError(sErrCode, sErrMsg)
'	msgBox(sErrMsg)		'// issue error message
'XRay oCellSplit

ExitNormal:

'//*-----------------restore error handling for caller-----------
	'// restore entry error settings.
	ErrLogSetModule(sErrCurrentMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
'//*----------end error handling setup---------------------------
			

	CheckSplitTrans = lRetValue
	
end function 	'// end CheckSplitTrans	7/4/20
'/**/
