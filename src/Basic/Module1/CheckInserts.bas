'// CheckInserts.bas
'//------------------------------------------------------------------
'// CheckInserts - Check target COA sheet(s) for insertion available.
'//		7/3/20.	wmk.
'//------------------------------------------------------------------

public function CheckInserts(psMonth As String, psAcct1 As String,_
		psAcct2 As String, rpoCat1Range As Object, rpoCat2Range As Object) as Integer

'//	Usage.	iVal = CheckInserts(sMonth, sAcct1, sAcct2, oCat1Range, oCat2Range)
'//
'//		sMonth = Month name string (e.g. "January")
'//		sAcct1 = COA from 1st line of transaction
'//		sAcct2 = COA from 2nd line of transaction
'//		rpoCat1Range = (returned)
'//		rpoCat2Range = (returned)
'//
'// Entry.	Spreadsheet has category sheets for all possible COA #s
'//
'//	Exit.	rplIns1Ptr = row for insertion of line 1 transaction
'//			rplIns2Ptr = row for insertion of line 2 transaction
'//
'// Calls.	GetInsRow.
'//
'//	Modification history.
'//	---------------------
'//	6/1/20.		wmk.	original code; stub
'//	6/2/20.		wmk.	code transported from within PlaceTransM; error handling
'//						set up
'//	6/8/20.		wmk.	bug fix where sCOA1Msg and sCOA2Msg not declared	
'// 7/3/20.		wmk.	bug fix in error handling using lGLCurrRow, needs lErrRow
'//
'//	Notes. Passed paramters are COA #s; they will be used to access the
'// correct account catgory sheets, for the month of the transaction
'//

'//	constants.
'// insert error codes.
const ERRCOANOTFOUND=-1
const ERROUTOFROOM=-2
const ERRNOCOAMONTH=-3
const sERRCOANOTFOUND="ERRCOANOTFOUND"
const sERROUTOFROOM="ERROUTOFROOM"
const sERRNOCOAMONTH="ERRNOCOAMONTH"
const sERRUNKNOWN="ERRUNKNOWN"

'//----------in Module error handling setup-------------------------------
'// LogError setup snippet.
const csERRUNKNOWN="ERRUNKNOWN"
const sMyName="CheckInserts"
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

'// process control constants.
const ERRCONTINUE=-1	'// error, but continue with next transaction
const ERRSTOP=-2		'// error, stop processing transactions

'//	local variables.
dim iRetValue As Integer	'// function returned value
dim sMonth As String		'// local month

'//	Object setup variables.
dim oDoc As Object			'// ThisComponent
dim oSheets As Object		'// Doc.getSheets()

'// COA Sheet processing variables.
dim sAcct As String			'// first line COA#
dim sAcctB As String		'// second line COA#
dim sAcctCat1 as String 	'// accounting category of 1st line of transaction
dim sAcctCat2 as String		'// account category of 2nd line of transaction
dim bSameAcctCat as boolean	'// account categories the same flag		
dim oCat1Sheet As Object	'// COA sheet for 1st transaction line
dim oCat2Sheet As Object	'// COA sheet for 2nd transaction line
dim bSameCOA As Boolean		'// COAs match flag
dim lCat1InsRow As Long		'// COA sheet insertion row, line 1
dim lCat2InsRow As Long		'// COA sheet insertion row, line 2
dim sBadCOA As String		'// first COA bad string
dim sBadCOA2 As String		'// second COA bad string
dim iStatus As Integer		'// processing status flag
dim sCOA1Msg As String		'// COA1 error message string
dim sCOA2Msg As String		'// COA2 error message string

	'// code.
	
	iRetValue = 0
	iStatus = 0				'// clear general status
'	rplIns1Ptr = -1			'// set bad pointers for return
'	rplIns2Ptr = -1
	sMonth = psMonth		'// set local copy of passed month
		
'//*------------error handling code initialization---------------
	'// preserve entry error settings.
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	sErrCurrentMod = ErrLogGetModule()
'//*	
	'// set local error settings.
	ErrLogSetModule(sMyName)
	'// keep entry sheet/cell selctions, since errors will be in GL
'	ErrLogSetSheet(poTransRange.Sheet)
'	ErrLogSetCellInfo(COLDATE, poTransRange.StartRow)
	ErrLogSetRecording(true)		'// enable log entries
'//*----------end error handling code initialization-------------

	
	CheckInserts = iRetValue
'	if true then
'		Exit Function
'	endif
'//---------------end stub code--------------------	
'//-----------------CheckInserts code starts here----------------------------------------
'//
'//	Usage.	iVal = CheckInserts(psMonth, psAcct1, psAcct2, rplIns1Ptr, rplIns2Ptr)
'//
'//		psMonth = Month name string (e.g. "January")
'//		psAcct1 = COA from 1st line of transaction
'//		psAcct2 = COA from 2nd line of transaction
'//		rplIns1Ptr = (returned)
'//		rplIns2Ptr = (returned)
'//

		'// get appropriate sheet names from COA account numbers
		iStatus = 0			'// clear processing exceptions
		sAcct = psAcct1		'// copy COAs to local vars
		sAcctB = psAcct2
		sAcctCat1 = GetTransSheetName(sAcct)
		sAcctCat2 = GetTransSheetName(sAcctB)
		bSameAcctCat = (StrComp(sAcctCat1, sAcctCat2) = 0)
dim iBrkpt As Integer
iBrkpt = 1

		'// set up oCat1Sheet as sheet object for first half of transaction
		'//			oCat2Sheet as sheet object for second half of transaction
'static oCat1Sheet as Object	'// account sheet object of 1st category in transaction
'static oCat2Sheet as Object	'// account sheet object of 2nd category in transaction
'dim lCat1InsRow as long		'// insert index for 1st category in transaction
'dim lCat2InsRow as long		'// insert index for 2nd cateory in transaction

'XRay Doc
		oDoc = ThisComponent
		oSheets = oDoc.getSheets()

'XRay oSheets
		'// check if same sheet, and just copy instance						'// mod051920
		oCat1Sheet = oSheets.getByName(sAcctCat1)
		if bSameAcctCat then
			oCat2Sheet = oCat1Sheet
		else
			oCat2Sheet = oSheets.getByName(sAcctCat2)
		endif
		
		rpoCat1Range.Sheet = oCat1Sheet.RangeAddress.Sheet
		rpoCat2Range.Sheet = oCat2Sheet.RangeAddress.Sheet
		
'XRay oCat1Sheet		
iBrkpt=1

		'// check here if COA#s different; if not, will bail out later		'// mod051920
		bSameCOA = StrComp(sAcct, sAcctB) = 0								'// mod051920
		lCat1InsRow = GetInsRow(oCat1Sheet, sAcct, sMonth)					'// mod051920

		if bSameCOA then													'// mod051920
			'// Note: might get away with setting same insertion row, since will
			'// insert 2 lines... for now, skipped
			lCat2InsRow = lCat1InsRow										'// mod051920
		else																'// mod051920
			lCat2InsRow = GetInsRow(oCat2Sheet, sAcctB, sMonth)				'// mod051920
		endif	'// end same COA conditional								'// mod051920
				
		'// add code to verify have a row in each sheet before committing to change
'dim lBadRow1 	as long		'// 1st acct insert row return value
'dim lBadRow2	as lont		'// 2nd acct insert row return value

		if lCat1InsRow < 0 OR lCat2InsRow < 0 then
			'// issue appropriate message and don't mark rows as processed
			sBadCOA = ""
			sBadCOA2 = ""
			'// handle transaction line 1 error.
			if lCat1InsRow < 0 then
				sBadCOA = sAcct
				sCOA1Msg = ""			'// clear line 1 message
				ErrLogSetSheet(oCat1Sheet.RangeAddress.Sheet)
				ErrLogSetCellInfo(COLDATE, lErrRow)
				Select Case lCat1InsRow
'				Case -1
				Case ERRCOANOTFOUND
					sCOA1Msg = "Account "+sBadCOA+" not found"
					sErrName = sERRCOANOTFOUND
'				Case -2
				Case ERROUTOFROOM
					sCOA1Msg = "Account "+sBadCOA+" month "+sMonth+" not enough rows"_
								+" to insert"
					sErrName = sERROUTOFROOM
'				Case -3
				Case ERRNOCOAMONTH
					sCOA1Msg = "Account "+sBadCOA+" month "+sMonth+" not found"
					sErrName = sERRNOCOAMONTH
				Case else
					sCOA1Msg = "Account "+sBadCOA+"Undocumented error"
					sErrName = sERRUNKNOWN
				end Select
				
				sErrMsg = sCOA1Msg
				Call LogError(sErrName, sErrMsg)
				
			endif	'// end error in 1st line conditional
			
			if lCat2InsRow < 0 then
				sBadCOA2 = sAcctB
				sCOA2Msg = ""			'// clear line 2 message
				Select Case lCat2InsRow
'				Case -1
				Case ERRCOANOTFOUND
					sCOA2Msg = "Account "+sBadCOA2+" not found"
					sErrName = sERRCOANOTFOUND
'				Case -2
				Case ERROUTOFROOM
					sCOA2Msg = "Account "+sBadCOA2+" month "+sMonth+" not enough rows"_
								+" to insert"
					sErrName = sERROUTOFROOM
'				Case -3
				Case ERRNOCOAMONTH
					sCOA2Msg = "Account "+sBadCOA2+" month "+sMonth+" not found"
					sErrName = sERRNOCOAMONTH
	
					sCOA2Msg = "Account "+sBadCOA2+"Undocumented error"
					sErrName = sERRUNKNOWN
				end Select
				
				sErrMsg = sCOA2Msg
				Call LogError(sErrName, sErrMsg)	'// log error and continue
				
			endif		'// end bad line 2 insert row conditional
			
			iStatus = ERRCONTINUE			'// flag error to continue with next transaction
'			msgBox(sCOA1Msg+CHR(13)+CHR(10) + sCOA2Msg)
			'// advance to next transaction
			GoTo AdvanceTrans
			
 		endif	'// end problem with either COA insert conditional
 		
 		'// no problems, set returned insertion rows
' 		rplIns1Ptr = lCat1InsRow
'		rplIns2Ptr = lCat2InsRow
 		rpoCat1Range.StartRow = lCat1InsRow
 		rpoCat2Range.StartRow = lCat2InsRow
 		rpoCat1Range.EndRow = lCat1InsRow
 		rpoCat2Range.EndRow = lCat2InsRow
 		rpoCat1Range.StartColumn = COLDATE
 		rpoCat2Range.StartColumn = COLDATE
 		rpoCat1Range.EndColumn = COLREF
 		rpoCat2Range.EndColumn = COLREF

AdvanceTrans:
'//*-----------------restore error handling for caller-----------
	'// restore entry error settings.
	ErrLogSetModule(sErrCurrentMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
'//*----------end error handling setup---------------------------

	'// iStatus = 0, no error
	'//				ERRCONTINUE (-1) - error; continue with next transaction
	'//				ERRSTOP (-2) - error; stop processing transactions
	iRetValue = iStatus				'// set returned status
	CheckInserts = iRetValue
	
end function 	'// end CheckInserts	7/3/20
'/**/

