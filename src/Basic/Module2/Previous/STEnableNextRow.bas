'// STEnableNextRow.bas
'//---------------------------------------------------------------
'// STEnableNextRow - Enable next ST dialog input row.
'//		6/27/20.	wmk.	05:00
'//---------------------------------------------------------------

public function STEnableNextRow(piCurrRow) As Integer

'//	Usage.	iVar = STEnableNextRow( iCurrRow )
'//
'//		iCurrRow = current input row index {0..9}
'//
'// Entry.	puoSTDialog = ST dialog object
'//			"AmtFld.n" or "AcctFld.n" just lost focus with enough to move on
'//			field "AmtFld.n-1" will be Enabled if not already when .9 <=9
'//
'//	Exit.	iVar = 0 - normal return
'//				 <> 0	- error in enabling next row
'//			expanded public input string arrays if next row beyond last
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'//	6/20/20.	wmk.	redimension input arrays to accept input
'// 6/27/20.	wmk.	open if sequence corrected; eliminated 2nd end
'//						function as fix
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	ST input field controls base names.
const	AMT_BASENAME="AmtFld"
const	ACCT_BASENAME="AcctFld"
const	DESC_BASENAME="DescFld"
const	REF_BASENAME="RefFld"

'//	local variables.
dim iRetValue 	As Integer
dim iNextIx		As Integer	'// next row index
dim sRowIx		As String	'// next index as String
dim oDlgControl	As Object	'// ST dialog object
dim sFldName	As String	'// field name

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	if piCurrRow >= 9 then
		iRetValue = 0
		GoTo NormalExit
	endif	'// end at end of input conditional

	'// check next row already enabled
	iNextIx = piCurrRow + 1
	if giSTLastIx >= iNextIx then
		iRetValue = 0
		GoTo NormalExit
	else
		giSTLastIx = iNextIx	'// advance last active row index
	endif	'// end next row already enabled conditional

	if iNextIx > UBound(gdSTAmts) then
		redim preserve gdSTAmts(iNextIx)
		redim preserve gsSTDescs(iNextIx)
		redim preserve gsSTCOAs(iNextIx)
		redim preserve gsSTRefs(iNextIx)
	endif
	
	iNextIx = iNextIx + 1				'// advance index to get row suffix
	sRowIx = Trim(Str(iNextIx))			'// get row string
	
	'// Enable controls in next row
	sFldName = DESC_BASENAME + sRowIx	'// description input fld
	oDlgControl = puoSTDialog.getControl(sFldName)
	oDlgControl.Model.Enabled = true
	
	sFldName = AMT_BASENAME + sRowIx	'// amount input fld
	oDlgControl = puoSTDialog.getControl(sFldName)
	oDlgControl.Model.Enabled = true

	sFldName = ACCT_BASENAME + sRowIx	'// COA account input fld
	oDlgControl = puoSTDialog.getControl(sFldName)
	oDlgControl.Model.Enabled = true

	sFldName = REF_BASENAME + sRowIx	'// reference input fld
	oDlgControl = puoSTDialog.getControl(sFldName)
	oDlgControl.Model.Enabled = true

	iRetValue = 0
	
NormalExit:
	STEnableNextRow = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STEnableNextRow - unprocessed error.")
	GoTo NormalExit
	
end function 	'// end STEnableNextRow		6/20/20
'/**/
