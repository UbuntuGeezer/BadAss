'// fSTAmtLostFocus.bas
'//---------------------------------------------------------------
'// fSTAmtLostFocus - Handle all AmtFld LostFocus events.
'//		6/22/20.	wmk.	08:00
'//---------------------------------------------------------------

public function fSTAmtLostFocus(psFldName As String) As Integer

'//	Usage.	iVar = fSTAmtLostFocus(sFldName)
'//
'//	Parameters.	sFldName = name of dialog field losing focus
'//				(e.g. "AmtFld1")
'// Entry.	AmtFld1..AmtFld10 lost focus
'//			gdSTAmts() = updated with newly entered amount
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gsSTAccumTot, gsSTTotalAmt updated with newly entered amount
'//			<Done> button activated if dialog entries complete
'//			next row fields activated if row complete
'//
'// Calls.	STFld6ToIndex, STUpdateAccumTot, CheckRowComplete, CheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; cloned from STAmtLostFocus.
'//	6/21/20.	wmk.	check for total reached before activating next row
'// 6/21/20.	wmk.	change to accept psFldName parameter; to be called by
'// 6/21/20.						1all STAmt.LostFocus events; changed totals check and
'// 6/21/20.						2activate <Done> conditional.
'//	6/22/20.	wmk.	change name to begin with "f" so can be distinguished
'//	6/22/20.						by sub with same name for testing.

'//	constants.

'//	local variables.
dim iRetValue 		As Integer			'// returned status
dim iStatus			As Integer			'// general status
dim iRowIx			As Integer			'// index of present row
dim oDlgControl		As Object			'// field ptr
dim bRowComplete	As Boolean			'// row complete flag
dim bDlgComplete	As Boolean			'// dialog complete flag
dim bTotalMet		As Boolean			'// split total met flag
dim sFldName 		As String			'// local copy of field name

	'// code.
	iRetValue = -1						'// set bad return
	ON ERROR GOTO ErrorHandler
'msgBox("In fSTAmtLostFocus entry values.." + CHR(13)+CHR(10) _
'	+ "gdSTAccumTot = " + gdSTAccumTot + CHR(13)+CHR(10) _
'	+ "gdSTTotalAmt = " + gdSTTotalAmt + CHR(13)+CHR(10) _
'	+ "gsSTObjFocus = '" + gsSTObjFocus + "'" )
	
	sFldName = psFldName				'// use local copy of field name
'msgBox("fSTAmtLostFocus - sFldName = '" + sFldName + "'")
	if Len(sFldName) < 7 then
		GoTo ErrorHandler
	endif	'// error - field name not set up for entry
	
	bDlgComplete = false
	iRowIx = STFld6ToIndex(sFldName)
	iStatus = STUpdateAccumTot(sFldName)
'msgBox("In fSTAmtLostFocus values after update.." + CHR(13)+CHR(10) _
'	+ "gdSTAccumTot = " + gdSTAccumTot + CHR(13)+CHR(10) _
'	+ "gdSTTotalAmt = " + gdSTTotalAmt )
	
	if iStatus < 0 then
		GoTo ErrorHandler
	endif	'// end error in updating accum total conditional

	bTotalMet = (iStatus = 1)
	bRowComplete = STRowComplete(iRowIx)
	if bRowComplete then
		bDlgComplete = STCheckComplete()
		oDlgControl = puoSTDialog.getControl("UnsplitTot")
'		if gdSTAccumTot < gdSTTotalAmt then		
		if oDlgControl.Value > 0 then		
			'// check for next row Enabled; if not, enable it
			iStatus = STEnableNextRow(iRowIx)
			if iStatus < 0 then
				GoTo ErrorHandler
			endif
		else
			bTotalMet = true
		endif '// end total not met yet conditional
		
	endif	'// end row complete conditional
	
	'// activate/deactivate <Done> control if dialog complete
	oDlgControl = puoSTDialog.getControl("DoneBtn")
'	bDlgComplete = (bDlgComplete AND bTotalMet)
	bDlgComplete = (bDlgComplete OR bTotalMet)
	if bDlgComplete then
		oDlgControl.Model.Enabled = true
	else
		oDlgControl.Model.Enabled = false
	endif
	iRetValue = 0
	
NormalExit:
	fSTAmtLostFocus = iRetValue
	exit function
	
ErrorHandler:
	msgBox("In fSTAmtLostFocus - unprocessed error")
	iRetValue = iStatus
	GoTo NormalExit
	
end function		'// end fSTAmtLostFocus	6/22/20
'/**/
