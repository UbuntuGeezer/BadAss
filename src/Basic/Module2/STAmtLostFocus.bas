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
'//	6/20/20.	wmk.	original code; cloned from STAmtLostFocus
'//	6/21/20.	wmk.	check for total reached before activating next row
'// 6/21/20.	wmk.	change to accept psFldName parameter; to be called by
'//						all STAmt.LostFocus events; changed totals check and
'//						activate <Done> conditional
'//	6/22/20.	wmk.	change name to begin with "f" so can be distinguished
'//						by sub with same name for testing

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
'// STAmtLostFocus.bas
'//---------------------------------------------------------------
'// STAmt4LostFocus - Handle AmtFld4 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt4LostFocus()

'//	Usage.	macro call or
'//			call STAmt4LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld4"

'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld4"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt4LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt4LostFocus	6/23/20
'/**/
'// STAmtLostFocus.bas
'//---------------------------------------------------------------
'// STAmtLostFocus - Handle all AmtFld LostFocus events.
'//		6/22/20.	wmk.	08:00
'//---------------------------------------------------------------

public sub STAmtLostFocus()

'//	Usage.	call STAmtLostFocus()
'//			or macro call
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//			gdSTAmts() = updated with newly entered amount
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gdSTAccumTot, gdSTTotalAmt updated with newly entered amount
'//			<Done> button activated if dialog entries complete
'//			next row fields activated if row complete
'//
'// Calls.	STFld6ToIndex, CheckRowComplete, CheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; cloned from STAmtLostFocus
'//	6/21/20.	wmk.	check for total reached before activating next row
'// 6/21/20.	wmk.	change to accept psFldName parameter; to be called by
'//						all STAmt.LostFocus events; changed totals check and
'//						activate <Done> conditional
'//	6/22/20.	wmk.	change name to begin with "f" so can be distinguished
'//						by sub with same name for testing

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
dim oHasFocus		As Object

	'// code.
	iRetValue = -1						'// set bad return
	ON ERROR GOTO ErrorHandler
'msgBox("In sub STAmtLostFocus entry values.." + CHR(13)+CHR(10) _
'	+ "gdSTAccumTot = " + gdSTAccumTot + CHR(13)+CHR(10) _
'	+ "gdSTTotalAmt = " + gdSTTotalAmt + CHR(13)+CHR(10) _
'	+ "gsSTObjFocus = '" + gsSTObjFocus + "'" )
	
	oHasFocus = puoSTDialog.getControl("HasFocus")
	sFldName = oHasFocus.Text	
'	sFldName = gsSTObjFocus				'// use local copy of field name
'msgBox("sub STAmtLostFocus - sFldName = '" + sFldName + "'")
	if Len(sFldName) < 7 then
		GoTo ErrorHandler
	endif	'// error - field name not set up for entry
	
	bDlgComplete = false
'	iRowIx = STFld6ToIndex(sFldName)
'	iStatus = STUpdateAccumTot(sFldName)
	iRowIx = STFld6ToIndex(gsSTObjFocus)
	iStatus = STUpdateAccumTot(gsSTObjFocus)
'msgBox("In sub STAmtLostFocus values after update.." + CHR(13)+CHR(10) _
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
'	fSTAmtLostFocus = iRetValue
	exit sub
	
ErrorHandler:
	msgBox("In sub STAmtLostFocus - unprocessed error")
	iRetValue = iStatus
	GoTo NormalExit
	
end sub		'// end STAmtLostFocus	6/22/20
'/**/
