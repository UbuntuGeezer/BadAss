'// STAcctLostFocus.bas
'//---------------------------------------------------------------
'// STAcctLostFocus - Handle any AcctFld LostFocus event.
'//		7/7/20.	wmk.	12:30
'//---------------------------------------------------------------

public function STAcctLostFocus() As Integer

'//	Usage.	macro call or
'//			call STAcctLostFocus()
'//
'//	[Parameters.	sFldName = name of dialog object lost focus]
'//
'// Entry.	AcctFld1..AcctFld10 lost focus
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gsCOAs(0) = text from AcctFld1..10
'//
'// Calls.	STFld7ToIndex
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; cloned from STAcct1LostFocus
'//	6/21/20.	wmk.	check for total reached before activating next row
'//	7/7/20.		wmk.	bug fix where total amount reached, but <Done> button
'//						deactivated
'//
'//	Notes. gsSTObjFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into.


'//	constants.

'//	local variables.
dim iRetValue 		As Integer			'// returned status
dim iStatus			As Integer			'// general status
dim iRowIx			As Integer			'// index of present row
dim oDlgControl		As Object			'// field ptr
dim bRowComplete	As Boolean			'// row complete flag
dim bDlgComplete	As Boolean			'// dialog complete flag
dim sFldName 		As String			'// local copy of field name

	'// code.
	iRetValue = -1						'// set bad return
	ON ERROR GOTO ErrorHandler
	sFldName = gsSTObjFocus				'// use local copy of field name
'	sFldName = psFldName				'// use local copy of field name
	oDlgControl = puoSTDialog.getControl(sFldName)
	iRowIx = STFld7ToIndex(sFldName)
	gsSTCOAs(iRowIx) = oDlgControl.Text

	bRowComplete = STRowComplete(iRowIx)
	if bRowComplete then
		bDlgComplete = STCheckComplete()

		if gdSTAccumTot < gdSTTotalAmt then		
			'// check for next row Enabled; if not, enable it
			iStatus = STEnableNextRow(iRowIx)
			if iStatus < 0 then
				GoTo ErrorHandler
			endif
		endif '// end total not met yet conditional
		
	endif	'// end row complete conditional
	
	'// activate/deactivate <Done> control if dialog complete
	oDlgControl = puoSTDialog.getControl("DoneBtn")
	if bDlgComplete then
		oDlgControl.Model.Enabled = true
	else
		oDlgControl.Model.Enabled = false
	endif
	
	iRetValue = 0
	gsSTObjFocus = ""					'// clear object focus tracker

NormalExit:
	STAcctLostFocus = iRetValue
	exit function
	
ErrorHandler:
	msgBox("In STAcctLostFocus - unprocessed error")
	GoTo NormalExit
	
end function		'// end STAcctLostFocus	6/21/20
'/**/
