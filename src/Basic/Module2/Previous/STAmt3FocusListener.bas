'// STAmt3FocusListener.bas
'//==============================---------------------------------
'// STAmt3FocusListener - Handle AmtFldn GotFocus event.
'//		6/23/20.	wmk.	08:30
'//---------------------------------------------------------------

public sub STAmt3FocusListener()

'//	Usage.	macro call or
'//			call STAmt2FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAmt1FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld3"

'//	local variables.
dim iStatus		As Integer			'// general status
dim oHasFocus	As Object
	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	oHasFocus = puoSTDialog.getControl("HasFocus")
	oHasFocus.Text = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
'// Note: code will loop forever in any GotFocus event if a msgBox
'// is executed, since the message box takes focus away from the
'// dialog implementing the listener, then refocuses on it...
'	msgBox("In STAmt3FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt3FocusListener	6/23/20
'/**/
