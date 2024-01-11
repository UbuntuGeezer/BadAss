'// STAmt2FocusListener.bas
'//---------------------------------------------------------------
'// STAmt2FocusListener - Handle AmtFldn GotFocus event.
'//		6/20/20.	wmk.	20:45
'//---------------------------------------------------------------

public sub STAmt2FocusListener()

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
const MY_NAME="AmtFld2"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt2FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt2FocusListener	6/20/20
'/**/
