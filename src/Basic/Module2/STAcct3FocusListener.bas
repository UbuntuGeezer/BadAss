'// STAcct3FocusListener.bas
'//---------------------------------------------------------------
'// STAcct3FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct3FocusListener()

'//	Usage.	macro call or
'//			call STAcct2FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld3"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct3FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct3FocusListener	6/20/20
'/**/
