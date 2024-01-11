'// STTotalAcctLostFocus.bas
'//---------------------------------------------------------------
'// STTotalAcctLostFocus - Handle TotalAcctFld LostFocus event.
'//		6/27/20.	wmk.	11:15
'//---------------------------------------------------------------

public sub STTotalAcctLostFocus()

'//	Usage.	macro call or
'//			call STTotalAcctLostFocus()
'//
'// Entry.	TotalAcctFld lost focus
'//			gbSTSplitCredits = true if Debit is parent account
'//							   false if Credit is parent acccount
'//			puoSTDialog = ST dialog object
'//
'//	Exit.	if gbSTSplitCredits = true; gsETAcct1 = user input from field
'//								false; gsETAcct2 = user input from field
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/27/20.	wmk.	original code


'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status
dim oDlgCtrl		As Object			'// dialog control

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "TotalAcctFld"	
	oDlgCtrl = puoSTDialog.getControl(gsSTObjFocus)
	if gbSTSplitCredits then
		gsETAcct1 = oDlgCtrl.Text
	else
		gsETAcct2 = oDlgCtrl.Text
	endif
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STTotalAcctLostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STTotalAcctLostFocus	6/27/20
'/**/
