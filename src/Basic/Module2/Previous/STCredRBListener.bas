'// STCredRBListener.bas
'//---------------------------------------------------------------
'// STCredRBListener - ST [Credit] radio button event handler.
'//		6/26/20.	wmk.	12:00
'//---------------------------------------------------------------

public sub STCredRBListener()

'//	Usage.	macro call or
'//			call STCredRBListener()
'//
'// Entry.	User clicked [Credit] radio button control
'//			gbSTSplitCredits = current split credits state
'//			gsETAcct1 = Debit account	
'//			gsETAcct2 = Credit account
'//
'//	Exit.	if clicked ON
'//			gbSTSplitCredits set false
'//			gbETDebitIsTotal set false
'//			if OFF, gbSTSplitCredits set true
'//			gbETDebitIsTotal set true
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.		wmk.	original code
'//`6/26/20.	wmk.	gsSTAcct1 or gsSTAcct2 used to set "TotalAcctFld"
'//						in ST dialog to display account being split
'//
'//	Notes.

'//	constants.

'//	local variables.
dim oSTCredRB		As Object		'// {Credit} radio button object
dim bState			As Boolean		'// .State of button
dim oSTDlgCtrl		As Object		'// control from "TotalAcctFld" in ST

	'// code.
	ON ERROR GOTO ErrorHandler
	oSTCredRB = puoSTDialog.getControl("CreditTotRB")
	oSTDlgCtrl = puoSTDialog.getControl("TotalAcctFld")
	bState = oSTCredRB.State
	if bState then	'// Credit button selected
		gbSTSplitCredits = false
		gbETDebitIsTotal = false
		oSTDlgCtrl.Text = gsETAcct2
	else
		gbSTSplitCredits = true
		gbETDebitIsTotal = true
		oSTDlgCtrl.Text = gsETAcct1
	endif	'// end Credit Total button State conditional

NormalExit:
	exit sub

ErrorHandler:
	msgBox("STCredRBListener - unprocessed error.")
	GoTo NormalExit

end sub		'// end STCredRBListener	6/26/20
'/**/
