'// STDebtRBListener.bas
'//---------------------------------------------------------------
'// STDebtRBListener - ST [Debit] radio button event handler.
'//		7/6/20.	wmk.	19:00
'//---------------------------------------------------------------

public sub STDebtRBListener()

'//	Usage.	macro call or
'//			call STDebtRBListener()
'//
'// Entry.	User clicked [Debit] radio button control
'//			puoSTDialog = Split Control dialog object
'//			gbSTSplitCredits = current split credits stat
'//			gsETAcct1 = Debit account	
'//			gsETAcct2 = Credit account
'//
'//	Exit.	if clicked ON
'//				gbSTSplitCredits set true
'//				gbETDebitIsTotal set true
'//				STSplitAcctFld = gsSTAcct1
'//			if OFF, gbSTSplitCredits set false
'/				gbETDebitIsTotal set false
'//				STSplitAcctFld - gsSTAcct2
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'//`6/26/20.	wmk.	gsSTAcct1 or gsSTAcct2 used to set "TotalAcctFld"
'//						in ST dialog to display account being split
'//	7/6/20.		wmk.	bug fix; code was checking oSTCredRB.State instead
'//						of oSTDebtRB.State; puoSTDailog corrected; gsETAcct1
'//						gsETAcct2 references corrected
'//
'//	Notes. 

'//	constants.

'//	local variables.
dim oSTDebtRB		As Object		'// {Debit} radio button object
dim bState			As Boolean		'// .State of button
dim oSTDlgCtrl		As Object		'// control from "SplitAcctFld" in ST

	'// code.
	ON ERROR GOTO ErrorHandler
	oSTDebtRB = puoSTDialog.getControl("DebitTotRB")
	oSTDlgCtrl = puoSTDialog.getControl("TotalAcctFld")
	bState = oSTDebtRB.State
	if bState then	'// Debit button selected
		gbSTSplitCredits = true
		gbETDebitIsTotal = true
		oSTDlgCtrl.Text = gsETAcct1
	else
		gbSTSplitCredits = false
		gbETDebitIsTotal = false
		oSTDlgCtrl.Text = gsETAcct2
	endif	'// end Debit Total button State conditional

NormalExit:
	exit sub

ErrorHandler:
	msgBox("STDebtRBListener - unprocessed error.")
	GoTo NormalExit

end sub		'// end STDebtRBListener	7/6/20
'/**/
