'// ETSplitRun.bas
'//---------------------------------------------------------------
'// ETSplitRun - [Split] button Execute() event handler.
'//		7/6/20.	wmk.	11:30
'//---------------------------------------------------------------

public sub ETSplitRun()

'//	Usage.	macro call or
'//			call ETSplitRun( )
'//
'// Entry.	User clicked [Split] control in ET dialog
'//			Standard library loaded, or else wouldn't be here
'//			puoSTDialog = defined as object for dialog
'//			gdSTTotalAmt = available for total from line 1 of transaction
'//
'//	Exit.	ST dialog run; on STdialog completion, global vars
'//			set with ST dialog entries
'//			gsSTDescs() = Description fields entered
'//			gsSTAmts() = Amount fields entered
'//			gsSTCOAs() = COA account fields entered
'//			gsSTRefs() = Reference fields entered
'//			gsSTTotalAcct = COA account for splitting
'//			
'//	ST Returned status = 	<Cancel> from ST dialog
'//							<Done> from ST dialog
'//			IF <Done>, activate [View Split] control
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code
'//	6/27/20.	wmk.	message boxes eliminated; TotalAcct picked up from
'//						ST dialog; disable ViewSplitBtn if user cancelled split
'//	7/6/20.		wmk.	change to use GlobalScope libraries/BadAss
'//
'//	Notes.	[Split] button should only activate after line 1 of ET
'// transaction has a nonzero value.
'//

'//	constants.

'//	local variables.
dim iStatus 	As Integer
dim	oDlgCtrl	As Object		'// ET control object to access Split RB
dim oDlgCtrlA	As Object		'// ET control object when accessing Debit/Credit

	'// code.
	ON ERROR GOTO ErrorHandler	

'// Initialize ET dialog.
	puoSTDialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.SplitDialog)

'//	set total amount previously set by ET AmtField1...
	gdSTTotalAmt = gdETAmount	
'	gdSTTotalAmt = 150.00

	iStatus = fSTDialogReset()		'// attempt to reset/initialize dialog
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
	
	Select Case puoSTDialog.Execute()
	Case 2		'// Done clicked
'		msgBox("Record & Finish clicked in Split Transaction")
		'// update DebitField or CreditField field with "split" based on flag
'//		gbSTSplitCredits = true; split CreditField; else split DebitField
		if gbSTSplitCredits then
			oDlgCtrl = puoETDialog.getControl("CreditField")
			oDlgCtrlA = puoETDialog.getControl("DebitField")
			oDlgCtrlA.Text = gsETAcct1		'// Debit is total
		else
			oDlgCtrl = puoETDialog.getControl("DebitField")
			oDlgCtrlA = puoETDialog.getControl("CreditField")
			oDlgCtrlA.Text = gsETAcct2		'// Credit is total
		endif	'// end credits are split conditional
		oDlgCtrl.Text = "split"

		
		'// enable [View Split] control
		oDlgCtrl = puoETDialog.getControl("ViewSplitBtn")
		oDlgCtrl.Model.Enabled = true
		
	Case 0		'// Cancel Split
		iStatus = STPubVarsReset()	'// reset everything
'		msgBox("ST dialog Cancel or FormClose clicked")
		if iStatus < 0 then
			GoTo ErrorHandler
		endif
		
		'// came back cancelled - deselect Split radio button
		oDlgCtrl = puoETDialog.getControl("SplitOption")
		oDlgCtrl.State = false
		oDlgCtrl = puoETDialog.getControl("ViewSplitBtn")
		oDlgCtrl.Model.Enabled = false
		
	Case else
		msgBox("unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gsETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

	'// clear instantiation of ET dialog
	puoSTDialog.dispose()

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("in ETSplitRun - error initializing Split dialog.")
	GoTo NormalExit
	
end sub		'// end ETSplitRun		7/6/20
'/**/
