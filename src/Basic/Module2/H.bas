'// H.bas
'//---------------------------------------------------------------
'// H - Split Transaction Dialog test procedure.
'//		8/26/22.	wmk.	10:24
'//----------------------------------------------------------------

public sub H()

'//	Usage.	macro call or
'//			call H()
'//
'// Entry.	Button click in dialog box		
'//			puoSTDialog = declared object for Split Transaction dialog
'//
'//	Exit.	gdSTTotalAmt = contrived total to split, say 150.00
'//			msgBox message displayed
'//
'// Calls.	STPubVarsReset
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code; adapted from G. bas.
'//	6/23/20.	wmk.	change to use fDialogReset.
'// 8/26/22.	wmk.	add intervening space to G. bas and H .bas in comments.
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim iStatus As Integer
dim dDummy As Double
dim sDummy As String

	'// code.
'	sDummy = Str(dDummy)
'	msgBox("string conversion of empty value is: " +sDummy)
'	GoTo NormalExit
	ON ERROR GOTO ErrorHandler	
	DialogLibraries.LoadLibrary("Standard")

'// Initialize ST dialog.
	puoSTDialog = CreateUnoDialog(DialogLibraries.Standard.SplitDialog)
'XRay puoSTDialog	
	'// emulate dialog being fired with Total to split initialized
	'// in public vars
	gdSTTotalAmt = 150.00
	iStatus = fSTDialogReset()		'// attempt to reset/initialize dialog
'	if iStatus < 0 then
'		GoTo ErrorHandler
'	endif
	dim oFld As Object
	oFld = puoSTDialog.getControl("HasFocusFld")
	oFld.Text = "Gotcha"
	Select Case puoSTDialog.Execute()
	Case 2		'// Record & Finish clicked
		msgBox("<Done>h clicked in Split Transaction")
	Case 0		'// Cancel
		iStatus = STPubVarsReset()	'// reset everything
		msgBox("ST dialog Cancel or FormClose clicked")
	Case else
		msgBox("unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gsETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

puoSTDialog.dispose()
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("in H - error initializing Split dialog.")
	GoTo NormalExit
	
end sub		'// end H	8/26/22
'/**/
