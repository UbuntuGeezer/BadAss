'// ETViewSplitRun.bas
'//---------------------------------------------------------------
'// ETViewSplitRun - [Split] button Execute() event handler.
'//		7/6/20.	wmk.	12:30
'//---------------------------------------------------------------

public sub ETViewSplitRun()

'//	Usage.	macro call or
'//			call ETViewSplitRun( )
'//
'// Entry.	User clicked [View Split] control in ET dialog
'//			Standard library loaded, or else wouldn't be here
'//			puoVSDialog = defined as object for dialog
'//			gdSTTotalAmt = available for total from line 1 of transaction
'//			gsSTDescs() = Description fields entered
'//			gsSTAmts() = Amount fields entered
'//			gsSTCOAs() = COA account fields entered
'//			gsRefs() = Reference fields entered
'//
'//	Exit.	VS dialog run; on VSDialog completion, global vars
'//			set with ST dialog entries
'//	ST Returned status = 	<Edit Split> from VS dialog
'//							<Done> from VS dialog
'//			IF <Done>, activate [View Split] control
'//
'// Calls.	fVSDialogPreset(), ETSplitRun
'//
'//	Modification history.
'//	---------------------
'//	6/26/20.	wmk.	original code; adapted from ETSplitRun;
'//						minimal code executed to test linkage
'//	6/27/20.	wmk.	initial bugs fixed; implement callout to
'//						ST dialog if <Edit Split> in VS
'//	7/6/20.		wmk.	change to GlobalScope/BadAss library
'//
'//	Notes.	[View Split] button should only activate after the user has
'//	completed a transaction split using ST dialog.


'//	constants.

'//	local variables.
dim iStatus 	As Integer
dim	oDlgCtrl	As Object		'// ET control object to access Split RB

	'// code.
	ON ERROR GOTO ErrorHandler	

'// Initialize ET dialog.
	puoVSDialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.ViewSplit)

if false then
GoTo JumpAround1
endif

	'// all public vars from the ET and ST dialogs are still current
	'// since this event is being handled for the ST dialog
	'//	fVSDialogPreset will initialize all the display fields in the
	'// VS dialog
	iStatus = fVSDialogPreset
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
	
JumpAround1:
	
	Select Case puoVSDialog.Execute()
	Case 2		'// <Edit Split> clicked
		msgBox("<Edit Split> clicked in View Split")
		puoVSDialog.dispose()
		GoTo RunSplitDlg
	Case 0		'// Cancel or Done
'		iStatus = STPubVarsReset()	'// reset everything
'		msgBox("VS dialog Done or FormClose clicked")
'		if iStatus < 0 then
'			GoTo ErrorHandler
'		endif
'		
'		'// came back cancelled - deselect Split radio button
'		oDlgCtrl = puoETDialog.getControl("SplitOption")
'		oDlgCtrl.State = false
		
	Case else
		msgBox("VS - unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gsETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

	'// clear instantiation of VS dialog
	puoVSDialog.dispose()

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("in ETViewSplitRun - error initializing View Split dialog.")
	GoTo NormalExit

RunSplitDlg:
	gbSTEditMode = true
	call ETSplitRun()
	GoTo NormalExit
	
end sub		'// end ETViewSplitRun		7/6/20
'/**/
