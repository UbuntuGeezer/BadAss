'// VSEditSplitBtnRun.bas
'//---------------------------------------------------------------------
'// VSEditSplitBtnRun - Event handler <Done> from View Split.
'//		6/27/20.	wmk.	17:00
'//---------------------------------------------------------------------

public sub VSEditSplitBtnRun()

'//	Usage.	macro call or
'//			call VSEditSplitBtnRun()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked the <Edit Split> button in VS dialog
'//			puoVSDialog = View Transaction dialog object
'//
'//	Exit.	VS dialog ended with flag = 2 (Edit Split)
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/27/20.	wmk.	original code
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Edit Split> cmd button in the View Split  dialog.

'//	constants.

'//	local variables.

	'// code.
	ON ERROR GoTo ErrorHandler
	gbSTEditMode = true				'// set ST edit mode flag
	
NormalExit:
	puoVSDialog.endDialog(2)		'// end dialog; same as Cancel
	exit sub
	
ErrorHandler:
	msgBox("VSEditSplitBtnRun - unprocessed error.")
	GoTo NormalExit
	
end sub		'// end VSEditSplitBtnRun	6/27/20
'/**/
