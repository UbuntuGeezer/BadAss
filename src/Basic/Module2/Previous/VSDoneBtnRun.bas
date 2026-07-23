'// VSDoneBtnRun.bas
'//---------------------------------------------------------------------
'// VSDoneBtnRun - Event handler <Done> from View Split.
'//		6/27/20.	wmk.	17:00
'//---------------------------------------------------------------------

public sub VSDoneBtnRun()

'//	Usage.	macro call or
'//			call VSDoneBtnRun()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked the [Record &' Finish] button in ET dialog
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	Transaction data recorded in GL sheet at user selection
'//			Transaction data in public vars cleared
'//			ET dialog ended with flag = 2 (user closed)
'//
'// Calls.	ETDialogRecord, ETPubVarsReset
'//
'//	Modification history.
'//	---------------------
'//	6/27/20.	wmk.	original code
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Done> cmd button in the View Split  dialog.

'//	constants.

'//	local variables.
dim oVSDoneBtn	As Object		'// Record &' Finish button
dim iStatus 		As Integer		'// general status

	'// code.
	ON ERROR GoTo ErrorHandler

NormalExit:
	puoVSDialog.endDialog(0)		'// end dialog; same as Cancel
	exit sub
	
ErrorHandler:
	msgBox("VSDoneBtnRun - unprocessed error.")
	GoTo NormalExit
	
end sub		'// end VSDoneBtnRun	6/17/20
'/**/
