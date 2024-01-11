'// ETRecdContListener.bas
'//---------------------------------------------------------------------
'// ETRecdContListener - Event handler <Record & Continue> from Enter Transaction.
'//		1/11/24.	wmk.	12:30
'//---------------------------------------------------------------------

public sub ETRecdContListener()

'//	Usage.	macro call or
'//			call ETRecdContListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked the [Record & Finish] button in ET dialog
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
'//	1/11/24.	wmk.	original; adapted from ETRecdDoneListener.
'//
'//	6/16/20.	wmk.	original code; stub
'//	6/17/20.	wmk.	completed; waiting on ETDialogRecord functional
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Record & Continue> cmd button in the Enter Transaction dialog.

'//	constants.

'//	local variables.
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button
dim iStatus 		As Integer		'// general status

	'// code.
	iStatus = -1		'// set error return
	ON ERROR GoTo ErrorHandler

	msgbox("ETRecdContListener stubbed.. - exiting."
	GOTO NormalExit
	
	'// record transaction
	iStatus = ETDialogRecord()
	if iStatus < 0 then
		GoToErrorHandler
	endif
	
	'// clear all fields entered and associated flags
	iStatus = ETPubVarsReset(1)	'// reset everything
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	puoETDialog.endDialog(2)			'// end dialog
	exit sub
	
ErrorHandler:
	GoTo NormalExit
	
end sub		'// end ETRecdContListener	1/11/24
'/**/
