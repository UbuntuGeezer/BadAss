'// ETRecdDoneListener.bas
'//---------------------------------------------------------------------
'// ETRecdDoneListener - Event handler <Record &' Continue> from Enter Transaction.
'//		7/23/26.	wmk.	13:20
'//---------------------------------------------------------------------

public sub ETRecdDoneListener()

'//	Usage.	macro call or
'//			call ETRecdDoneListener()
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
'//	--------------------
'// 7/23/26.    wmk.    ilineNo var added for error tracking.
'// 7/23/26.	wmk.	(automated) Modification History sorted.
'// 9/7/22.     wmk.	msgbox added to error handling. 
'// 6/17/20.	wmk.	completed; waiting on ETDialogRecord functional 
'// 6/16/20.	wmk.	original code; stub 
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Record &' Continue> cmd button in the Enter Transaction dialog.

'//	constants.

'//	local variables.
dim oETRecordBtn	As Object		'// Record &' Finish button
dim oETRecordCont	As Object		'// Record &' Continue button
dim iStatus 		As Integer		'// general status
dim ilineNo         As Integer      '// error tracking line number

	'// code.
	iStatus = -1		'// set error return
	ON ERROR GoTo ErrorHandler

	'// record transaction
    ilineNo = 46
    iStatus = ETDialogRecord()
	if iStatus < 0 then
		GoToErrorHandler
	endif
	
	'// clear all fields entered and associated flags
    ilineNo = 53
	iStatus = ETPubVarsReset(1)	'// reset everything
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
    ilineNo = 60
	puoETDialog.endDialog(2)			'// end dialog
	exit sub
	
ErrorHandler:
    msgbox("ETRecdDoneListener - (line " + iLineNo + " unprocessed error.")
	GoTo NormalExit
	
end sub		'// end ETRecdDoneListener	7/23/26.
'/**/
