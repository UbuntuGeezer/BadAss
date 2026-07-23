'// STCancelBtnRun.bas
'//---------------------------------------------------------------
'// STCancelBtnRun - ST dialog <Cancel> button Execute() event handler.
'//		6/27/20.	wmk.	15:00
'//---------------------------------------------------------------

public sub STCancelBtnRun()

'//	Usage.	macro call or
'//			call STCancelBtnRun()
'//
'// Entry.	puoSTDialog = ST dialog object
'//
'//	Exit.	ST dialog ended, returning exit code 0
'//			all ST-related vars cleared
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.		wmk.	original code
'//	6/27/20.	wmk.	gbSTEditMode flag set false
'//
'//	Notes. It is assumed that if the user hits <Cancel>, any
'// split information should be abandoned
'//

'//	constants.

'//	local variables.
dim iStatus		As Integer

	'// code.
	'// clear all stored array values and exit with cancel
	gbSTEditMode = false
	iStatus = STPubVarsReset()
	puoSTDialog.endDialog(0)
	
end sub		'// end STCancelBtnRun	6/27/20
'/**/
