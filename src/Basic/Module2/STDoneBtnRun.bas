'// STDoneBtnRun.bas
'//---------------------------------------------------------------
'// STDoneBtnRun - ST dialog <Done> button Execute() event handler.
'//		6/24/20.	wmk.
'//---------------------------------------------------------------

public sub STDoneBtnRun()

'//	Usage.	macro call or
'//			call STDoneBtnRun()
'//
'// Entry.	puoSTDialog = ST dialog object
'//
'//	Exit.	ST dialog ended, returning exit code 2
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.		wmk.	original code
'//
'//	Notes.
'//

'//	constants.

'//	local variables.

	'// code.
	puoSTDialog.endDialog(2)
	
end sub		'// end STDoneBtnRun	6/24/20
'/**/
