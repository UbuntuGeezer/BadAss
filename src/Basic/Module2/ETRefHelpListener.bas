'// ETRefHelpListener.bas
'//---------------------------------------------------------------------
'// ETRefHelpListener - Event handler <?> from Enter Transaction.
'//		6/16/20.	wmk.	18:00
'//---------------------------------------------------------------------

public sub ETRefHelpListener()

'//	Usage.	macro call or
'//			call ETRefHelpListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <reference> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	msgBox displayed with Reference Help information
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Referemce> text field in the Enter Transaction dialog.

'//	constants.

'//	local variables.

	'// code.
	msgBox("Reference field: This field can be used similar to a memo" _
		+ CHR(13)+CHR(10) + "field on a check. It may contain a notation" _
		+ CHR(13)+CHR(10) + "such as 'Monthly', return credit, etc." _
		+ CHR(13)+CHR(10) + "and is limited to 12 characters.")

end sub		'// end ETRefHelpListener	6/16/20
'/**/
