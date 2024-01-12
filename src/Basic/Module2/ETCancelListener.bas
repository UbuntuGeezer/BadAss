'// ETCancelListener.bas
'//---------------------------------------------------------------------
'// ETCancelListener - Event handler <Cancel> from Enter Transaction.
'//		6/15/20.	wmk.	16:30
'//---------------------------------------------------------------------

public sub ETCancelListener()

'//	Usage.	macro call or
'//			call ETCancelListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked <Cancel> in Enter Transaction dialog
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	ET dialog ended
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.	wmk.	original code
'//	6/15/20.	wmk.	code modified to emulate Cancel type button
'//
'//	Notes. This sub is the linked macro to the Execute() event linked
'// to the <Cancel> button in the COA selection dialog. If the <Cancel>
'// button is also defined as type = 2, CANCEL it will drop the dialog
'// from the screen with returned value 0. If not, this event handler
'// be called and will emulate that process.

'//	constants.

'//	local variables.

	'// code.
	puoETDialog.endDialog(0)	'// end dialog as though cancelled	

end sub		'// end ETCancelListener	6/15/20
'/**/
