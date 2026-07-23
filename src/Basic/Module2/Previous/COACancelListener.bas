'// COACancelListener.bas
'//--------------------------------------------------------------------
'// COACancelListener - Event handler when <Cancel> from COA selection.
'//		9/7/22.	wmk.	17:26
'//--------------------------------------------------------------------

public sub COACancelListener()

'//	Usage.	macro call or
'//			call COACancelListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked <Cancel> in COA dialog
'//			puoCOADialog = COA dialog box object
'//
'//	Exit.	pusCOASelected = ""
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'// 9/7/22.		wmk.	this control behaves like "Cancel"; the
'//				 dialog has been dropped and disposed.
'// Legacy mods.
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
	pusCOASelected = ""		'// clear COA selection string
'	msgBox("In Module2/COACancelListener" + CHR(13)+CHR(10) _
'        + "'"+pusCOASelected+"'" + " selected")
'	puoCOADialog.endDialog(0)	'// end dialog as though cancelled	

end sub		'// end COACancelListener	6/13/20
'/**/
