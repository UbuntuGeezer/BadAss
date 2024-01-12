'// COAListBoxListener.bas
'//--------------------------------------------------------------------
'// COAListBoxListener - Event handler when <Cancel> from COA selection.
'//		6/15/20.	wmk.	12:30
'//--------------------------------------------------------------------

public sub COAListBoxListener()

'//	Usage.	macro call or
'//			call COAListBoxListener()
'//
'// Entry.	normal entry from event handler for dialog where
'//			user selecting COA account from listbox
'//			puoCOASelectBtn = <Select> button object in dialog
'//	Exit.	pusCOASelected = ""
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.		wmk.	original code
'//
'//	Notes. This sub is the linked macro to the Execute() event linked
'// to the <COAList> listbox in the COA selection dialog. This event
'// activates the <Select> button in the dialog.

'//	constants.

'//	local variables.

	'// code.
	puoCOASelectBtn.Model.Enabled = true	'// enable <Select> button

end sub		'// end COAListBoxListener	6/15/20
'/**/
