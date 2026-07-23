'// COASelectListener.bas
'//--------------------------------------------------------------------
'// COASelectListener - Event handler when <Select> from COA selection.
'//		9/7/22.	wmk.	17:24
'//--------------------------------------------------------------------

public sub COASelectListener()

'//	Usage.	macro call or
'//			call COASelectListener()
'//
'// Entry.	puoCOAListBox.SelectedItem = item selection from dialog
'//			puoCOADialog = COA dialog box object
'//
'//	Exit.	pusCOASelected = puoCOAListBox.SelectedItem
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'// 9/7/22.		wmk.	this event is fired as if "OK" were selected;
'//				 the dialog box has been ended normally and closed.
'// Legacy mods.
'//	6/13/20.	wmk.	original code
'//	6/14/20.	wmk.	close dialog code added
'// 6/17/20.	wmk.	debug msgBox disabled
'//
'//	Notes. This sub is the linked macro to the Execute() event linked
'// to the <Select> button in the COA selection dialog. The <Select>
'// button is also defined as type = 1 (OK), and will drop the dialog.

'//	constants.

'//	local variables.

	'// code.
	ON ERROR GOTO ErrorHandler
	pusCOASelected = puoCOAListBox.SelectedItem
'	msgBox("In Module2/COASelectListener" + CHR(13)+CHR(10) _
'        + pusCOASelected + " selected")
	puoCOADialog.endDialog(1)				'// close dialog when Select
	puoCOADialog.Dispose					'// release instance
NormalExit:
    exit sub
ErrorHandler:
	msgbox("COASelectListener - unprocessed error.")
	GOTO NormalExit
end sub		'// end COASelectListener	6/17/20
'/**/
