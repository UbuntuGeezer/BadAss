'// COASelectListener.bas
'//--------------------------------------------------------------------
'// COASelectListener - Event handler when <Select> from COA selection.
'//		6/17/20.	wmk.	07:45
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
'//	6/13/20.	wmk.	original code
'//	6/14/20.	wmk.	close dialog code added
'// 6/17/29.	wmk.	debug msgBox disabled
'//
'//	Notes. This sub is the linked macro to the Execute() event linked
'// to the <Select> button in the COA selection dialog. The <Select>
'// button is also defined as type = 0, and will drop allow the dialog
'// to continue.

'//	constants.

'//	local variables.

	'// code.
	pusCOASelected = puoCOAListBox.SelectedItem
'	msgBox("In Module2/COASelectListener" + CHR(13)+CHR(10) _
'        + pusCOASelected + " selected")
	puoCOADialog.endDialog(2)				'// close dialog when Select
	
end sub		'// end COASelectListener	6/17/20
'/**/
