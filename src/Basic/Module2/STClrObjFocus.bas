'// STClrObjFocus.bas
'//---------------------------------------------------------------
'// STClrObjFocus - Clear control name with focus inside object.
'//		6/22/20.	wmk.	10:00
'//---------------------------------------------------------------

public function STClrObjFocus() As Integer

'//	Usage.	iVar = STClrObjFocus()
'//
'// Entry.	puoSTDialog = ST dialog object
'//			puoSTDialog.<HasFocusObject>.Text = available
'//
'//	Exit.	iVar = 0 - control name cleared in dialog
'//				 < 0 - error attempting to store active control namne
'//			puoSTDialog.<HasFocus-Object>.Text = ""
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6//20.		wmk.	original code
'//
'//	Notes. STSetObjFocus should be called by any <control>Listener
'//	event when the <control> gets focus. This will provide an easy
'// way for a control to get its own name during processing. The
'// STClrObjFocus function should be called when the <control> loses
'// focus so other controls may reuse the dialog field "HasFocus".
'// A third method STGetObjFocus will retrieve the current "HasFocus"
'// .Text field for use during the control processing.
'//

'//	constants.

'//	local variables.
dim iRetValue As Any
dim oDlgControl		As Object		'// dialog control
dim sFocusName		As String		'// HasFocus text control field name

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	oDlgControl = puoSTDialog.getControl("HasFocusFld")
	sFocusName = oDlgControl.Text
	oDlgControl.Text = ""
	iRetValue = 0			'// set normal return
	
NormalExit:	
	STClrObjFocus = iRetValue
	exit function

ErrorHandler:
	msgBox("In STClrFocus - unprocessed error")
	GoTo NormalExit
	
end function 	'// end STClrObjFocus	6/22/20
'/**/
