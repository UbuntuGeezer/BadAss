'// STGetObjFocus.bas
'//---------------------------------------------------------------
'// STGetObjFocus - Get control name with focus from dialog.
'//		6/22/20.	wmk.	11:00
'//---------------------------------------------------------------

public function STGetObjFocus() As String

'//	Usage.	sVar = STGetObjFocus()
'//
'// Entry.	puoSTDialog = ST dialog object
'//			puoSTDialog.<HasFocusObject>.Text = available
'//
'//	Exit.	sVar = control name from "HasFocus".Text in dialog
'//				 = "" if error, or not previously set
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
dim sRetValue As String
dim oDlgControl		As Object		'// dialog control
dim sFocusName		As String		'// HasFocus text control field name

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	oDlgControl = puoSTDialog.getControl("HasFocusFld")
	sFocusName = oDlgControl.Text
msgBox("In STGetObjFocus sFocusName = '" + sFocusName + "'")
	sRetValue = sFocusName
	
NormalExit:	
	STGetObjFocus = sRetValue
	exit function

ErrorHandler:
	sRetValue = ""		'// abnormal return value
	msgBox("In STGetObjFocus - unprocessed error")
	GoTo NormalExit
	
end function 	'// end STGetObjFocus	6/22/20
'/**/
