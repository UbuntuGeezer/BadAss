'// STSetObjFocus.bas
'//---------------------------------------------------------------
'// STSetObjFocus - Set control name with focus inside object.
'//		6/22/20.	wmk.	10:00
'//---------------------------------------------------------------

public function STSetObjFocus(psControl As String) As Integer

'//	Usage.	iVar = STSetObjFocus( sControl )
'//
'//		sControl = dialog control name with focus
'//					(e.g. "AmtFld1")
'//
'// Entry.	puoSTDialog = ST dialog object
'//			puoSTDialog.<HasFocusObject>.Text = available
'//
'//	Exit.	iVar = 0 - control name stored in dialog
'//				 < 0 - error attempting to store active control namne
'//			puoSTDialog.<HasFocus-Object>.Text = psControl
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.		wmk.	original code
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
msgBox("In STSetObjFocus.. calling .getControl(..")
	oDlgControl = puoSTDialog.getControl("HasFocusFld")
msgBox("In STSetObjFocus back from .getControl on entry.." + CHR(13)+CHR(10) _
+	"stored name = '" + oDlgControl.Text + "'")
	sFocusName = oDlgControl.Text

'	if Len(sFocusName) > 0 then
'		iRetValue = -2		'// set busy flag
'		GoTo ErrorHandler
'	endif		'// end HasFocus in use conditional

	oDlgControl.Text = psControl		'// set control name
msgBox("In STSetObjFocus sFocusName = '" + psControl + "'")
	
	iRetValue = 0			'// set normal return
	
NormalExit:	
	STSetObjFocus = iRetValue
	exit function

ErrorHandler:
	Select Case iRetValue
	Case -2
		msgBox("In STSetObjFocus - 'HasFocus' field busy")
	Case else
		msgBox("In STSetObjFocus - unprocessed error")
	End Select
	
	GoTo NormalExit
	
end function 	'// end STSetObjFocus	6/22/20
'/**/
