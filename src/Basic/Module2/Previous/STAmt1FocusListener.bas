'// STAmt1FocusListener.bas
'//---------------------------------------------------------------
'// STAmt1FocusListener - Handle AmtFldn GotFocus event.
'//		6/23/20.	wmk.	05:00
'//---------------------------------------------------------------

public sub STAmt1FocusListener()

'//	Usage.	macro call or
'//			call STAmt1FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAcct3FocusListener
'// 6/22/20.	wmk.	call to STSetObjFocus to set this control name
'//						in dialog for later retrieval
'//	6/23/20.	wmk.	reverted code to near original; eliminated
'//						STSetObjFocus call, since throws dialog into
'//						infinite loop when setting the field in the
'//						dialog, the "Got Focus" event gets re-invoked
'//						when returning from the "HasFocus" control;
'//						this may be conquerable with a flag in the 
'//						publics that indicates recursive callback and
'//						just exits, or skips ahead to code following
'//						STSetObject call
'//
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.

'//	local variables.
dim iStatus		As Integer			'// general status
dim oDlgControl	As Object			'// dialog control object

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler

'	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	gsSTObjFocus = "AmtFld1"
	
	if true then
		GoTo NormalExit
	endif
'//---------------------------------------------------------------------------------	
	oDlgControl = puoSTDialog.getControl("HasFocusFld")
msgBox("In STAmt1FocusListener back from .getControl on entry.." + CHR(13)+CHR(10) _
+	"stored name = '" + oDlgControl.Text + "'")
'	sFocusName = oDlgControl.Text

'	if Len(sFocusName) > 0 then
'		iRetValue = -2		'// set busy flag
'		GoTo ErrorHandler
'	endif		'// end HasFocus in use conditional

	oDlgControl.Text = "AmtFld1"		'// set control name
msgBox("In STAmt1FocusListener ready to exit.." + CHR(13)+CHR(10) _
+	"stored name = '" + oDlgControl.Text + "'")
	GoTo NormalExit
	
'//------------------------------------------------------------------------------
	iStatus = STSetObjFocus("AmtFld1")
msgBox("In STAmt1FocusListener, back from STSetObjectFocus..." + CHR(13)+CHR(10) _
+	"iStatus = " + iStatus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

	iStatus = 0
	GoTo NormalExit
'//--------------------------------------------------------------	
ErrorHandler:
	msgBox("In STAmt1FocusListener - unprocessed error")

NormalExit:
	exit sub		
	
end sub		'// end STAmt1FocusListener	6/23/20
'/**/
