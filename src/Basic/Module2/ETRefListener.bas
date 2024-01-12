'// ETRefListener.bas
'//---------------------------------------------------------------------
'// ETRefListener - Event handler <Reference> from Enter Transaction.
'//		6/16/20.	wmk.	18:00
'//---------------------------------------------------------------------

public sub ETRefListener()

'//	Usage.	macro call or
'//			call ETRefListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <reference> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gsETRef = entered reference
'//			gbETRefEntered = true if nonempty string
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
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETRefText	As Object			'// reference text field

'//	local variables.

	'// code.
	oETRefText = puoETDialog.getControl("RefField")
	gsETRef = oETRefText.Text
	gbETRefEntered = (Len(gsETRef)>0)

end sub		'// end ETRefListener	6/16/20
'/**/
