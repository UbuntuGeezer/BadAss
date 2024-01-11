'// ETDescListener.bas
'//---------------------------------------------------------------------
'// ETDescListener - Event handler <description> from Enter Transaction.
'//		6/17/20.	wmk.	10:15
'//---------------------------------------------------------------------

public sub ETDescListener()

'//	Usage.	macro call or
'//			call ETDescListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <description> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gsETDescription = entered transaction description
'//			gbETDescEntered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.	wmk.	original code
'//	6/17/20.	wmk.	bug fix enabling Select button
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <date> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETDescText		As Object		'// Desc text field
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button

'//	local variables.

	'// code.
	oETDescText = puoETDialog.getControl("DescField")
	gsETDescription = oETDescText.Text
	gbETDescEntered = (Len(gsETDescription)>0)
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETDescListener	6/17/20
'/**/
