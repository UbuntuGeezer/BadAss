'// ETDateListener.bas
'//---------------------------------------------------------------------
'// ETDateListener - Event handler <date> from Enter Transaction.
'//		6/17/20.	wmk.	10:45
'//---------------------------------------------------------------------

public sub ETDateListener()

'//	Usage.	macro call or
'//			call ETDateListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <date> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gdtETDate = entered date
'//			gbETDateEntered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.	wmk.	original code
'//	6/17/20.	wmk.	bug fix enabling Select buttons
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <date> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETDateText		As Object		'// Date text field
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button

'//	local variables.

	'// code.
	oETDateText = puoETDialog.getControl("DateField1")
	gsETDate = oETDateText.Text
	gbETDateEntered = (Len(gsETDate)>0)
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETDateListener	6/17/20
'/**/
