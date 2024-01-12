'// ETCOA2Listener.bas
'//---------------------------------------------------------------------
'// ETCOA2Listener - Event handler <COA1 account> from Enter Transaction.
'//		6/17/20.	wmk.	10:30
'//---------------------------------------------------------------------

public sub ETCOA2Listener()

'//	Usage.	macro call or
'//			call ETCOA2Listener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <COA2 account> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gsETAcct2 = entered account
'//			gbETCOA2Entered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code
'//	6/17/20.	wmk.	bug fix enabling Select buttons
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Credit account> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETCOA2Text	As Object			'// account text field
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button

'//	local variables.

	'// code.
	oETCOA2Text = puoETDialog.getControl("CreditField")
	gsETAcct2 = oETCOA2Text.Text
	gbETCOA2Entered = (Len(gsETAcct2)>0)
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETCOA2Listener	6/17/20
'/**/
