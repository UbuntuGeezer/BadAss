'// ETCOA1Listener.bas
'//---------------------------------------------------------------------
'// ETCOA1Listener - Event handler <COA1 account> from Enter Transaction.
'//		6/17/20.	wmk.	14:30
'//---------------------------------------------------------------------

public sub ETCOA1Listener()

'//	Usage.	macro call or
'//			call ETCOA1Listener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <COA1 account> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gsETAcct1 = entered amount
'//			gbETCOA1Entered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.	wmk.	original code.
'//	6/17/20.	wmk.	bug fix enabling [Record..] buttons
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <debit> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETCOA1Text	As Object			'// account text field
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button

'//	local variables.

	'// code.
	oETCOA1Text = puoETDialog.getControl("DebitField")
	gsETAcct1 = oETCOA1Text.Text
	gbETCOA1Entered = (Len(gsETAcct1)>0)
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETCOA1Listener	6/17/20
'/**/
