'// ETAmountListener.bas
'//---------------------------------------------------------------------
'// ETAmountListener - Event handler <amount> from Enter Transaction.
'//		6/24/20.	wmk.	17:45
'//---------------------------------------------------------------------

public sub ETAmountListener()

'//	Usage.	macro call or
'//			call ETAmountListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <amount> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gdETAmount = entered amount
'//			gbETAmtEntered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.	wmk.	original code
'// 6/17/20.	wmk.	bug fix enabling <record> buttons
'//	6/24/20.	wmk.	change to use type Double gdETAmount
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <debit> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETDlgCtrl		As Object		'// reusable field object
dim oETRecordBtn	As Object		'// Record &' Finish button
dim oETRecordCont	As Object		'// Record &' Continue button

'//	local variables.

	'// code.
	oETDlgCtrl = puoETDialog.getControl("AmtField1")
	gdETAmount = oETDlgCtrl.getValue()
	gbETAmtEntered = (Len(gdETAmount)>0)
	
	'// Enable SplitOption if amount entered > 0
	if gbETAmtEntered then
		oETDlgCtrl = puoETDialog.getControl("SplitOption")
		oETDlgCtrl.Model.Enabled = true
	endif
	
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETAmountListener	6/24/20
'/**/
