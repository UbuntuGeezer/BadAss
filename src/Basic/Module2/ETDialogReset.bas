'// ETDialogReset.bas
'//---------------------------------------------------------------
'// ETDialogReset - Reset all ET dialog fields and flags.
'//		6/17/20.	wmk.	12:00
'//---------------------------------------------------------------

public function ETDialogReset(piOpt As Integer) As Integer

'//	Usage.	iVar = ETDialogReset( iOpt )
'//
'//		iOpt = 0 - reset all fields except Date
'//			 <> 0 - reset ALL fields
'//
'// Entry.	ET dialog public field values and flags defined
'//			puoETDialog = ET dialog object
'//			gbETDateEntered = date field entered flag
'//			gsETDate = Date field string
'/
'//	Exit.	iVar = 0 - normal return
'//				 <> 0 - error in resetting fields
'//
'// Calls.	ETPubVarsReset, ETDlgCtrlsReset
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code; stub
'// 6/17/20.	wmk.	move code to ETPubVarsReset and ETDlgCtrlsReset
'//
'//	Notes. Allows call to preserve date field for repetitive entries

'//	constants.

'//	local variables.
dim iRetValue 	As Integer		'// returned value
dim iOption 	As Integer		'// local copy of initialize option

	'// code.
	
	iRetValue = 0
	iOption = piOpt
	ON ERROR GOTO ErrorHandler
	
	'// clear all stored values
	iRetValue = ETPubVarsReset(iOption)	'// clear field storage and flags
	if iRetValue <> 0 then
		GoTo ErrorHandler
	endif

if 1 = 0 then
	'// clear dialog object fields.
	iRetValue = ETDlgCtrlsReset(iOption)	'// clear dialog fields
	if iRetValue <> 0 then
		GoTo ErrorHandler
	endif
endif

	'// if reset all fields clear these; DateField1, DebitField, CreditField.
	if iOption = 1 then
dim oETDate			As Object
		oETDate =  puoETDialog.getControl("DateField1")
		oETDate.Text = ""
	
dim oETDebit		As Object
		oETDebit =  puoETDialog.getControl("DebitField")
		oETDebit.Text = ""

dim oETCredit		As Object
		oETCredit =  puoETDialog.getControl("CreditField")
		oETCredit.Text = ""

	endif

	'// Always clear description, amount, reference, splitoption.	
dim oETDescText		As Object
	oETDescText = puoETDialog.getControl("DescField")
	oETDescText.Text = ""		'// clear description
dim oETAmount		As Object
	oETAmount = puoETDialog.getControl("AmtField1")
	oETAmount.Text = ""			'// clear amount
dim oETRefText		As Object
	oETRefText = puoETDialog.getControl("RefField")
	oETRefText.Text = ""
dim oETSplit		As Object
	oETSplit = puoETDialog.getControl("SplitOption")
	oETSplit.Value = false
	
NormalExit:
	ETDialogReset = iRetValue
	exit function
	
ErrorHandler:
	iRetValue = -1		'// set error return
	GoTo NormalExit
	

end function 	'// end ETDialogReset	6/17/20
'/**/
