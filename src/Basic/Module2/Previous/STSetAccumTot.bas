'// STSetAccumTot.bas
'//---------------------------------------------------------------
'// STSetAccumTot - Get accumulated splits total.
'//		6/18/20.	wmk.	18:00
'//---------------------------------------------------------------

public function STSetAccumTot(pdTotal As Double) As Integer

'//	Usage.	iVar = STSetAccumTot(dTotal)
'//
'//	Paramters.	dTotal = new total to set in AccumTot field
'//
'//	Entry.	gdAccumTot = allocated for accumulated split total
'//			puoSTDialog	= loaded ST dialog object
'//
'//	Exit.	iVar = 0 - normal return
'//				   -1 - error setting accumulated total
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.		wmk.	original code
'//
'//	Notes. gdAccumTot is initialized to 0, then updated as user enters
'//	split entries
'//

'//	constants.

'//	local variables.
dim iRetValue	As Integer		'// returned value
dim oSTAccumTotFld	As Object	'// accumulated total text field

	'// code.
	ON ERROR GOTO ErrorHandler
	iStatus = 0					'// normal return
	gdAccumTot = pdTotal	

	'// update the AccumTot text box in the ST dialog
	oSTAccumTotFld = puoSTDialog.getControl("AccumTot")
	oSTAccumTotFld.Value = pdAccumTotal
	
NormalExit:
	STSetAccumTot = iStatus
	exit function
	
ErrorHandler:	
	iRetValue = -1
	GoTo NormalExit
end function		'// end STSetAccumTot		6/18/20
'/**/
