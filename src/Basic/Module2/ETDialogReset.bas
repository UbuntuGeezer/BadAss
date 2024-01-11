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

	'// clear dialog object fields.
	iRetValue = ETDlgCtrlsReset(iOption)	'// clear dialog fields
	if iRetValue <> 0 then
		GoTo ErrorHandler
	endif
	
NormalExit:
	ETDialogReset = iRetValue
	exit function
	
ErrorHandler:
	iRetValue = -1		'// set error return
	GoTo NormalExit
	

end function 	'// end ETDialogReset	6/17/20
'/**/
