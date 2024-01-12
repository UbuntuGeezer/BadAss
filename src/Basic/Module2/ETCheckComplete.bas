'// ETCheckComplete.bas
'//---------------------------------------------------------------------
'// ETCheckComplete - Check all Enter Transaction dialog fields entered.
'//		6/15/20.	wmk.	17:30
'//---------------------------------------------------------------------

public function ETCheckComplete() As Boolean

'//	Usage.	bVar = ETCheckComplete()
'//
'// Entry.
'//		gbETDateEntered = date field entered flag
'//		gbETDescEntered = description entered flag
'//		gbETAmtEntered As Boolean			'// amount entered flag
'//		gbETCOA1Entered As Boolean			'// Debit COA entered flag
'//		gbETCOA2Entered As Boolean			'// Credit COA entered flag
'//
'//	Exit.	bVal = true if all entry flags are true
'//				   false otherwise
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.		wmk.	original code
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim bRetValue As Boolean

	'// code.
	
	bRetValue = gbETDateEntered AND gbETDescEntered AND gbETAmtEntered _
				 AND gbETCOA1Entered AND gbETCOA2Entered
	
	ETCheckComplete = bRetValue

end function 	'// end ETCheckComplete		6/15/20
'/**/
