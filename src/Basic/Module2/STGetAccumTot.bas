'// STGetAccumTot.bas
'//---------------------------------------------------------------
'// STGetAccumTot - Get accumulated splits total.
'//		6/18/20.	wmk.
'//---------------------------------------------------------------

public function STGetAccumTot() As Double

'//	Usage.	dVar = STGetAccumTot()
'//
'//	Entry.	gdAccumTot = accumulated split total
'//
'//	Exit.	dVar = accumulated splits total
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
dim dRetValue	As Double		'// returned value

	'// code.
	dRetValue = gdAccumTot		'// return accumulated total
	STGetAccumTot = dRetValue
	
end function		'// end STGetAccumTot		6/18/20
'/**/
