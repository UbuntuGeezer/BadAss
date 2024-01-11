'// STRowComplete.bas
'//---------------------------------------------------------------
'// STRowComplete - Check if entry row has minimum information.
'//		6/20/20.	wmk.	20:00
'//---------------------------------------------------------------

public function STRowComplete(piRowIx As Integer) As Boolean

'//	Usage.	bVal = STRowComplete(iRowIx)
'//
'//	Parameters.	iRowIx = row to check for minimum information
'//
'//	Entry	gdSTAmts(iRowIx) has stored value
'//			gsSTCOAs(iRowIx) has 4 digits
'// Entry.	giSTLastIx = index of last stored row from dialog
'//			gdSTAmts(0..giSTLastIx0 = stored amounts from dialog
'//			gsSTCOAs(0..giSTLastIx) = stored account #s from dialog
'//
'//	Exit.	bVal = true if ALL of the following:
'//					gdSTAmts(0..STLastIx) = non-empty values
'//					gsSTCOAs(0..STLastIx) = non-empty values
'//					gdSTTotalAmt = gdSTAccumTot
'//				   false otherwise
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	bug fix with return value not being set
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim bRetValue As Boolean		'// returned flag
dim bRowComplete	As Boolean	'// interim tests flag

	'// code.
	bRowComplete = (Len(gsSTCOAs(piRowIx)) = 4)
	bRowComplete = bRowComplete AND (gdSTAmts(piRowix) >= 0)	
	STRowComplete = bRowComplete
'STRowComplete = true	'//$
end function 	'// end STRowComplete	6/20/20
'/**/
