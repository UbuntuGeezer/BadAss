'// STCheckComplete.bas
'//---------------------------------------------------------------
'// STCheckComplete - Check if dialog has minimum information.
'//		7/7/20.	wmk.	12:40
'//---------------------------------------------------------------

public function STCheckComplete() As Boolean

'//	Usage.	bVal = STCheckComplete()
'//
'// Entry.	giSTLastIx = index of last stored row from dialog
'//			gdSTAmts(0..giSTLastIx0 = stored amounts from dialog
'//			gsSTCOAs(0..giSTLastIx) = stored account #s from dialog
'//
'//	Exit.	bVal = true if ALL of the following:
'//					gdSTAmts(0..STLastIx) = non-empty values
'//					gsSTCOAs(0..STLastIx) = each exactly 4 chars
'//					gdSTTotalAmt = gdSTAccumTot
'//				   false otherwise
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'//	6/21/20.	wmk.	modified to check dialog UnsplitTot = 0.
'//	7/7/20.		wmk.	bug fix where fractional cents making
'//						UnsplitTot <> 0; changed to string comparison

'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim bRetValue As Boolean		'// returned flag
dim bComplete	As Boolean	'// interim tests flag.
dim i			As Integer		'// loop index
dim oUnsplit	As Object		'// unsplit field from dialot

	'// code.
	bComplete = true
	for i = 0 to giSTLastIx
		bComplete = bComplete AND (Len(gsSTCOAs(i)) = 4)
		bComplete = bComplete AND (gdSTAmts(i) >= 0)
		if NOT bComplete then
			exit for
		endif	'// end failed conditional	
	next i		'// loop for all entered rows

	oUnSplit = puoSTDialog.getControl("UnsplitTot")	
	bComplete = bComplete AND _
			(StrComp(Str(gdSTTotalAmt),Str(gdSTAccumTot))=0)
'	bComplete = bComplete AND (oUnSplit.Value = 0.)
	STCheckComplete = bComplete

end function 	'// end STCheckComplete	7/7/20
'/**/
