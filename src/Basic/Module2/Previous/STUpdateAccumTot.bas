'// STUpdateAccumTot.bas
'//---------------------------------------------------------------
'// STUpdateAccumTot - Update accumulated split total.
'//		7/7/20.	wmk.	12:00
'//---------------------------------------------------------------

public function STUpdateAccumTot(psAmtFldName As String) As Integer

'//	Usage.	iVar = STUpdateAccumTot( sAmtFldName )
'//
'// Parameters.	sAmtFldName dialog amount field name (e.g. "AmtFld1")
'//
'// Entry.	This routine should be called on the "LostFocus" event
'//			on any AmtFld in the dialog, with the fieldname passed
'//			as the parameter.	
'//			puoSTDialog = Split Transaction dialog object
'//
'//	Exit.	iVar = 0 - totals updated; limit not reached
'//				   1 - totals updated; limit reached
'//				 < 0 - error updating totals
'//			gdStAmts(i) = matching index i updated with new value from dialog
'//			gdSTAccumTot = updated with any old value from array
'//					subtracted and new value from dialog added
'//			puoSTDialog.UnsplitTot.Value = updated with any old value
'//					from array added and new value from dialog subtracted
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'//	6/21/20.	wmk.	change to check split total met and return 1
'//	7/7/20.		wmk.	change to string comparison to allow for values
'//						equal except fractional cents; properly update
'//						totals without exception; fix bug where storing string
'//						into gdSTAmts array instead of value
'//
'//	Notes. The Accumulated total field is initialized to 0 at dialog
'//	startup, then is updated whenever the user adds a split amount to
'// any of the AmtFld1..AmtFld10 fields. STUpdateAccumTot should be
'// called whenever an AmtFldn object changes; The dialog UnsplitTot
'// field is also updated to hold the tracking
'//iAmtFldIx = index of ST dialog AmtFld to add to
'//			accumulator. If the amount field in the dialog has
'//			been cleared, any current value in the psSTAmts array
'//			will be subtracted from the accumulated total, the
'//			new value from the dialog box will be both stored in 
'//			the gsSTAmts array and added to the accumulated total.
'//			Then the Accumulated Splits Total box will be updated
'//			in the dialog objects.

'//	constants.

'//	local variables.
dim iRetValue 	As Integer		'// returned value
dim iAmtFldIx	As Integer		'// amounts array index
dim sAmtFld		As String		'// amount field local name
dim oAmtFld		As Object		'// amount field from dialog
dim oAccumTot	As Object		'// accumulated total field
dim dUnsplit	As Double		'// unsplit current amount

	'// code.

	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sAmtFld = psAmtFldName		'// get local copy of name
	
	'// convert name to array index
	iAmtFldIx = STFld6ToIndex(sAmtFld)
	if iAmtFldIx < 0 then
		GoTo ErrorHandler
	endif
	
	'//	adjust both accumulators by old value
	gdSTAccumTot = gdSTAccumTot - gdSTAmts(iAmtFldIx)
	oAccumTot = puoSTDialog.getControl("UnsplitTot")
	dUnsplit = oAccumTot.Value
	oAccumTot.Value = dUnsplit + gdSTAmts(iAmtFldIx)
	
	
	'// get new value from split transaction dialog and add
	'// into accumulator
	oAmtFld = puoSTDialog.getControl(sAmtFld)
	gdSTAccumTot = gdSTAccumTot + oAmtFld.Value
	
	if gdSTAccumTot > gdSTTotalAmt then
		if StrComp(Str(gdSTAccumTot),Str(gdSTTotalAmt))= 0 then
			gdSTAccumTot = gdSTTotalAmt
		else
			msgBox("Splits with new amount exceeds Total to split")
			iRetValue = 0
			GoTo NormalExit		'// bail out without updating array or control
		endif
	endif
	
	'// store new value in gdSTAmts array for this index
	gdStAmts(iAmtFldIx) = oAmtFld.Value
	
	'// update transaction dialog Accumulated Splits and Unsplit field
	oAccumTot = puoSTDialog.getControl("AccumTot")
	oAccumTot.Value = gdSTAccumTot
	oAccumTot = puoSTDialog.getControl("UnsplitTot")
	dUnsplit = oAccumTot.Value
	oAccumTot.Value = dUnsplit - oAmtFld.Value
		
'	if gdSTAccumTot = gdSTTotalAmt then
	if StrComp(Str(gdSTAccumTot),Str(gdSTTotalAmt))=0 then
		iRetValue = 1
	else
		iRetValue = 0
	endif	'// end split total met conditional
	
NormalExit:
	STUpdateAccumTot = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STUpdateAccumTot - unprocessed error.")
	GoTo NormalExit
	
end function		'// end STUpdateAccumTot	7/7/20
'/**/
