'// STUpdateRefArray.bas
'//---------------------------------------------------------------
'// STUpdateRefArray - Update public Reference array.
'//		6/24/20.	wmk.	06:45
'//---------------------------------------------------------------

public function STUpdateRefArray(psFldName As String) As Integer

'//	Usage.	iVar = STUpdateRefArray( sFldName )
'//
'// 		sFldName dialog amount field name (e.g. "RefFld1")
'//
'// Entry.	This routine should be called on the "LostFocus" event
'//			on any REfFldn in the dialog, with the fieldname passed
'//			as the parameter.	
'//			puoSTDialog = Split Transaction dialog object
'//
'//	Exit.	iVar = 0 - Ref array updated ok
'//				 < 0 - error updating Ref Array
'//			gsSTRefs(i) = matching index i updated with new string
'//					 from dialog
'//
'// Calls.	STFld7ToIndex
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; adapted from STUpdateDescArray
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue 	As Integer		'// returned value
dim iDlgFldIx	As Integer		'// amounts array index
dim sDlgFld		As String		'// amount field local name
dim oDlgFld		As Object		'// amount field from dialog

	'// code.
'//WARNING: IT APPEARS THAT THE STACK GETS MESSED UP WHEN GOING 2 LEVELS
'// DOWN IN AN EVENT HANDLER... THIS ROUTINE WILL REMAIN IN PLACE, BUT
'// CANNOT BE CALLED DOWNSTREAM BY "LOSTFOCUS" EVENT HANDLING...
'// Experiment with making this a parameterless function and see if the
'// situation improves...

	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sDlgFld = psFldName		'// get local copy of name
msgBox("In STUpdateRefArray.. psFldName = '" + psFldName + "'" + CHR(13)+CHR(10) _
+	"sDlgFld = '" + sDlgFld + "'")	
	'// convert name to array index
	iDlgFldIx = STFld6ToIndex(sDlgFld)
msgBox(".. iDlgFldIx = " + iDlgFldIx + "'")	
	if iDlgFldIx < 0 then
		GoTo ErrorHandler
	endif
	
	'// get new field dialog and store in array

msgBox(".. invoking .getControl")	
	
	oDlgFld = puoSTDialog.getControl(sDlgFld)
msgBox(".. back from .getControl")	
	
	gsSTRefs(iDlgFldIx) = oDlgFld.Text
msgBox(".. gsSTRefs() value stored")	
	
	iRetValue = 0			'// normal return
	GoTo NormalExit
	
ErrorHandler:
	msgBox("STUpdateRefArray - unprocessed error.")

NormalExit:
msgBox("Exiting STUpdateRefArray; iRetValue = " + iRetValue)
	STUpdateRefArrray = iRetValue
	
end function		'// end STUpdateRefArray	6/24/20
'/**/
