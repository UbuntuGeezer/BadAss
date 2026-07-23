'// STUpdateCOAArray.bas
'//---------------------------------------------------------------
'// STUpdateCOAArray - Update public COAs array.
'//		6/24/20.	wmk.	07:15
'//---------------------------------------------------------------

public function STUpdateCOAArray(psFldName As String) As Integer

'//	Usage.	iVar = STUpdateCOAArray( sFldName )
'//
'// 		sFldName dialog amount field name (e.g. "DescFld1")
'//
'// Entry.	This routine should be called on the "LostFocus" event
'//			on any DescFldn in the dialog, with the fieldname passed
'//			as the parameter.	
'//			puoSTDialog = Split Transaction dialog object
'//
'//	Exit.	iVar = 0 - Desc array updated ok
'//				 < 0 - error updating Desc Array
'//			gdSTCOAs(i) = matching index i updated with new string
'//					 from dialog
'//
'// Calls.	STFld6ToIndex
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

	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sDlgFld = psFldName		'// get local copy of name
	
	'// convert name to array index
	iDlgFldIx = STFld6ToIndex(sDlgFld)
	if iDlgFldIx < 0 then
		GoTo ErrorHandler
	endif
	
	'// get new field dialog and store in array
	oDlgFld = puoSTDialog.getControl(sDlgFld)
	gdSTCOAs(iDlgFldIx) = oDlgFld.Text
	
	iRetValue = 0			'// normal return
	
NormalExit:
	STUpdateCOAArrray = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STUpdateCOAArray - unprocessed error.")
	GoTo NormalExit
	
end function		'// end STUpdateCOAArray	6/24/20
'/**/
