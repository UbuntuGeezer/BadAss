'// STAmtLostFocus.bas
'//---------------------------------------------------------------
'// STAmt4LostFocus - Handle AmtFld4 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt4LostFocus()

'//	Usage.	macro call or
'//			call STAmt4LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld4"

'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld4"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt4LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt4LostFocus	6/23/20
'/**/
