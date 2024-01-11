'// STDesc10LostFocus.bas
'//---------------------------------------------------------------
'// STDesc10LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STDesc10LostFocus()

'//	Usage.	macro call or
'//			call STDesc10LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld10"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc10LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc10LostFocus	6/25/20
'/**/
