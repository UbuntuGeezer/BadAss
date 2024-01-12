'// STDesc2LostFocus.bas
'//---------------------------------------------------------------
'// STDesc2LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	05:15
'//---------------------------------------------------------------

public sub STDesc2LostFocus()

'//	Usage.	macro call or
'//			call STDesc2LostFocus()
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
	gsSTObjFocus = "DescFld2"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc2LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc2LostFocus	6/25/20
'/**/
