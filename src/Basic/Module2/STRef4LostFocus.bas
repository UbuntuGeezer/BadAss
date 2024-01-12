'// STRef4LostFocus.bas
'//---------------------------------------------------------------
'// STRef4LostFocus - Handle RefFldn LostFocus event.
'//		6/25/20.	wmk.	13:30
'//---------------------------------------------------------------

public sub STRef4LostFocus()

'//	Usage.	macro call or
'//			call STRef4LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsRefs(0) = text from RefFld1..10
'//
'// Calls.	STRefLostFocus
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
	'// use generic handler
	gsSTObjFocus = "RefFld4"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef4LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef4LostFocus	6/25/20
'/**/
