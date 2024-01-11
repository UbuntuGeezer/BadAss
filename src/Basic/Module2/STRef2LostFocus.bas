'// STRef2LostFocus.bas
'//---------------------------------------------------------------
'// STRef2LostFocus - Handle AcctFldn LostFocus event.
'//		6/24/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STRef2LostFocus()

'//	Usage.	macro call or
'//			call STRef2LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsCOAs(0) = text from AcctFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; adapted from STRef2LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// use generic handler
	gsSTObjFocus = "RefFld2"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef2LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef2LostFocus	6/24/20
'/**/
