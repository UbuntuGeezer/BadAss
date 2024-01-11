'// STRef1LostFocus.bas
'//---------------------------------------------------------------
'// STRef1LostFocus - Handle AcctFldn LostFocus event.
'//		6/24/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STRef1LostFocus()

'//	Usage.	macro call or
'//			call STRef1LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsCOAs(0) = text from AcctFld1..10
'//
'// Calls.	STAcctLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	bug fix storing COA; add code to check/enable
'//						next row if this row complete
'//	6/24/20.	wmk.	common code moved to STAcctLostFocus; generic
'//						call set up to use common code function
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "RefFld1"					'// clear object focus tracker
	iStatus = STRefLostFocus("RefFld1")
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef1LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef1LostFocus	6/24/20
'/**/
