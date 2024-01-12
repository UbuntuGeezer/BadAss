'// STRefLostFocus.bas
'//---------------------------------------------------------------
'// STRefLostFocus - Handle any RefFld LostFocus event.
'//		8/26/22.	wmk.	16:04
'//---------------------------------------------------------------

public function STRefLostFocus(psFldName As String) As Integer

'//	Usage.	macro call or
'//			call STRefLostFocus(sFieldName)
'//
'//	Parameters.	sFldName = name of dialog object lost focus
'//
'// Entry.	RerFld1..RefFld10 lost focus
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gsSTRefs(n) = text from RefFld1..10
'//
'// Calls.	STFld6ToIndex.
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; cloned from STRef1LostFocus.
'// 8/26/22.	wmk.	comments corrected.
'//
'//	Notes. gsSTObjFocus will be used by the [...] button Execute
'//	event to determine which Ref array string to store the selected
'// Ref into.


'//	constants.

'//	local variables.
dim iRetValue 		As Integer			'// returned status
dim iStatus			As Integer			'// general status
dim iRowIx			As Integer			'// index of present row
dim oDlgControl		As Object			'// field ptr
dim bRowComplete	As Boolean			'// row complete flag
dim bDlgComplete	As Boolean			'// dialog complete flag
dim sFldName 		As String			'// local copy of field name

	'// code.
	iRetValue = -1						'// set bad return
	ON ERROR GOTO ErrorHandler
'// Note: the callout to STUpdateRefArray was mysteriously getting messed up
'// and not exiting properly. ODDLY, this may have been because of an unhandled
'// error in the AcctFld handlers where STAcctLostFocus was removed from the code
'// inadvertantly, but still called and not caught in the runtime error handlers...

	sFldName = psFldName				'// use local copy of field name
	oDlgControl = puoSTDialog.getControl(sFldName)
	iRowIx = STFld6ToIndex(sFldName)
	gsSTRefs(iRowIx) = oDlgControl.Text
	
	iRetValue = 0

NormalExit:
	STRefLostFocus = iRetValue
	exit function
	
ErrorHandler:
	msgBox("In STRefLostFocus - unprocessed error")
	GoTo NormalExit
	
end function		'// end STRefLostFocus	8/26/22
'/**/
