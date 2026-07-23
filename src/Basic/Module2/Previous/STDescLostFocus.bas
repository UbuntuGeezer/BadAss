'// STDescLostFocus.bas
'//---------------------------------------------------------------
'// STDescLostFocus - Handle any DescFld LostFocus event.
'//		6/25/20.	wmk.	05:30
'//---------------------------------------------------------------

public function STDescLostFocus(psFldName As String) As Integer

'//	Usage.	macro call or
'//			call STDescLostFocus()
'//
'//	[Parameters.	sFldName = name of dialog object lost focus]
'//
'// Entry.	AcctFld1..AcctFld10 lost focus
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gsCOAs(0) = text from AcctFld1..10
'//
'// Calls.	STFld7ToIndex
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; cloned from STRefLostFocus
'//
'//	Notes. gsSTObjFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into.


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
	sFldName = psFldName				'// use local copy of field name
	oDlgControl = puoSTDialog.getControl(sFldName)
	iRowIx = STFld7ToIndex(sFldName)
	gsSTDescs(iRowIx) = oDlgControl.Text
	
	iRetValue = 0

NormalExit:
	STDescLostFocus = iRetValue
	exit function
	
ErrorHandler:
	msgBox("In STDescLostFocus - unprocessed error")
	GoTo NormalExit
	
end function		'// end STDescLostFocus	6/25/20
'/**/
