'// STCOAListBtnListener.bas
'//---------------------------------------------------------------
'// STCOAListBtnListener - Handle COA List Button event.
'//		6/18/20.	wmk.
'//---------------------------------------------------------------

public sub STCOAListBtnListener()

'//	Usage.	macro call or
'//			call STCOAListBtnListener()
'//
'// Entry.	gsSTAcctFocus = name of last COA object to get focus
'//			puoSTDialog = ST Dialog object
'//
'//	Exit.	gsSTCOAs(n-1) updated with COA selection from GetCOA	
'//
'// Calls.	GetCOA
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.		wmk.	original code
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim iNameLen	As Integer	'// length of name string
dim sFldName	As String	'// full field name from gsSTAcctFocus var
dim sAcctIx		As String	'// account index extracted from field name
dim iDigCount	As Integer	'// field name digit count
dim iArrayIx	As Integer	'// array index into gsSTCOAs string array
dim oSTDlgCtrl	As Object	'// account field in STdialog
dim sCOANum		As String	'// COA from GetCOA

	'// code.
	ON ERROR GOTO ErrorHandler
	sFldName = gsSTAcctFocus		'// get "AcctFldxx" string
	iNameLen = Len(sFldName)		'// get length, either 8 or 9
	Select Case iNameLen
	Case 8
		iDigCount = 1
	Case 9
		iDigCount = 2
	Case else
		iDigCount = 1
	End Select
	
	sAcctIx = Right(sFldName,iDigCount)
	iArrayIx = Val(sAcctIx) - 1
	
	sCOANum = GetCOA(1)		'// get COA from dialog
	gsSTCOAs(iArrayIx) = sCOANum
	oSTDlgCtrl = puoSTDialog.getControl(sFldName)
	oSTDlgCtrl.Text = sCOANum
	
NormalExit:
	exit sub

ErrorHandler:
	msgBox("Error getting COA from list...")
	GoTo NormalExit
end sub		'// end STCOAListBtnListener		6/18/20
'/**/
