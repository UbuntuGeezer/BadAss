'// ET.bas
'//---------------------------------------------------------------
'// ET - Enter Transaction Dialog run procedure.
'//		7/7/20.	wmk.	23:00	
'//----------------------------------------------------------------

public sub ET()

'//	Usage.	macro call or
'//			call ET()
'//
'//		Run ET dialog
'//
'// Entry.	Button click in dialog box		
'//			puoETDialog = declared object for COA dialog
'//
'//	Exit.	msgBox message displayed
'//
'// Calls.	ETPubVarsReset
'//
'//	Modification history.
'//	---------------------
'//	7/6/20.		wmk.	original code; adapted from G in BadAss Module2
'//	7/7/20.		wmk.	ensure dialog library loaded
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim iStatus As Integer

'// code.
	GlobalScope.BasicLibraries.LoadLibrary("BadAss")
	GlobalScope.DialogLibraries.LoadLibrary("BadAss")
	
'// Initialize ET dialog.
	puoETDialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.TransEntry)

	Select Case puoETDialog.Execute()
	Case 2		'// Record & Finish clicked
'		msgBox("Record & Finish clicked in Enter Transaction")
		'// iStatus = ETDialogRecord... called by [Record & Finish]
	Case 0		'// Cancel
'		pusCOASelected = ""		'// since cancel, Execute() event skipped
		iStatus = ETPubVarsReset(1)	'// reset everything
'		msgBox("Cancel or FormClose clicked")
	Case else
		msgBox("ET.bas - unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gdETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

puoETDialog.dispose()
	
end sub		'// end ET	7/7/20
'/**/
