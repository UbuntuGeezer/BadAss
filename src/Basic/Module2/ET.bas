'// ET.bas
'//---------------------------------------------------------------
'// ET - Enter Transaction Dialog run procedure.
'//		9/8/22.	wmk.	13:47	
'//----------------------------------------------------------------

public sub ET()

'//	Usage.	macro call or
'//			call ET()
'//
'//		Run ET dialog
'//
'// Entry.	Button click in dialog box		
'//			puoETDialog = declared object for COA dialog
'//			pbETActive = flag to indicate ET dialog is active
'//			pbETDebitList = ET in COA debit list flag
'//			pbETCreditList = ET in COA credit list flag
'//
'//	Exit.	msgBox message displayed
'//
'// Calls.	ETPubVarsReset
'//
'//	Modification history.
'//	---------------------
'// 9/7/22.		wmk.	ETRecdDoneListener called when "Record" (OK) clicked.
'// 9/8/22.		wmk.	pbETActive, pbETDebitList, pbETCreditList flags
'//				 supported for COADialog results.
'// Legacy mods.
'//	7/6/20.		wmk.	original code; adapted from G in BadAss Module2
'//	7/7/20.		wmk.	ensure dialog library loaded
'//
'//	Notes.	pbInETDialog is a flag set for the COADialog to indicate it is being
'// invoked by the ET dialog, as opposed to the GA macro. The COADialog then
'// knows to set the COADialog.DebitField or COADialog.CreditField as opposed
'// to the selected cell in the sheet it has been invoked on by the GA
'// macro (ctl-4 hotkey).
'//

'//	constants.

'//	local variables.
dim iStatus As Integer
dim oDocument	As Object
dim oSel		As Object
dim oRange		As Object
dim iSheet		As Integer

'// code.
	GlobalScope.BasicLibraries.LoadLibrary("BadAss")
	GlobalScope.DialogLibraries.LoadLibrary("BadAss")
	oDocument = ThisComponent
	oSel = oDocument.getCurrentSelection()
	oRange = oSel.RangeAddress
	iSheet = oRange.Sheet
		
'// Initialize ET dialog.
	puoETDialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.TransEntry)
if( 1 = 0 ) then
Xray puoETDialog
endif
	pbETActive=true
	pbETDebitList = false
	pbETCreditList = false
	Select Case puoETDialog.Execute()
	Case 1		'// Record &' Finish clicked
		ETRecdDoneListener()
'		msgBox("Record &' Finish clicked in Enter Transaction")
		'// iStatus = ETDialogRecord... called by [Record &' Finish]
	Case 0		'// Cancel
'		pusCOASelected = ""		'// since cancel, Execute() event skipped
		iStatus = ETPubVarsReset(1)	'// reset everything
'		msgBox("Cancel or FormClose clicked")
	Case else
'		iStatus = ETPubVarsReset(1)	'// reset everything
'		msgBox("ET - unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gdETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)
'puoETDialog.endDialog(2)
'puoETDialog.dispose()
	
end sub		'// end ET	9/8/22
'/**/
