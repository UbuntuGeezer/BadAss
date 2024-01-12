'// RunCOADlg.bas
'//---------------------------------------------------------------
'// RunCOADlg - Run COA Dialog.
'//		7/5/20.	wmk.	15:00
'//---------------------------------------------------------------

public sub RunCOADlg()

'//	Usage.	macro call or
'//			call RunCOADialog()
'//
'// Entry.	Button click in dialog box		
'//			puoCOADialog = declared object for COA dialog
'//			puoCOAListBox = list box object
'//			puoSelectBtn = declared object for <Select> button
'//			pusCOASelected = reserved for COA string selected
'//			current Component has sheet "Chart of Accounts" with COA table
'//			BadAss library exists
'//
'//	Exit.	msgBox message displayed
'//			pusCOASelected = COA selected
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.	wmk.	original code
'//	6/14/20.	wmk.	call LoadCOAs and run form added; code
'//						modified to use puoCOADialog object
'//	6/15/20.	wmk.	name of COA list dialog changed; conbtrols
'//						labeled as SelectButton, CancelButton
'//	7/5/20.		wmk.	Entry dependencies updated
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
'dim oListBox As com.sun.star.awt.UnoControlListBox
dim oListBox As Object

Dim Dlg As Object
dim iCOACount As Integer

'// code.
	ON ERROR GOTO ErrorHandler
	BasicLibraries.LoadLibrary("Standard")
	BasicLibraries.LoadLibrary("BadAss")

'// Initialize COA Dialog
	puoCOADialog = CreateUnoDialog(DialogLibraries.BadAss.COADialog)
	puoCOAListBox =puoCOADialog.getControl("COAList")
	puoCOASelectBtn = puoCOADialog.getControl("SelectBtn")
	puoCOASelectBtn.Model.Enabled = false			'// disable button until user selects item
	iCOACount = LoadCOAs()
	puoCOAListBox.addItems(pusCOAList, 0)

	Select Case puoCOADialog.Execute()
	Case 2		'// Select clicked
		msgBox("Select clicked in Chart of Accounts")
	Case 0		'// Cancel
		pusCOASelected = ""		'// since cancel, Execute() event skipped
		msgBox("Cancel or FormClose clicked")
	Case else
		msgBox("unevaluated dialog return")
	End Select
	msgBox("In Module2/F" + CHR(13)+CHR(10) _
    	    + "'"+pusCOASelected+"'" + " selected")

NormalExit:
	puoCOAListBox.dispose()
	exit sub
	
ErrorHandler:
	msgBox("RunCOADlg - unprocessed error.")
	GoTo NormalExit
	
end sub		'// end RunCOADlg	7/5/20
'/**/
