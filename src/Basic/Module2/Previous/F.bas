'// F.bas
'//---------------------------------------------------------------
'// F - First attempt at dialog box processing. (COADialog)
'//		6/15/20.	wmk.	17:00
'//---------------------------------------------------------------

public sub F()

'//	Usage.	macro call or
'//			call F()
'//
'//		<parameters description>
'//
'// Entry.	Button click in dialog box		
'//			puoCOADialog = declared object for COA dialog
'//			puoCOAListBox = list box object
'//			puoSelectBtn = declared object for <Select> button
'//			pusCOASelected = reserved for COA string selected
'//			current Component has sheet "Chart of Accounts" with COA table
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
	DialogLibraries.LoadLibrary("Standard")

'// Initialize COA Dialog
	puoCOADialog = CreateUnoDialog(DialogLibraries.Standard.COADialog)
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

puoCOAListBox.dispose()
	
end sub		'// end F	6/15/20
'/**/
