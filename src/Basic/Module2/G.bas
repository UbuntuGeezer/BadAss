'// G.bas
'//---------------------------------------------------------------
'// G - Enter Transaction Dialog test procedure.
'//		8/26/22.	wmk.	10:26	
'//----------------------------------------------------------------

public sub G()

'//	Usage.	macro call or
'//			call G()
'//
'//		Run ET dialog...
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
'//	6/15/20.	wmk.	original code.
'// 6/16/10.	wmk.	message boxes for verifification.
'//	6/17/20.	wmk.	add ETPubVars reset call on Cancel.
'//	6/25/20.	wmk.	gsETAmount corrected to new gdETAmount.
'//	6/27/20.	wmk.	msgBoxes commented to deactivate.
'// 8/26/22.	wmk.	eliminate references to G. bas.
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
'dim oListBox As com.sun.star.awt.UnoControlListBox
dim iStatus As Integer

'// code.
DialogLibraries.LoadLibrary("Standard")

'// Initialize ET dialog.
	puoETDialog = CreateUnoDialog(DialogLibraries.Standard.TransEntry)

	Select Case puoETDialog.Execute()
	Case 2		'// Record & Finish clicked
'		msgBox("Record & Finish clicked in Enter Transaction")
		'// iStatus = ETDialogRecord... called by [Record & Finish]
	Case 0		'// Cancel
'		pusCOASelected = ""		'// since cancel, Execute() event skipped
		iStatus = ETPubVarsReset(1)	'// reset everything
'		msgBox("Cancel or FormClose clicked")
	Case else
		msgBox("G- unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gdETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

puoETDialog.dispose()
	
end sub		'// end G	8/26/22
'/**/

