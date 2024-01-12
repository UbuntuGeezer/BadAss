'// GA.bas
'//---------------------------------------------------------------
'// GA - Store user selected COA in first cell in selected range
'// (Shortcut for GetCOA(1)
'//		7/5/20.	wmk.	22:00
'//---------------------------------------------------------------

public sub GA()

'//	Usage.	macro call or
'//			call GA()
'//
'// Entry.	pbCOADlgExists = true if COA dialog previously initialized
'//							 false if initialized here
'//			BadAss library loaded (automatic event on document open)
'//
'//	Exit.	if user did not cancel dialog, user selected COA stored at
'//			first cell in currently selected range
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	7/5/20.		wmk.	original code; adapted from GetCOA
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim sRetValue As String		'// returned COA
dim iStatus As Integer		'// general status
dim iCOACount As Integer	'// COA list length
dim	oDoc		As Object	'// ThisComponent
dim oSel		As Object	'// CellRangeAddress; current selection
dim oAcct		As Object	'// COA cell user selected
dim lColumn		As Long		'// selected column
dim lRow		As Long		'// selected row
dim oSheet		As Object	'// selected sheet
dim iSheetIx	As Integer	'/// sheet index

	'// code.
	ON ERROR GOTO ErrorHandler
	oDoc = ThisComponent
	oSel = oDoc.getCurrentSelection()
	iSheetIx = oSel.RangeAddress.Sheet
	oSheet = oDoc.Sheets.getByIndex(iSheetIx)
	lColumn = oSel.RangeAddress.StartColumn
	lRow = oSel.RangeAddress.StartRow
			
'		iStatus = LoadStdLibrary()		'// ensure dialogs
'DialogLibraries.LoadLibrary("Standard")
'		iStatus = LoadBadAssLib()		'// ensure dialogs
'
'		if iStatus < 0 then
'			GoTo ErrorHandler
'		endif

	'// create dialog if not yet done
	if NOT pbCOADlgExists then
		puoCOADialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.COADialog)
		pbCOADlgExists = true
	endif	'// end dialog non-existent conditional
		'// load list with COA strings
	puoCOAListBox = puoCOADialog.getControl("COAList")
	puoCOASelectBtn = puoCOADialog.getControl("SelectBtn")
	puoCOASelectBtn.Model.Enabled = false			'// disable button until user selects item
	iCOACount = LoadCOAs()
	puoCOAListBox.addItems(pusCOAList, 0)

	'// run dialog and get user selection
	Select Case puoCOADialog.Execute()
	Case 2		'// Select clicked
'		msgBox("COA Selected: " + pusCOASelected)
'			msgBox("Select clicked in Chart of Accounts")
		'// store COA in first cell of oSel range
		oAcct = oSheet.getCellByPosition(lColumn, lRow)
		oAcct.String = Left(pusCOASelected,4)
		oAcct.HoriJustify = CJUST
		
	Case 0		'// Cancel
		pusCOASelected = ""		'// since cancel, Execute() event skipped
'			msgBox("Cancel or FormClose clicked")
	Case else
		msgBox("unevaluated COA dialog return")
	End Select

		
NormalExit:

'msgBox("In GA" + CHR(13)+CHR(10) _
'        + "'"+sRetValue+"'" + " selected")

	exit sub
	
ErrorHandler:
	msgBox("In GA - unprocessed error.")
	GoTo NormalExit
	
end sub		'// end GA		7/5/20
'/**/
