'// GetCOA.bas
'//---------------------------------------------------------------
'// GetCOA - Return user selected COA.
'//		7/6/20.	wmk.	11:00
'//---------------------------------------------------------------

public function GetCOA(piFromOpt As Integer) As String

'//	Usage.	sVar = GetCOA(iFromOpt)
'//
'//		iFromPrompt = USELAST (0) - use COA from last user selection
'//					  SELECTNEW (1) - use dialog to have user get COA
'//
'// Entry.	pusCOASelected = COA full string from last dialog
'//			pbCOADlgExists = true if COA dialog previously initialized
'//							 false if initialized here
'//
'//	Exit.	sVar = first 4 chars of COA user selected;
'//				   if USELAST option, from pusCOASelected on entry
'//				   if SELECTNEW option, from Chart of Accounts dialog
'//
'// Calls.	[LoadStdLibrary] LoadBadAssLib
'//
'//	Modification history.
'//	---------------------
'//	6/14/20.	wmk.	original code
'//	6/16/20.	wmk.	update to use new COADialog (formerly ListBox);
'//	6/17/20.	wmk.	correct load dialog code to initialize
'//						puoCOASelectBtn; debug msgBox calls removed'//
'//	7/5/20.		wmk.	change to load BadAss library
'//	7/6/20.		wmk.	make dialog reload unconditional
'//
'//	Notes. <Insert notes here>
'//

'//	constants.
const USELAST=0		'// use last COA selected
const SELECTNEW=1	'// select new COA from dialog

'//	local variables.
dim sRetValue As String		'// returned COA
dim iStatus As Integer		'// general status
dim iCOACount As Integer	'// COA list length
static sLastCOA As String	'// last COA selected

	'// code.
	
	sRetValue = ""
	iStatus = 0
	if piFromOpt = USELAST then
'		sRetValue = Left(pusCOASelected, 4)
		sRetValue = sLastCOA
		GoTo NormalExit
	else
		iStatus = LoadStdLibrary()		'// ensure dialogs
'DialogLibraries.LoadLibrary("Standard")
		iStatus = LoadBadAssLib()		'// ensure dialogs

		if iStatus < 0 then
			GoTo ErrorHandler
		endif

		'// create dialog if not yet done
'		if NOT pbCOADlgExists then
			puoCOADialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.COADialog)
			pbCOADlgExists = true
			'// load list with COA strings
			puoCOAListBox = puoCOADialog.getControl("COAList")
			puoCOASelectBtn = puoCOADialog.getControl("SelectBtn")
			puoCOASelectBtn.Model.Enabled = false			'// disable button until user selects item
			iCOACount = LoadCOAs()
			puoCOAListBox.addItems(pusCOAList, 0)
'		endif	'// end dialog non-existent conditional

		'// run dialog and get user selection
		Select Case puoCOADialog.Execute()
		Case 2		'// Select clicked
'			msgBox("COA Selected: " + pusCOASelected)
'			msgBox("Select clicked in Chart of Accounts")
			sLastCOA = pusCOASelected
		Case 0		'// Cancel
			pusCOASelected = ""		'// since cancel, Execute() event skipped
'			msgBox("Cancel or FormClose clicked")
		Case else
			msgBox("unevaluated COA dialog return")
		End Select

		sRetValue = Left(pusCOASelected, 4)
		
	endif	'// end use last COA conditional
	
NormalExit:

'msgBox("In Module2/GetCOA" + CHR(13)+CHR(10) _
'        + "'"+sRetValue+"'" + " selected")

	GetCOA = sRetValue
	exit function
	
ErrorHandler:
	sRetValue = ""
	GoTo NormalExit
	
end function 	'// end GetCOA		7/6/20
'/**/
