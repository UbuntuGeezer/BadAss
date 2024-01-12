'// STPubVarsReset.bas
'//---------------------------------------------------------------------
'// STPubVarsReset - Reset public variables for Split Transaction dialog.
'//		6/27/20.	wmk.	14:45
'//---------------------------------------------------------------------

public function STPubVarsReset() As Integer

'//	Usage.	iVar = STPubVarsReset()
'//
'// Entry.
'//		puoSTDialog	= ST dialog object
'//		gbSTEditMode = true if <Edit Split> in progress; false otherwise
'//
'// ST dialog stored values.
'public gdSTAccumTot		As Double		'// split accumulated total
'public gdSTTotalAmt		As Double		'// total amount to split
'public gbSTSplitCredits	As Boolean		'// credit values are splits flag
'public gsSTDescs(0)		As String		'// array of split descriptions
'public gdSTAmts(0)		As Double		'// array of amounts
'public gdSTCOAs(0)		As String		'// array of COAs
'//		giLastIX		As Integer		'// last active row index for input
'//		gbSTEditMode = true if <Edit Split> in progress; false otherwise
'//
'//	Exit.	iVar = 0 - normal return
'//				   -1 - error in reset
'//				   +1 - reset skipped; edit in progress
'//			public vars initialized ONLY if gbSTEditMode = false
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	giLastIx added
'//	6/20/20.	wmk.	gsSTRefs initialized
'//	6/27/20.	wmk.	gbSTEditMode flag tested before resetting vars
'//
'//	Notes.	When these publics are cleared, the dialog controls should
'// also be reset by a call to STDlgCtrlsReset so its values are correct.


'//	constants.

'//	local variables.
dim iRetValue As Integer

	'// code.
	
	iRetValue = -1		'// set bad return
	ON ERROR GOTO ErrorHandler
	
	'// only initialize if not dialog not being loaded for edit mode
	if gbSTEditMode then
		iRetValue = 1	'// set edit mode; initialization skipped
		GoTo NormalExit
	endif
	
	'// gdsTotalAmt = itself; this should have been set up by dialog
	'//   invoking the Split Transaction dialog
	gdSTAccumTot = 0.			'// clear accumulated total
	giSTLastIx = 0				'// last active row index
	gbSTSplitCredits = true		'// assume Debits as total
	gsSTAcctFocus = ""			'// clear account object focus
	gsSTObjFocus = ""			'// object name with focus
	redim gsSTDescs(0)	'// scrap descriptions array
	redim gdSTAmts(0)	'// scrap amounts array
	redim gsSTCOAs(0)	'// scrap COAs array
	redim gsSTRefs(0)	'// scrap refs array
	iRetValue = 0

NormalExit:
	STPubVarsReset = iRetValue
	exit function

ErrorHandler:
	msgBox("STPubVarsReset - unprocessed error.")
	GoTo NormalExit
	
end function 	'// end STPubVarsReset		6/27/20
'/**/
