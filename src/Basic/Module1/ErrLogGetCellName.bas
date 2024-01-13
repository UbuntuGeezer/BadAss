'// ErrLogGetCellName.bas
'//---------------------------------------------------------------
'// ErrLogGetCellName - Get cell name from error log globals.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public Function ErrLogGetCellName() As String

'//	Usage.	sCellName = ErrLogGetCellName()
'//
'// Entry.	goLogRange = current cell tracking
'//
'//	Exit.	sCellName = "XYY" where
'//				X = column letter converted from goLogRange.StartColumn
'//				Y = row number converted from goLogRange.StartRow
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code; stub
'//
'//	Notes. Complemented by ErrLogGetModule()
'//

'//	local variables
dim sRetStr As String		'// returned cell string

	'// code.
	sRetStr = ""
	ErrLogGetCellName = sRetStr	'// return current name
	
end Function		'// end ErrLogGetCellName
'/**/
