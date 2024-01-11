'// ErrLogGetSheet.bas
'//---------------------------------------------------------------
'// ErrLogGetSheet - Get sheet index from error log globals.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public Function ErrLogGetSheet() As Integer

'//	Usage.	iSheetIx = ErrLogGetSheet()
'//
'//			iSheet = new sheet index to set in goErrRangeAddress.Sheet global
'//
'// Entry.	goErrRangeAddress.Sheet = current error focus sheet index
'//
'//	Exit.	iSheetIx = goErrRangeAddress.Sheet = new error focus sheet index
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. Complemented by ErrLogSetSheet(iSheet)
'//

'//	local variables.

	'// code.
	
	ErrLogGetSheet = goErrRangeAddress.Sheet	'// return global sheet index
	
end Function		'// end ErrLogGetSheet
'/**/

