'// ErrLogGetSheet.bas
'//---------------------------------------------------------------
'// ErrLogGetSheet - Get sheet index from error log globals.
'//		wmk. 4/20/25.
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
'// 4/20/25.	wmk.	end Function > end function.
'//	5/26/20		wmk.	original code.
'//
'//	Notes. Complemented by ErrLogSetSheet(iSheet)
'//

'//	local variables.

	'// code.
	
	ErrLogGetSheet = goErrRangeAddress.Sheet	'// return global sheet index
	
end function		'// end ErrLogGetSheet		4/20/25.
'/**/
