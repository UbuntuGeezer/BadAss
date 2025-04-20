'// ErrLogSetSheet.bas
'//---------------------------------------------------------------
'// ErrLogSetSheet - Set sheet index in error log globals.
'//		4/20/25.	wmk.
'//---------------------------------------------------------------

public Function ErrLogSetSheet(piSheet as Integer) As Void

'//	Usage.	ErrLogSetModule(iSheet)
'//
'//			iSheet = new sheet index to set in goErrRangeAddress.Sheet global
'//
'// Entry.	goErrRangeAddress.Sheet = current error focus sheet index
'//
'//	Exit.	goErrRangeAddress.Sheet = new error focus sheet index
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'// 4/20/25.	wmk.	end Function > end function.
'//	5/26/20		wmk.	original code.
'//
'//	Notes. Complemented by ErrLogGetSheet()
'//

'//	local variables.

	'// code.
	goErrRangeAddress.Sheet = piSheet	'// reset global sheet index
	
end function		'// end ErrLogSetSheet		4/20/25.
'/**/
