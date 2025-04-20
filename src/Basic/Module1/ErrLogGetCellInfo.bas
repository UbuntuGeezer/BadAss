'// ErrLogGetCellInfo.bas
'//-----------------------------------------------------------------
'// ErrLogGetCellInfo - Get cell information from error log globals.
'//		wmk. 4/20/25.
'//-----------------------------------------------------------------

public Function ErrLogGetCellInfo(rplColumn As Long, rplRow As Long) As Void

'//	Usage.	ErrLogGetCellInfo(lColumn, lRow)
'//
'// Entry.	goErrLogRange = current cell tracking
'//					.StartColumn = starting column index
'//					.StartRow = starting row index
'//
'//	Exit.	lColumn = goErrLogRange.StartColumn
'//			lRow = goErrLogRange.StartRow
'//
'//	Modification history.
'//	---------------------
'// 4/20/25.	wmk.	end Function > end function.
'//	5/26/20		wmk.	original code
'//
'//	Notes. Complemented by ErrLogSetInfo(lColumn, lRow)
'//

'//	local variables

	'// code.
	rplColumn = goErrRangeAddress.StartColumn
	rplRow = goErrRangeAddress.StartRow
	
end function		'// end ErrLogGetCellnfo	4/20/25.
'/**/
