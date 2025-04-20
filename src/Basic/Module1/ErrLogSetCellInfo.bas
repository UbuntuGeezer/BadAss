'// ErrLogSetCellInfo.bas
'//-----------------------------------------------------------------
'// ErrLogSetCellInfo - Set cell information in error log globals.
'//		4/20/25.	wmk.
'//-----------------------------------------------------------------

public Function ErrLogSetCellInfo(plColumn As Long, plRow As Long) As Void

'//	Usage.	ErrLogSetCellInfo(lColumn, lRow)
'//
'// Entry.	goLogRange = current cell tracking
'//					.StartColumn = starting column index
'//					.StartRow = starting row index
'//
'//	Exit.	lColumn = goLogRange.StartColumn
'//			lRow = goLogRange.StartRow
'//
'//	Modification history.
'//	---------------------
'// 4/20/25.	wmk.	end Function > end function.
'//	5/26/20		wmk.	original code.
'//
'//	Notes. Complemented by ErrLogGetInfo(lColumn, lRow)
'//

'//	local variables

	'// code.
	goErrRangeAddress.StartColumn = plColumn
	goErrRangeAddress.StartRow = plRow
	
end function		'// end ErrLogSetCellnfo	4/20/25.
'/**/
