'// ErrLogDisable.bas
'//---------------------------------------------------------------
'// ErrLogDisable - Disable error logging.
'//		4/20/25.	wmk.
'//---------------------------------------------------------------

public Function ErrLogDisable() As Void

'//	Usage.	macro call or
'//			call ErrLogDisable
'//
'// Entry.	<entry conditions>
'//
'//	Exit.	gbErrLogging = false
'//			goErrRangeAddress object released
'//			gsErrMsg = ""
'//			gsErrName = ""
'//			gsErrModule = ""
'//			gsErrSheet = ""
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'// 4/20/25.	wmk.	end Function > end function.
'// 5/26/20.	wmk.	change from sub to function.
'//	5/25/20.	wmk.	original code.

'//	constants.

'//	local variables.

	'// code.
	'// if not Basic, add code to release goErrRangeAddress conditionally
	
	gbErrLogging = false
	gsErrMsg = ""
	gsErrName = ""
	gsErrModule = ""
	gsErrSheet = ""
	
end function		'// end ErrLogDisable	4/20/25.
'/**/
