'// ErrLogDisable.bas
'//---------------------------------------------------------------
'// ErrLogDisable - Disable error logging.
'//		wmk. 5/26/20.
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
'//	5/25/20.	wmk.	original code
'// 5/26/20.	wmk.	change from sub to function

'//	constants.

'//	local variables.

	'// code.
	'// if not Basic, add code to release goErrRangeAddress conditionally
	
	gbErrLogging = false
	gsErrMsg = ""
	gsErrName = ""
	gsErrModule = ""
	gsErrSheet = ""
	
end Function		'// end ErrLogDisable
'/**/

