'// ErrLogSetModule.bas
'//---------------------------------------------------------------
'// ErrLogSetModule - Set module name in error log globals.
'//		4/20/25.	wmk.
'//---------------------------------------------------------------

public Function ErrLogSetModule(psName as String) As Void

'//	Usage.	ErrLogSetModule(sName)
'//
'//			sName = new name to set in error module caller global
'//
'// Entry.	gsErrModule = current module name global
'//
'//	Exit.	gsErrModule = new error module caller name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'// 4/20/25.	wmk.	end Function > end function.
'//	5/26/20		wmk.	original code.
'//
'//	Notes. Complemented by ErrLogGetModule()
'//

'//	local variables.

	'// code.
	gsErrModule = psName	'// reset global module name
	
end function		'// end ErrLogSetModule		4/20/25.
'/**/
