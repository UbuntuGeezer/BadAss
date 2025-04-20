'// ErrLogGetModule.bas
'//---------------------------------------------------------------
'// ErrLogGetModule - Get module name in error log globals.
'//		wmk. 4/20/25.
'//---------------------------------------------------------------

public Function ErrLogGetModule() As String

'//	Usage.	sCurrModule = ErrLogGetModule()
'//
'// Entry.	gsErrModule = current module name global
'//
'//	Exit.	sCurrModule = gsErrModule
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
	ErrLogGetModule = gsErrModule		'// return current name
	
end function		'// end ErrLogGetModule		4/20/25.
'/**/
