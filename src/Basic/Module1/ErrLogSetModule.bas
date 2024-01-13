'// ErrLogSetModule.bas
'//---------------------------------------------------------------
'// ErrLogSetModule - Set module name in error log globals.
'//		wmk. 5/26/20.
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
'//	5/26/20		wmk.	original code
'//
'//	Notes. Complemented by ErrLogGetModule()
'//

'//	local variables.

	'// code.
	gsErrModule = psName	'// reset global module name
	
end Function		'// end ErrLogSetModule
'/**/
