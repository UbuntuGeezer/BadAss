'// LoadBadAssLib.bas
'//---------------------------------------------------------------
'// LoadBadAssLib - Load BadAss Dialog library.
'//		7/6/20.	wmk.	12:45
'//---------------------------------------------------------------

public function LoadBadAssLib() As Integer

'//	Usage.	iVal = LoadBadAssLib()
'//
'// Entry.	pbDlgLibLoaded = true if Standard dialogs already loaded
'//							 false otherwisd
'//
'//	Exit.	iVal = 0 - normal return
'//				   -1 - error encountered in load
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/30/20.		wmk.	original code
'//	7/6/20.		wmk.	BadAss load unconditional; change to BasicLibraries
'//
'//	Notes. Adapted from LoadStdLibrary to use My Macros library BadAss.
'//

'//	constants.

'//	local variables.
dim iRetValue As Integer

	'// code.
	
	iRetValue = -1		'// set abnormal return
	ON ERROR GOTO ErrorHandler
	
'	if NOT pbDlgLibLoaded then
		GlobalScope.BasicLibraries.LoadLibrary("BadAss")
		pbDlgLibLoaded = true
'	endif	'// end library not loaded conditional

	iRetValue = 0		'// set normal return
	
NormalExit:
	LoadBadAssLib = iRetValue
	exit function
	
ErrorHandler:
	GoTo NormalExit

end function 	'// end LoadBadAssLib	7/6/20
'/**/
