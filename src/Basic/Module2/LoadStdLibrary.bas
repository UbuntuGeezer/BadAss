'// LoadStdLibrary.bas
'//---------------------------------------------------------------
'// LoadStdLibrary - Load Standard Dialog library.
'//		7/5/20.	wmk.	22:00
'//---------------------------------------------------------------

public function LoadStdLibrary() As Integer

'//	Usage.	iVal = LoadStdLibrary()
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
'//	6/14/20.	wmk.	original code
'//	7/5/20.		wmk.	bug fix; LoadStdLib return corrected to LoadStdLibrary
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim iRetValue As Integer

	'// code.
	
	iRetValue = -1		'// set abnormal return
	ON ERROR GOTO ErrorHandler
	
	if NOT pbDlgLibLoaded then
		DialogLibraries.LoadLibrary("Standard")
		pbDlgLibLoaded = true
	endif	'// end library not loaded conditional

	iRetValue = 0		'// set normal return
	
NormalExit:
	LoadStdLibrary = iRetValue
	exit function
	
ErrorHandler:
	GoTo NormalExit

end function 	'// end LoadStdLibrary	7/5/20
'/**/
