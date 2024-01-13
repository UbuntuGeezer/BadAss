'// ErrLogGetDisplay.bas
'//---------------------------------------------------------------
'// ErrLogGetDisplay - Get msgBox error messaging on/off status.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public function ErrLogGetDisplay() As Boolean

'//	Usage.	bMsgDisplay = ErrLogGetDisplay()
'//
'// Entry.	gbErrDisplay = current msgBox error messaging status
'//
'//	Exit.	bMsgDisplay = gbErrDisplay
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'// Notes. gbErrDisplay flag is used by LogError to determine if
'//	error messages are displayed in a msgBox onscreen

	'// code.
	bRetVal = gbErrDisplay
	ErrLogGetDisplay = bRetVal

end function 	'// end ErrLogGetDisplay
'/**/
