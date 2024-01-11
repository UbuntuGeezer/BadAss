'// ErrLogSetDisplay.bas
'//---------------------------------------------------------------
'// ErrLogSetDisplay - Set msgBox error messaging on/off.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public function ErrLogSetDisplay(pbDisplayOn As Boolean) As Void

'//	Usage.	ErrLogSetDisplay(bDisplayOn)
'//
'//		bDisplayOn = true to turn on msgBpx error messaging
'//					 false to turn off msgBox error messaging
'//
'// Entry.	gbErrDisplay = current msgBox error messaging status
'//
'//	Exit.	gbErrDisplay  = bDisplayOn
'//
'//	Modification history.
'//	---------------------
'//	5/26/20.	wmk.	original code
'//
'// Notes. gbErrDisplay flag is used by LogError to determine if
'//	error messages are displayed in a msgBox onscreen

	'// code.
	gbErrDisplay = pbDisplayOn

end function 	'// end ErrLogSetDisplay
'/**/

