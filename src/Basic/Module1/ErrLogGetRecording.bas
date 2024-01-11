'// ErrLogGetRecording.bas
'//-----------------------------------------------------------------------
'// ErrLogGetRecording - Get error recording status from error log globals
'//		wmk. 5/26/20.
'//-----------------------------------------------------------------------

public function ErrLogGetRecording() As Boolean

'//	Usage.	bLogRecording = ErrLogGetRecording()
'//
'// Entry.	gbErrRecording = current log recording status
'//
'//	Exit.	bLogRecording = gbErrRecording
'//							true if log recording on
'//							false if log recording off
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. LogError uses the gbErrRecording flag to determine if Error Log
'// sheet (ERRLOGSHEET) entries are to be made.

'//	local variables.
dim bRetValue	as Boolean

	'// code.
	
	bRetValue = gbErrRecording
	ErrLogGetRecording = bRetValue

end function 	'// end ErrLogGetRecording
'/**/

