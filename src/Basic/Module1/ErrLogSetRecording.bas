'// ErrLogSetRecording.bas
'//---------------------------------------------------------------
'// ErrLogSetRecording - Set error recording status in error log globals
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public function ErrLogSetRecording(pbRecordingOn as Boolean) As Void

'//	Usage.	ErrLogSetRecording(bRecordingOn)
'//
'//		bRecordingOn = true to turn on log recording in ERRLOGSHEET
'//					   false to turn off log recording in ERRLOGSHEET
'//
'// Entry.	gbErrRecording = current ERRLOGSHEET recording status
'//
'//	Exit.	gbErrRecording = bRecordingOn
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. LogError uses the gbErrRecording flag to determine if Error Log
'// sheet (ERRLOGSHEET) entries are to be made.
'//	Eventually, this should issue a message to the Error Log
'// sheet reporting the recording on/off event.

	'// code.
	gbErrRecording = pbRecordingOn

end function 	'// end ErrLogSetRecording
'/**/

