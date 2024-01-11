'// LogError.bas
'//---------------------------------------------------------------
'// LogError - Make error log entry to ERRLOGSHEET.
'//		wmk. 6/2/20.
'//---------------------------------------------------------------

public sub LogError(psErrName, psErrMsg)

'//	Usage.	macro call or
'//			call LogError(sErrName, sErrMsg)
'//
'//		sErrName = error code name (e.g. "ERRBADVALS")
'//		sErrMsg = error message string
'//
'// Entry.	gbErrDisplay = true if msgBox to be displayed
'//						   false otherwise
'//			gbErrRecording = true if to write ERRLOGSHEET entry
'//							 false otherwise
'//
'//	Exit.	gsErrName = psErrName
'//			gsErrMsg = psErrMsg
'//
'// Calls.	ErrLogGetCellInfo, ErrLogGetModule, ErrLogMakeEntry
'//
'//	Modification history.
'//	---------------------
'//	5/26/20.	wmk.	original code; issue message in box; set
'//						up call to ErrMakeEntry
'//	6/2/20.		wmk.	added caller module name to msgBox	
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim lErrCol As Long			'// error column index
dim lErrRow As Long			'// error row index
dim lErrSheet as Integer	'// error sheet index
dim sErrCol As String	'// string conversions
dim sErrRow As String
dim sErrSheet As String
dim iStatus As Integer		'// recording status

	'// code.
	gsErrName = psErrName
	gsErrMsg =  psErrMsg
	ErrLogGetCellInfo(lErrCol, lErrRow)		'// extract column and row information
	lErrSheet = ErrLogGetSheet()
	sErrSheet = Str(lErrSheet)
	sErrCol = Str(lErrCol)
	sErrRow = Str(lErrRow)	

	if gbErrDisplay then
		msgBox("In LogError, Called by '" + ErrLogGetModule() + CHR(13)+CHR(10)_
			+ "Error Name: "+psErrName+CHR(13)+CHR(10)_
			+ "Sheet index: "+sErrSheet+CHR(13)+CHR(10)_
			+ "At cell index: "+sErrCol+":"+sErrRow+CHR(13)+CHR(10)_
			+psErrMsg)
	endif	'// end display message conditional
	
	if gbErrRecording then
		iStatus = ErrLogMakeEntry()
	endif	'// end log recording active conditional
	
			
end sub		'// end LogError	6/2/20
'/**/

