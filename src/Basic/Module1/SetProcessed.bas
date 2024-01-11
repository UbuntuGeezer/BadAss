'// SetProcessed.bas
'//---------------------------------------------------------------
'// SetProcessed - Set double-entry transaction processed.
'//		6/1/20.	wmk.
'//---------------------------------------------------------------

public function SetProcessed(oTransRange as Object) as Integer

'//	Usage.	iVar = SetProcessed(oTransRange)
'//
'//		oTransRange = RangeAddress of double-entry transaction
'//
'// Entry.	<entry conditions>
'//
'//	Exit.	iVar = 0 if no error; -1 otherwise
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/1/20.	wmk.	original code; stub
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim iRetValue As Integer

	'// code.
	
	iRetValue = 0		'// set normal return
	
	SetProcessed = iRetValue

end function 	'// end SetProcessed  6/1/20
'/**/

