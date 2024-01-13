'// E.bas
'//------------------------------------------------------------------
'// E - stub function to test error handling capabilities.
'//		5/27/20.	wmk.
'//------------------------------------------------------------------

public function E() as void

'//	Usage.	=E()
'//
'// Entry.	public const DEVDEBUG set = true/false
'//			most tests coded here depend upon the =D() reference being
'//			in the same row as the data to be tested, since the called
'//			subs/functions will look at the current range selection.
'//			When the user presses <Enter> after completing the formula
'//			referencing D(), the range selection is the cell in which
'//			the formula occurs.
'//
'//	Exit.	varies according to Calls.
'//			if DEVDEBUG=false, exits without doing anything
'//
'//	Calls.	varies according to test needed.
'//			ErrLogSetup, LogError, ErrLogDisable
'/
'//	Modification history.
'//	---------------------
'//	5/25/20.	wmk.	original code
'//	5/27/20.	wmk.	DEVDEBUG constant used to activate/deactivate
'//
'//	Notes. If DEVDEBUG=true, D acts to test whatever functions/subs
'// its code is set up for. Usually the <Insert notes here>

'// local variables. (defined as needed)
dim Doc As Object			'// this component
dim oSel As Object			'// current selection
dim oRange As Object		'// CellRangeAddress object

	'// code.
	if NOT DEVDEBUG then
		Exit Function
	endif
	
	Doc = ThisComponent
	oSel = Doc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	Call ErrLogSetup(oRange, "Test sub E")
	Call LogError("TESTERROR", "Test sub E - error")
	Call ErrLogDisable
					
end function	'// end E
'/**/
