'// D.bas
'//------------------------------------------------------------------
'// D - stub function to enable calls to development subs/functions.
'//		5/27/20.	wmk.
'//------------------------------------------------------------------

public function D() as void

'//	Usage.	=D()
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
'//			PlaceSplitTrans
'/
'//	Modification history.
'//	---------------------
'//	5/13/20		wmk.	original code
'//	5/16/20.	wmk.	documentation; test calls to GetTransSheetName
'//	5/22/20.	wmk.	test call to PlaceSplitTrans
'//	5/27/20.	wmk.	DEVDEBUG constant used to activate/deactivate
'//
'//	Notes. If DEVDEBUG=true, D acts to test whatever functions/subs
'// its code is set up for. Usually the <Insert notes here>
'//

'// local variables. (defined as needed)
dim Doc as object		'// current component
dim oSel as object		'// current selection
dim GLSheet as object	'// current sheet selected
dim oRange as object	'// current selection within sheet
dim oSheets As Object		'// sheets array of objects
Dim oGLSheet As Object		'// GL sheet object
dim iSheetIx as Integer		'// sheet index
dim sGLSheet as String		'// GL sheet name 	

dim	lGLStartRow	as long
dim	lGLEndRow	as long
dim	lGLRowCount	as long
dim	lGLNewRow	as long
dim	lGLCurrRow	as long
dim	lNRowsSelected	as long
dim	lTransCount	as long
dim	bOddRowCount	as boolean

dim iBrkpt as integer
dim lSplitEnd	as long		'// end row of split transaction
dim oCellDate	as Object

	'// code.
	'// NEVER REMOVE THIS CONDITIONAL.
	if NOT DEVDEBUG then
		Exit Function
	endif
	
	Doc = ThisComponent
	oSel = Doc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet				'// get sheet index value
	oGLSheet = Doc.Sheets(iSheetIx)		'// set up sheet object
'XRay oGLSheet		'// xray to get .sun. reference
	sGLSheet = oGLSheet.CodeName			'// get this sheet name

iBrkpt=1	

	'// set row bounds from selection information
	lGLStartRow = oRange.StartRow
	lGLEndRow = oRange.EndRow
	lGLRowCount = lGLEndRow-lGLStartRow
	lGLNewRow = lGLStartRow - 2			'// kludge new row for loop
	lGLCurrRow = lGLStartRow - 2		'// current GL processing row preset for increment
	lNRowsSelected = lGLEndRow+1 - lGLStartRow
	lTransCount = lNRowsSelected/2	'// get transaction count (2 rows per)
	bOddRowCount = lTransCount*2 <> lNRowsSelected

	lSplitEnd = PlaceSplitTrans(oGLSheet, lGLStartRow)
	
	if lSplitEnd < 0 then
		oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLStartRow)
		oCellDate.Text.CellBackColor = YELLOW
		msgBox("Split transaction - bad transaction; check data")
	endif
	
'	lSplitEnd = PlaceSplitTrans(oGLSheet, lGLStartRow)	'// check static var
	
iBrkPt=1
				
end function	'// end D
'/**/
