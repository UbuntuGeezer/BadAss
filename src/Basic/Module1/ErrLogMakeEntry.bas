'// ErrLogMakeEntry.bas
'//---------------------------------------------------------------
'// ErrLogMakeEntry - record entry in ERRLOGSHEET spreadsheet.
'//		wmk. 6/6/20. 09:45
'//---------------------------------------------------------------

public function ErrLogMakeEntry() As Integer

'//	Usage.	iStatus = ErrLogMakeEntry()
'//
'// Entry.	error handling global values set
'//
'//			ERRLOGSHEET = (string) name of error log sheet
'//
'//	Exit.	iStatus = 0 if log entry made successfully
'//					< 0 if failed
'//
'// Calls.	SetSheetDate(oLogRange, FMTDATETIME)
'//
'//	Modification history.
'//	---------------------
'//	5/26/20.	wmk.	original code
'//	5/28/20.	wmk.	time stamp column added
'//	6/3/20.		wmk.	add date/time stamp to error log sheet heading.
'// 6/6/20.		wmk.	update SetSheetDate call with format spec
'//
'//	Notes. <Insert notes here>
'//	Sheet	Column	Row	Module	ERRCODE	Message

'//	constants.
const COLTIME=0			'// Time Stamp column 
const COLSHEET=1		'// Sheet column
const COLCOLUMN=2		'// Column column
const COLROW=3			'// Row column
const COLMODULE=4		'// Module column
const COLERRCODE=5		'// ERRCODE column
const COLERRMSG=6		'// Message column
const ERRINSROW=5		'// insertion row index
'const FMTDATETIME=50	'// numeric value of mm/dd/yyy hh:mm:ss spec

'//	local variables.
dim iBrkPt as Integer		'// easy breakpoint spotter
dim iRetValue As Integer	'// return value
dim iStatus As Integer		'// general status
dim oDoc as Object			'// ThisComponent
dim oRange as Object		'// range address of this document
dim oLogSheet as Object		'// sheet ERRLOGSHEET
dim oSheets as Object		'// Doc.Sheets()
dim iSheetIx as Integer		'// sheet index of ERRLOGSHEET
dim oSel as Object
dim sSheet As String		'// sheet name for ERRLOGSHEET
dim oCellSheet As Object	'// Sheet cell
dim oCellTime As Object		'// Time stamp cell
dim oCellColumn As Object	'// Column cell
dim oCellRow As Object		'// Row cell
dim oCellModule As Object	'// Module cell
dim oCellErrCode As Object	'// Error Code cell
dim oCellErrMsg As Object	'// Error Message cell
dim oRangeAdd As Object		'// range address object
dim oErrSheet As Object		'// sheet error occurred in

	'// code.
	
	iRetValue = -1	'// set initial failure
	ON ERROR GoTo BailOut
	oDoc = ThisComponent
	oSheets = oDoc.getSheets()			'// get sheet collection
	oSel = oDoc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet				'// get sheet index value
'	oGLSheet = Doc.Sheets(iSheetIx)		'// set up sheet object
'XRay oGLSheet		'// xray to get .sun. reference
'	sGLSheet = oGLSheet.CodeName			'// get this sheet name

'XRay oSheets	
	sSheet = ERRLOGSHEET
	oLogSheet = oSheets.getByName(sSheet)
	
	'// see if ERRLOGSHEET exists and set up insertion pointers for
	'// new row
	oRangeAdd = new com.sun.star.table.CellRangeAddress
	oRangeAdd.Sheet = oLogSheet.RangeAddress.Sheet
	oRangeAdd.StartColumn = COLTIME
	oRangeAdd.EndColumn = COLERRMSG
	oRangeAdd.StartRow = ERRINSROW
	oRangeAdd.EndRow = ERRINSROW
	
	'// insert new row into log sheet
	oLogSheet.insertCells(oRangeAdd,_
						com.sun.star.sheet.CellInsertMode.ROWS)
	
	'// set values in new row fields from global error variables
iBrkPt=1
	oCellTime = oLogSheet.getCellByPosition(COLTIME, ERRINSROW)		'// Time stamp cell
	oCellSheet = oLogSheet.getCellByPosition(COLSHEET, ERRINSROW)	'// Sheet cell
	oCellColumn = oLogSheet.getCellByPosition(COLCOLUMN, ERRINSROW)	'// Column cell
	oCellRow = oLogSheet.getCellByPosition(COLROW, ERRINSROW)	'// Row cell
	oCellModule = oLogSheet.getCellByPosition(COLMODULE, ERRINSROW)	'// Module cell
	oCellErrCode = oLogSheet.getCellByPosition(COLERRCODE, ERRINSROW)'// Error Code cell
	oCellErrMsg = oLogSheet.getCellByPosition(COLERRMSG, ERRINSROW)	'// Error Message cell
	oErrSheet = oSheets.getByIndex(goErrRangeAddress.Sheet)

	oCellTime.setValue(Now())					'// time stamp
	oCellTime.NumberFormat = FMTDATETIME		'// set number format
	oCellSheet.String = oErrSheet.CodeName		'// set sheet name
	oCellColumn.String = Str(goErrRangeAddress.StartColumn)
	oCellRow.String = Str(goErrRangeAddress.StartRow)
	oCellModule.String = gsErrModule
	oCellErrCode.String = gsErrName
	oCellErrMsg.String = gsErrMsg
	
	'// set sheet modification date in log sheet header
	iStatus = SetSheetDate(oRangeAdd, FMTDATETIME)
	
	'// set iRetValue = 0 for success

	iRetValue = 0		'// if make it here, success...

BailOut:

	if iRetValue < 0 then	'// if problem accessing ERRLOGSHEET
		msgBox("Log Sheet '"+ERRLOGSHEET+"' not found!")
	endif
	
	ErrLogMakeEntry = iRetValue

end function 	'// end ErrLogMakeEntry	6/3/20
'/**/
