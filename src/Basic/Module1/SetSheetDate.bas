'// SetSheetDate.bas
'//---------------------------------------------------------------
'// SetSheetDate - Set date/time stamp in common sheet date field.
'//		6/8/20.	wmk.
'//---------------------------------------------------------------

public function SetSheetDate(poRange As Object, piFmt As Integer) As Integer

'//	Usage.	iVar = SetSheetDate(oRange, iFmt)
'//
'//		oRange = RangeAddress of sheet to set date/time stamp in
'//		iFmt = NumberFormat for date/time stamp {FMTDATETIME, MMDDYY}
'//
'// Entry.	system clock running
'//
'//	Exit.	Sheet at Sheet index within oRange has Now() date/time
'//			stamp set in Cell column=COLDATEMOD, row=ROWDATEMOD
'//			iVar = 0 - no error
'//				   -1 - untrapped error occurred; date/time stamp
'//						setting probably failed
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/3/20.		wmk.	original code
'//	6/6/20.		wmk.	date/time format parameter added	
'//	6/8/20.		wmk.	add LTGRN background color to new date
'//	Notes.


'//	constants.
const COLDATEMOD=2		'// corresponds to "C"
const ROWDATEMOD=1		'// corresponds to "2" for cell C2
'const FMTDATETIME=50	'// numeric value of mm/dd/yyy hh:mm:ss spec

'//	local variables.
dim iRetValue As Integer
dim oDoc As Object		'// ThisComponent
dim oSheets As Object	'// Doc.getSheets()
dim oSheet As Object	'// target sheet
dim iSheetIx As Integer	'// target sheet index
dim oCellTime as Object	'// date/time modified cell
dim iDateFormat As Integer	'// date/time format to set

	'// code.
	
	iRetValue = 0
	ON ERROR GoTo Abnormal
	iSheetIx = poRange.Sheet	'// extract sheet index from passed range
	oDoc = ThisComponent		'// access target sheet using range information
	oSheets = oDoc.getSheets()
	oSheet = oSheets.getByIndex(iSheetIx)
	oCellTime = oSheet.getCellByPosition(COLDATEMOD, ROWDATEMOD)
	
	oCellTime.setValue(Now())					'// time stamp
	
	Select Case piFmt
	CASE FMTDATETIME
		iDateFormat = FMTDATETIME
		
	CASE MMDDYY
		iDateFormat = MMDDYY
		
	CASE else
		oDateFormat = MMDDYY
		
	End Select
	
	oCellTime.NumberFormat = iDateFormat		'// set number format
	oCellTime.CellBackColor = LTGREEN			'// set changed color
	GoTo NormalExit
	
Abnormal:
	iRetValue = -1		'// set abnormal return code
	
NormalExit:
	SetSheetDate = iRetValue

end function 	'// end SetSheetDate	6/8/20
'/**/

