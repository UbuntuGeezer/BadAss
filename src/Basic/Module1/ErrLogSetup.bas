'// ErrLogSetup.bas
'//---------------------------------------------------------------
'// ErrLogSetup - set up error logging to ERRLOGSHEET.
'//		wmk. 5/26/20.
'//---------------------------------------------------------------

public sub ErrLogSetup(poCellRange as Object, psModule as String)

'//	Usage.	macro call or
'//			call ErrLogSetup( oCellRange, sModule)
'//
'//			oCellRange = .CellRangeAddress object with following
'//							.Sheet = sheet index
'//							.StartColumn = starting column of range
'//							.EndColumn = ending column of range
'//							.StartRow = starting row of range
'//							.EndRow = ending row of range
'//			sModule = (string) name of sub/function for logging
'//
'// Entry.	none.
'//
'//	Exit.	global variables set for ErrLogError
'//			gbErrLogging = true
'//			gsErrModule = sModule
'//			gsErrMsg = ""
'//			gsErrName = ""
'//			gsErrSheet = ""
'//			goErrRangeAddress = oCellRange object values
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	5/26/20		wmk.	original code
'//
'//	Notes. This sub should be the first call from any sub/function
'//	desiring debug/error logging. The gbErrLogging var should be
'// tested before calling the LogError sub to ensure that ErrLogSetup
'//	has been called and errors within the error logging will be avoided.
'// 

'//	constants.

'//	local variables.
Static bBeenSetup as boolean	'// very 1st time flag
dim oDoc	As Object		'// this component
dim oSheets As Object		'// Sheets() array
dim oCurrSheet as Object	'// current sheet

	'// code.
	if IsNull(poCellRange) then
		msgBox("ErrLogSetup - null pointer passed for cell range.")
		GoTo BailOut
	endif

	if IsNull(psModule) then
		msgBox("ErrLogSetup - null string passed for module name.")
		GoTo BailOut
	endif

	if NOT bBeenSetup then
		bBeenSetup = true
		goErrRangeAddress = new com.sun.star.table.CellRangeAddress
	endif
		
	'// new module creates fresh instance
	if StrComp(psModule, gsErrModule) <> 0 then
'XRay goErrRangeAddress	'// no .Release method
		goErrRangeAddress = new com.sun.star.table.CellRangeAddress
	endif 		'// end new module conditional

	oDoc = ThisComponent
	oCurrSheet = oDoc.Sheets(poCellRange.Sheet)		'// set up sheet object
	gsErrSheet = oCurrSheet.CodeName				'// set up sheet name
	goErrRangeAddress.Sheet = oCurrSheet.RangeAddress.Sheet
	goErrRangeAddress.StartColumn = poCellRange.StartColumn
	goErrRangeAddress.EndColumn = poCellRange.EndColumn
	goErrRangeAddress.StartRow = poCellRange.StartRow
	goErrRangeAddress.EndRow = poCellRange.EndRow
	gbErrDisplay = true							'// display loggin on
	gbErrRecording = false						'// log sheet recording on/off flag
	gbErrLogging = true							'// enable error logging
	gsErrMsg = ""
	gsErrName = ""
	gsErrModule = psModule						'// initialize module name
	
Bailout:

end sub		'// end ErrLogSetup
'/**/
