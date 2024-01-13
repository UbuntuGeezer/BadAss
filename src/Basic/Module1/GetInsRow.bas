'// GetInsRow.bas
'//------------------------------------------------------------------
'// GetInsRow - Get insertion row given sheet, COA account, month.
'//		wmk. 6/8/20. 15:00
'//------------------------------------------------------------------

function GetInsRow( poSheet, psAcctCat, psMonth) as long

'//	Usage.	lInsRow = GetInsRow( oSheet, sAcctCat, sMonth)
'//
'//		oSheet = Sheet object from Doc.Sheets[]
'//		sAcctCat = accounting COA to search for
'//		sMonth = month to search for within category
'//
'//	Exit.	lInsRow = >= 0; row index at which to insert new transaction
'//							returned index is where COA and month found + INSBIAS
'//					= -1 COA not found at all
'//					= -2 COA and month found, but not enough room to insert
'//					= -3 COA found, but month not found before COA changed
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	5/17/20		wmk.	original code; stub.
'//	5/18/20.	wmk.	code fleshed out
'//	5/20/20.	wmk.	vars compliant with EXPLICIT
'// 5/23/20.	wmk.	clarify calling sequence documentation
'// 6/8/20.		wmk.	mod to stop checking if have encountered a COA
'//						that is > COA looking for and flag not found

'//
'//	Notes. Does row-by-row search for line containing the COA account number
'// in Column A, and the month in ColumnB
'//

'//	constants. (see module-wide constants)

'//	local variables.
dim il			as Long		'// loop index
dim iBrkPt		as Integer	'// easy breakpoint marker
dim	j			as integer	'// short loop index
dim lRowIndex	as Long		'// returned row index
dim sAcctCat	as String	'// localized reference to psAcctCat
dim sMonth		as String	'// localized reference to psMonth
dim bCOAMatched	as Boolean	'// found COA flag
dim bNoMatch	as Boolean	'// this one not COA looking for flag

'// sheet information variables.
dim lStartRow	as Long		'// sheet starting row (should be 0)
dim lStartColumn	as Long		'// sheet starting column
dim lEndRow		as Long		'// sheet end row
dim lEndColumn		as Long		'// sheet end column
dim oRange		as Object	'// RangeAddress object
Static oCellDate	as Object
Static oCellTrans	as Object
dim sRowDate	as String	'// string from Date column
dim sRowTrans	as String	'// string from Transaction column
dim sScanDate	as String	'// Date field from current scan row
'dim oRow		as Object	'// row cells from search row index

	'// code.

	lRowIndex = -1		'// set not found
	sAcctCat = psAcctCat
	sMonth = psMonth
	oRange = poSheet.RangeAddress
	lStartRow = oRange.StartRow
	lStartColumn = oRange.StartColumn
	lEndRow = oRange.EndRow
	lEndColumn = oRange.EndColumn
	bCOAMatched = false		'// set COA not found yet	

iBrkpt = 1
'XRay poSheet
	
	'// for each row 0 to sheet last row
	for il = lStartRow to lEndRow

		'// search for month first to ease debugging, since there are so many
		'// less month matches...		

		'// check Transaction field for month match
		
		oCellTrans = poSheet.getCellByPosition(COLTRANS, il)		
		sRowTrans = oCellTrans.String

'// debugging
if Len(sRowTRans)<>Len(sMonth) then
iBrkPt=0
else
iBrkPt=1
endif
iBrkPt = StrComp(sRowTrans, sMonth)
'//---end debugging

		if StrComp(sRowTrans, sMonth) <> 0 then		'// if not month match, skip out
			GoTo AdvanceRow
		endif

iBrkPt=1	

		'// check Date field for COA account match
		oCellDate = poSheet.getCellByPosition(COLDATE, il)
		sRowDate = oCellDate.String
		if Len(sRowDate) <> COALEN then 	'// if not correct COA account length, skip out
			GoTo AdvanceRow
		endif
	
iBrkPt=1	
		bNoMatch = StrComp(sRowDate, sAcctCat) <> 0	
		if bNoMatch then	'// if not a match
'			if bCOAMatched then	'// but matched before, COA changed
'				lInsertRow = -3		'// no month match for account
'				exit For
'			else
'				GoTo AdvanceRow		'// keep looking for matching COA
'			endif	'// end found COA before condtional
			
			'// check to see if past COA looking for; if so, stop searching
			'// and flag COANOTOFUND
			if StrComp(sRowDate, sAcctCat) > 0 then					'// mod060820
				exit for											'// mod060820
			endif													'// mod060820
			
			
			GoTo AdvanceRow
		endif	'// end no match conditional

		bCOAMatched = true
		lRowIndex = il
		
		'// Check next rows to be within INSBIAS on this account number
		for j = 1 to INSBIAS
			oCellDate = poSheet.getCellByPosition(COLDATE, il+j)
			sScanDate = oCellDate.String
			if Len(sScanDate) = COALEN then		'// found another account
				lRowIndex = -2		'// regardless of account, not enough room for insert
				exit For
			endif
		next j		'// end look-ahead loop

		'// check if enough room to insert
		if lRowIndex >= 0 then
			lRowIndex = il + INSBIAS			'// have insertion row index
		endif
		exit For
		'//	if match exit loop with lRowIndex = index row + INSBIAS(=3)
		
AdvanceRow:	
	next il		'// end loop on sheet rows

	'// set return value
	GetInsRow = lRowIndex		

iBrkpt = 1	

end function 	'// end GetInsRow	6/8/20
'/**/
