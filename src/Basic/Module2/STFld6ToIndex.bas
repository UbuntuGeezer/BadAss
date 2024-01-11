'// STFld6ToIndex.bas
'//---------------------------------------------------------------------
'// STFld6ToIndex - Convert field with 6 char base name to array index.
'//		6/19/20.	wmk.
'//---------------------------------------------------------------------

public function STFld6ToIndex(psFldName As String) As Integer

'//	Usage.	iVar = STFld6ToIndex( sFldName )
'//
'//		sFldName = name of field (e.g. "AmtFld1")
'//
'// Entry.	gdSTAmtFldsn() array contains amount values entered
'//
'//	Exit.	iVar = 0-based index into arrays corresponding to
'//					<object-name>n
'//					subtracts 1 from n to obtain index
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue As Integer	'// returned index value
dim iNameLen	As Integer	'// passed name length
dim sFldName	As String	'// passed field name
dim iDigCount	As Integer	'// digit count in name
dim sAcctIx		As String	'// extracted end digit(s)
dim iArrayIx	As Integer	'// converted array index

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
'	sFldName = psFldName
	iNameLen = Len(psFldName)		'// get length, either 7 or 8
	Select Case iNameLen
	Case 7
		iDigCount = 1
	Case 8
		iDigCount = 2
	Case else
		iDigCount = 1
	End Select
	
	sAcctIx = Right(psFldName,iDigCount)
	iArrayIx = Val(sAcctIx) - 1
	if iArrayIx < 0 then
		GoTo ErrorHandler
	endif	'//end negative index calculated conditional

	iRetValue = iArrayIx
	
NormalExit:
	STFld6ToIndex = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STFld6ToIndex - unprocessed error converting '" + psFldName _
			+ "' to index")
	GoTo NormalExit
	
end function 	'// end STFld6ToIndex	6/19/20
'/**/
