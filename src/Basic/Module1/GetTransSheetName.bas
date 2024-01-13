'// GetTransSheetName.bas
'//---------------------------------------------------------------
'// GetTransSheetName - Get sheet name where transaction belongs.
'//		wmk. 5/20/20.
'//---------------------------------------------------------------

function GetTransSheetName( sAcct as String ) as String

'//	Usage.	sSheetName = GetTransSheetName( sAcct )
'//
'//		sAcct = account number for transaction
'//				1xxx - Asset account number
'//				2xxx - Liability account number
'//				3xxx - Income account number
'//				4xxx - Expenses account number
'//				5xxx - Expenses account number	(non-deductible)
'//				9xxx - Expenses account number	(uncategorized)
'//
'//	Exit.	sSheetName = <target> = <result>
'//
'//	Modification history.
'//	---------------------
'//	5/15/20		wmk.	original code
'//	5/20/20.	wmk.	var dims to comply with EXPLICIT
'//
'//	Notes. <Insert notes here>
'//

'//	constants. (see module-wide constants)

'//	local variables.

dim sName	as String	'// returned string
dim sDigit 	as String	'// extracted 1st digit

	'// code.
	
	sName = ""					'// default is empty string
	sDigit = Left(sAcct, 1)		'// trigger on leftmost digit
	Select Case sDigit
	
		Case "1"
			sName = ASTSHEET
		Case "2"
			sName = LIASHEET
		Case "3"
			sName = INCSHEET
		Case "4"
			sName = EXPSHEET
		Case "5"
			sName = EXPSHEET
		Case "9"
			sName = EXPSHEET
			
	end Select	'// end 1st digit case statement

'// Note: the Switch statement must account for all values of sDigit or else
'// a runtime error occurs... when "6" was passed, function failed.
'	sName = Switch( sDigit="1",ASTSHEET, sDigit="2",LIASHEET, sDigit="3",INCSHEET,_
'				sDigit="4",EXPSHEET, sDigit="5",EXPSHEET, sDigit="9",EXPSHEET)		

	GetTransSheetName = sName					'// return name string

'//--------------Debugging
'msgBox( "In GetTransSheetName, returning sheet name: "+sName)
'//-----------end Debugging

end function 	'// end GetTransSheetName
'/**/
