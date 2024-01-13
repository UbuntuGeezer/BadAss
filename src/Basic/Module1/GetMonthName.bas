'// GetMonthName.bas
'//---------------------------------------------------------------
'// GetMonthName - return month name given date number.
'//		wmk. 5/20/20.
'//---------------------------------------------------------------

function GetMonthName(psDate as String) as String

'//	Usage.	sMonth = GetMonthName( sDate )
'//
'//		sDate = date string "mm/dd/yy" or compatible format
'//
'//	Exit.	sMonth = appropriate month from list
'//			"January", "February", "March", "April", "May", "June"
'//			"July", "August", "September", "October", "November", "December"
'//
'//	Modification history.
'//	---------------------
'//	5/15/20		wmk.	original code
'//	5/20/20.	wmk.	vars compliant with EXPLICIT
'//
'//	Notes. <Insert notes here>
'//

'//	constants. (see module-wide constants)

'//	local variables.

dim nMonth as integer	'// number of month
dim sName as String		'// returned name of month

	'// code.

	sName = ""	'// default to empty string
	nMonth = Month(DateValue(psDate))
	
	Select Case nMonth
	Case 1
		sName = MON1
	Case 2
		sName = MON2
	Case 3
		sName = MON3
	Case 4
		sName = MON4
	Case 5
		sName = MON5
	Case 6
		sName = MON6
	Case 7
		sName = MON7
	Case 8
		sName = MON8
	Case 9
		sName = MON9
	Case 10
		sName = MON10
	Case 11
		sName = MON11
	Case 12
		sName = MON12

	end Select	'// end month number case
	
	GetMonthName = sName	'// set return value
	
'//--------------Debugging
'msgBox( "In GetMonthName, returning month: "+sName)
'//-----------end Debugging
	
end function 	'// end GetMonthName 
'/**/
