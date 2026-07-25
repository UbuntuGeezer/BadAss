'// DayName.bas
'// DayName - Return text name of day from week day number
public Function DayName(nDay as integer) as string

'//	Modification history.
'//	---------------------
'//	12/8/15.	wmk.	Added "??" return value if nDay type is unkown.

'DayName = "Monday"
dim strDay as string

'//	Code.

select case nDay

case 1
	strDay = "Sunday"
case 2
	strDay ="Monday"
case 3
	strDay = "Tuesday"
case 4
	strDay = "Wednesday"
case 5
	strDay = "Thursday"
case 6
	strDay = "Friday"
case 7
	strDay = "Saturday"
	
case else
	strDay = "??"
	
end select


DayName = strDay
end function	'// end Dayname
'/**/
