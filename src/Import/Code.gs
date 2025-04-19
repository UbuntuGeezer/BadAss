// Code.gs - AppScript code converted from OpenOffice.Calc macros.
//	8/11/24.	wmk.

// DayName.bas
// DayName - Return text name of day from week day number.
function DayName(nDay){

//	Modification history.
//	---------------------
//	7/30/24.	wmk.	original for Apps Script.

var DayName = "Monday"
var strDay

//	Code.

switch(nDay){

case 1:
	strDay = "Sunday";
  break;
case 2:
	strDay ="Monday";
  break;
case 3:
	strDay = "Tuesday";
  break;
case 4:
	strDay = "Wednesday";
  break;
case 5:
	strDay = "Thursday";
  break;
case 6:
	strDay = "Friday";
  break;
case 7:
	strDay = "Saturday";
	break;
default:
	strDay = "??";
  break;
}	



return strDay
} // end Dayname
//**/
