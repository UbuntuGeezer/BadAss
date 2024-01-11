'// Module1Hdr.bas
REM  *****  BASIC  *****
'// Module1Hdr - BadAss Library Module1 header.
'// Module1 (Beta Release) - Local macros for Chase Activity Download Spreadsheet and
'//	transaction processing in Production Alpha Clean Accounting spreadsheet.
'// 8/22/22.	wmk.	'/**/ and // <module-name>.bas bracketing lines added; application
'//				 extended to general banking activity in Accounting subsystem.
'// Legacy Mods.
'//	5/14/20.	wmk.	original code with mods
'//	5/16/20.	wmk.	new functions and subs to span sheets for
'//						transaction mapping to accounts sheets
'//	5/17/20.	wmk.	correct test constant EXPSHEET to match this
'//						document 'Expense Accounts_2'
'//	5/18/20.	wmk.	row data processing constants added.
'//	5/19/20.	wmk.	sheet name constants changed to match production
'//	5/20/20.	wmk.	added OPTION EXPLICIT for code control
'//	5/22/20.	wmk.	added PlaceSplitTrans function.
'//	5/23/20.	wmk.	added BeanField function; bug fix with ASTSHEET
'//	5/25/20.	wmk.	added errhandling.bas and associated subs
'//	5/26/20.	wmk.	errhandling.bas update resinserted
'//	6/28/20.	wmk.	Beta Release to BadAss Library

'// Code protection.
OPTION EXPLICIT				'// every variable must be declared before use

'/**/

