'// ToDo.bas
'//---------------------------------------------------------------
'// ToDo - move user into ToDo sheet, if exists.
'//		16/26.	wmk.
'//---------------------------------------------------------------

public function ToDo() As String

'// Usage.	svar = ToDo()
'//
'// Calls. MoveToSheet.
'//
'// Exit. user placed in ToDo sheet if exists.
'//		svar = '=ToDo()
'//
'// Notes. This function places the user in the ToDo sheet and returns the
'// string "-=ToDo()". If called from a sheet cell, this changes the cell to
'// a comment (thus avoiding recursion if AutoCalc is on).
'//

'// code.
	ON ERROR GOTO ErrHandler
	MoveToSheet("ToDo")

NormalReturn:
	ToDo = "'=ToDo()"
	exit function
	
end function		'// end ToDo	1/6/26.
'/**/
