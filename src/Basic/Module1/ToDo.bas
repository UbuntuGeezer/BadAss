'// ToDo.bas
'//---------------------------------------------------------------
'// ToDo - move user into ToDo sheet, if exists.
'//		7/4/26.	wmk.
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

'// local variables.
dim nlineNo     As Integer      '// line number for error tracking

'// code.
	ON ERROR GOTO ErrHandler
    nLineNo = 26
    MoveToSheet("ToDo")

NormalReturn:
	ToDo = "'=ToDo()"
	exit function

ErrHandler:
    msgbox ("ToDo - (line " + nLineNo + " unprocessed error.")
end function		'// end ToDo	7/24/26.
'/**/
