'// ETCOA1ListListener.bas
'//--------------------------------------------------------------------
'// ETCOA1ListListener - Process user click on <...> with COA1 account.
'//		6/17/20.	wmk.	07:45
'//--------------------------------------------------------------------

public sub ETCOA1ListListener()

'//	Usage.	macro call or
'//			call ETCOA1ListListener()
'//
'// Entry.	User clicked on <...> next to COA1 field
'//			<DebitField> in ET dialog is text field for result
'//
'//	Exit.	If user selected a COA from the COADialog then
'//			gsETAcct1 = the COA # selected
'//			DebitField.Text = the COA # selected
'//			otherwise the above remain unchanged
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code
'//	6/17/20.	wmk.	bug fix setting selection in text box
'//
'//	Notes. If the user does not select a new COA from GetCOA, the
'//	current COA for the Debit field (gsETAcct1) will remain unchanged.
'//

'//	constants.

'//	local variables.
dim oETCOA1Text	As Object			'// account text field
dim sAcct1		As String			'// returned account from GetCOA

	'// code.
	sAcct1 = ""
	ON ERROR GOTO ErrorHandler
	
	pbETDebitList=true
	pbETCreditList=false	
	sAcct1 = GetCOA(1)			'// get COA from dialog
	if Len(sAcct1)>0 then
	    puoETDialog.Model.DebitField.Text = sAcct1
'		oETCOA1Text = puoETDialog.getControl("DebitField")
'		oETCOA1Text.Text = sAcct1
		gsETAcct1 = sAcct1
	endif
	
NormalExit:
	exit sub
	
ErrorHandler:
	gsETAcct1 = sAcct1
	GoTo NormalExit
	
end sub		'// end ETCOA1ListListener	6/17/20
'/**/
