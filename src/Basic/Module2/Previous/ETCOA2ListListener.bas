'// ETCOA2ListListener.bas
'//--------------------------------------------------------------------
'// ETCOA2ListListener - Process user click on <...> with COA2 account.
'//		6/17/20.	wmk.	10:15
'//--------------------------------------------------------------------

public sub ETCOA2ListListener()

'//	Usage.	macro call or
'//			call ETCOA2ListListener()
'//
'// Entry.	User clicked on <...> next to COA2 field
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
'//	6/17/20.	wmk.	original code; cloned from ETCOA1ListListener
'//
'//	Notes. If the user does not select a new COA from GetCOA, the
'//	current COA for the Credit field (gsETAcct2) will remain unchanged.
'//

'//	constants.

'//	local variables.
dim oETCOA2Text	As Object			'// account text field
dim sAcct2		As String			'// returned account from GetCOA

	'// code.
	sAcct2 = ""
	ON ERROR GOTO ErrorHandler

	pbETDebitList=false
	pbETCreditList=true
	sAcct2 = GetCOA(1)			'// get COA from dialog
	if Len(sAcct2)>0 then
		puoETDialog.Model.CreditField.Text = sAcct2
'		oETCOA2Text = puoETDialog.getControl("CreditField")
'		oETCOA2Text.Text = sAcct2
		gsETAcct2 = sAcct2
	endif
	
NormalExit:
	exit sub
	
ErrorHandler:
	gsETAcct2 = sAcct2
	GoTo NormalExit
	
end sub		'// end ETCOA2ListListener	6/17/20
'/**/
