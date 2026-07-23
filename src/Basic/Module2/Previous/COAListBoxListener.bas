'// COAListBoxListener.bas
'//--------------------------------------------------------------------
'// COAListBoxListener - Event handler when <Cancel> from COA selection.
'//		9/8/22.	wmk.	16:01
'//--------------------------------------------------------------------

public sub COAListBoxListener()

'//	Usage.	macro call or
'//			call COAListBoxListener()
'//
'// Entry.	normal entry from event handler for dialog where
'//			user selecting COA account from listbox
'//			puoCOASelectBtn = <Select> button object in dialog
'//			pbETActive = true if ET dialog active (implies update
'//			 of ET dialog .DebitField or .CreditField
'//			pbETDebitList = true; update .DebitField
'//			pbETCreditList = true; update .CreditField
'//
'//	Exit.	pusCOASelected = ""
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'// 9/7/22.		wmk.	pusCOASelected, puoETDialog.DebitField.Text = selection.
'// 9/8/22.		wmk.	pbETActive flag used to distinguish between ET call
'//				 and ledger sheet call.
'// Legacy mods.
'//	6/13/20.		wmk.	original code.
'//
'//	Notes. This sub is the linked macro to the StatusChanged() event linked
'// to the <COAList> listbox in the COA selection dialog. This event
'// activates the <Select> button in the dialog.

'//	constants.

'//	local variables.

	'// code.
	puoCOASelectBtn.Model.Enabled = true	'// enable <Select> button

	if not pbETActive then
dim oDocument	AS Object
dim oSel		AS Object
dim oRange		AS Object
dim oCell		AS Object
dim iSheetIx	AS Integer
	  oDocument=ThisComponent
	  oSel = oDocument.getCurrentSelection()
	  oRange = oSel.RangeAddress
	  iSheetIx = oRange.Sheet
	  oCell = oDocument.Sheets(iSheetIx).getCellByPosition(oRange.StartColumn,oRange.StartRow)
  	  oCell.String = puoCOAListBox.SelectedItem
'	endif	'// not ET calling
	else	'//    if pbETActive then
    	if pbETDebitList then
	 puoETDialog.Model.DebitField.Text = left(puoCOAListBox.SelectedItem,4)
	 	else
	 puoETDialog.Model.CreditField.Text = left(puoCOAListBox.SelectedItem,4)
	 	endif
	endif		'// pbETActive
end sub		'// end COAListBoxListener	9/8/22
'/**/
