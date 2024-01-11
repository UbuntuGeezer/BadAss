'// Module2Hdr.bas
'// Module2 code (complete with publics_2.bas) for dialogs
'//	Module2 subs/functions index.	7/5/20.
'//
'//	Main - Module2 Main - clear pbDlgLibLoaded flag.
'// F - First attempt at dialog box processing. (COADialog)
'// RunCOADlg - Run COA Dialog.
'// G - Enter Transaction Dialog test procedure.
'// H - Split Transaction Dialog test procedure.
'// ETAmountListener - Event handler <amount> from Enter Transaction.
'// ETPubVarsReset - Reset all ET dialog public vars and flags
'// ETCancelListener - Event handler <Cancel> from Enter Transaction.
'//
'/**/
'// publics2.bas
'//---------------publics_2.bas---------------------------------
'// Module2 public vars and constants
'//		7/6/20.		wmk.	07:15
'//
'//	Modification History.
'// ---------------------
'//	6/17/20.	wmk.	arrays added for splitting transactions
'//	6/18/20.	wmk.	publics added for SplitDialog handling
'// 6/20/20.	wmk.	gsSTRefs added for reference field support;
'//						gsSTObjFocus added to attempt code generalization
'//	6/24/20.	wmk		gdETAmount added; gsETAmount deleted for numeric
'//						field support on data entry; gbETSplitTrans 
'//						spelling corrected
'//	6/26/20.	wmk.	VS dialog support added
'//	6/27/20.	wmk.	MEDGRAY color constant added; gbSTEditMode flag
'//						added to initialize with public vars set from
'//						previous dialog instance
'//	7/5/20.		wmk.	subs/functions documentation added to header
'//	7/6/20.		wmk.	const COA_COLROWS added

'// color constants.
public const MEDGRAY=13421722
public const NOFILL = -1

'// chart of accounts access and control (LoadCOAs function)
'public goCOARangePtr As com.sun.star.table.CellAddress
public const	COA_SHEET="Chart of Accounts"
public const	COA_COLSTART=3		'//A9
public const	COA_ROWSTART=7
public const	COA_COLROWS=5		'// column with COA row count
public const	sCOA_1STROWID="BEGIN COA"
public const	sCOA_LASTROWID="END OF COA"
public const	COA_NASSETS=14		'// assets count
public const	COA_NLIABILITIES=20	'// liabilities count
public const	COA_NINCOME=13		'// income count
public const	COA_NEXPENSES=29	'// expenses count

'// Enter Transaction dialog publics.
public puoETDialog		As Object		'// ET dialog object
public gbETDateEntered	As Boolean		'// date field entered flag
public gbETDescEntered	As Boolean		'// description entered flag
public gbETAmtEntered	As Boolean		'// amount entered flag
public gbETCOA1Entered	As Boolean		'// Debit COA entered flag
public gbETCOA2Entered	As Boolean		'// Credit COA entered flag
public gbETRefEntered 	As Boolean		'// refereince entered flag
public gbETSplitTrans	As Boolean		'// split transaction flag
public gsETSplitCOAs (0) As String		'// list of COAs for split
public gsETSplitAmts (0)				'// list of amounts in split
public gsETSplitDescs (0)				'// list of descriptions in split
public gbETDebitIsTotal As Boolean		'// Debit is total; credits split
'//	ET dialog stored values.
public gsETDate As String				'// date entered
public gsETDescription As String		'// description entered
'public gsETAmount As String			'// debit and credit amt entered
public gdETAmount As Double				'// debit and credit amt entered
public gsETAcct1 As String				'// COA1 account
public gsETAcct2 As String				'// COA2 account
public gsETRef As String				'// reference text

'// Split Transaction dialog publics.
public puoSTDialog		As Object		'// ST dialog object
public gbSTEditMode		As Boolean		'// ST edit mode flag

'// ST dialog stored values.
public gdSTAccumTot		As Double		'// split accumulated total
public gdSTTotalAmt		As Double		'// total amount to split
public gbSTSplitCredits	As Boolean		'// credit values are splits flag
public giSTLastIx		As Integer		'// last active row index for input
public gsSTAcctFocus	As String		'// COA field name got Focus
public gsSTObjFocus		As String		'// set when object gets focus
public gsSTDescs(0)		As String		'// array of split descriptions
public gdSTAmts(0)		As Double		'// array of amounts
public gsSTCOAs(0)		As String		'// array of COAs
public gsSTRefs(0)		As String		'// array of Refs

'// View Split dialog publics.
public puoVSDialog		As Object		'// VS dialog object

'// COA dialog publics.
public pbDlgLibLoaded	As Boolean	'// dialogs loaded flag
public pbCOADlgExists	As Boolean	'// COA dialog exists flag
public puoCOADialog 	As Object	'// COA dialog box
public pusCOAList(0)	As String	'// COA Account strings
public puoCOASelectBtn	As Object	'// <Select> button

'// vars for dialogues interface to accounting package

public puoCOAListBox As Object		'// COA list box object in active form
public pusCOASelected As String		'// COA selected from list box (..Listener.bas)

OPTION EXPLICIT				'// code protection

'//--------------end publics_2.bas----7/6/20--------------------------
'/**/
'// Main.bas
'// COPY CODE FROM MAIN TO COASelectListener, COACancelListener
'//---------------------------------------------------------------------
'//	Main - Module2 Main - clear pbDlgLibLoaded flag.
'//			7/5/20.	wmk.	
'//---------------------------------------------------------------------

Sub Main
	pbDlgLibLoaded = false
End Sub		'// end Main	7/5/20
'/**/
'// COACancelListener.bas
'//--------------------------------------------------------------------
'// COACancelListener - Event handler when <Cancel> from COA selection.
'//		6/15/20.	wmk.	16:30
'//--------------------------------------------------------------------

public sub COACancelListener()

'//	Usage.	macro call or
'//			call COACancelListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked <Cancel> in COA dialog
'//			puoCOADialog = COA dialog box object
'//
'//	Exit.	pusCOASelected = ""
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.	wmk.	original code
'//	6/15/20.	wmk.	code modified to emulate Cancel type button
'//
'//	Notes. This sub is the linked macro to the Execute() event linked
'// to the <Cancel> button in the COA selection dialog. If the <Cancel>
'// button is also defined as type = 2, CANCEL it will drop the dialog
'// from the screen with returned value 0. If not, this event handler
'// be called and will emulate that process.

'//	constants.

'//	local variables.

	'// code.
	pusCOASelected = ""		'// clear COA selection string
'	msgBox("In Module2/COACancelListener" + CHR(13)+CHR(10) _
'        + "'"+pusCOASelected+"'" + " selected")
	puoCOADialog.endDialog(0)	'// end dialog as though cancelled	

end sub		'// end COACancelListener	6/13/20
'/**/
'// COAListBoxListener.bas
'//--------------------------------------------------------------------
'// COAListBoxListener - Event handler when <Cancel> from COA selection.
'//		6/15/20.	wmk.	12:30
'//--------------------------------------------------------------------

public sub COAListBoxListener()

'//	Usage.	macro call or
'//			call COAListBoxListener()
'//
'// Entry.	normal entry from event handler for dialog where
'//			user selecting COA account from listbox
'//			puoCOASelectBtn = <Select> button object in dialog
'//	Exit.	pusCOASelected = ""
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.		wmk.	original code
'//
'//	Notes. This sub is the linked macro to the Execute() event linked
'// to the <COAList> listbox in the COA selection dialog. This event
'// activates the <Select> button in the dialog.

'//	constants.

'//	local variables.

	'// code.
	puoCOASelectBtn.Model.Enabled = true	'// enable <Select> button

end sub		'// end COAListBoxListener	6/15/20
'/**/
'// COASelectListener.bas
'//--------------------------------------------------------------------
'// COASelectListener - Event handler when <Select> from COA selection.
'//		6/17/20.	wmk.	07:45
'//--------------------------------------------------------------------

public sub COASelectListener()

'//	Usage.	macro call or
'//			call COASelectListener()
'//
'// Entry.	puoCOAListBox.SelectedItem = item selection from dialog
'//			puoCOADialog = COA dialog box object
'//
'//	Exit.	pusCOASelected = puoCOAListBox.SelectedItem
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.	wmk.	original code
'//	6/14/20.	wmk.	close dialog code added
'// 6/17/29.	wmk.	debug msgBox disabled
'//
'//	Notes. This sub is the linked macro to the Execute() event linked
'// to the <Select> button in the COA selection dialog. The <Select>
'// button is also defined as type = 0, and will drop allow the dialog
'// to continue.

'//	constants.

'//	local variables.

	'// code.
	pusCOASelected = puoCOAListBox.SelectedItem
'	msgBox("In Module2/COASelectListener" + CHR(13)+CHR(10) _
'        + pusCOASelected + " selected")
	puoCOADialog.endDialog(2)				'// close dialog when Select
	
end sub		'// end COASelectListener	6/17/20
'/**/
'// ET.bas
'//---------------------------------------------------------------
'// ET - Enter Transaction Dialog run procedure.
'//		7/7/20.	wmk.	23:00	
'//----------------------------------------------------------------

public sub ET()

'//	Usage.	macro call or
'//			call ET()
'//
'//		Run ET dialog
'//
'// Entry.	Button click in dialog box		
'//			puoETDialog = declared object for COA dialog
'//
'//	Exit.	msgBox message displayed
'//
'// Calls.	ETPubVarsReset
'//
'//	Modification history.
'//	---------------------
'//	7/6/20.		wmk.	original code; adapted from G in BadAss Module2
'//	7/7/20.		wmk.	ensure dialog library loaded
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim iStatus As Integer

'// code.
	GlobalScope.BasicLibraries.LoadLibrary("BadAss")
	GlobalScope.DialogLibraries.LoadLibrary("BadAss")
	
'// Initialize ET dialog.
	puoETDialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.TransEntry)

	Select Case puoETDialog.Execute()
	Case 2		'// Record & Finish clicked
'		msgBox("Record & Finish clicked in Enter Transaction")
		'// iStatus = ETDialogRecord... called by [Record & Finish]
	Case 0		'// Cancel
'		pusCOASelected = ""		'// since cancel, Execute() event skipped
		iStatus = ETPubVarsReset(1)	'// reset everything
'		msgBox("Cancel or FormClose clicked")
	Case else
		msgBox("ET.bas - unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gdETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

puoETDialog.dispose()
	
end sub		'// end ET	7/7/20
'/**/
'// ETAmountListener.bas
'//---------------------------------------------------------------------
'// ETAmountListener - Event handler <amount> from Enter Transaction.
'//		6/24/20.	wmk.	17:45
'//---------------------------------------------------------------------

public sub ETAmountListener()

'//	Usage.	macro call or
'//			call ETAmountListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <amount> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gdETAmount = entered amount
'//			gbETAmtEntered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.	wmk.	original code
'// 6/17/20.	wmk.	bug fix enabling <record> buttons
'//	6/24/20.	wmk.	change to use type Double gdETAmount
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <debit> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETDlgCtrl		As Object		'// reusable field object
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button

'//	local variables.

	'// code.
	oETDlgCtrl = puoETDialog.getControl("AmtField1")
	gdETAmount = oETDlgCtrl.getValue()
	gbETAmtEntered = (Len(gdETAmount)>0)
	
	'// Enable SplitOption if amount entered > 0
	if gbETAmtEntered then
		oETDlgCtrl = puoETDialog.getControl("SplitOption")
		oETDlgCtrl.Model.Enabled = true
	endif
	
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETAmountListener	6/24/20
'/**/
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
	
	sAcct1 = GetCOA(1)			'// get COA from dialog
	if Len(sAcct1)>0 then
		oETCOA1Text = puoETDialog.getControl("DebitField")
		oETCOA1Text.Text = sAcct1
		gsETAcct1 = sAcct1
	endif
	
NormalExit:
	exit sub
	
ErrorHandler:
	gsETAcct1 = sAcct1
	GoTo NormalExit
	
end sub		'// end ETCOA1ListListener	6/17/20
'/**/
'// ETCOA1Listener.bas
'//---------------------------------------------------------------------
'// ETCOA1Listener - Event handler <COA1 account> from Enter Transaction.
'//		6/17/20.	wmk.	14:30
'//---------------------------------------------------------------------

public sub ETCOA1Listener()

'//	Usage.	macro call or
'//			call ETCOA1Listener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <COA1 account> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gsETAcct1 = entered amount
'//			gbETCOA1Entered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.	wmk.	original code.
'//	6/17/20.	wmk.	bug fix enabling [Record..] buttons
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <debit> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETCOA1Text	As Object			'// account text field
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button

'//	local variables.

	'// code.
	oETCOA1Text = puoETDialog.getControl("DebitField")
	gsETAcct1 = oETCOA1Text.Text
	gbETCOA1Entered = (Len(gsETAcct1)>0)
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETCOA1Listener	6/17/20
'/**/
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
	
	sAcct2 = GetCOA(1)			'// get COA from dialog
	if Len(sAcct2)>0 then
		oETCOA2Text = puoETDialog.getControl("CreditField")
		oETCOA2Text.Text = sAcct2
		gsETAcct2 = sAcct2
	endif
	
NormalExit:
	exit sub
	
ErrorHandler:
	gsETAcct2 = sAcct2
	GoTo NormalExit
	
end sub		'// end ETCOA2ListListener	6/17/20
'/**/
'// ETCOA2Listener.bas
'//---------------------------------------------------------------------
'// ETCOA2Listener - Event handler <COA1 account> from Enter Transaction.
'//		6/17/20.	wmk.	10:30
'//---------------------------------------------------------------------

public sub ETCOA2Listener()

'//	Usage.	macro call or
'//			call ETCOA2Listener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <COA2 account> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gsETAcct2 = entered account
'//			gbETCOA2Entered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code
'//	6/17/20.	wmk.	bug fix enabling Select buttons
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Credit account> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETCOA2Text	As Object			'// account text field
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button

'//	local variables.

	'// code.
	oETCOA2Text = puoETDialog.getControl("CreditField")
	gsETAcct2 = oETCOA2Text.Text
	gbETCOA2Entered = (Len(gsETAcct2)>0)
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETCOA2Listener	6/17/20
'/**/
'// ETCancelListener.bas
'//---------------------------------------------------------------------
'// ETCancelListener - Event handler <Cancel> from Enter Transaction.
'//		6/15/20.	wmk.	16:30
'//---------------------------------------------------------------------

public sub ETCancelListener()

'//	Usage.	macro call or
'//			call ETCancelListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked <Cancel> in Enter Transaction dialog
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	ET dialog ended
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.	wmk.	original code
'//	6/15/20.	wmk.	code modified to emulate Cancel type button
'//
'//	Notes. This sub is the linked macro to the Execute() event linked
'// to the <Cancel> button in the COA selection dialog. If the <Cancel>
'// button is also defined as type = 2, CANCEL it will drop the dialog
'// from the screen with returned value 0. If not, this event handler
'// be called and will emulate that process.

'//	constants.

'//	local variables.

	'// code.
	puoETDialog.endDialog(0)	'// end dialog as though cancelled	

end sub		'// end ETCancelListener	6/15/20
'/**/
'// ETCheckComplete.bas
'//---------------------------------------------------------------------
'// ETCheckComplete - Check all Enter Transaction dialog fields entered.
'//		6/15/20.	wmk.	17:30
'//---------------------------------------------------------------------

public function ETCheckComplete() As Boolean

'//	Usage.	bVar = ETCheckComplete()
'//
'// Entry.
'//		gbETDateEntered = date field entered flag
'//		gbETDescEntered = description entered flag
'//		gbETAmtEntered As Boolean			'// amount entered flag
'//		gbETCOA1Entered As Boolean			'// Debit COA entered flag
'//		gbETCOA2Entered As Boolean			'// Credit COA entered flag
'//
'//	Exit.	bVal = true if all entry flags are true
'//				   false otherwise
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.		wmk.	original code
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim bRetValue As Boolean

	'// code.
	
	bRetValue = gbETDateEntered AND gbETDescEntered AND gbETAmtEntered _
				 AND gbETCOA1Entered AND gbETCOA2Entered
	
	ETCheckComplete = bRetValue

end function 	'// end ETCheckComplete		6/15/20
'/**/
'// ETDateListener.bas
'//---------------------------------------------------------------------
'// ETDateListener - Event handler <date> from Enter Transaction.
'//		6/17/20.	wmk.	10:45
'//---------------------------------------------------------------------

public sub ETDateListener()

'//	Usage.	macro call or
'//			call ETDateListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <date> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gdtETDate = entered date
'//			gbETDateEntered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.	wmk.	original code
'//	6/17/20.	wmk.	bug fix enabling Select buttons
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <date> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETDateText		As Object		'// Date text field
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button

'//	local variables.

	'// code.
	oETDateText = puoETDialog.getControl("DateField1")
	gsETDate = oETDateText.Text
	gbETDateEntered = (Len(gsETDate)>0)
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETDateListener	6/17/20
'/**/
'// ETDescListener.bas
'//---------------------------------------------------------------------
'// ETDescListener - Event handler <description> from Enter Transaction.
'//		6/17/20.	wmk.	10:15
'//---------------------------------------------------------------------

public sub ETDescListener()

'//	Usage.	macro call or
'//			call ETDescListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <description> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gsETDescription = entered transaction description
'//			gbETDescEntered = true if nonempty string
'//
'// Calls.	ETCheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.	wmk.	original code
'//	6/17/20.	wmk.	bug fix enabling Select button
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <date> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETDescText		As Object		'// Desc text field
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button

'//	local variables.

	'// code.
	oETDescText = puoETDialog.getControl("DescField")
	gsETDescription = oETDescText.Text
	gbETDescEntered = (Len(gsETDescription)>0)
	bFormComplete = ETCheckComplete()
	if bFormComplete then
		oETRecordBtn = puoETDialog.getControl("RecordBtn")
		oETRecordBtn.Model.Enabled = true
		oETRecordCont = puoETDialog.getControl("RecordContBtn")
		oETRecordCont.Model.Enabled = true
	endif	

end sub		'// end ETDescListener	6/17/20
'/**/
'// ETDialogRecord.bas
'//---------------------------------------------------------------
'// ETDialogRecord - Record entered fields from ET Dialog.
'//		6/27/20.	wmk.	21:30
'//---------------------------------------------------------------

public function ETDialogRecord() As Integer

'//	Usage.	iVar = ETDialogRecord()
'//
'// Entry.	ET Dialog public vars have entered values
'//
'//	Exit.	Double entry transaction recorded in GL sheet at
'//			user selected position
'//			gbETSplitTrans = true if split transaction
'//							 false if regular transaction
'//			gbETDebitIsTotal = true if credits split;
'//							   false if debits split
'//
'// Calls.	ETDlgSplitRecord
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code; stub
'//	6/17/20.	wmk.	code fleshed out gradually
'//	6/24/20.	wmk.	change to record numeric debit and credit amount
'//						from gdETAmount
'//	6/25/20.	wmk.	bug fix eliminating gsETAmount which was replaced by
'//						gdETAmount 6/24
'// 6/27/20.	wmk.	ensure background in all but date fields is NOFILL
'//
'//	Notes. When the ET dialog completes, either the DebitField or CreditField
'// may have been set to "split". If that is the case, ETDlgSplitRecord will
'// be called to record the split transaction.
'//

'//	constants.

'//----------------constants borrowed from header.bas-------------------
const LTGREEN=10092390		'// decimal value of LTGREEN color

'// column index values and string lengths for column data			
'public const	INSBIAS=3		'// insert line count bias
'public const	COALEN=4		'// length of COA field strings
const COLDATE=0		'// Date - column A
const COLTRANS=1	'// Transaction - column B
const COLDEBIT=2	'// Debit - column C
const COLCREDIT=3	'// Credit - column D
const COLACCT=4		'// COA Account - column E
const COLREF=5		'// Reference - column F

'// cell formatting constants.
const DEC2=123		'// number format for (x,xxx.yy)			'// mod052020
const MMDDYY=37		'// date format mm/dd/y						'// mod052020
const FMTDATETIME=50	'// date/time format mm/dd/yyy hh:mm:ss 	'// mod060620
const LJUST=1		'// left-justify HoriJustify				'// mod052020
const CJUST=2		'// center HoriJustify						'// mod052020
const RJUST=3		'// right-justify HoriJustify				'// mod052320
const MAXTRANSL=50	'// maximum transaction text length			'// mod052020

'//----------------end constants borrowed from publics.bas--------------

'//	local variables
dim iRetValue As Integer
dim i		As Integer
dim iStatus	As Integer			'// general status

'// GL Sheet access variables for inserting transaction
dim oDoc		As Object		'// ThisComponent
dim oSel		As Object		'// current user selection
dim oRange		As Object		'// cell range selection info
dim iSheetIx	As Integer		'// sheet index
dim oGLSheet	As Object		'// GL sheet object
dim lGLRow		As Long			'// GL 1st line row
dim lGLCurrRow	As Long			'// current new transaction row

'// GL Transaction field cells
dim	oCellDate 	As Object		'// Date field
dim	oCellTrans 	As Object		'// Description field
dim	oCellDebit 	As Object		'// Debit field
dim	oCellCredit	As Object		'// Credit field
dim	oCellAcct	As Object		'// COA field
dim	oCellRef	As Object		'// Reference field
dim oDate		As Object		'// oCellDate.Text

	'// code.
	
	iRetValue = 0
	ON ERROR GOTO ErrorHandler

	'// determine if split transaction and handle
	gbETDebitIsTotal = (StrComp(gsETAcct2,"split")=0)
	gbETSplitTrans = (gbETDebitIsTotal OR (StrComp(gsETAcct1,"split")=0))
	if gbETSplitTrans then
		iStatus = ETDlgSplitRecord()
		if iStatus < 0 then
			GoTo ErrorHandler
		else
			GoTo NormalExit
		endif
	endif	'// end split transaction conditional
	
	'// set up access to GL on user selection
	oDoc = ThisComponent				'// set up for sheet access within document
	oSel = oDoc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet	'// get sheet index value
'	oSheets = oDoc.getSheets()
	oGLSheet = oDoc.Sheets.getByIndex(iSheetIx)	'// get GL sheet
	lGLRow = oRange.StartRow			'// set row pointer to 1st row
	
	'// set up and insert 2 rows at first row pointed to by user
	oGLSheet.Rows.insertByIndex(lGLRow, 2)	'// insert transaction rows
	'// add code here to update sheet date modified
	iStatus = SetSheetDate(oRange, MMDDYY)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif	'// end abnormal return from SetSheetDate
	
	'// for each row, copy transaction information, placing debit amt
	'// in Debit column of 1st row, credit amt in Credit column of
	'// 2nd row, all other columns in same position
	lGLCurrRow = lGLRow
	for i = 0 to 1
		'// extract key ledger fields for setting in new entries
		oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
		oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
		oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
		oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
		oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
		oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		
		'// set Date field in transaction line		
		oCellDate.setValue(DateValue(gsETDate))		'// set date field to value
		oCellDate.NumberFormat = MMDDYY				'// set mm/dd/yy format
		oCellDate.HoriJustify = RJUST				'// right-justify date
		oCellDate.CellBackColor = LTGREEN			'// set light green background
		
		'// add code here to set remaining transaction cells
		oCellTrans.String = gsETDescription
		oCellTrans.HoriJustify = LJUST
		oCellTrans.CellBackColor = NOFILL
		
		'// Debit amt is in first line; credit in second
		if i = 0 then
			oCellDebit.String = Str(gdETAmount)
			oCellDebit.setValue(gdETAmount)							'//	mod062420
			oCellDebit.NumberFormat = DEC2
			oCellCredit.String = ""
			oCellAcct.String = gsETAcct1
		else
			oCellDebit.String = ""
			oCellCredit.String = Str(gdETAmount)
			oCellCredit.setValue(Val(gdETAmount))					'//	mod062420
			oCellCredit.NumberFormat = DEC2
			oCellAcct.String = gsETAcct2
		endif
		oCellDebit.HoriJustify = RJUST
		oCellDebit.CellBackColor = NOFILL
		oCellCredit.HoriJustify = RJUST
		oCellCredit.CellBackColor = NOFILL
		oCellAcct.HoriJustify = CJUST
		oCellAcct.CellBackColor = NOFILL
				
		oCellRef.String = gsETRef 
		oCellRef.HoriJustify = CJUST
		oCellRef.CellBackColor = NOFILL
		
		lGLCurrRow = lGLCurrRow + 1	'// advance to next transaction row
		
	next i		'// advance transaction row

	GoTo NormalExit
	
NormalExit:
	ETDialogRecord = iRetValue
	exit function
	
ErrorHandler:
	iRetValue = -1
	GoTo NormalExit
	

end function 	'// end ETDialogRecord	6/25/20
'/**/
'// ETDialogReset.bas
'//---------------------------------------------------------------
'// ETDialogReset - Reset all ET dialog fields and flags.
'//		6/17/20.	wmk.	12:00
'//---------------------------------------------------------------

public function ETDialogReset(piOpt As Integer) As Integer

'//	Usage.	iVar = ETDialogReset( iOpt )
'//
'//		iOpt = 0 - reset all fields except Date
'//			 <> 0 - reset ALL fields
'//
'// Entry.	ET dialog public field values and flags defined
'//			puoETDialog = ET dialog object
'//			gbETDateEntered = date field entered flag
'//			gsETDate = Date field string
'/
'//	Exit.	iVar = 0 - normal return
'//				 <> 0 - error in resetting fields
'//
'// Calls.	ETPubVarsReset, ETDlgCtrlsReset
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code; stub
'// 6/17/20.	wmk.	move code to ETPubVarsReset and ETDlgCtrlsReset
'//
'//	Notes. Allows call to preserve date field for repetitive entries

'//	constants.

'//	local variables.
dim iRetValue 	As Integer		'// returned value
dim iOption 	As Integer		'// local copy of initialize option

	'// code.
	
	iRetValue = 0
	iOption = piOpt
	ON ERROR GOTO ErrorHandler
	
	'// clear all stored values
	iRetValue = ETPubVarsReset(iOption)	'// clear field storage and flags
	if iRetValue <> 0 then
		GoTo ErrorHandler
	endif

	'// clear dialog object fields.
	iRetValue = ETDlgCtrlsReset(iOption)	'// clear dialog fields
	if iRetValue <> 0 then
		GoTo ErrorHandler
	endif
	
NormalExit:
	ETDialogReset = iRetValue
	exit function
	
ErrorHandler:
	iRetValue = -1		'// set error return
	GoTo NormalExit
	

end function 	'// end ETDialogReset	6/17/20
'/**/
'// ETDlgSplitRecord.bas
'//---------------------------------------------------------------
'// ETDlgSplitRecord - Record split transaction from ET Dialog.
'//		7/7/20.	wmk.	12:00	
'//---------------------------------------------------------------

public function ETDlgSplitRecord() As Integer

'//	Usage.	iVar = ETDlgSplitRecord()
'//
'// Entry.	ET Dialog public vars have entered values
'//			gbETDebitIsTotal = true if credits are split
'//							   false if debits are split
'//
'//	Exit.	Double entry split transaction recorded in GL sheet at
'//			user selected position
'//
'// Calls.	SetSumFormula
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; adapted from ETDialogRecord
'//	6/26/20.	wmk.	bug fix where both debit and credit splits being
'//						recorded in Credit column
'//	6/27/20.	wmk.	bug fix where transaction rows being set to
'//						color of row above where split inserted; need
'//						to be set to NOFILL
'// 7/7/10.		wmk.	bug fix where total line background color not being
'//						set to NOFILL; bug fix where description centered if
'//						line above was centered
'//
'//	Notes. When the ET dialog completes, either the DebitField or CreditField
'// may have been set to "split". If that is the case, ETDlgSplitRecord will
'// be called to record the split transaction.
'//

'//	constants.

'//----------------constants borrowed from header.bas-------------------
const LTGREEN=10092390		'// decimal value of LTGREEN color

'// column index values and string lengths for column data			
'public const	INSBIAS=3		'// insert line count bias
'public const	COALEN=4		'// length of COA field strings
const COLDATE=0		'// Date - column A
const COLTRANS=1	'// Transaction - column B
const COLDEBIT=2	'// Debit - column C
const COLCREDIT=3	'// Credit - column D
const COLACCT=4		'// COA Account - column E
const COLREF=5		'// Reference - column F

'// cell formatting constants.
const DEC2=123		'// number format for (x,xxx.yy)			'// mod052020
const MMDDYY=37		'// date format mm/dd/y						'// mod052020
const FMTDATETIME=50	'// date/time format mm/dd/yyy hh:mm:ss 	'// mod060620
const LJUST=1		'// left-justify HoriJustify				'// mod052020
const CJUST=2		'// center HoriJustify						'// mod052020
const RJUST=3		'// right-justify HoriJustify				'// mod052320
const MAXTRANSL=50	'// maximum transaction text length			'// mod052020

'//----------------end constants borrowed from publics.bas--------------

'//	local variables
dim iRetValue As Integer
dim i		As Integer
dim iStatus	As Integer			'// general status
dim sSumFormula	As String		'// split sum formula

'// GL Sheet access variables for inserting transaction
dim oDoc		As Object		'// ThisComponent
dim oSel		As Object		'// current user selection
dim oRange		As Object		'// cell range selection info
dim iSheetIx	As Integer		'// sheet index
dim oGLSheet	As Object		'// GL sheet object
dim lGLRow		As Long			'// GL 1st line row
dim lGLCurrRow	As Long			'// current new transaction row
dim lGLBottomRow	As Long		'// bottom row of split transaction
dim iInsertCnt	As Integer		'// insertion row count
dim iSplitLim	As Integer		'// splits limit 0-based

'// GL Transaction field cells
dim	oCellDate 	As Object		'// Date field
dim	oCellTrans 	As Object		'// Description field
dim	oCellDebit 	As Object		'// Debit field
dim	oCellCredit	As Object		'// Credit field
dim	oCellAcct	As Object		'// COA field
dim	oCellRef	As Object		'// Reference field
dim oDate		As Object		'// oCellDate.Text

	'// code.
	
	iRetValue = 0
	ON ERROR GOTO ErrorHandler

	'// set up access to GL on user selection
	oDoc = ThisComponent				'// set up for sheet access within document
	oSel = oDoc.getCurrentSelection()	'// get current cell selection(s) info
	oRange = oSel.RangeAddress			'// extract range information
	iSheetIx = oRange.Sheet	'// get sheet index value
	oGLSheet = oDoc.Sheets.getByIndex(iSheetIx)	'// get GL sheet
	lGLRow = oRange.StartRow			'// set row pointer to 1st row
	lGLCurrRow = lGLRow
	
	'// rows to insert = upper bound of amounts + 1, +2 for upper and
	'// lower "split" divider rows +1 for total row
	iSplitLim = uBound(gdSTAmts)
	iInsertCnt = iSplitLim + 4
	lGLBottomRow = lGLCurrRow + iInsertCnt - 1	
	'// set up and insert iInsertCnt rows at first row pointed to by user
	oGLSheet.Rows.insertByIndex(lGLRow, iInsertCnt)	'// insert transaction rows
	oRange.EndRow = lGLBottomRow			'// set for SetSumFormula call
		
	'// update sheet date modified
	iStatus = SetSheetDate(oRange, MMDDYY)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif	'// end abnormal return from SetSheetDate

'/*--------------------------------------------------------------------------
	'// create top row "split" divider
	'// extract key ledger fields for setting in new entries
	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
	oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
	oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
	oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
	oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		
	'// set Date field in transaction line		
	oCellDate.setValue(DateValue(gsETDate))		'// set date field to value
	oCellDate.NumberFormat = MMDDYY				'// set mm/dd/yy format
	oCellDate.HoriJustify = RJUST				'// right-justify date
	oCellDate.CellBackColor = LTGREEN			'// set light green background

	'// set description field to ET description
	'// gsETDescription = entered description
	oCellTrans.String = gsETDescription
	oCellTrans.CellBackColor = LTGREEN			'// color entire 1st row
	oCellTrans.HoriJustify = LJUST
	
	'//	set debit and credit amount fields to ""
	oCellDebit.String = ""
	oCellDebit.CellBackColor = LTGREEN
	oCellCredit.String = ""
	oCellCredit.CellBackColor = LTGREEN
	
	'// set account and ref fields to "-" and "split"
	oCellAcct.String = "-"
	oCellAcct.HoriJustify = CJUST
	oCellAcct.CellBackColor = LTGREEN
	oCellRef.String = "split"
	oCellRef.HoriJustify = CJUST
	oCellRef.CellBackColor = LTGREEN
'*/----------------------------------------------------------------------------

	lGLCurrRow = lGLCurrRow + 1		'// advance to total row
'/*------------------enter total row-------------------------------------------
	'// extract key ledger fields for setting in new entries
	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
	oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
	oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
	oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
	oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		
	'// set Date field in transaction line		
	oCellDate.setValue(DateValue(gsETDate))		'// set date field to value
	oCellDate.NumberFormat = MMDDYY				'// set mm/dd/yy format
	oCellDate.HoriJustify = RJUST				'// right-justify date
	oCellDate.CellBackColor = LTGREEN			'// set light green background

	'// set description field to ET description
	'// gsETDescription = entered description
	oCellTrans.String = gsETDescription
	oCellTrans.CellBackColor = NOFILL
	oCellTrans.HoriJustify = LJUST

	'// set Total in debit or credit and appropriate COA
	if gbETDebitIsTotal then		'// credits split
		oCellDebit.Value = gdETAmount
		oCellDebit.NumberFormat = DEC2
		oCellCredit.String = ""
		oCellAcct.String = gsETAcct1
	else		'// debits split
		oCellCredit.Value = gdETAmount
		oCellCredit.NumberFormat = DEC2
		oCellDebit.String = ""
		oCellAcct.String = gsETAcct2
	endif		'// end debit is total conditional
	oCellDebit.HoriJustify = RJUST
	oCellDebit.CellBackColor = NOFILL
	oCellCredit.HoriJustify = RJUST
	oCellCredit.CellBackColor = NOFILL
	oCellAcct.HoriJustify = CJUST
	oCellAcct.CellBackColor = NOFILL
		
	'// set ref field
	oCellRef.String = gsETRef
	oCellRef.HoriJustify = CJUST
	oCellRef.CellBackColor = NOFILL
'*/---------end Total row------------------------------------------------------
 if false then
 GoTo JumpAhead
 endif
 
'/*----------------------------------------------------------------------------
	'// for each row, copy transaction information, placing debit amt
	'// in Debit column of 1st row, credit amt in Credit column of
	'// 2nd row, all other columns in same position
	lGLCurrRow = lGLCurrRow + 1		'// advance past total row
	for i = 0 to iSplitLim
		'// extract key ledger fields for setting in new entries
		oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
		oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
		oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
		oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
		oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
		oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		
		'// set Date field in transaction line		
		oCellDate.setValue(DateValue(gsETDate))		'// set date field to value
		oCellDate.NumberFormat = MMDDYY				'// set mm/dd/yy format
		oCellDate.HoriJustify = RJUST				'// right-justify date
		oCellDate.CellBackColor = LTGREEN			'// set light green background
		
		'// set Description field
		oCellTrans.String = gsSTDescs(i)
		oCellTrans.HoriJustify = LJUST
		oCellTrans.CellBackColor = NOFILL
		
		'// set amount in proper column
		if gbETDebitIsTotal then		'// credits split
			oCellDebit.String = ""
			oCellCredit.String = Str(gdSTAmts(i))
			oCellCredit.setValue(gdSTAmts(i))					
			oCellCredit.NumberFormat = DEC2
			oCellAcct.String = gsSTCOAs(i)
		else		'// debits split
			oCellDebit.String = Str(gdSTAmts(i))
			oCellCredit.String = ""
			oCellDebit.setValue(gdSTAmts(i))					
			oCellDebit.NumberFormat = DEC2
			oCellAcct.String = gsSTCOAs(i)
		endif		'// end debit is total conditional
	
		oCellDebit.HoriJustify = RJUST
		oCellDebit.CellBackColor = NOFILL
		oCellCredit.HoriJustify = RJUST
		oCellCredit.CellBackColor = NOFILL
		oCellAcct.HoriJustify = CJUST
		OCellAcct.CellBackColor = NOFILL
		
		oCellRef.String = gsSTRefs(i) 
		oCellRef.HoriJustify = CJUST
		oCellRef.CellBackColor = NOFILL
		
		lGLCurrRow = lGLCurrRow + 1	'// advance to next transaction row
		
	next i		'// advance split transaction row
'*/-----------------------------------------------------------------------------
 if false then
 GoTo JumpAhead
 endif

'/*--------------------------------------------------------------------------
	'// create bottom row "split" divider with total
	'// extract key ledger fields for setting in new entries
	oCellDate = oGLSheet.getCellByPosition(COLDATE, lGLCurrRow)
	oCellTrans = oGLSheet.getCellByPosition(COLTRANS, lGLCurrRow)
	oCellDebit = oGLSheet.getCellByPosition(COLDEBIT, lGLCurrRow)
	oCellCredit = oGLSheet.getCellByPosition(COLCREDIT, lGLCurrRow)
	oCellAcct = oGLSheet.getCellByPosition(COLACCT, lGLCurrRow)
	oCellRef = oGLSheet.getCellByPosition(COLREF, lGLCurrRow)
		
	'// set Date field in transaction line		
	oCellDate.setValue(DateValue(gsETDate))		'// set date field to value
	oCellDate.NumberFormat = MMDDYY				'// set mm/dd/yy format
	oCellDate.HoriJustify = RJUST				'// right-justify date
	oCellDate.CellBackColor = LTGREEN			'// set light green background

	'// set description field to ET description
	'// gsETDescription = entered description
	oCellTrans.String = gsETDescription
	oCellTrans.CellBackColor = LTGREEN			'// color entire last row
	oCellTrans.HoriJustify = LJUST
	
	'//	set debit and credit amount fields to ""
	oCellDebit.String = ""
	oCellDebit.CellBackColor = LTGREEN
	oCellCredit.String = ""
	oCellCredit.CellBackColor = LTGREEN
	
	'// set account field to appropriate SUM formula
	sSumFormula = SetSumFormula(oRange, gbETDebitIsTotal)
	oCellAcct.String = sSumFormula
	oCellAcct.SetFormula(sSumFormula)
	oCellAcct.NumberFormat = DEC2
	oCellAcct.HoriJustify = CJUST
	oCellAcct.CellBackColor = LTGREEN
	
	'// set ref field to "split"
	oCellRef.String = "split"
	oCellRef.HoriJustify = CJUST
	oCellRef.CellBackColor = LTGREEN
'*/----------------------------------------------------------------------------
JumpAhead:
	iRetValue = 0
	
NormalExit:
	ETDlgSplitRecord = iRetValue
	exit function
	
ErrorHandler:
	iRetValue = -1
	GoTo NormalExit
	

end function 	'// end ETDlgSplitRecord	7/7/20
'/**/
'// ETPubVarsReset.bas
'//--------------------------------------------------------------
'// ETPubVarsReset - Reset all ET dialog public vars and flags
'//		6/24/20.	wmk.	21:00
'//---------------------------------------------------------------

public function ETPubVarsReset(piOpt As Integer) As Integer

'//	Usage.	iVar = ETPubVarsReset( iOpt )
'//
'//		iOpt = 0 - reset all vars and flags except Date
'//			 <> 0 - reset ALL fields
'//
'// Entry.	called by <Cancel> handling code where <Cancel> button
'//			is type 2, and Listener bypassed
'//			ET dialog public field values and flags defined
'//			gbETDateEntered = date field entered flag
'//			gsETDate = Date field string
'/
'//	Exit.	iVar = 0 - normal return
'//				 <> 0 - error in resetting fields
'//			gsET... vars all cleared to ""
'//			gbET... flags al set to false
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/17/20.	wmk.	original code
'//	6/24/20.	wmk.	change to type Double gdETAmount from String'
'//						dbETDebitIsTotal initialized true
'//
'//	Notes. Allows call to preserve date field for repetitive entries
'// 6/24/20 WARNING: CHANGE ETSplitAmts ARRAY TO TYPE DOUBLE FROM STRING
'// FOR COMPATIBILITY WITH ST DIALOG RETURNED VALUES...

'//	constants.

'//	local variables.
dim iRetValue 	As Integer		'// returned value

	'// code.
	
	iRetValue = 0
	ON ERROR GOTO ErrorHandler
	
	'// conditionally reset Date field.
	if piOpt <> 0 then
		gbETDateEntered = false		
		gsETDate = ""
	endif

'// clear all fields and stored values
	'// clear field entry flags
	gbETDescEntered = false		'// description entered flag
	gbETAmtEntered = false		'// amount entered flag
	gbETCOA1Entered = false		'// Debit COA entered flag
	gbETCOA2Entered = false		'// Credit COA entered flag
	gbETRefEntered = false		'// refereince entered flag
	gbETSplitTrans = false		'// split transaction flag
	gbETDebitIsTotal = true		'// Debit is total; credits split
	'//	clear ET dialog stored values.
	gsETDescription = ""		'// description entered
	gdETAmount = 0. 			'// debit and credit amt entered
	gsETAcct1  = ""			'// COA1 account
	gsETAcct2  = ""			'// COA2 account
	gsETRef  = ""			'// reference text
 redim gsETSplitCOAs (0)		'// list of COAs in split
 redim gsETSplitAmts (0)		'// list of amounts in split
 redim gsETSplitDescs (0)		'// list of descriptions in split
	gsETSPlitCOAs(0) = ""
	gsETSplitAmts(0) = ""
	gsETSplitDescs(0) = ""
	
NormalExit:
	ETPubVarsReset = iRetValue
	exit function
	
ErrorHandler:
	iRetValue = -1		'// set error return
	GoTo NormalExit
	

end function 	'// end ETPubVarsReset	6/24/20
'/**/
'// ETRecdDoneListener.bas
'//---------------------------------------------------------------------
'// ETRecdDoneListener - Event handler <Record & Continue> from Enter Transaction.
'//		6/17/20.	wmk.	14:15
'//---------------------------------------------------------------------

public sub ETRecdDoneListener()

'//	Usage.	macro call or
'//			call ETRecdDoneListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked the [Record & Finish] button in ET dialog
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	Transaction data recorded in GL sheet at user selection
'//			Transaction data in public vars cleared
'//			ET dialog ended with flag = 2 (user closed)
'//
'// Calls.	ETDialogRecord, ETPubVarsReset
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code; stub
'//	6/17/20.	wmk.	completed; waiting on ETDialogRecord functional
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Record & Continue> cmd button in the Enter Transaction dialog.

'//	constants.

'//	local variables.
dim oETRecordBtn	As Object		'// Record & Finish button
dim oETRecordCont	As Object		'// Record & Continue button
dim iStatus 		As Integer		'// general status

	'// code.
	iStatus = -1		'// set error return
	ON ERROR GoTo ErrorHandler

	'// record transaction
	iStatus = ETDialogRecord()
	if iStatus < 0 then
		GoToErrorHandler
	endif
	
	'// clear all fields entered and associated flags
	iStatus = ETPubVarsReset(1)	'// reset everything
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	puoETDialog.endDialog(2)			'// end dialog
	exit sub
	
ErrorHandler:
	GoTo NormalExit
	
end sub		'// end ETRecdDoneListener	6/17/20
'/**/
'// ETRefHelpListener.bas
'//---------------------------------------------------------------------
'// ETRefHelpListener - Event handler <?> from Enter Transaction.
'//		6/16/20.	wmk.	18:00
'//---------------------------------------------------------------------

public sub ETRefHelpListener()

'//	Usage.	macro call or
'//			call ETRefHelpListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <reference> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	msgBox displayed with Reference Help information
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Referemce> text field in the Enter Transaction dialog.

'//	constants.

'//	local variables.

	'// code.
	msgBox("Reference field: This field can be used similar to a memo" _
		+ CHR(13)+CHR(10) + "field on a check. It may contain a notation" _
		+ CHR(13)+CHR(10) + "such as 'Monthly', return credit, etc." _
		+ CHR(13)+CHR(10) + "and is limited to 12 characters.")

end sub		'// end ETRefHelpListener	6/16/20
'/**/
'// ETRefListener.bas
'//---------------------------------------------------------------------
'// ETRefListener - Event handler <Reference> from Enter Transaction.
'//		6/16/20.	wmk.	18:00
'//---------------------------------------------------------------------

public sub ETRefListener()

'//	Usage.	macro call or
'//			call ETRefListener()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user changed the <reference> field in Enter Transaction
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	gsETRef = entered reference
'//			gbETRefEntered = true if nonempty string
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/16/20.	wmk.	original code
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Referemce> text field in the Enter Transaction dialog.

'//	constants.
dim bFormComplete	As Boolean		'// all fields complete flag
dim oETRefText	As Object			'// reference text field

'//	local variables.

	'// code.
	oETRefText = puoETDialog.getControl("RefField")
	gsETRef = oETRefText.Text
	gbETRefEntered = (Len(gsETRef)>0)

end sub		'// end ETRefListener	6/16/20
'/**/
'// ETSplitRun.bas
'//---------------------------------------------------------------
'// ETSplitRun - [Split] button Execute() event handler.
'//		7/6/20.	wmk.	11:30
'//---------------------------------------------------------------

public sub ETSplitRun()

'//	Usage.	macro call or
'//			call ETSplitRun( )
'//
'// Entry.	User clicked [Split] control in ET dialog
'//			Standard library loaded, or else wouldn't be here
'//			puoSTDialog = defined as object for dialog
'//			gdSTTotalAmt = available for total from line 1 of transaction
'//
'//	Exit.	ST dialog run; on STdialog completion, global vars
'//			set with ST dialog entries
'//			gsSTDescs() = Description fields entered
'//			gsSTAmts() = Amount fields entered
'//			gsSTCOAs() = COA account fields entered
'//			gsSTRefs() = Reference fields entered
'//			gsSTTotalAcct = COA account for splitting
'//			
'//	ST Returned status = 	<Cancel> from ST dialog
'//							<Done> from ST dialog
'//			IF <Done>, activate [View Split] control
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code
'//	6/27/20.	wmk.	message boxes eliminated; TotalAcct picked up from
'//						ST dialog; disable ViewSplitBtn if user cancelled split
'//	7/6/20.		wmk.	change to use GlobalScope libraries/BadAss
'//
'//	Notes.	[Split] button should only activate after line 1 of ET
'// transaction has a nonzero value.
'//

'//	constants.

'//	local variables.
dim iStatus 	As Integer
dim	oDlgCtrl	As Object		'// ET control object to access Split RB
dim oDlgCtrlA	As Object		'// ET control object when accessing Debit/Credit

	'// code.
	ON ERROR GOTO ErrorHandler	

'// Initialize ET dialog.
	puoSTDialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.SplitDialog)

'//	set total amount previously set by ET AmtField1...
	gdSTTotalAmt = gdETAmount	
'	gdSTTotalAmt = 150.00

	iStatus = fSTDialogReset()		'// attempt to reset/initialize dialog
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
	
	Select Case puoSTDialog.Execute()
	Case 2		'// Done clicked
'		msgBox("Record & Finish clicked in Split Transaction")
		'// update DebitField or CreditField field with "split" based on flag
'//		gbSTSplitCredits = true; split CreditField; else split DebitField
		if gbSTSplitCredits then
			oDlgCtrl = puoETDialog.getControl("CreditField")
			oDlgCtrlA = puoETDialog.getControl("DebitField")
			oDlgCtrlA.Text = gsETAcct1		'// Debit is total
		else
			oDlgCtrl = puoETDialog.getControl("DebitField")
			oDlgCtrlA = puoETDialog.getControl("CreditField")
			oDlgCtrlA.Text = gsETAcct2		'// Credit is total
		endif	'// end credits are split conditional
		oDlgCtrl.Text = "split"

		
		'// enable [View Split] control
		oDlgCtrl = puoETDialog.getControl("ViewSplitBtn")
		oDlgCtrl.Model.Enabled = true
		
	Case 0		'// Cancel Split
		iStatus = STPubVarsReset()	'// reset everything
'		msgBox("ST dialog Cancel or FormClose clicked")
		if iStatus < 0 then
			GoTo ErrorHandler
		endif
		
		'// came back cancelled - deselect Split radio button
		oDlgCtrl = puoETDialog.getControl("SplitOption")
		oDlgCtrl.State = false
		oDlgCtrl = puoETDialog.getControl("ViewSplitBtn")
		oDlgCtrl.Model.Enabled = false
		
	Case else
		msgBox("unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gsETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

	'// clear instantiation of ET dialog
	puoSTDialog.dispose()

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("in ETSplitRun - error initializing Split dialog.")
	GoTo NormalExit
	
end sub		'// end ETSplitRun		7/6/20
'/**/
'// ETViewSplitRun.bas
'//---------------------------------------------------------------
'// ETViewSplitRun - [Split] button Execute() event handler.
'//		7/6/20.	wmk.	12:30
'//---------------------------------------------------------------

public sub ETViewSplitRun()

'//	Usage.	macro call or
'//			call ETViewSplitRun( )
'//
'// Entry.	User clicked [View Split] control in ET dialog
'//			Standard library loaded, or else wouldn't be here
'//			puoVSDialog = defined as object for dialog
'//			gdSTTotalAmt = available for total from line 1 of transaction
'//			gsSTDescs() = Description fields entered
'//			gsSTAmts() = Amount fields entered
'//			gsSTCOAs() = COA account fields entered
'//			gsRefs() = Reference fields entered
'//
'//	Exit.	VS dialog run; on VSDialog completion, global vars
'//			set with ST dialog entries
'//	ST Returned status = 	<Edit Split> from VS dialog
'//							<Done> from VS dialog
'//			IF <Done>, activate [View Split] control
'//
'// Calls.	fVSDialogPreset(), ETSplitRun
'//
'//	Modification history.
'//	---------------------
'//	6/26/20.	wmk.	original code; adapted from ETSplitRun;
'//						minimal code executed to test linkage
'//	6/27/20.	wmk.	initial bugs fixed; implement callout to
'//						ST dialog if <Edit Split> in VS
'//	7/6/20.		wmk.	change to GlobalScope/BadAss library
'//
'//	Notes.	[View Split] button should only activate after the user has
'//	completed a transaction split using ST dialog.


'//	constants.

'//	local variables.
dim iStatus 	As Integer
dim	oDlgCtrl	As Object		'// ET control object to access Split RB

	'// code.
	ON ERROR GOTO ErrorHandler	

'// Initialize ET dialog.
	puoVSDialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.ViewSplit)

if false then
GoTo JumpAround1
endif

	'// all public vars from the ET and ST dialogs are still current
	'// since this event is being handled for the ST dialog
	'//	fVSDialogPreset will initialize all the display fields in the
	'// VS dialog
	iStatus = fVSDialogPreset
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
	
JumpAround1:
	
	Select Case puoVSDialog.Execute()
	Case 2		'// <Edit Split> clicked
		msgBox("<Edit Split> clicked in View Split")
		puoVSDialog.dispose()
		GoTo RunSplitDlg
	Case 0		'// Cancel or Done
'		iStatus = STPubVarsReset()	'// reset everything
'		msgBox("VS dialog Done or FormClose clicked")
'		if iStatus < 0 then
'			GoTo ErrorHandler
'		endif
'		
'		'// came back cancelled - deselect Split radio button
'		oDlgCtrl = puoETDialog.getControl("SplitOption")
'		oDlgCtrl.State = false
		
	Case else
		msgBox("VS - unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gsETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

	'// clear instantiation of VS dialog
	puoVSDialog.dispose()

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("in ETViewSplitRun - error initializing View Split dialog.")
	GoTo NormalExit

RunSplitDlg:
	gbSTEditMode = true
	call ETSplitRun()
	GoTo NormalExit
	
end sub		'// end ETViewSplitRun		7/6/20
'/**/
'// F.bas
'//---------------------------------------------------------------
'// F - First attempt at dialog box processing. (COADialog)
'//		6/15/20.	wmk.	17:00
'//---------------------------------------------------------------

public sub F()

'//	Usage.	macro call or
'//			call F()
'//
'//		<parameters description>
'//
'// Entry.	Button click in dialog box		
'//			puoCOADialog = declared object for COA dialog
'//			puoCOAListBox = list box object
'//			puoSelectBtn = declared object for <Select> button
'//			pusCOASelected = reserved for COA string selected
'//			current Component has sheet "Chart of Accounts" with COA table
'//
'//	Exit.	msgBox message displayed
'//			pusCOASelected = COA selected
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.	wmk.	original code
'//	6/14/20.	wmk.	call LoadCOAs and run form added; code
'//						modified to use puoCOADialog object
'//	6/15/20.	wmk.	name of COA list dialog changed; conbtrols
'//						labeled as SelectButton, CancelButton
'//	7/5/20.		wmk.	Entry dependencies updated
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
'dim oListBox As com.sun.star.awt.UnoControlListBox
dim oListBox As Object

Dim Dlg As Object
dim iCOACount As Integer

'// code.
	DialogLibraries.LoadLibrary("Standard")

'// Initialize COA Dialog
	puoCOADialog = CreateUnoDialog(DialogLibraries.Standard.COADialog)
	puoCOAListBox =puoCOADialog.getControl("COAList")
	puoCOASelectBtn = puoCOADialog.getControl("SelectBtn")
	puoCOASelectBtn.Model.Enabled = false			'// disable button until user selects item
	iCOACount = LoadCOAs()
	puoCOAListBox.addItems(pusCOAList, 0)

	Select Case puoCOADialog.Execute()
	Case 2		'// Select clicked
		msgBox("Select clicked in Chart of Accounts")
	Case 0		'// Cancel
		pusCOASelected = ""		'// since cancel, Execute() event skipped
		msgBox("Cancel or FormClose clicked")
	Case else
		msgBox("unevaluated dialog return")
	End Select
	msgBox("In Module2/F" + CHR(13)+CHR(10) _
    	    + "'"+pusCOASelected+"'" + " selected")

puoCOAListBox.dispose()
	
end sub		'// end F	6/15/20
'/**/
'// fSTAmtLostFocus.bas
'//---------------------------------------------------------------
'// fSTAmtLostFocus - Handle all AmtFld LostFocus events.
'//		6/22/20.	wmk.	08:00
'//---------------------------------------------------------------

public function fSTAmtLostFocus(psFldName As String) As Integer

'//	Usage.	iVar = fSTAmtLostFocus(sFldName)
'//
'//	Parameters.	sFldName = name of dialog field losing focus
'//				(e.g. "AmtFld1")
'// Entry.	AmtFld1..AmtFld10 lost focus
'//			gdSTAmts() = updated with newly entered amount
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gsSTAccumTot, gsSTTotalAmt updated with newly entered amount
'//			<Done> button activated if dialog entries complete
'//			next row fields activated if row complete
'//
'// Calls.	STFld6ToIndex, STUpdateAccumTot, CheckRowComplete, CheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; cloned from STAmtLostFocus
'//	6/21/20.	wmk.	check for total reached before activating next row
'// 6/21/20.	wmk.	change to accept psFldName parameter; to be called by
'//						all STAmt.LostFocus events; changed totals check and
'//						activate <Done> conditional
'//	6/22/20.	wmk.	change name to begin with "f" so can be distinguished
'//						by sub with same name for testing

'//	constants.

'//	local variables.
dim iRetValue 		As Integer			'// returned status
dim iStatus			As Integer			'// general status
dim iRowIx			As Integer			'// index of present row
dim oDlgControl		As Object			'// field ptr
dim bRowComplete	As Boolean			'// row complete flag
dim bDlgComplete	As Boolean			'// dialog complete flag
dim bTotalMet		As Boolean			'// split total met flag
dim sFldName 		As String			'// local copy of field name

	'// code.
	iRetValue = -1						'// set bad return
	ON ERROR GOTO ErrorHandler
'msgBox("In fSTAmtLostFocus entry values.." + CHR(13)+CHR(10) _
'	+ "gdSTAccumTot = " + gdSTAccumTot + CHR(13)+CHR(10) _
'	+ "gdSTTotalAmt = " + gdSTTotalAmt + CHR(13)+CHR(10) _
'	+ "gsSTObjFocus = '" + gsSTObjFocus + "'" )
	
	sFldName = psFldName				'// use local copy of field name
'msgBox("fSTAmtLostFocus - sFldName = '" + sFldName + "'")
	if Len(sFldName) < 7 then
		GoTo ErrorHandler
	endif	'// error - field name not set up for entry
	
	bDlgComplete = false
	iRowIx = STFld6ToIndex(sFldName)
	iStatus = STUpdateAccumTot(sFldName)
'msgBox("In fSTAmtLostFocus values after update.." + CHR(13)+CHR(10) _
'	+ "gdSTAccumTot = " + gdSTAccumTot + CHR(13)+CHR(10) _
'	+ "gdSTTotalAmt = " + gdSTTotalAmt )
	
	if iStatus < 0 then
		GoTo ErrorHandler
	endif	'// end error in updating accum total conditional

	bTotalMet = (iStatus = 1)
	bRowComplete = STRowComplete(iRowIx)
	if bRowComplete then
		bDlgComplete = STCheckComplete()
		oDlgControl = puoSTDialog.getControl("UnsplitTot")
'		if gdSTAccumTot < gdSTTotalAmt then		
		if oDlgControl.Value > 0 then		
			'// check for next row Enabled; if not, enable it
			iStatus = STEnableNextRow(iRowIx)
			if iStatus < 0 then
				GoTo ErrorHandler
			endif
		else
			bTotalMet = true
		endif '// end total not met yet conditional
		
	endif	'// end row complete conditional
	
	'// activate/deactivate <Done> control if dialog complete
	oDlgControl = puoSTDialog.getControl("DoneBtn")
'	bDlgComplete = (bDlgComplete AND bTotalMet)
	bDlgComplete = (bDlgComplete OR bTotalMet)
	if bDlgComplete then
		oDlgControl.Model.Enabled = true
	else
		oDlgControl.Model.Enabled = false
	endif
	iRetValue = 0
	
NormalExit:
	fSTAmtLostFocus = iRetValue
	exit function
	
ErrorHandler:
	msgBox("In fSTAmtLostFocus - unprocessed error")
	iRetValue = iStatus
	GoTo NormalExit
	
end function		'// end fSTAmtLostFocus	6/22/20
'/**/
'// fVSDialogPreset.bas
'//---------------------------------------------------------------
'// fVSDialogPreset - Preset fields in VS dialog.
'//		6/27/20.	wmk.	08:45
'//---------------------------------------------------------------

public function fVSDialogPreset() As Integer	

'//	Usage.	iVar = fVSDialogPreset()
'//
'// Entry.	gsSTDescs() = descriptions array
'//			gdSTAmts() = split amounts array
'//			gsSTCOAs() = COA accounts array
'//			gsSTRefs() = references array
'//			gbSplitCredits = true if gdSTAmts are Credits;
'//							 false if gdSTAmts are Debits
'//			gdSTTotalAmt = split total
'//			VSDialog fields:
'// DebitTotRB - Debit is total radio button
'// CreditTotRB - Credit is total radio button
'// TotalAmt - total transaction amount text box (numeric, read-only)
'//	TotalAcctFld - account COA transaction being split from
'// DescFld1 - description field 1
'// ...
'// DescFld10 - description field 10
'//	AmtFld1 - split amount field 1	(numeric)
'// ...
'// AmtFld10 - split amount field 10 (numeric)
'// AcctFld1 - COA account field 1
'//	...
'// AcctFld10 - COA account field 10
'// DoneBtn - [Done] button
'// EditBtn - [Reset Form] button
'//
'//	Exit.	iVar = 0 - no error in preset
'//				 < 0 - error in preset
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/26/20.	wmk.	original code
'//	6/27/20.	wmk.	initial bug fixes; inactive cell backgrounds to MEDGRAY
'//
'//	Notes.


'//	constants.

'//	local variables.
dim iRetValue As Integer
dim oVSDlgCtrl 	As Object
dim oVSDlgCtrl1 As Object		'// 2nd control for when needed
dim i			As Integer		'// loop index
dim iActiveLimit	As Integer	'// active index limit in arrays
dim sFldNum		As String		'// field number
dim sCTrlName	As String		'// control name
dim oVSStyle	As Object		'// to access FieldColor

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler

	'// set State of Debit/Credit radio buttons based on gbSTSplitCredits
	oVSDlgCtrl = puoVSDialog.getControl("DebitTotRB")
	oVSDlgCtrl1 = puoVSDialog.getControl("CreditTotRB")
	if gbSTSplitCredits then
		oVSDlgCtrl.State = true
		oVSDlgCtrl1.Model.Enabled = false
	else
		oVSDlgCtrl.State = false
		oVSDlgCtrl.Model.Enabled = false
		oVSDlgCtrl1.Model.Enabled = true
	endif
	
	'// set Total field to gdSTTotalAmt
	oVSDlgCtrl = puoVSDialog.getControl("TotalAmt")
	oVSDlgCtrl.Value = gdSTTotalAmt

	'// set account split from field
	oVSDlgCtrl = puoVSDialog.getControl("TotalAcctFld")
	if gbSTSplitCredits then
		oVSDlgCtrl.Text = gsETAcct1	'// debit account
	else
		oVSDlgCtrl.Text = gsETAcct2	'// credit account
	endif
	
	'//	set all active Description, Amount, and COA fields for all
	'// array values up to subscript UBound and set fields enabled
	'// so easily visible. All fields are read-only.
	'// set all inactive Description, Amount, and COA fields empty
	'// and set Enabled to false to darken fields
	iActiveLimit = uBound(gdSTAmts)
	for i = 0 to iActiveLimit
		sFldNum = Trim(Str(i+1))
		sCtrlName = "DescFld" + sFldNum
		oVSDlgCtrl = puoVSDialog.getControl(sCtrlName)
		oVSDlgCtrl.Text = gsSTDescs(i)
		oVSDlgCtrl.Model.Enabled = true
		
		sCtrlName = "AmtFld" + sFldNum
		oVSDlgCtrl = puoVSDialog.getControl(sCtrlName)
		oVSDlgCtrl.Text = Str(gdSTAmts(i))
		oVSDlgCtrl.setValue(gdSTAmts(i))
		oVSDlgCtrl.Model.Enabled = true
		
		sCtrlName = "AcctFld" + sFldNum
		oVSDlgCtrl = puoVSDialog.getControl(sCtrlName)
		oVSDlgCtrl.Text = gsSTCOAs(i)
		oVSDlgCtrl.Model.Enabled = true
		
		sCtrlName = "RefFld" + sFldNum
		oVSDlgCtrl = puoVSDialog.getControl(sCtrlName)
		oVSDlgCtrl.Text = gsSTRefs(i)
		oVSDlgCtrl.Model.Enabled = true
		
	next i	'// advance to next control name for each field

	'// set remaining fields Enabled = false with no data
	if iActiveLimit < 9 then
		for i = iActiveLimit+1 to 9
			sFldNum = Trim(Str(i+1))
			sCtrlName = "DescFld" + sFldNum
			oVSDlgCtrl = puoVSDialog.getControl(sCtrlName)
			oVSStyle = oVSDlgCtrl.StyleSettings
			oVSDlgCtrl.Model.Enabled = false
			oVSStyle.FieldColor = MEDGRAY

			sCtrlName = "AmtFld" + sFldNum
			oVSDlgCtrl = puoVSDialog.getControl(sCtrlName)
			oVSStyle = oVSDlgCtrl.StyleSettings
			oVSDlgCtrl.Model.Enabled = false
			oVSStyle.FieldColor = MEDGRAY
		
			sCtrlName = "AcctFld" + sFldNum
			oVSDlgCtrl = puoVSDialog.getControl(sCtrlName)
			oVSStyle = oVSDlgCtrl.StyleSettings
			oVSDlgCtrl.Model.Enabled = false
			oVSStyle.FieldColor = MEDGRAY
		
			sCtrlName = "RefFld" + sFldNum
			oVSDlgCtrl = puoVSDialog.getControl(sCtrlName)
			oVSStyle = oVSDlgCtrl.StyleSettings
			oVSDlgCtrl.Model.Enabled = false
			oVSStyle.FieldColor = MEDGRAY
		next i
		
	endif		'// end active limit < 9 conditional
	

	iRetValue = 0	'// normal return
	
NormalExit:	
	fVSDialogPreset = iRetValue
	exit function
	
ErrorHandler:
	msgBox("fVSDialogPreset - i = " + i + " unprocessed error.")
	GoTo NormalExit
	
end function 	'// end fVSDialogPreset	6/27/20
'/**/
'// G.bas
'//---------------------------------------------------------------
'// G - Enter Transaction Dialog test procedure.
'//		8/26/22.	wmk.	10:26	
'//----------------------------------------------------------------

public sub G()

'//	Usage.	macro call or
'//			call G()
'//
'//		Run ET dialog...
'//
'// Entry.	Button click in dialog box		
'//			puoETDialog = declared object for COA dialog
'//
'//	Exit.	msgBox message displayed
'//
'// Calls.	ETPubVarsReset
'//
'//	Modification history.
'//	---------------------
'//	6/15/20.	wmk.	original code.
'// 6/16/10.	wmk.	message boxes for verifification.
'//	6/17/20.	wmk.	add ETPubVars reset call on Cancel.
'//	6/25/20.	wmk.	gsETAmount corrected to new gdETAmount.
'//	6/27/20.	wmk.	msgBoxes commented to deactivate.
'// 8/26/22.	wmk.	eliminate references to G. bas.
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
'dim oListBox As com.sun.star.awt.UnoControlListBox
dim iStatus As Integer

'// code.
DialogLibraries.LoadLibrary("Standard")

'// Initialize ET dialog.
	puoETDialog = CreateUnoDialog(DialogLibraries.Standard.TransEntry)

	Select Case puoETDialog.Execute()
	Case 2		'// Record & Finish clicked
'		msgBox("Record & Finish clicked in Enter Transaction")
		'// iStatus = ETDialogRecord... called by [Record & Finish]
	Case 0		'// Cancel
'		pusCOASelected = ""		'// since cancel, Execute() event skipped
		iStatus = ETPubVarsReset(1)	'// reset everything
'		msgBox("Cancel or FormClose clicked")
	Case else
		msgBox("G- unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gdETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

puoETDialog.dispose()
	
end sub		'// end G	8/26/22
'/**/

'// GA.bas
'//---------------------------------------------------------------
'// GA - Store user selected COA in first cell in selected range
'// (Shortcut for GetCOA(1)
'//		7/5/20.	wmk.	22:00
'//---------------------------------------------------------------

public sub GA()

'//	Usage.	macro call or
'//			call GA()
'//
'// Entry.	pbCOADlgExists = true if COA dialog previously initialized
'//							 false if initialized here
'//			BadAss library loaded (automatic event on document open)
'//
'//	Exit.	if user did not cancel dialog, user selected COA stored at
'//			first cell in currently selected range
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	7/5/20.		wmk.	original code; adapted from GetCOA
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim sRetValue As String		'// returned COA
dim iStatus As Integer		'// general status
dim iCOACount As Integer	'// COA list length
dim	oDoc		As Object	'// ThisComponent
dim oSel		As Object	'// CellRangeAddress; current selection
dim oAcct		As Object	'// COA cell user selected
dim lColumn		As Long		'// selected column
dim lRow		As Long		'// selected row
dim oSheet		As Object	'// selected sheet
dim iSheetIx	As Integer	'/// sheet index

	'// code.
	ON ERROR GOTO ErrorHandler
	oDoc = ThisComponent
	oSel = oDoc.getCurrentSelection()
	iSheetIx = oSel.RangeAddress.Sheet
	oSheet = oDoc.Sheets.getByIndex(iSheetIx)
	lColumn = oSel.RangeAddress.StartColumn
	lRow = oSel.RangeAddress.StartRow
			
'		iStatus = LoadStdLibrary()		'// ensure dialogs
'DialogLibraries.LoadLibrary("Standard")
'		iStatus = LoadBadAssLib()		'// ensure dialogs
'
'		if iStatus < 0 then
'			GoTo ErrorHandler
'		endif

	'// create dialog if not yet done
	if NOT pbCOADlgExists then
		puoCOADialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.COADialog)
		pbCOADlgExists = true
	endif	'// end dialog non-existent conditional
		'// load list with COA strings
	puoCOAListBox = puoCOADialog.getControl("COAList")
	puoCOASelectBtn = puoCOADialog.getControl("SelectBtn")
	puoCOASelectBtn.Model.Enabled = false			'// disable button until user selects item
	iCOACount = LoadCOAs()
	puoCOAListBox.addItems(pusCOAList, 0)

	'// run dialog and get user selection
	Select Case puoCOADialog.Execute()
	Case 2		'// Select clicked
'		msgBox("COA Selected: " + pusCOASelected)
'			msgBox("Select clicked in Chart of Accounts")
		'// store COA in first cell of oSel range
		oAcct = oSheet.getCellByPosition(lColumn, lRow)
		oAcct.String = Left(pusCOASelected,4)
		oAcct.HoriJustify = CJUST
		
	Case 0		'// Cancel
		pusCOASelected = ""		'// since cancel, Execute() event skipped
'			msgBox("Cancel or FormClose clicked")
	Case else
		msgBox("unevaluated COA dialog return")
	End Select

		
NormalExit:

'msgBox("In GA" + CHR(13)+CHR(10) _
'        + "'"+sRetValue+"'" + " selected")

	exit sub
	
ErrorHandler:
	msgBox("In GA - unprocessed error.")
	GoTo NormalExit
	
end sub		'// end GA		7/5/20
'/**/
'// GetCOA.bas
'//---------------------------------------------------------------
'// GetCOA - Return user selected COA.
'//		7/6/20.	wmk.	11:00
'//---------------------------------------------------------------

public function GetCOA(piFromOpt As Integer) As String

'//	Usage.	sVar = GetCOA(iFromOpt)
'//
'//		iFromPrompt = USELAST (0) - use COA from last user selection
'//					  SELECTNEW (1) - use dialog to have user get COA
'//
'// Entry.	pusCOASelected = COA full string from last dialog
'//			pbCOADlgExists = true if COA dialog previously initialized
'//							 false if initialized here
'//
'//	Exit.	sVar = first 4 chars of COA user selected;
'//				   if USELAST option, from pusCOASelected on entry
'//				   if SELECTNEW option, from Chart of Accounts dialog
'//
'// Calls.	[LoadStdLibrary] LoadBadAssLib
'//
'//	Modification history.
'//	---------------------
'//	6/14/20.	wmk.	original code
'//	6/16/20.	wmk.	update to use new COADialog (formerly ListBox);
'//	6/17/20.	wmk.	correct load dialog code to initialize
'//						puoCOASelectBtn; debug msgBox calls removed'//
'//	7/5/20.		wmk.	change to load BadAss library
'//	7/6/20.		wmk.	make dialog reload unconditional
'//
'//	Notes. <Insert notes here>
'//

'//	constants.
const USELAST=0		'// use last COA selected
const SELECTNEW=1	'// select new COA from dialog

'//	local variables.
dim sRetValue As String		'// returned COA
dim iStatus As Integer		'// general status
dim iCOACount As Integer	'// COA list length
static sLastCOA As String	'// last COA selected

	'// code.
	
	sRetValue = ""
	iStatus = 0
	if piFromOpt = USELAST then
'		sRetValue = Left(pusCOASelected, 4)
		sRetValue = sLastCOA
		GoTo NormalExit
	else
		iStatus = LoadStdLibrary()		'// ensure dialogs
'DialogLibraries.LoadLibrary("Standard")
		iStatus = LoadBadAssLib()		'// ensure dialogs

		if iStatus < 0 then
			GoTo ErrorHandler
		endif

		'// create dialog if not yet done
'		if NOT pbCOADlgExists then
			puoCOADialog = CreateUnoDialog(GlobalScope.DialogLibraries.BadAss.COADialog)
			pbCOADlgExists = true
			'// load list with COA strings
			puoCOAListBox = puoCOADialog.getControl("COAList")
			puoCOASelectBtn = puoCOADialog.getControl("SelectBtn")
			puoCOASelectBtn.Model.Enabled = false			'// disable button until user selects item
			iCOACount = LoadCOAs()
			puoCOAListBox.addItems(pusCOAList, 0)
'		endif	'// end dialog non-existent conditional

		'// run dialog and get user selection
		Select Case puoCOADialog.Execute()
		Case 2		'// Select clicked
'			msgBox("COA Selected: " + pusCOASelected)
'			msgBox("Select clicked in Chart of Accounts")
			sLastCOA = pusCOASelected
		Case 0		'// Cancel
			pusCOASelected = ""		'// since cancel, Execute() event skipped
'			msgBox("Cancel or FormClose clicked")
		Case else
			msgBox("unevaluated COA dialog return")
		End Select

		sRetValue = Left(pusCOASelected, 4)
		
	endif	'// end use last COA conditional
	
NormalExit:

'msgBox("In Module2/GetCOA" + CHR(13)+CHR(10) _
'        + "'"+sRetValue+"'" + " selected")

	GetCOA = sRetValue
	exit function
	
ErrorHandler:
	sRetValue = ""
	GoTo NormalExit
	
end function 	'// end GetCOA		7/6/20
'/**/
'// H.bas
'//---------------------------------------------------------------
'// H - Split Transaction Dialog test procedure.
'//		8/26/22.	wmk.	10:24
'//----------------------------------------------------------------

public sub H()

'//	Usage.	macro call or
'//			call H()
'//
'// Entry.	Button click in dialog box		
'//			puoSTDialog = declared object for Split Transaction dialog
'//
'//	Exit.	gdSTTotalAmt = contrived total to split, say 150.00
'//			msgBox message displayed
'//
'// Calls.	STPubVarsReset
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code; adapted from G. bas.
'//	6/23/20.	wmk.	change to use fDialogReset.
'// 8/26/22.	wmk.	add intervening space to G. bas and H .bas in comments.
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim iStatus As Integer
dim dDummy As Double
dim sDummy As String

	'// code.
'	sDummy = Str(dDummy)
'	msgBox("string conversion of empty value is: " +sDummy)
'	GoTo NormalExit
	ON ERROR GOTO ErrorHandler	
	DialogLibraries.LoadLibrary("Standard")

'// Initialize ST dialog.
	puoSTDialog = CreateUnoDialog(DialogLibraries.Standard.SplitDialog)
'XRay puoSTDialog	
	'// emulate dialog being fired with Total to split initialized
	'// in public vars
	gdSTTotalAmt = 150.00
	iStatus = fSTDialogReset()		'// attempt to reset/initialize dialog
'	if iStatus < 0 then
'		GoTo ErrorHandler
'	endif
	dim oFld As Object
	oFld = puoSTDialog.getControl("HasFocusFld")
	oFld.Text = "Gotcha"
	Select Case puoSTDialog.Execute()
	Case 2		'// Record & Finish clicked
		msgBox("<Done>h clicked in Split Transaction")
	Case 0		'// Cancel
		iStatus = STPubVarsReset()	'// reset everything
		msgBox("ST dialog Cancel or FormClose clicked")
	Case else
		msgBox("unevaluated dialog return")
	End Select

'msgbox("Date entered: " + gsETDate)
'msgbox("Decription: " + gsETDescription)
'msgBox("Amount: " + gsETAmount)
'msgBox("Debit Account: " + gsETAcct1)
'msgBox("Credit Account: " + gsETAcct2)
'msgBox("Reference: " + gsETRef)

puoSTDialog.dispose()
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("in H - error initializing Split dialog.")
	GoTo NormalExit
	
end sub		'// end H	8/26/22
'/**/
'// LoadBadAssLib.bas
'//---------------------------------------------------------------
'// LoadBadAssLib - Load BadAss Dialog library.
'//		7/6/20.	wmk.	12:45
'//---------------------------------------------------------------

public function LoadBadAssLib() As Integer

'//	Usage.	iVal = LoadBadAssLib()
'//
'// Entry.	pbDlgLibLoaded = true if Standard dialogs already loaded
'//							 false otherwisd
'//
'//	Exit.	iVal = 0 - normal return
'//				   -1 - error encountered in load
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/30/20.		wmk.	original code
'//	7/6/20.		wmk.	BadAss load unconditional; change to BasicLibraries
'//
'//	Notes. Adapted from LoadStdLibrary to use My Macros library BadAss.
'//

'//	constants.

'//	local variables.
dim iRetValue As Integer

	'// code.
	
	iRetValue = -1		'// set abnormal return
	ON ERROR GOTO ErrorHandler
	
'	if NOT pbDlgLibLoaded then
		GlobalScope.BasicLibraries.LoadLibrary("BadAss")
		pbDlgLibLoaded = true
'	endif	'// end library not loaded conditional

	iRetValue = 0		'// set normal return
	
NormalExit:
	LoadBadAssLib = iRetValue
	exit function
	
ErrorHandler:
	GoTo NormalExit

end function 	'// end LoadBadAssLib	7/6/20
'/**/
'// LoadCOAs.bas
'//---------------------------------------------------------------
'// LoadCOAs - Load COA account#s for list boxes.
'//		7/6/20.	wmk.	07:30
'//---------------------------------------------------------------

public function LoadCOAs() as Integer

'//	Usage.	iVal = LoadCOAs()
'//
'// Entry.	pusCOA_Assets = array allocated for Asset COAs
'//			pusCOA_Liabilities = array for Liability COAs
'//			pusCOA_Income = array for Income COAs
'//			pusCOA_Expenses = array for Expense COAs
'//			pusCOAList = array for COA list for listbox
'//			COA_SHEET = sheet name of COA worksheet
'//			COA_COLSTART = starting column of COA loadable list
'//			COA_ROWSTART = starting row of COA loadable list
'//			sCOA_1STROWID = string to match verifying 1st row COA list
'//			sCOA_LASTROWID = string to match verifying last row COA list
'//			COA_NASSETS = Assets COA entry count
'//
'//	Exit.	iVal = 0 - no errors encountered
'//				  -1 - bad load
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.	wmk.	original code; stub
'//	6/14/20.	wmk.	code completed.
'//	6/30/20.	wmk.	comments updated.
'//	7/6/20.		wmk.	modified to get COA entry count from 2 columns to
'//						right of "BEGIN COA" (COA_COLROWS)
'//
'//	Notes. Initial version knows how many rows to expect by public constants
'//	set internally for each accounting category. Future version can be
'// generalized to load "n" rows where "n" is specified at the head of
'// the COA table in the Chart of Accounts.


'//	constants.
const ERRNOTCOATBL=-2		'// COA table not where expected

'//	local variables.
dim iRetValue As Integer
dim iAsset As Integer		'// asset COA counter
dim iLiability As Integer	'// liability COA counter
dim iIncome As Integer		'// income COA counter
dim iExpense As Integer		'// expense COA counter
dim oDoc As Object			'// component
dim oSheets As Object		'// sheets array
dim oCOASheet As Object		'// COA sheet
dim oCOACell As Object		'// COA cell general object
dim oCOACntCell	As Object	'// cell with COA count
dim bIsCOATable As Boolean	'// is COA table flag
dim sCOAList(0) As String	'// COA list array

dim iRowsExpected As Integer		'// expected row count
dim i	As Integer					'// loop count
dim oAcctCell As Object				'// account name cell
dim bTableEndFound As Boolean		'// end of table found
dim lCurrRow As Long				'// current row index
dim iBrkPt As Integer				'// breakpoint finder


'//----------in sub/function error handling setup-------------------------------
'// Error Handling setup snippet.
	const sMyName="LoadCOAs"
	dim lLogRow as long		'// cell row working on
	dim lLogCol as Long		'// cell column working on
	dim iLogSheetIx as Integer	'// sheet index module working on
	dim oLogRange as new com.sun.star.table.CellRangeAddress

'// LogError setup snippet.
	const csERRUNKNOWN="ERRUNKNOWN"
'// add additional error code strings here...
	Dim sErrName as String		'// error name for LogError
	Dim sErrMsg As String		'// error message for LogError

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sErrName = "ERRUNKNOWN"
	sErrMsg = "In LoadCOAs - unprocessed error."

	'// set up access to COA sheet
	oDoc = ThisComponent		'// get into document
	oSheets = oDoc.getSheets()
	oCOASheet = oSheets.getByName(COA_SHEET)
	
'//-----------in sub/function error handling initialization-------------
'//	dim oDoc as Object		'// ThisComponent
'//	dim oSheets as Object	'// oDoc.Sheets
'//	code to put sheet index into oLogRange.Sheet
	oLogRange.Sheet = oCOASheet.RangeAddress.Sheet
	lLogCol = COA_COLSTART
	lLogRow = COA_ROWSTART
	oLogRange.StartColumn = lLogCol
	oLogRange.StartRow = lLogRow
	oLogRange.EndRow = lLogRow
	oLogRange.EndColumn = lLogCol
	Call ErrLogSetup(oLogRange, sMyName)
	ErrLogSetRecording(true)
'//-----------end in sub/function error handling initialization---------

	
	iAsset = 0			'// clear counters
	iLiability = 0
	iIncome = 0
	iExpense = 0
	
	'// get COA table 1st row and verify
	oCOACell = oCOASheet.getCellByPosition(COA_COLSTART, COA_ROWSTART)
	bIsCOATable = (StrComp(Trim(oCOACell.String), sCOA_1STROWID)=0)
	if NOT bIsCOATable then
		iRetValue = ERRNOTCOATBL
		sErrName = "ERRNOTCOATBL"
		sErrMsg = "LoadCOAs/Invalid COA table"
		GoTo ErrorHandler
	endif

	oCOACntCell = oCOASheet.getCellByPosition(COA_COLROWS, COA_ROWSTART)

'	iRowsExpected = COA_NASSETS + COA_NLIABILITIES + COA_NINCOME _
'					+ COA_NEXPENSES
	iRowsExpected = oCOACntCell.Value
	if iRowsExpected < 1 then
		sErrName = "ERRCOACOUNT"
		sErrMsg = "LoadCOAs/Invalid COA count in table"
		GOTO ErrorHandler
	endif
redim pusCOAList(iRowsExpected-1)		'// minus 1, since 0-based

	lCurrRow = COA_ROWSTART + 1
	'// loop on rows until termination string found "****"
	for i = 0 to iRowsExpected
		oAcctCell = oCOASheet.getCellByPosition(COA_COLSTART, lCurrRow)
		bTableEndFound = (StrComp(oAcctCell.String, sCOA_LASTROWID)=0)
		if bTableEndFound then
			if i <> iRowsExpected then
				iRetValue = ERRCOACOUNT
				sErrName = "ERRCOACOUNT"
				sErrMsg = "LoadCOAs/Table entry count mismatch"
				GoTo ErrorHandler
			else
				exit for
			endif	'// end wrong row count conditional			
		endif	'// end table end found conditional
	
		'// combine and store COA strings
		oCOACell = oCOASheet.getCellByPosition(COA_COLSTART+1, lCurrRow)
		pusCOAList(i) = oCOACell.String +" " + oAcctCell.String
		lCurrRow = lCurrRow + 1
	next i

	'// check for premature table end
	if i < iRowsExpected then
		iRetValue = ERRCOACOUNT
		sErrName = "ERRCOACOUNT"
		sErrMsg = "LoadCOAs/Table entry count mismatch"
		GoTo ErrorHandler
	endif
	
iBrkPt = 1		'// no errors		
					
	iRetValue = 0		'// set normal return
	
NormalExit:
	call ErrLogDisable()		'// disable error logging
	iRetValue = iRowsExpected	'// return row count
	LoadCOAs = iRetValue
	exit function
	
ErrorHandler:
	Call LogError(sErrName, sErrMsg)
	GoTo NormalExit
	
end function 	'// end LoadCOAs	7/6/20
'/**/
'// LoadStdLibrary.bas
'//---------------------------------------------------------------
'// LoadStdLibrary - Load Standard Dialog library.
'//		7/5/20.	wmk.	22:00
'//---------------------------------------------------------------

public function LoadStdLibrary() As Integer

'//	Usage.	iVal = LoadStdLibrary()
'//
'// Entry.	pbDlgLibLoaded = true if Standard dialogs already loaded
'//							 false otherwisd
'//
'//	Exit.	iVal = 0 - normal return
'//				   -1 - error encountered in load
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/14/20.	wmk.	original code
'//	7/5/20.		wmk.	bug fix; LoadStdLib return corrected to LoadStdLibrary
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim iRetValue As Integer

	'// code.
	
	iRetValue = -1		'// set abnormal return
	ON ERROR GOTO ErrorHandler
	
	if NOT pbDlgLibLoaded then
		DialogLibraries.LoadLibrary("Standard")
		pbDlgLibLoaded = true
	endif	'// end library not loaded conditional

	iRetValue = 0		'// set normal return
	
NormalExit:
	LoadStdLibrary = iRetValue
	exit function
	
ErrorHandler:
	GoTo NormalExit

end function 	'// end LoadStdLibrary	7/5/20
'/**/
'// PlaceSplitTotal.bas
'//------------------------------------------------------------------
'// PlaceSplitTotal - Itemize split transaction in total COA account.
'//		7/8/20.	wmk.	15:00
'//------------------------------------------------------------------

public function PlaceSplitTotal(poGLRange As Object, psAcct As String) As Integer

'//	Usage.	iVal = PlaceSplitTotal(oGLRange, sAcct)
'//
'//		oGLRange = RangeAddress information for GL sheet
'//		sAcct = Total COA account from which items split
'//
'// Entry.	oGLRange.Sheet = index of ledger sheet
'//					.StartColumn = index of starting column
'//					.StartRow = index of first "split" row
'//					.EndColumn = index of ending column
'//					.EndRow = index of last "split" row
'//
'//	Exit.	iVal = 0 - normal return
'//				   ERRCONTINUE - nonfatal error; continue processing
'//				   ERRSTOP - fatal error; stop processing transactions
'//
'// Calls.	GetMonthName, GetTransSheetName, GetInsRow, SetSumFormula
'//
'//	Modification history.
'//	---------------------
'//	6/4/20.		wmk.	original code
'//	6/6/20.		wmk.
'//	6/27/20.	wmk.	superfluous "if" removed
'//	7/8/20.		wmk.	code to remove extra row in source COA transaction	
'//
'//	Notes.
'// This function adds lines to the splitting account category
'// then copies lines from the original "split" transaction
'// except for the Total transaction line; then moves the
'// nonempty Debit or Credit range into the opposing column
'// and ensures that the SUM formula in the last "split" row
'// is corrected to sum the moved data.
'//

'//	constants.
const FORMULA=3		'// FORMULA type on cell

'//	local variables.
dim iRetValue As Integer		'// returned value
dim oDoc As Object				'// ThisComponent
dim oSheets As Object			'// Doc.getSheets()
dim iSheetIx As Integer			'// sheet index of ledger
dim oGLSheet As Object			'// ledger sheet object
dim sSplitAcct As String		'// COA of account splitting from
dim iBrkPt As Integer			'// easy breakpoint

'//	GL transaction variables.
dim oGLCellDate	As Object		'// Date cell from GL
dim sStyle	As String			'// cell style
dim sDate	As String			'// date field string
dim sMonth	As String			'// month name of date (e.g. "January")
dim l1stRow	As Long				'// first "split" row
dim lLastRow As Long			'// last "split" row
dim lRowCount	As Long			'// transaction row count between "split"s

'//	COA category sheet variables.
dim sAcctCat2 As String			'// COA account category total transaction
dim oCat2Sheet As Object		'// COA account sheet for total
dim lCat2InsRow	As Long			'// COA sheet insertion row
dim oDebitCell as Object
dim bMoveDebits as Boolean
dim oFormulaCell As Object
dim sFormula As String
dim sFormStart as String
dim lCat2CurrRow as Long	'// insertion sheet current row
dim li as Long				'// long loop counter
dim oDateCell as Object		'// Date cell object

	'// code.
	sSplitAcct = psAcct			'// copy COA account string
	
	iRetValue = 0
	
	PlaceSplitTotal = iRetValue
	
'	if true then
'		Exit Function
'	endif

'//---------------begin PlaceSplitTotal here------------------------
'//	poGLSheet is the ledger sheet with the whole transaction
'// sSplitAcct is the account that has the total COA
		'//==========================================================
'//-----------------error handling setup-----------------------------
dim sErrCallerMod 	As String	'// caller routine name
dim iErrSheetIx		As Integer	'// caller sheet index
dim lErrColumn		As Long		'// caller error focus column
dim lErrRow			As Long		'// caller error focus row
dim sMyName			As String	'// this routine name
	'//*-----------------preserve error handling for caller-----------
	'// preserve caller error settings.
	sErrCallerMod = ErrLogGetModule()
	iErrSheetIx = ErrLogGetSheet()
	ErrLogGetCellInfo(lErrColumn, lErrRow)
	'// set up new error settings.
	sMyName = "PlaceSplitTotal"
	ErrLogSetModule(sMyName)
	ErrLogSetSheet(poGLRange.Sheet)
	ErrLogSetCellInfo(poGLRange.StartColumn, poGLRange.StartRow)
'//*----------end error handling setup-------------------------------

		'//	insertCells...
'// entry.	sSplitAcct is the COA from the total line entry
'//         oGLCellDate.Text.String is the Date field from the totals
'//			 line entry
'//			oSheets object is the Sheets[] array this component

	oDoc = ThisComponent			'// set component and access objects
	oSheets = oDoc.getSheets()
	iSheetIx = poGLRange.Sheet
	oGLSheet = oSheets.getByIndex(iSheetIx)
	
	'// extract date information from ledger entry
	l1stRow = poGLRange.StartRow	'// first row of "split"
	lLastRow = poGLRange.EndRow		'// last row of "split"
	lRowCount = poGLRange.EndRow - l1stRow - 1			'// transaction lines with COAs
	oGLCellDate = oGLSheet.getCellByPosition(COLDATE, l1stRow)

	'// set up parameters for GetTransSheetName and GetInsRow
	sStyle = oGLCellDate.Text.String
	sDate = oGLCellDate.Text.String
	sMonth = GetMonthName(sDate)		'// get month search name
	sAcctCat2 = GetTransSheetName(sSplitAcct)	'// get sheet name matching COA

	'// access target COA sheet and set up insertion row
	oCat2Sheet = oSheets.getByName(sAcctCat2)
iBrkpt=1		'// oCat2Sheet.RangeAddress.Sheet = sheet index
ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)	'// set sheet for errors
	lCat2InsRow = GetInsRow(oCat2Sheet, sSplitAcct, sMonth)

	if lCat2InsRow < 0 then
		sErrName = "ERRCOAINSERT"
		sErrMsg = "Unable to insert Total rows in split account"
		GoTo ErrHandler
	endif
		
	
Dim oCOARange As New com.sun.star.table.CellRangeAddress

	'// set up oCOARange object
	'// copy entire transaction just like if doing it manually
	'// move filled values to other column (Debit > Credit) or
	'//		(Credit to Debit)
	'// delete total row from transaction
	'// set sum formula in COA field of last "split" row
	
	'// at this point l1stRow = 1st row with "split"
	'//               lLastRow = last row with "split
	'//				  set lStartColumn = COLDATE
	'//				  set lEndColumn = COLREF
	'//				  set sheet number (integer) = index of GL sheet
	'//				  set target start row at lCat2InsRow
	'//				  set target start column to COLDATE
	'//				  set number of rows to lRowCount + 1
	'// Doc = ThisComponent
	'// Sheet = Doc.Sheets(0)
	'//	oSel = Doc.getCurrentSelection()
	'//	oRange = oSel.RangeAddress			'// extract range information
iBrkPt=1		
	oCOARange.Sheet = oCat2Sheet.RangeAddress.Sheet
	oCOARange.StartColumn = COLDATE
	oCOARange.EndColumn = COLREF
	oCOARange.StartRow = lCat2InsRow	'// we are inserting into the 2nd sheet
	oCOARange.EndRow = lCat2InsRow + lRowCount + 2	'// start+transaction rowcount+1
ErrLogSetCellInfo(COLDATE, lCat2InsRow)
	oCat2Sheet.insertCells(oCOARange,_
							com.sun.star.sheet.CellInsertMode.ROWS)
dim lSplitEnd As Long		'// save split end in target sheet
	lSplitEnd = oCOARange.EndRow
		
	'// now copy original transaction split range address to same range
	'// overwriting blank rows;
dim oTargetCellAddress As new com.sun.star.table.CellAddress
	oTargetCellAddress.Sheet = oCOARange.Sheet
	oTargetCellAddress.Column = COLDATE
	oTargetCellAddress.Row = lCat2InsRow
		
dim oGLCellRange As new com.sun.star.table.CellRangeAddress
	oGLCellRange.Sheet = poGLRange.Sheet
	oGLCellRange.StartColumn = COLDATE
	oGLCellRange.EndColumn = COLREF
	oGLCellRange.StartRow = l1stRow
	oGLCellRange.EndRow = lLastRow

'// ErrLogGetSheet - Get sheet index from error log globals.
ErrLogSetSheet(oTargetCellAddress.Sheet)
ErrLogSetCellInfo(oTargetCellAddress.Column, oTargetCellAddress.Row)

	'// copy the entire split transaction to the COA sheet
	'// then delete the total row;
	oGLSheet.copyRange(oTargetCellAddress, oGLCellRange)
		
	'// use removeCells to remove Total row
	'// uses same parameters as insertCells; only change Row range
	oCOARange.Sheet = oTargetCellAddress.Sheet	
	oCOARange.StartRow = lCat2InsRow + 1	'// total row
	oCOARange.EndRow = oCOARange.StartRow
	
ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)
ErrLogSetCellInfo(COLDATE, oCOARange.EndRow)
	
	oCat2Sheet.removeRange(oCOARange,_
						com.sun.star.sheet.CellDeleteMode.ROWS)
	lSplitEnd = lSplitEnd - 1		'// adjust split end for removed row
		
iBrkPt=1
							
	'// then move the Debit/Credit values to the other column
	'// change background color of colored cells to LTGREEN
	'// test which column is .Type empty cells
	'// then move the Debit/Credit values to the empty column
	oDebitCell = oCat2Sheet.getCellByPosition(COLDEBIT,_
					oCOARange.StartRow)
	oTargetCellAddress.Row = oDebitCell.CellAddress.Row
	oCOARange.EndRow = lSplitEnd - 2	'// 2 back; no Total row and not end "split"

	bMoveDebits = (oDebitCell.Type <> EMPTY)
		'	EMPTY,VALUE,TEXT,FORMULA

	'// if there is something in the Debits column; the Debits move to Credits		
	if bMoveDebits then	'// set range to debits column
		oTargetCellAddress.Column = COLCREDIT
		oCOARange.StartColumn = COLDEBIT
		oCOARange.EndColumn = COLDEBIT
	else				'// otherwise, the Credits move to Debits
		oTargetCellAddress.Column = COLDEBIT
		oCOARange.StartColumn = COLCREDIT
		oCOARange.EndColumn = COLCREDIT
	endif	'// end move Debits conditional

'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
ErrLogSetCellInfo(oTargetCellAddress.Column, oTargetCellAddress.Row)

	'// move the column containing values to opposing column
	oCat2Sheet.moveRange(oTargetCellAddress, oCOARange)

		
	'// set .formula on COLACCT of last line to
	'//	 "=SUM()" previous cells above in either Debit or Credit

ErrLogSetCellInfo(COLACCT, lSplitEnd-1)

	'// row is now lSplitEnd-1 since Total row deleted	
	oFormulaCell = oCat2Sheet.getCellByPosition(COLACCT, lSplitEnd-1)
		
	'// Always force SUM formula into last row COLACCT field
'	if oFormulaCell.Type = FORMULA then

	'// set up RangeAddress object for SetSumFormula
dim oNewSplitRange as new com.sun.star.table.CellRangeAddress
	oNewSplitRange.Sheet = oCat2Sheet.RangeAddress.Sheet
	oNewSplitRange.StartRow = lCat2InsRow		'// first "split" row
	oNewSplitRange.EndRow = lSplitEnd-1			'// last "split" row (minus 1, row deleted)
	oNewSplitRange.StartColumn = COLDATE		'// Date column
	oNewSplitRange.EndColumn = COLREF			'// Reference column
	sFormula = oFormulaCell.String			'// just for kicks

ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)
ErrLogSetCellInfo(COLACCT, oNewSplitRange.EndRow)
			
	sFormula = SetSumFormula(oNewSplitRange, bMoveDebits)		
	oFormulaCell.String = sFormula			'// set formula in cell
	oFormulaCell.setFormula(sFormula)
	'// loop changing Date fields' color to LTGREEN
ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)
ErrLogSetCellInfo(COLDATE, oNewSplitRange.StartRow)
	lRowCount = oNewSplitRange.EndRow - oNewSplitRange.StartRow + 1
	lCat2CurrRow = oNewSplitRange.StartRow
	for li = 0 to lRowCount-1
		oDateCell = oCat2Sheet.getCellByPosition(COLDATE, lCat2CurrRow)
		oDateCell.CellBackColor = LTGREEN
		lCat2CurrRow = lCat2CurrRow + 1	
	next li					

'// add code to remove superfluous row at bottom
	'// row to remove is at lSplitEnd								'// mod070820
	'// use removeCells to remove extraneous						'// mod070820
	'// uses same parameters as insertCells; only change Row range	'// mod070820
	oCOARange.Sheet = oTargetCellAddress.Sheet						'// mod070820
	oCOARange.StartRow = lSplitEnd	'// blank row					'// mod070820
	oCOARange.EndRow = lSplitEnd									'// mod070820
	
ErrLogSetSheet(oCat2Sheet.RangeAddress.Sheet)						'// mod070820
ErrLogSetCellInfo(COLDATE, lSplitEnd)								'// mod070820	
	oCat2Sheet.removeRange(oCOARange,_
						com.sun.star.sheet.CellDeleteMode.ROWS) 		'// mod070820

		'// update sheet modified date cell							'// mod060620
		oCOARange.Sheet = oCat2Sheet.RangeAddress.Sheet				'// mod060620
		Call SetSheetDate(oCOARange, MMDDYY)						'// mod060620
		oDateCell = oCat2Sheet.getCellByPosition(COLDEBIT, DATEROW)	'// mod060620
		oDateCell.CellBackColor = LTGREEN							'// mod060620

	GoTo NormalExit

ErrHandler:
	'// here because insertion row returned <0; unable to insert rows
	'// iStatus is bad value returned
	iRetValue = iStatus
	Select Case iStatus
	Case ERRCOANOTFOUND
		sErrCode = csERRCOANOTFOUND
		sErrMsg = "Account "+sBadCOA2+" not found"
	Case ERROUTOFROOM
		sErrCode = csERROUTOFROOM
		sErrMsg = "Account "+sBadCOA2+" month "+sMonth+" not enough rows"_
					+" to insert"
	Case ERRCOACHANGED
		sErrCode = csCOACHANGED
		sErrMsg = "Account "+sBadCOA2+" month "+sMonth+" not found"
	Case else
		sErrCode = csERRUNKNOWN
		sErrMsg = "Account "+sBadCOA2+"Undocumented error"
	end Select

	Call LogError(sErrCode, sErrMsg)			
	
'//-----------------end PlaceSplitTotal here----------------------------
NormalExit:
	'//*-----------------restore error handling for caller-----------
	'// restore entry error settings.
	ErrLogSetModule(sErrCallerMod)
	ErrLogSetSheet(iErrSheetIx)
	ErrLogSetCellInfo(lErrColumn, lErrRow)
	'//*----------end error handling restore----------------------------

	'// iStatus = 0, no error
	'//				ERRCONTINUE (-1) - error; continue with next transaction
	'//				ERRSTOP (-2) - error; stop processing transactions

	PlaceSplitTotal = iRetValue

end Function		'// end PlaceSplitTotal		7/8/20
'/**/
'// RunCOADlg.bas
'//---------------------------------------------------------------
'// RunCOADlg - Run COA Dialog.
'//		7/5/20.	wmk.	15:00
'//---------------------------------------------------------------

public sub RunCOADlg()

'//	Usage.	macro call or
'//			call RunCOADialog()
'//
'// Entry.	Button click in dialog box		
'//			puoCOADialog = declared object for COA dialog
'//			puoCOAListBox = list box object
'//			puoSelectBtn = declared object for <Select> button
'//			pusCOASelected = reserved for COA string selected
'//			current Component has sheet "Chart of Accounts" with COA table
'//			BadAss library exists
'//
'//	Exit.	msgBox message displayed
'//			pusCOASelected = COA selected
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/13/20.	wmk.	original code
'//	6/14/20.	wmk.	call LoadCOAs and run form added; code
'//						modified to use puoCOADialog object
'//	6/15/20.	wmk.	name of COA list dialog changed; conbtrols
'//						labeled as SelectButton, CancelButton
'//	7/5/20.		wmk.	Entry dependencies updated
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
'dim oListBox As com.sun.star.awt.UnoControlListBox
dim oListBox As Object

Dim Dlg As Object
dim iCOACount As Integer

'// code.
	ON ERROR GOTO ErrorHandler
	BasicLibraries.LoadLibrary("Standard")
	BasicLibraries.LoadLibrary("BadAss")

'// Initialize COA Dialog
	puoCOADialog = CreateUnoDialog(DialogLibraries.BadAss.COADialog)
	puoCOAListBox =puoCOADialog.getControl("COAList")
	puoCOASelectBtn = puoCOADialog.getControl("SelectBtn")
	puoCOASelectBtn.Model.Enabled = false			'// disable button until user selects item
	iCOACount = LoadCOAs()
	puoCOAListBox.addItems(pusCOAList, 0)

	Select Case puoCOADialog.Execute()
	Case 2		'// Select clicked
		msgBox("Select clicked in Chart of Accounts")
	Case 0		'// Cancel
		pusCOASelected = ""		'// since cancel, Execute() event skipped
		msgBox("Cancel or FormClose clicked")
	Case else
		msgBox("unevaluated dialog return")
	End Select
	msgBox("In Module2/F" + CHR(13)+CHR(10) _
    	    + "'"+pusCOASelected+"'" + " selected")

NormalExit:
	puoCOAListBox.dispose()
	exit sub
	
ErrorHandler:
	msgBox("RunCOADlg - unprocessed error.")
	GoTo NormalExit
	
end sub		'// end RunCOADlg	7/5/20
'/**/
'// STAcct10FocusListener.bas
'//---------------------------------------------------------------
'// STAcct10FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct10FocusListener()

'//	Usage.	macro call or
'//			call STAcct10FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld10"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct10FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct10FocusListener	6/20/20
'/**/
'// STAcct1FocusListener.bas
'//---------------------------------------------------------------
'// STAcct1FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	15:45
'//---------------------------------------------------------------

public sub STAcct1FocusListener()

'//	Usage.	macro call or
'//			call STAcct1FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld1"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct1FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct1FocusListener	6/20/20
'/**/
'// STAcct1LostFocus.bas
'//---------------------------------------------------------------
'// STAcct1LostFocus - Handle AcctFldn LostFocus event.
'//		6/24/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAcct1LostFocus()

'//	Usage.	macro call or
'//			call STAcct1LostFocus()
'//
'// Entry.	AcctFld1..AcctFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "AcctFld1")
'//
'//	Exit.	gsCOAs(0) = text from AcctFld1..10
'//
'// Calls.	STAcctLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	bug fix storing COA; add code to check/enable
'//						next row if this row complete
'//	6/24/20.	wmk.	common code moved to STAcctLostFocus; generic
'//						call set up to use common code function
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into.


'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "AcctFld1"					'// clear object focus tracker
	iStatus = STAcctLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct1LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct1LostFocus	6/24/20
'/**/
'// STAcct2FocusListener.bas
'//---------------------------------------------------------------
'// STAcct2FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct2FocusListener()

'//	Usage.	macro call or
'//			call STAcct2FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld2"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct2FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct2FocusListener	6/20/20
'/**/
'// STAcct3FocusListener.bas
'//---------------------------------------------------------------
'// STAcct3FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct3FocusListener()

'//	Usage.	macro call or
'//			call STAcct2FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld3"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct3FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct3FocusListener	6/20/20
'/**/
'// STAcct4FocusListener.bas
'//---------------------------------------------------------------
'// STAcct4FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct4FocusListener()

'//	Usage.	macro call or
'//			call STAcct2FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld4"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct4FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct4FocusListener	6/20/20
'/**/
'// STAcct5FocusListener.bas
'//---------------------------------------------------------------
'// STAcct5FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct5FocusListener()

'//	Usage.	macro call or
'//			call STAcct5FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld5"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct5FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct5FocusListener	6/20/20
'/**/
'// STAcct6FocusListener.bas
'//---------------------------------------------------------------
'// STAcct6FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct6FocusListener()

'//	Usage.	macro call or
'//			call STAcct6FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld6"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct6FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct6FocusListener	6/20/20
'/**/
'// STAcct7FocusListener.bas
'//---------------------------------------------------------------
'// STAcct7FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct7FocusListener()

'//	Usage.	macro call or
'//			call STAcct7FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld7"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct7FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct7FocusListener	6/20/20
'/**/
'// STAcct8FocusListener.bas
'//---------------------------------------------------------------
'// STAcct8FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct8FocusListener()

'//	Usage.	macro call or
'//			call STAcct8FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld8"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct8FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct8FocusListener	6/20/20
'/**/
'// STAcct9FocusListener.bas
'//---------------------------------------------------------------
'// STAcct9FocusListener - Handle AcctFldn GotFocus event.
'//		6/20/20.	wmk.	16:00
'//---------------------------------------------------------------

public sub STAcct9FocusListener()

'//	Usage.	macro call or
'//			call STAcct9FocusListener()
'//
'// Entry.	AcctFld1..AcctFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AcctFld object name
'//			gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	modified and reduced to set global control name
'//						MY_NAME local constant introduced
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AcctFld9"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAcct9FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAcct9FocusListener	6/20/20
'/**/
'// STAcctLostFocus.bas
'//---------------------------------------------------------------
'// STAcctLostFocus - Handle any AcctFld LostFocus event.
'//		7/7/20.	wmk.	12:30
'//---------------------------------------------------------------

public function STAcctLostFocus() As Integer

'//	Usage.	macro call or
'//			call STAcctLostFocus()
'//
'//	[Parameters.	sFldName = name of dialog object lost focus]
'//
'// Entry.	AcctFld1..AcctFld10 lost focus
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gsCOAs(0) = text from AcctFld1..10
'//
'// Calls.	STFld7ToIndex
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; cloned from STAcct1LostFocus
'//	6/21/20.	wmk.	check for total reached before activating next row
'//	7/7/20.		wmk.	bug fix where total amount reached, but <Done> button
'//						deactivated
'//
'//	Notes. gsSTObjFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into.


'//	constants.

'//	local variables.
dim iRetValue 		As Integer			'// returned status
dim iStatus			As Integer			'// general status
dim iRowIx			As Integer			'// index of present row
dim oDlgControl		As Object			'// field ptr
dim bRowComplete	As Boolean			'// row complete flag
dim bDlgComplete	As Boolean			'// dialog complete flag
dim sFldName 		As String			'// local copy of field name

	'// code.
	iRetValue = -1						'// set bad return
	ON ERROR GOTO ErrorHandler
	sFldName = gsSTObjFocus				'// use local copy of field name
'	sFldName = psFldName				'// use local copy of field name
	oDlgControl = puoSTDialog.getControl(sFldName)
	iRowIx = STFld7ToIndex(sFldName)
	gsSTCOAs(iRowIx) = oDlgControl.Text

	bRowComplete = STRowComplete(iRowIx)
	if bRowComplete then
		bDlgComplete = STCheckComplete()

		if gdSTAccumTot < gdSTTotalAmt then		
			'// check for next row Enabled; if not, enable it
			iStatus = STEnableNextRow(iRowIx)
			if iStatus < 0 then
				GoTo ErrorHandler
			endif
		endif '// end total not met yet conditional
		
	endif	'// end row complete conditional
	
	'// activate/deactivate <Done> control if dialog complete
	oDlgControl = puoSTDialog.getControl("DoneBtn")
	if bDlgComplete then
		oDlgControl.Model.Enabled = true
	else
		oDlgControl.Model.Enabled = false
	endif
	
	iRetValue = 0
	gsSTObjFocus = ""					'// clear object focus tracker

NormalExit:
	STAcctLostFocus = iRetValue
	exit function
	
ErrorHandler:
	msgBox("In STAcctLostFocus - unprocessed error")
	GoTo NormalExit
	
end function		'// end STAcctLostFocus	6/21/20
'/**/
'// STAmt10FocusListener.bas
'//---------------------------------------------------------------
'// STAmt10FocusListener - Handle AmtFldn GotFocus event.
'//		6/20/20.	wmk.	20:45
'//---------------------------------------------------------------

public sub STAmt10FocusListener()

'//	Usage.	macro call or
'//			call STAmt10FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAcct3FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld10"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt10FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt10FocusListener	6/20/20
'/**/
'// STAmt10LostFocus.bas
'//---------------------------------------------------------------
'// STAmt10LostFocus - Handle AmtFld10 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt10LostFocus()

'//	Usage.	macro call or
'//			call STAmt10LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld10"
'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

dim bTotalMet		As Boolean		'// accum = split flag
	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld10"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt10LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt10LostFocus	6/23/20
'/**/
'// STAmt1FocusListener.bas
'//---------------------------------------------------------------
'// STAmt1FocusListener - Handle AmtFldn GotFocus event.
'//		6/23/20.	wmk.	05:00
'//---------------------------------------------------------------

public sub STAmt1FocusListener()

'//	Usage.	macro call or
'//			call STAmt1FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAcct3FocusListener
'// 6/22/20.	wmk.	call to STSetObjFocus to set this control name
'//						in dialog for later retrieval
'//	6/23/20.	wmk.	reverted code to near original; eliminated
'//						STSetObjFocus call, since throws dialog into
'//						infinite loop when setting the field in the
'//						dialog, the "Got Focus" event gets re-invoked
'//						when returning from the "HasFocus" control;
'//						this may be conquerable with a flag in the 
'//						publics that indicates recursive callback and
'//						just exits, or skips ahead to code following
'//						STSetObject call
'//
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.

'//	local variables.
dim iStatus		As Integer			'// general status
dim oDlgControl	As Object			'// dialog control object

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler

'	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	gsSTObjFocus = "AmtFld1"
	
	if true then
		GoTo NormalExit
	endif
'//---------------------------------------------------------------------------------	
	oDlgControl = puoSTDialog.getControl("HasFocusFld")
msgBox("In STAmt1FocusListener back from .getControl on entry.." + CHR(13)+CHR(10) _
+	"stored name = '" + oDlgControl.Text + "'")
'	sFocusName = oDlgControl.Text

'	if Len(sFocusName) > 0 then
'		iRetValue = -2		'// set busy flag
'		GoTo ErrorHandler
'	endif		'// end HasFocus in use conditional

	oDlgControl.Text = "AmtFld1"		'// set control name
msgBox("In STAmt1FocusListener ready to exit.." + CHR(13)+CHR(10) _
+	"stored name = '" + oDlgControl.Text + "'")
	GoTo NormalExit
	
'//------------------------------------------------------------------------------
	iStatus = STSetObjFocus("AmtFld1")
msgBox("In STAmt1FocusListener, back from STSetObjectFocus..." + CHR(13)+CHR(10) _
+	"iStatus = " + iStatus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

	iStatus = 0
	GoTo NormalExit
'//--------------------------------------------------------------	
ErrorHandler:
	msgBox("In STAmt1FocusListener - unprocessed error")

NormalExit:
	exit sub		
	
end sub		'// end STAmt1FocusListener	6/23/20
'/**/
'// STAmt1LostFocus.bas
'//---------------------------------------------------------------
'// STAmt1LostFocus - Handle AmtFld1 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt1LostFocus()

'//	Usage.	macro call or
'//			call STAmt1LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld1"

'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld1"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt1LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt1LostFocus	6/23/20
'/**/
'// STAmt2FocusListener.bas
'//---------------------------------------------------------------
'// STAmt2FocusListener - Handle AmtFldn GotFocus event.
'//		6/20/20.	wmk.	20:45
'//---------------------------------------------------------------

public sub STAmt2FocusListener()

'//	Usage.	macro call or
'//			call STAmt2FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAmt1FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld2"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt2FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt2FocusListener	6/20/20
'/**/
'// STAmt2LostFocus.bas
'//---------------------------------------------------------------
'// STAmt2LostFocus - Handle AmtFld2 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt2LostFocus()

'//	Usage.	macro call or
'//			call STAmt2LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld2"

'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld2"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt2LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt2LostFocus	6/23/20
'/**/
'// STAmt3FocusListener.bas
'//==============================---------------------------------
'// STAmt3FocusListener - Handle AmtFldn GotFocus event.
'//		6/23/20.	wmk.	08:30
'//---------------------------------------------------------------

public sub STAmt3FocusListener()

'//	Usage.	macro call or
'//			call STAmt2FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAmt1FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld3"

'//	local variables.
dim iStatus		As Integer			'// general status
dim oHasFocus	As Object
	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
	oHasFocus = puoSTDialog.getControl("HasFocus")
	oHasFocus.Text = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
'// Note: code will loop forever in any GotFocus event if a msgBox
'// is executed, since the message box takes focus away from the
'// dialog implementing the listener, then refocuses on it...
'	msgBox("In STAmt3FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt3FocusListener	6/23/20
'/**/
'// STAmt3LostFocus.bas
'//---------------------------------------------------------------
'// STAmt3LostFocus - Handle AmtFld3 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt3LostFocus()

'//	Usage.	macro call or
'//			call STAmt3LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld3"

'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld3"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt3LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt3LostFocus	6/23/20
'/**/
'// STAmt4FocusListener.bas
'//---------------------------------------------------------------
'// STAmt4FocusListener - Handle AmtFldn GotFocus event.
'//		6/20/20.	wmk.	20:45
'//---------------------------------------------------------------

public sub STAmt4FocusListener()

'//	Usage.	macro call or
'//			call STAmt4FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAcct3FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld4"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt4FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt4FocusListener	6/20/20
'/**/
'// STAmtLostFocus.bas
'//---------------------------------------------------------------
'// STAmt4LostFocus - Handle AmtFld4 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt4LostFocus()

'//	Usage.	macro call or
'//			call STAmt4LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld4"

'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld4"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt4LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt4LostFocus	6/23/20
'/**/
'// STAmt5FocusListener.bas
'//---------------------------------------------------------------
'// STAmt5FocusListener - Handle AmtFldn GotFocus event.
'//		6/20/20.	wmk.	20:45
'//---------------------------------------------------------------

public sub STAmt5FocusListener()

'//	Usage.	macro call or
'//			call STAmt5FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAcct3FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld5"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt5FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt5FocusListener	6/20/20
'/**/
'// STAmt5LostFocus.bas
'//---------------------------------------------------------------
'// STAmt5LostFocus - Handle AmtFld5 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt5LostFocus()

'//	Usage.	macro call or
'//			call STAmt5LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld5"
'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

dim bTotalMet		As Boolean		'// accum = split flag
	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld5"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt5LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt5LostFocus	6/23/20
'/**/
'// STAmt6FocusListener.bas
'//---------------------------------------------------------------
'// STAmt6FocusListener - Handle AmtFldn GotFocus event.
'//		6/20/20.	wmk.	20:45
'//---------------------------------------------------------------

public sub STAmt6FocusListener()

'//	Usage.	macro call or
'//			call STAmt6FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAcct3FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld6"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt6FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt6FocusListener	6/20/20
'/**/
'// STAmt6LostFocus.bas
'//---------------------------------------------------------------
'// STAmt6LostFocus - Handle AmtFld6 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt6LostFocus()

'//	Usage.	macro call or
'//			call STAmt6LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld6"
'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

dim bTotalMet		As Boolean		'// accum = split flag
	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld6"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt6LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt6LostFocus	6/23/20
'/**/
'// STAmt7FocusListener.bas
'//---------------------------------------------------------------
'// STAmt7FocusListener - Handle AmtFldn GotFocus event.
'//		6/20/20.	wmk.	20:45
'//---------------------------------------------------------------

public sub STAmt7FocusListener()

'//	Usage.	macro call or
'//			call STAmt7FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAcct3FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld7"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt7FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt7FocusListener	6/20/20
'/**/
'// STAmt7LostFocus.bas
'//---------------------------------------------------------------
'// STAmt7LostFocus - Handle AmtFld3 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt7LostFocus()

'//	Usage.	macro call or
'//			call STAmt7LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld7"
'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

dim bTotalMet		As Boolean		'// accum = split flag
	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld7"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt7LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt7LostFocus	6/23/20
'/**/
'// STAmt8FocusListener.bas
'//---------------------------------------------------------------
'// STAmt8FocusListener - Handle AmtFldn GotFocus event.
'//		6/20/20.	wmk.	20:45
'//---------------------------------------------------------------

public sub STAmt8FocusListener()

'//	Usage.	macro call or
'//			call STAmt8FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAcct3FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld8"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt8FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt8FocusListener	6/20/20
'/**/
'// STAmt8LostFocus.bas
'//---------------------------------------------------------------
'// STAmt8LostFocus - Handle AmtFld8 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt8LostFocus()

'//	Usage.	macro call or
'//			call STAmt8LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld8"
'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

dim bTotalMet		As Boolean		'// accum = split flag
	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld8"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt8LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt8LostFocus	6/23/20
'/**/
'// STAmt9FocusListener.bas
'//---------------------------------------------------------------
'// STAmt9FocusListener - Handle AmtFldn GotFocus event.
'//		6/20/20.	wmk.	20:45
'//---------------------------------------------------------------

public sub STAmt9FocusListener()

'//	Usage.	macro call or
'//			call STAmt9FocusListener()
'//
'// Entry.	AmtFld1..AmtFld10 got focus
'//
'//	Exit.	gsSTObjFocus = ST dialog AmtFld object name
'//		??	gsSTAcctFocus = ST dialog AcctFld object name
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; adapted from STAcct3FocusListener
'//
'//	Notes. gsSTAcctFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into. The gsSTObjFocus name will be cleared when the control
'// loses focus; but the gsSTAcctFocus will remain set until a different
'// acct object is accessed


'//	constants.
const MY_NAME="AmtFld9"

'//	local variables.
dim iStatus		As Integer			'// general status

	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	gsSTObjFocus = MY_NAME
'	gsSTAcctFocus = MY_NAME
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt9FocusListener - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt9FocusListener	6/20/20
'/**/
'// STAmt9LostFocus.bas
'//---------------------------------------------------------------
'// STAmt9LostFocus - Handle AmtFld9 LostFocus event.
'//		6/23/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STAmt9LostFocus()

'//	Usage.	macro call or
'//			call STAmt9LostFocus()
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//
'//	Exit.	gdSTAccumTot = updated with enterd value if within split bounds
'//
'// Calls.	fSTAmtLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	activate/deactivate <Done> control code fixed
'//	6/21/20.	wmk.	MY_NAME const added
'//	6/22/20.	wmk.	eliminate MY_NAME const; use STGetObjFocus to retrieve
'//						this control name in dialog
'//	6/23/20.	wmk.	rewritten; uses fSTAmtLostFocus
'//
'//	Notes.

'//	constants.
'const MY_NAME="AmtFld9"
'//	local variables.
dim iStatus		As Integer		'// general status
dim sFldName	As String		'// name of this field

dim bTotalMet		As Boolean		'// accum = split flag
	'// code.
	iStatus = -1
	ON ERROR GOTO ErrorHandler
	sFldName = "AmtFld9"
	
	iStatus = fSTAmtLostFocus(sFldName)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif
		
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STAmt9LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STAmt9LostFocus	6/23/20
'/**/
'// STAmtLostFocus.bas
'//---------------------------------------------------------------
'// STAmtLostFocus - Handle all AmtFld LostFocus events.
'//		6/22/20.	wmk.	08:00
'//---------------------------------------------------------------

public sub STAmtLostFocus()

'//	Usage.	call STAmtLostFocus()
'//			or macro call
'//
'// Entry.	AmtFld1..AmtFld10 lost focus
'//			gdSTAmts() = updated with newly entered amount
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gdSTAccumTot, gdSTTotalAmt updated with newly entered amount
'//			<Done> button activated if dialog entries complete
'//			next row fields activated if row complete
'//
'// Calls.	STFld6ToIndex, CheckRowComplete, CheckComplete
'//
'//	Modification history.
'//	---------------------
'//	6/20/20.	wmk.	original code; cloned from STAmtLostFocus
'//	6/21/20.	wmk.	check for total reached before activating next row
'// 6/21/20.	wmk.	change to accept psFldName parameter; to be called by
'//						all STAmt.LostFocus events; changed totals check and
'//						activate <Done> conditional
'//	6/22/20.	wmk.	change name to begin with "f" so can be distinguished
'//						by sub with same name for testing

'//	constants.

'//	local variables.
dim iRetValue 		As Integer			'// returned status
dim iStatus			As Integer			'// general status
dim iRowIx			As Integer			'// index of present row
dim oDlgControl		As Object			'// field ptr
dim bRowComplete	As Boolean			'// row complete flag
dim bDlgComplete	As Boolean			'// dialog complete flag
dim bTotalMet		As Boolean			'// split total met flag
dim sFldName 		As String			'// local copy of field name
dim oHasFocus		As Object

	'// code.
	iRetValue = -1						'// set bad return
	ON ERROR GOTO ErrorHandler
'msgBox("In sub STAmtLostFocus entry values.." + CHR(13)+CHR(10) _
'	+ "gdSTAccumTot = " + gdSTAccumTot + CHR(13)+CHR(10) _
'	+ "gdSTTotalAmt = " + gdSTTotalAmt + CHR(13)+CHR(10) _
'	+ "gsSTObjFocus = '" + gsSTObjFocus + "'" )
	
	oHasFocus = puoSTDialog.getControl("HasFocus")
	sFldName = oHasFocus.Text	
'	sFldName = gsSTObjFocus				'// use local copy of field name
'msgBox("sub STAmtLostFocus - sFldName = '" + sFldName + "'")
	if Len(sFldName) < 7 then
		GoTo ErrorHandler
	endif	'// error - field name not set up for entry
	
	bDlgComplete = false
'	iRowIx = STFld6ToIndex(sFldName)
'	iStatus = STUpdateAccumTot(sFldName)
	iRowIx = STFld6ToIndex(gsSTObjFocus)
	iStatus = STUpdateAccumTot(gsSTObjFocus)
'msgBox("In sub STAmtLostFocus values after update.." + CHR(13)+CHR(10) _
'	+ "gdSTAccumTot = " + gdSTAccumTot + CHR(13)+CHR(10) _
'	+ "gdSTTotalAmt = " + gdSTTotalAmt )
	
	if iStatus < 0 then
		GoTo ErrorHandler
	endif	'// end error in updating accum total conditional

	bTotalMet = (iStatus = 1)
	bRowComplete = STRowComplete(iRowIx)
	if bRowComplete then
		bDlgComplete = STCheckComplete()
		oDlgControl = puoSTDialog.getControl("UnsplitTot")
'		if gdSTAccumTot < gdSTTotalAmt then		
		if oDlgControl.Value > 0 then		
			'// check for next row Enabled; if not, enable it
			iStatus = STEnableNextRow(iRowIx)
			if iStatus < 0 then
				GoTo ErrorHandler
			endif
		else
			bTotalMet = true
		endif '// end total not met yet conditional
		
	endif	'// end row complete conditional
	
	'// activate/deactivate <Done> control if dialog complete
	oDlgControl = puoSTDialog.getControl("DoneBtn")
'	bDlgComplete = (bDlgComplete AND bTotalMet)
	bDlgComplete = (bDlgComplete OR bTotalMet)
	if bDlgComplete then
		oDlgControl.Model.Enabled = true
	else
		oDlgControl.Model.Enabled = false
	endif
	iRetValue = 0
	
NormalExit:
'	fSTAmtLostFocus = iRetValue
	exit sub
	
ErrorHandler:
	msgBox("In sub STAmtLostFocus - unprocessed error")
	iRetValue = iStatus
	GoTo NormalExit
	
end sub		'// end STAmtLostFocus	6/22/20
'/**/

'// STCOAListBtnListener.bas
'//---------------------------------------------------------------
'// STCOAListBtnListener - Handle COA List Button event.
'//		6/18/20.	wmk.
'//---------------------------------------------------------------

public sub STCOAListBtnListener()

'//	Usage.	macro call or
'//			call STCOAListBtnListener()
'//
'// Entry.	gsSTAcctFocus = name of last COA object to get focus
'//			puoSTDialog = ST Dialog object
'//
'//	Exit.	gsSTCOAs(n-1) updated with COA selection from GetCOA	
'//
'// Calls.	GetCOA
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.		wmk.	original code
'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim iNameLen	As Integer	'// length of name string
dim sFldName	As String	'// full field name from gsSTAcctFocus var
dim sAcctIx		As String	'// account index extracted from field name
dim iDigCount	As Integer	'// field name digit count
dim iArrayIx	As Integer	'// array index into gsSTCOAs string array
dim oSTDlgCtrl	As Object	'// account field in STdialog
dim sCOANum		As String	'// COA from GetCOA

	'// code.
	ON ERROR GOTO ErrorHandler
	sFldName = gsSTAcctFocus		'// get "AcctFldxx" string
	iNameLen = Len(sFldName)		'// get length, either 8 or 9
	Select Case iNameLen
	Case 8
		iDigCount = 1
	Case 9
		iDigCount = 2
	Case else
		iDigCount = 1
	End Select
	
	sAcctIx = Right(sFldName,iDigCount)
	iArrayIx = Val(sAcctIx) - 1
	
	sCOANum = GetCOA(1)		'// get COA from dialog
	gsSTCOAs(iArrayIx) = sCOANum
	oSTDlgCtrl = puoSTDialog.getControl(sFldName)
	oSTDlgCtrl.Text = sCOANum
	
NormalExit:
	exit sub

ErrorHandler:
	msgBox("Error getting COA from list...")
	GoTo NormalExit
end sub		'// end STCOAListBtnListener		6/18/20
'/**/
'// STCancelBtnRun.bas
'//---------------------------------------------------------------
'// STCancelBtnRun - ST dialog <Cancel> button Execute() event handler.
'//		6/27/20.	wmk.	15:00
'//---------------------------------------------------------------

public sub STCancelBtnRun()

'//	Usage.	macro call or
'//			call STCancelBtnRun()
'//
'// Entry.	puoSTDialog = ST dialog object
'//
'//	Exit.	ST dialog ended, returning exit code 0
'//			all ST-related vars cleared
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.		wmk.	original code
'//	6/27/20.	wmk.	gbSTEditMode flag set false
'//
'//	Notes. It is assumed that if the user hits <Cancel>, any
'// split information should be abandoned
'//

'//	constants.

'//	local variables.
dim iStatus		As Integer

	'// code.
	'// clear all stored array values and exit with cancel
	gbSTEditMode = false
	iStatus = STPubVarsReset()
	puoSTDialog.endDialog(0)
	
end sub		'// end STCancelBtnRun	6/27/20
'/**/
'// STCheckComplete.bas
'//---------------------------------------------------------------
'// STCheckComplete - Check if dialog has minimum information.
'//		7/7/20.	wmk.	12:40
'//---------------------------------------------------------------

public function STCheckComplete() As Boolean

'//	Usage.	bVal = STCheckComplete()
'//
'// Entry.	giSTLastIx = index of last stored row from dialog
'//			gdSTAmts(0..giSTLastIx0 = stored amounts from dialog
'//			gsSTCOAs(0..giSTLastIx) = stored account #s from dialog
'//
'//	Exit.	bVal = true if ALL of the following:
'//					gdSTAmts(0..STLastIx) = non-empty values
'//					gsSTCOAs(0..STLastIx) = each exactly 4 chars
'//					gdSTTotalAmt = gdSTAccumTot
'//				   false otherwise
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'//	6/21/20.	wmk.	modified to check dialog UnsplitTot = 0.
'//	7/7/20.		wmk.	bug fix where fractional cents making
'//						UnsplitTot <> 0; changed to string comparison

'//
'//	Notes.
'//

'//	constants.

'//	local variables.
dim bRetValue As Boolean		'// returned flag
dim bComplete	As Boolean	'// interim tests flag.
dim i			As Integer		'// loop index
dim oUnsplit	As Object		'// unsplit field from dialot

	'// code.
	bComplete = true
	for i = 0 to giSTLastIx
		bComplete = bComplete AND (Len(gsSTCOAs(i)) = 4)
		bComplete = bComplete AND (gdSTAmts(i) >= 0)
		if NOT bComplete then
			exit for
		endif	'// end failed conditional	
	next i		'// loop for all entered rows

	oUnSplit = puoSTDialog.getControl("UnsplitTot")	
	bComplete = bComplete AND _
			(StrComp(Str(gdSTTotalAmt),Str(gdSTAccumTot))=0)
'	bComplete = bComplete AND (oUnSplit.Value = 0.)
	STCheckComplete = bComplete

end function 	'// end STCheckComplete	7/7/20
'/**/
'// STClrObjFocus.bas
'//---------------------------------------------------------------
'// STClrObjFocus - Clear control name with focus inside object.
'//		6/22/20.	wmk.	10:00
'//---------------------------------------------------------------

public function STClrObjFocus() As Integer

'//	Usage.	iVar = STClrObjFocus()
'//
'// Entry.	puoSTDialog = ST dialog object
'//			puoSTDialog.<HasFocusObject>.Text = available
'//
'//	Exit.	iVar = 0 - control name cleared in dialog
'//				 < 0 - error attempting to store active control namne
'//			puoSTDialog.<HasFocus-Object>.Text = ""
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6//20.		wmk.	original code
'//
'//	Notes. STSetObjFocus should be called by any <control>Listener
'//	event when the <control> gets focus. This will provide an easy
'// way for a control to get its own name during processing. The
'// STClrObjFocus function should be called when the <control> loses
'// focus so other controls may reuse the dialog field "HasFocus".
'// A third method STGetObjFocus will retrieve the current "HasFocus"
'// .Text field for use during the control processing.
'//

'//	constants.

'//	local variables.
dim iRetValue As Any
dim oDlgControl		As Object		'// dialog control
dim sFocusName		As String		'// HasFocus text control field name

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	oDlgControl = puoSTDialog.getControl("HasFocusFld")
	sFocusName = oDlgControl.Text
	oDlgControl.Text = ""
	iRetValue = 0			'// set normal return
	
NormalExit:	
	STClrObjFocus = iRetValue
	exit function

ErrorHandler:
	msgBox("In STClrFocus - unprocessed error")
	GoTo NormalExit
	
end function 	'// end STClrObjFocus	6/22/20
'/**/
'// STCredRBListener.bas
'//---------------------------------------------------------------
'// STCredRBListener - ST [Credit] radio button event handler.
'//		6/26/20.	wmk.	12:00
'//---------------------------------------------------------------

public sub STCredRBListener()

'//	Usage.	macro call or
'//			call STCredRBListener()
'//
'// Entry.	User clicked [Credit] radio button control
'//			gbSTSplitCredits = current split credits state
'//			gsETAcct1 = Debit account	
'//			gsETAcct2 = Credit account
'//
'//	Exit.	if clicked ON
'//			gbSTSplitCredits set false
'//			gbETDebitIsTotal set false
'//			if OFF, gbSTSplitCredits set true
'//			gbETDebitIsTotal set true
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.		wmk.	original code
'//`6/26/20.	wmk.	gsSTAcct1 or gsSTAcct2 used to set "TotalAcctFld"
'//						in ST dialog to display account being split
'//
'//	Notes.

'//	constants.

'//	local variables.
dim oSTCredRB		As Object		'// {Credit} radio button object
dim bState			As Boolean		'// .State of button
dim oSTDlgCtrl		As Object		'// control from "TotalAcctFld" in ST

	'// code.
	ON ERROR GOTO ErrorHandler
	oSTCredRB = puoSTDialog.getControl("CreditTotRB")
	oSTDlgCtrl = puoSTDialog.getControl("TotalAcctFld")
	bState = oSTCredRB.State
	if bState then	'// Credit button selected
		gbSTSplitCredits = false
		gbETDebitIsTotal = false
		oSTDlgCtrl.Text = gsETAcct2
	else
		gbSTSplitCredits = true
		gbETDebitIsTotal = true
		oSTDlgCtrl.Text = gsETAcct1
	endif	'// end Credit Total button State conditional

NormalExit:
	exit sub

ErrorHandler:
	msgBox("STCredRBListener - unprocessed error.")
	GoTo NormalExit

end sub		'// end STCredRBListener	6/26/20
'/**/
'// STDebtRBListener.bas
'//---------------------------------------------------------------
'// STDebtRBListener - ST [Debit] radio button event handler.
'//		7/6/20.	wmk.	19:00
'//---------------------------------------------------------------

public sub STDebtRBListener()

'//	Usage.	macro call or
'//			call STDebtRBListener()
'//
'// Entry.	User clicked [Debit] radio button control
'//			puoSTDialog = Split Control dialog object
'//			gbSTSplitCredits = current split credits stat
'//			gsETAcct1 = Debit account	
'//			gsETAcct2 = Credit account
'//
'//	Exit.	if clicked ON
'//				gbSTSplitCredits set true
'//				gbETDebitIsTotal set true
'//				STSplitAcctFld = gsSTAcct1
'//			if OFF, gbSTSplitCredits set false
'/				gbETDebitIsTotal set false
'//				STSplitAcctFld - gsSTAcct2
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'//`6/26/20.	wmk.	gsSTAcct1 or gsSTAcct2 used to set "TotalAcctFld"
'//						in ST dialog to display account being split
'//	7/6/20.		wmk.	bug fix; code was checking oSTCredRB.State instead
'//						of oSTDebtRB.State; puoSTDailog corrected; gsETAcct1
'//						gsETAcct2 references corrected
'//
'//	Notes. 

'//	constants.

'//	local variables.
dim oSTDebtRB		As Object		'// {Debit} radio button object
dim bState			As Boolean		'// .State of button
dim oSTDlgCtrl		As Object		'// control from "SplitAcctFld" in ST

	'// code.
	ON ERROR GOTO ErrorHandler
	oSTDebtRB = puoSTDialog.getControl("DebitTotRB")
	oSTDlgCtrl = puoSTDialog.getControl("TotalAcctFld")
	bState = oSTDebtRB.State
	if bState then	'// Debit button selected
		gbSTSplitCredits = true
		gbETDebitIsTotal = true
		oSTDlgCtrl.Text = gsETAcct1
	else
		gbSTSplitCredits = false
		gbETDebitIsTotal = false
		oSTDlgCtrl.Text = gsETAcct2
	endif	'// end Debit Total button State conditional

NormalExit:
	exit sub

ErrorHandler:
	msgBox("STDebtRBListener - unprocessed error.")
	GoTo NormalExit

end sub		'// end STDebtRBListener	7/6/20
'/**/
'// STDesc10LostFocus.bas
'//---------------------------------------------------------------
'// STDesc10LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STDesc10LostFocus()

'//	Usage.	macro call or
'//			call STDesc10LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld10"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc10LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc10LostFocus	6/25/20
'/**/
'// STDesc1LostFocus.bas
'//---------------------------------------------------------------
'// STDesc1LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	05:15
'//---------------------------------------------------------------

public sub STDesc1LostFocus()

'//	Usage.	macro call or
'//			call STDesc1LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld1"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc1LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc1LostFocus	6/25/20
'/**/
'// STDesc2LostFocus.bas
'//---------------------------------------------------------------
'// STDesc2LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	05:15
'//---------------------------------------------------------------

public sub STDesc2LostFocus()

'//	Usage.	macro call or
'//			call STDesc2LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld2"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc2LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc2LostFocus	6/25/20
'/**/
'// STDesc3LostFocus.bas
'//---------------------------------------------------------------
'// STDesc3LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	05:15
'//---------------------------------------------------------------

public sub STDesc3LostFocus()

'//	Usage.	macro call or
'//			call STDesc3LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld3"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc3LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc3LostFocus	6/25/20
'/**/
'// STDesc4LostFocus.bas
'//---------------------------------------------------------------
'// STDesc4LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STDesc4LostFocus()

'//	Usage.	macro call or
'//			call STDesc4LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld4"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc4LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc4LostFocus	6/25/20
'/**/
'// STDesc5LostFocus.bas
'//---------------------------------------------------------------
'// STDesc5LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STDesc5LostFocus()

'//	Usage.	macro call or
'//			call STDesc5LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld5"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc5LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc5LostFocus	6/25/20
'/**/
'// STDesc6LostFocus.bas
'//---------------------------------------------------------------
'// STDesc6LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STDesc6LostFocus()

'//	Usage.	macro call or
'//			call STDesc6LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld6"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc6LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc6LostFocus	6/25/20
'/**/
'// STDesc7LostFocus.bas
'//---------------------------------------------------------------
'// STDesc7LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STDesc7LostFocus()

'//	Usage.	macro call or
'//			call STDesc7LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld7"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc7LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc7LostFocus	6/25/20
'/**/
'// STDesc8LostFocus.bas
'//---------------------------------------------------------------
'// STDesc8LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STDesc8LostFocus()

'//	Usage.	macro call or
'//			call STDesc8LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld8"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc8LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc8LostFocus	6/25/20
'/**/
'// STDesc9LostFocus.bas
'//---------------------------------------------------------------
'// STDesc9LostFocus - Handle DescFldn LostFocus event.
'//		6/25/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STDesc9LostFocus()

'//	Usage.	macro call or
'//			call STDesc4LostFocus()
'//
'// Entry.	DescFld1..DescFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "DescFld1")
'//
'//	Exit.	gsSTDescs(0) = text from DescFld1..10
'//
'// Calls.	STDescLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "DescFld9"					'// clear object focus tracker
	iStatus = STDescLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STDesc9LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STDesc9LostFocus	6/25/20
'/**/
'// STDescLostFocus.bas
'//---------------------------------------------------------------
'// STDescLostFocus - Handle any DescFld LostFocus event.
'//		6/25/20.	wmk.	05:30
'//---------------------------------------------------------------

public function STDescLostFocus(psFldName As String) As Integer

'//	Usage.	macro call or
'//			call STDescLostFocus()
'//
'//	[Parameters.	sFldName = name of dialog object lost focus]
'//
'// Entry.	AcctFld1..AcctFld10 lost focus
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gsCOAs(0) = text from AcctFld1..10
'//
'// Calls.	STFld7ToIndex
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; cloned from STRefLostFocus
'//
'//	Notes. gsSTObjFocus will be used by the [...] button Execute
'//	event to determine which Acct array string to store the selected
'// COA into.


'//	constants.

'//	local variables.
dim iRetValue 		As Integer			'// returned status
dim iStatus			As Integer			'// general status
dim iRowIx			As Integer			'// index of present row
dim oDlgControl		As Object			'// field ptr
dim bRowComplete	As Boolean			'// row complete flag
dim bDlgComplete	As Boolean			'// dialog complete flag
dim sFldName 		As String			'// local copy of field name

	'// code.
	iRetValue = -1						'// set bad return
	ON ERROR GOTO ErrorHandler
	sFldName = psFldName				'// use local copy of field name
	oDlgControl = puoSTDialog.getControl(sFldName)
	iRowIx = STFld7ToIndex(sFldName)
	gsSTDescs(iRowIx) = oDlgControl.Text
	
	iRetValue = 0

NormalExit:
	STDescLostFocus = iRetValue
	exit function
	
ErrorHandler:
	msgBox("In STDescLostFocus - unprocessed error")
	GoTo NormalExit
	
end function		'// end STDescLostFocus	6/25/20
'/**/
'// STDialogReset.bas
'//---------------------------------------------------------------
'// STDialogReset - Reset Split Transaction Dialog.
'//		7/7/20.	wmk.	07:15
'//---------------------------------------------------------------

public sub STDialogReset()

'//	Usage.	call STDialogReset()
'//			or macro call
'//
'// Entry.	puoSTDialog = Split Transaction dialog object
'//
'//	Exit.	iVar = 0 - normal return.
'//				 <> 0 = abnormal return
'//
'// Calls.	STPubVarsReset, STDlgCtrlsReset
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'//	6/23/20.	wmk.	sub to support <Reset Form> execute() event
'//	7/7/20.		wmk.	remove reference to deleted "HasFocus" field
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue 	As Integer
dim i			As Integer
dim oSTDlgCtrl	As Object
dim sFldNum		As String
dim sCtrlName	As String

	'// code.
	
	iRetValue = -1		'// abnormal return
	ON ERROR GOTO ErrorHandler
'XRay puoSTDialog
	gdSTAccumTot = 0.			'// clear accumulated total
	giSTLastIx = 0				'// last active row index
	gbSTSplitCredits = true		'// assume Debits as total
	gsSTAcctFocus = ""			'// clear account object focus
	gsSTObjFocus = ""			'// object name with focus
	redim gsSTDescs(0)	'// scrap descriptions array
	redim gdSTAmts(0)	'// scrap amounts array
	redim gsSTCOAs(0)	'// scrap COAs array
	redim gsSTRefs(0)	'// scrap refs array

	'// set State of Debit/Credit radio buttons to Debit.State = true
	oSTDlgCtrl = puoSTDialog.getControl("DebitTotRB")
	oSTDlgCtrl.State = true
	
	'// set Total field to gdSTTotalAmt
	oSTDlgCtrl = puoSTDialog.getControl("TotalAmt")
	oSTDlgCtrl.Value = gdSTTotalAmt

	if false then
	GoTo JumpAround
	endif
		
	'// clear all 10 Description, Amount, and COA fields
	'// and set .Enabled attribute to false for last 9 of each
	for i = 0 to 9
		sFldNum = Trim(Str(i+1))
		sCtrlName = "DescFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = ""
		if i > 0 then
			oSTDlgCtrl.Model.Enabled = false
		endif
		
		if false then 
		GoTo ContFor
		endif
		
		sCtrlName = "AmtFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = ""
		if i > 0 then
			oSTDlgCtrl.Model.Enabled = false
		endif
		
		sCtrlName = "AcctFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = ""
		if i > 0 then
			oSTDlgCtrl.Model.Enabled = false
		endif
		
		sCtrlName = "RefFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = ""
		if i > 0 then
			oSTDlgCtrl.Model.Enabled = false
		endif
ContFor:		
	next i	'// advance to next control name for each field
JumpAround:
	oSTDlgCtrl = puoSTDialog.getControl("DescFld1")
	oSTDlgCtrl.Text = ""
	oSTDlgCtrl = puoSTDialog.getControl("AmtFld1")
	oSTDlgCtrl.Text = ""
	
'	oSTDlgCtrl = puoSTDialog.getControl("HasFocusFld")
'	oSTDlgCtrl.Text = "MyohMy"		
	'// clear Accumulated Total field to 0; set Unsplit field to gdSTTotalAmt
	oSTDlgCtrl = puoSTDialog.getControl("AccumTot")
	oSTDlgCtrl.Value = 0.
	oSTDlgCtrl = puoSTDialog.getControl("UnsplitTot")
	oSTDlgCtrl.Value = gdSTTotalAmt
	
	'// set DoneBtn Enabled to false
	oSTDlgCtrl = puoSTDialog.getControl("DoneBtn")
	oSTDlgCtrl.Model.Enabled = false
	

	if true then
		GoTo NormalExit
	endif	
	'// reset dialog public variables
	iRetValue = STPubVarsReset()
	if iRetValue < 0 then
		GoTo ErrorHandler
	endif	'// end bad controls reset conditional
	
	'// reset dialog controls
	iRetValue = STDlgCtrlsReset()
	if iRetValue < 0 then
		GoTo ErrorHandler
	endif	'// end bad controls reset conditional
	
	iRetValue = 0		'// no errors - exit
	
NormalExit:
	
	exit sub
	
ErrorHandler:
	msgBox("STDialogReset - i = " + i + " ..unprocessed error.")
	GoTo NormalExit
	
end sub 	'// end STDialogReset		7/7/20
'/**/
'// STDialogReset.bas
'//---------------------------------------------------------------
'// fSTDialogReset - Reset Split Transaction Dialog.
'//		6/23/20.	wmk.
'//---------------------------------------------------------------

public function fSTDialogReset() as Integer

'//	Usage.	iVar = fSTDialogReset()
'//
'// Entry.	puoSTDialog = Split Transaction dialog object
'//
'//	Exit.	iVar = 0 - normal return.
'//				 <> 0 = abnormal return
'//
'// Calls.	STPubVarsReset, STDlgCtrlsReset
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'//	6/23/20.	wmk.	name change to fSTDialogReset to support
'//						STDialogReset sub for event processing
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue As Integer

	'// code.
	
	iRetValue = -1		'// abnormal return
	
	'// reset dialog public variables
	iRetValue = STPubVarsReset()
	if iRetValue < 0 then
		GoTo ErrorHandler
	endif	'// end bad controls reset conditional
	
	'// reset dialog controls
	iRetValue = STDlgCtrlsReset()
	if iRetValue < 0 then
		GoTo ErrorHandler
	endif	'// end bad controls reset conditional
	
	iRetValue = 0		'// no errors - exit
	
NormalExit:
	fSTDialogReset = iRetValue
	exit function
	
ErrorHandler:
	msgBox("fSTDialogReset - unprocessed error.")
	GoTo NormalExit
	
end function 	'// end fSTDialogReset		6/18/20
'/**/
'// STDlgCtrlsReset.bas
'//---------------------------------------------------------------
'// STDlgCtrlsReset - Reset Split Dialog controls.
'//		6/27/20.	wmk.	15:00
'//---------------------------------------------------------------

public function STDlgCtrlsReset() As Integer

'//	Usage.	iVar = STDlgCtrlsReset()
'//
'// Entry.	puoSTDialog = Split Transaction dialog object
'//			gbETDebitIsTotal = Debit is total flag from ET dialot
'//			gsETAcct1 = Debit account COA
'//			gsETAcct2 = Credit account COA
'//			gbSTEditMode = true if <Edit Split>  being started
'//							false otherwise		
'//
'//	"SplitDialog" object fields as follows:
'// DebitTotRB - Debit is total radio button
'// CreditTotRB - Credit is total radio button
'// TotalAmt - total transaction amount text box (numeric, read-only)
'//	TotalAcctFld - account COA transaction being split from
'// DescFld1 - description field 1
'// ...
'// DescFld10 - description field 10
'//	AmtFld1 - split amount field 1	(numeric)
'// ...
'// AmtFld10 - split amount field 10 (numeric)
'// AcctFld1 - COA account field 1
'//	...
'// AcctFld10 - COA account field 10
'// COAListBtn - COA list [...] button
'// AccumTot - splits accumulated total	(numeric, read-only)
'// UnsplitTot - unsplit remaining amount (numeric, read-only)
'// DoneBtn - [Done] button
'// ResetBtn - [Reset Form] button
'//	CancelBtn - [Cancel] button
'//
'//	Exit.	iVar = 0 - normal return
'//				<> 0 - error return
'//			if gbSTEditMode = true on entry; iVar will be the returned
'//			status from STDlgCtrlsRestore
'//
'// Calls.	STDlgCtrlsRestore
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code; stub
'//	6/23/20.	wmk.	code fleshed out; loop added, but problematic
'//	6/25/20.	wmk.	bug fixes where loop calls were to .getObject
'//						instead of .getControl; new field "TotalAcctFld"
'//						set from ET gsETAcct1 or gsETAcct2 dependent
'//						on flag gbETDebitIsTotal; corrected loop to use
'//						oSTDlgCtrl instead of oDlgControl
'//	6/27/20.	wmk.	<Edit Split> support to restore previously
'//						entered values
'//
'//	Notes. When intializing some of the ST fields, the values are
'//	picked up from the ET dialog public vars.


'//	constants.

'//	local variables.
dim iRetValue 	As Integer
dim oSTDlgCtrl 	As Object
dim i			As Integer		'// loop index
dim sFldNum		As String		'// field number
dim sCTrlName	As String		'// control name

	'// code.
	
	iRetValue = -1		'// set error return
	ON ERROR GOTO ErrorHandler

	'// check for <Edit Split> and restore controls if so
	if gbSTEditMode then
		iRetValue = STDlgCtrlsRestore()
		GoTo NormalExit
	endif
	
	'// set State of Debit/Credit radio buttons to Debit.State = true
	oSTDlgCtrl = puoSTDialog.getControl("DebitTotRB")
	oSTDlgCtrl.State = true
	
	'// set Total field to gdSTTotalAmt
	oSTDlgCtrl = puoSTDialog.getControl("TotalAmt")
	oSTDlgCtrl.Value = gdSTTotalAmt

if false then
GoTo JumpAround
endif
	'// set account split from field
	oSTDlgCtrl = puoSTDialog.getControl("TotalAcctFld")
	if gbSTSplitCredits then
		oSTDlgCtrl.Text = gsETAcct1	'// debit account
	else
		oSTDlgCtrl.Text = gsETAcct2	'// credit account
	endif
	
if false then
GoTo JumpAround
endif
	
	'// clear all 10 Description, Amount, and COA fields
	'// and set .Enabled attribute to false for last 9 of each
	for i = 0 to 9
		sFldNum = Trim(Str(i+1))
		sCtrlName = "DescFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = ""
		if i > 0 then
			oSTDlgCtrl.Model.Enabled = false
		endif
		
		sCtrlName = "AmtFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = ""
		if i > 0 then
			oSTDlgCtrl.Model.Enabled = false
		endif
		
		sCtrlName = "AcctFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = ""
		if i > 0 then
			oSTDlgCtrl.Model.Enabled = false
		endif
		
		sCtrlName = "RefFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = ""
		if i > 0 then
			oSTDlgCtrl.Model.Enabled = false
		endif
		
	next i	'// advance to next control name for each field
	
JumpAround:
	'// clear Accumulated Total field to 0; set Unsplit field to gdSTTotalAmt
	oSTDlgCtrl = puoSTDialog.getControl("AccumTot")
	oSTDlgCtrl.Value = 0.
	oSTDlgCtrl = puoSTDialog.getControl("UnsplitTot")
	oSTDlgCtrl.Value = gdSTTotalAmt
	
	'// set DoneBtn Enabled to false
	oSTDlgCtrl = puoSTDialog.getControl("DoneBtn")
	oSTDlgCtrl.Model.Enabled = false
	
	iRetValue = 0
	
NormalExit:	
	STDlgCtrlsReset = iRetValue
	exit function

ErrorHandler:
msgBox("STDlgCtrlsReset - i = " + i + ".. unprocessed error.")
	GoTo NormalExit
	
end function 	'// end STDlgCtrlsReset		6/27/20
'/**/
'// STDlgCtrlsRestore.bas
'//---------------------------------------------------------------
'// STDlgCtrlsRestore - Restore fields in ST dialog (Edit).
'//		6/27/20.	wmk.	16:45
'//---------------------------------------------------------------

public function STDlgCtrlsRestore() As Integer	

'//	Usage.	iVar = STDlgCtrlsRestore()
'//
'// Entry.	gsSTDescs() = descriptions array
'//			gdSTAmts() = split amounts array
'//			gsSTCOAs() = COA accounts array
'//			gsSTRefs() = references array
'//			gbSplitCredits = true if gdSTAmts are Credits;
'//							 false if gdSTAmts are Debits
'//			gdSTTotalAmt = split total
'//			VSDialog fields:
'// DebitTotRB - Debit is total radio button
'// CreditTotRB - Credit is total radio button
'// TotalAmt - total transaction amount text box (numeric, read-only)
'//	TotalAcctFld - account COA transaction being split from
'// DescFld1 - description field 1
'// ...
'// DescFld10 - description field 10
'//	AmtFld1 - split amount field 1	(numeric)
'// ...
'// AmtFld10 - split amount field 10 (numeric)
'// AcctFld1 - COA account field 1
'//	...
'// AcctFld10 - COA account field 10
'// UnsplitTot - unsplit total amount
'// AccumTot - accumulated split sums
'// DoneBtn - [Done] button
'// EditBtn - [Reset Form] button
'//
'//	Exit.	iVar = 0 - no error in preset
'//				 < 0 - error in preset
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/26/20.	wmk.	original code
'//	6/27/20.	wmk.	initial bug fixes; inactive cell backgrounds to MEDGRAY
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue As Integer
dim oSTDlgCtrl	As Object
dim oSTDlgCtrl1 As Object		'// 2nd control for when needed
dim i			As Integer		'// loop index
dim iActiveLimit	As Integer	'// active index limit in arrays
dim sFldNum		As String		'// field number
dim sCtrlName	As String		'// control name

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler

	'// set State of Debit/Credit radio buttons based on gbSTSplitCredits
	oSTDlgCtrl = puoSTDialog.getControl("DebitTotRB")
	oSTDlgCtrl1 = puoSTDialog.getControl("CreditTotRB")
	if gbSTSplitCredits then
		oSTDlgCtrl.State = true
		oSTDlgCtrl.Model.Enabled = true
		oSTDlgCtrl1.Model.Enabled = true
	else
		oSTDlgCtrl.State = false
		oSTDlgCtrl.Model.Enabled = true
		oSTDlgCtrl1.Model.Enabled = true
	endif
	
	'// set Total field to gdSTTotalAmt
	oSTDlgCtrl = puoSTDialog.getControl("TotalAmt")
	oSTDlgCtrl.Value = gdSTTotalAmt
	oSTDlgCtrl = puoSTDialog.getControl("AccumTot")
	oSTDlgCtrl.Value = gdSTTotalAmt
	
	'// set Unplit Total field to 0, since it was all 
	'// split before
	oSTDlgCtrl = puoSTDialog.getControl("UnsplitTot")
	oSTDlgCtrl.Value = 0. 

	'// set account split from field
	oSTDlgCtrl = puoSTDialog.getControl("TotalAcctFld")
	if gbSTSplitCredits then
		oSTDlgCtrl.Text = gsETAcct1	'// debit account
	else
		oSTDlgCtrl.Text = gsETAcct2	'// credit account
	endif
	
	'//	set all active Description, Amount, and COA fields for all
	'// array values up to subscript UBound and set fields enabled
	'// so easily visible. All fields are read-only.
	'// set all inactive Description, Amount, and COA fields empty
	'// and set Enabled to false to darken fields
	iActiveLimit = uBound(gdSTAmts)
	for i = 0 to iActiveLimit
		sFldNum = Trim(Str(i+1))
		sCtrlName = "DescFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = gsSTDescs(i)
		oSTDlgCtrl.Model.Enabled = true
		
		sCtrlName = "AmtFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = Str(gdSTAmts(i))
		oSTDlgCtrl.setValue(gdSTAmts(i))
		oSTDlgCtrl.Model.Enabled = true
		
		sCtrlName = "AcctFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = gsSTCOAs(i)
		oSTDlgCtrl.Model.Enabled = true
		
		sCtrlName = "RefFld" + sFldNum
		oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
		oSTDlgCtrl.Text = gsSTRefs(i)
		oSTDlgCtrl.Model.Enabled = true
		
	next i	'// advance to next control name for each field

	'// set remaining fields Enabled = false with no data
	if iActiveLimit < 9 then
		for i = iActiveLimit+1 to 9
			sFldNum = Trim(Str(i+1))
			sCtrlName = "DescFld" + sFldNum
			oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
			oSTDlgCtrl.Model.Enabled = false

			sCtrlName = "AmtFld" + sFldNum
			oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
			oSTDlgCtrl.Model.Enabled = false
		
			sCtrlName = "AcctFld" + sFldNum
			oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
			oSTDlgCtrl.Model.Enabled = false
		
			sCtrlName = "RefFld" + sFldNum
			oSTDlgCtrl = puoSTDialog.getControl(sCtrlName)
			oSTDlgCtrl.Model.Enabled = false
		next i
		
	endif		'// end active limit < 9 conditional
	

	iRetValue = 0	'// normal return
	
NormalExit:	
	STDlgCtrlsRestore = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STDlgCtrlsRestore - sCtrlName = " + sCtrlName + " unprocessed error.")
	GoTo NormalExit
	
end function 	'// end STDlgCtrlsRestore	6/27/20
'/**/
'// STDoneBtnRun.bas
'//---------------------------------------------------------------
'// STDoneBtnRun - ST dialog <Done> button Execute() event handler.
'//		6/24/20.	wmk.
'//---------------------------------------------------------------

public sub STDoneBtnRun()

'//	Usage.	macro call or
'//			call STDoneBtnRun()
'//
'// Entry.	puoSTDialog = ST dialog object
'//
'//	Exit.	ST dialog ended, returning exit code 2
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.		wmk.	original code
'//
'//	Notes.
'//

'//	constants.

'//	local variables.

	'// code.
	puoSTDialog.endDialog(2)
	
end sub		'// end STDoneBtnRun	6/24/20
'/**/
'// STEnableNextRow.bas
'//---------------------------------------------------------------
'// STEnableNextRow - Enable next ST dialog input row.
'//		6/27/20.	wmk.	05:00
'//---------------------------------------------------------------

public function STEnableNextRow(piCurrRow) As Integer

'//	Usage.	iVar = STEnableNextRow( iCurrRow )
'//
'//		iCurrRow = current input row index {0..9}
'//
'// Entry.	puoSTDialog = ST dialog object
'//			"AmtFld.n" or "AcctFld.n" just lost focus with enough to move on
'//			field "AmtFld.n-1" will be Enabled if not already when .9 <=9
'//
'//	Exit.	iVar = 0 - normal return
'//				 <> 0	- error in enabling next row
'//			expanded public input string arrays if next row beyond last
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'//	6/20/20.	wmk.	redimension input arrays to accept input
'// 6/27/20.	wmk.	open if sequence corrected; eliminated 2nd end
'//						function as fix
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	ST input field controls base names.
const	AMT_BASENAME="AmtFld"
const	ACCT_BASENAME="AcctFld"
const	DESC_BASENAME="DescFld"
const	REF_BASENAME="RefFld"

'//	local variables.
dim iRetValue 	As Integer
dim iNextIx		As Integer	'// next row index
dim sRowIx		As String	'// next index as String
dim oDlgControl	As Object	'// ST dialog object
dim sFldName	As String	'// field name

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	if piCurrRow >= 9 then
		iRetValue = 0
		GoTo NormalExit
	endif	'// end at end of input conditional

	'// check next row already enabled
	iNextIx = piCurrRow + 1
	if giSTLastIx >= iNextIx then
		iRetValue = 0
		GoTo NormalExit
	else
		giSTLastIx = iNextIx	'// advance last active row index
	endif	'// end next row already enabled conditional

	if iNextIx > UBound(gdSTAmts) then
		redim preserve gdSTAmts(iNextIx)
		redim preserve gsSTDescs(iNextIx)
		redim preserve gsSTCOAs(iNextIx)
		redim preserve gsSTRefs(iNextIx)
	endif
	
	iNextIx = iNextIx + 1				'// advance index to get row suffix
	sRowIx = Trim(Str(iNextIx))			'// get row string
	
	'// Enable controls in next row
	sFldName = DESC_BASENAME + sRowIx	'// description input fld
	oDlgControl = puoSTDialog.getControl(sFldName)
	oDlgControl.Model.Enabled = true
	
	sFldName = AMT_BASENAME + sRowIx	'// amount input fld
	oDlgControl = puoSTDialog.getControl(sFldName)
	oDlgControl.Model.Enabled = true

	sFldName = ACCT_BASENAME + sRowIx	'// COA account input fld
	oDlgControl = puoSTDialog.getControl(sFldName)
	oDlgControl.Model.Enabled = true

	sFldName = REF_BASENAME + sRowIx	'// reference input fld
	oDlgControl = puoSTDialog.getControl(sFldName)
	oDlgControl.Model.Enabled = true

	iRetValue = 0
	
NormalExit:
	STEnableNextRow = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STEnableNextRow - unprocessed error.")
	GoTo NormalExit
	
end function 	'// end STEnableNextRow		6/20/20
'/**/
'end function	'//$
'/**/
'// STFld6ToIndex.bas
'//---------------------------------------------------------------------
'// STFld6ToIndex - Convert field with 6 char base name to array index.
'//		6/19/20.	wmk.
'//---------------------------------------------------------------------

public function STFld6ToIndex(psFldName As String) As Integer

'//	Usage.	iVar = STFld6ToIndex( sFldName )
'//
'//		sFldName = name of field (e.g. "AmtFld1")
'//
'// Entry.	gdSTAmtFldsn() array contains amount values entered
'//
'//	Exit.	iVar = 0-based index into arrays corresponding to
'//					<object-name>n
'//					subtracts 1 from n to obtain index
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue As Integer	'// returned index value
dim iNameLen	As Integer	'// passed name length
dim sFldName	As String	'// passed field name
dim iDigCount	As Integer	'// digit count in name
dim sAcctIx		As String	'// extracted end digit(s)
dim iArrayIx	As Integer	'// converted array index

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
'	sFldName = psFldName
	iNameLen = Len(psFldName)		'// get length, either 7 or 8
	Select Case iNameLen
	Case 7
		iDigCount = 1
	Case 8
		iDigCount = 2
	Case else
		iDigCount = 1
	End Select
	
	sAcctIx = Right(psFldName,iDigCount)
	iArrayIx = Val(sAcctIx) - 1
	if iArrayIx < 0 then
		GoTo ErrorHandler
	endif	'//end negative index calculated conditional

	iRetValue = iArrayIx
	
NormalExit:
	STFld6ToIndex = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STFld6ToIndex - unprocessed error converting '" + psFldName _
			+ "' to index")
	GoTo NormalExit
	
end function 	'// end STFld6ToIndex	6/19/20
'/**/
'// STFld7ToIndex.bas
'//---------------------------------------------------------------------
'// STFld7ToIndex - Convert field with 7 char base name to array index.
'//		6/19/20.	wmk.	10:00
'//---------------------------------------------------------------------

public function STFld7ToIndex(psFldName As String) As Integer

'//	Usage.	iVar = STFld7ToIndex( sFldName )
'//
'//		sFldName = name of field (e.g. "AmtFld1")
'//
'// Entry.	gdSTAmtFldsn() array contains amount values entered
'//
'//	Exit.	iVar = 0-based index into arrays corresponding to
'//					<object-name>n
'//					subtracts 1 from n to obtain index
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue As Integer	'// returned index value
dim iNameLen	As Integer	'// passed name length
dim sFldName	As String	'// passed field name
dim iDigCount	As Integer	'// digit count in name
dim sAcctIx		As String	'// extracted end digit(s)
dim iArrayIx	As Integer	'// converted array index

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sFldName = psFldName
	iNameLen = Len(sFldName)		'// get length, either 7 or 8
	Select Case iNameLen
	Case 8
		iDigCount = 1
	Case 9
		iDigCount = 2
	Case else
		iDigCount = 1
	End Select
	
	sAcctIx = Right(sFldName,iDigCount)
	iArrayIx = Val(sAcctIx) - 1
	if iArrayIx < 0 then
		GoTo ErrorHandler
	endif	'//end negative index calculated conditional

	iRetValue = iArrayIx
	
NormalExit:
	STFld7ToIndex = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STFld7ToIndex - unprocessed error converting '" + psFldName _
			+ "' to index")
	GoTo NormalExit
	
end function 	'// end STFld7ToIndex	6/19/20
'/**/
'// STGetAccumTot.bas
'//---------------------------------------------------------------
'// STGetAccumTot - Get accumulated splits total.
'//		6/18/20.	wmk.
'//---------------------------------------------------------------

public function STGetAccumTot() As Double

'//	Usage.	dVar = STGetAccumTot()
'//
'//	Entry.	gdAccumTot = accumulated split total
'//
'//	Exit.	dVar = accumulated splits total
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.		wmk.	original code
'//
'//	Notes. gdAccumTot is initialized to 0, then updated as user enters
'//	split entries
'//

'//	constants.

'//	local variables.
dim dRetValue	As Double		'// returned value

	'// code.
	dRetValue = gdAccumTot		'// return accumulated total
	STGetAccumTot = dRetValue
	
end function		'// end STGetAccumTot		6/18/20
'/**/
'// STGetObjFocus.bas
'//---------------------------------------------------------------
'// STGetObjFocus - Get control name with focus from dialog.
'//		6/22/20.	wmk.	11:00
'//---------------------------------------------------------------

public function STGetObjFocus() As String

'//	Usage.	sVar = STGetObjFocus()
'//
'// Entry.	puoSTDialog = ST dialog object
'//			puoSTDialog.<HasFocusObject>.Text = available
'//
'//	Exit.	sVar = control name from "HasFocus".Text in dialog
'//				 = "" if error, or not previously set
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6//20.		wmk.	original code
'//
'//	Notes. STSetObjFocus should be called by any <control>Listener
'//	event when the <control> gets focus. This will provide an easy
'// way for a control to get its own name during processing. The
'// STClrObjFocus function should be called when the <control> loses
'// focus so other controls may reuse the dialog field "HasFocus".
'// A third method STGetObjFocus will retrieve the current "HasFocus"
'// .Text field for use during the control processing.
'//

'//	constants.

'//	local variables.
dim sRetValue As String
dim oDlgControl		As Object		'// dialog control
dim sFocusName		As String		'// HasFocus text control field name

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	oDlgControl = puoSTDialog.getControl("HasFocusFld")
	sFocusName = oDlgControl.Text
msgBox("In STGetObjFocus sFocusName = '" + sFocusName + "'")
	sRetValue = sFocusName
	
NormalExit:	
	STGetObjFocus = sRetValue
	exit function

ErrorHandler:
	sRetValue = ""		'// abnormal return value
	msgBox("In STGetObjFocus - unprocessed error")
	GoTo NormalExit
	
end function 	'// end STGetObjFocus	6/22/20
'/**/
'// STPubVarsReset.bas
'//---------------------------------------------------------------------
'// STPubVarsReset - Reset public variables for Split Transaction dialog.
'//		6/27/20.	wmk.	14:45
'//---------------------------------------------------------------------

public function STPubVarsReset() As Integer

'//	Usage.	iVar = STPubVarsReset()
'//
'// Entry.
'//		puoSTDialog	= ST dialog object
'//		gbSTEditMode = true if <Edit Split> in progress; false otherwise
'//
'// ST dialog stored values.
'public gdSTAccumTot		As Double		'// split accumulated total
'public gdSTTotalAmt		As Double		'// total amount to split
'public gbSTSplitCredits	As Boolean		'// credit values are splits flag
'public gsSTDescs(0)		As String		'// array of split descriptions
'public gdSTAmts(0)		As Double		'// array of amounts
'public gdSTCOAs(0)		As String		'// array of COAs
'//		giLastIX		As Integer		'// last active row index for input
'//		gbSTEditMode = true if <Edit Split> in progress; false otherwise
'//
'//	Exit.	iVar = 0 - normal return
'//				   -1 - error in reset
'//				   +1 - reset skipped; edit in progress
'//			public vars initialized ONLY if gbSTEditMode = false
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	giLastIx added
'//	6/20/20.	wmk.	gsSTRefs initialized
'//	6/27/20.	wmk.	gbSTEditMode flag tested before resetting vars
'//
'//	Notes.	When these publics are cleared, the dialog controls should
'// also be reset by a call to STDlgCtrlsReset so its values are correct.


'//	constants.

'//	local variables.
dim iRetValue As Integer

	'// code.
	
	iRetValue = -1		'// set bad return
	ON ERROR GOTO ErrorHandler
	
	'// only initialize if not dialog not being loaded for edit mode
	if gbSTEditMode then
		iRetValue = 1	'// set edit mode; initialization skipped
		GoTo NormalExit
	endif
	
	'// gdsTotalAmt = itself; this should have been set up by dialog
	'//   invoking the Split Transaction dialog
	gdSTAccumTot = 0.			'// clear accumulated total
	giSTLastIx = 0				'// last active row index
	gbSTSplitCredits = true		'// assume Debits as total
	gsSTAcctFocus = ""			'// clear account object focus
	gsSTObjFocus = ""			'// object name with focus
	redim gsSTDescs(0)	'// scrap descriptions array
	redim gdSTAmts(0)	'// scrap amounts array
	redim gsSTCOAs(0)	'// scrap COAs array
	redim gsSTRefs(0)	'// scrap refs array
	iRetValue = 0

NormalExit:
	STPubVarsReset = iRetValue
	exit function

ErrorHandler:
	msgBox("STPubVarsReset - unprocessed error.")
	GoTo NormalExit
	
end function 	'// end STPubVarsReset		6/27/20
'/**/
'// STRef10LostFocus.bas
'//---------------------------------------------------------------
'// STRef10LostFocus - Handle RefFldn LostFocus event.
'//		6/25/20.	wmk.	13:30
'//---------------------------------------------------------------

public sub STRef10LostFocus()

'//	Usage.	macro call or
'//			call STRef10LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsRefs(0) = text from RefFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "RefFld10"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef10LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef10LostFocus	6/25/20
'/**/

'// STRef1LostFocus.bas
'//---------------------------------------------------------------
'// STRef1LostFocus - Handle AcctFldn LostFocus event.
'//		6/24/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STRef1LostFocus()

'//	Usage.	macro call or
'//			call STRef1LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsCOAs(0) = text from AcctFld1..10
'//
'// Calls.	STAcctLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.	wmk.	original code
'// 6/19/20.	wmk.	code modified to handle either focus event
'//	6/20/20.	wmk.	bug fix storing COA; add code to check/enable
'//						next row if this row complete
'//	6/24/20.	wmk.	common code moved to STAcctLostFocus; generic
'//						call set up to use common code function
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "RefFld1"					'// clear object focus tracker
	iStatus = STRefLostFocus("RefFld1")
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef1LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef1LostFocus	6/24/20
'/**/
'// STRef2LostFocus.bas
'//---------------------------------------------------------------
'// STRef2LostFocus - Handle AcctFldn LostFocus event.
'//		6/24/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STRef2LostFocus()

'//	Usage.	macro call or
'//			call STRef2LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsCOAs(0) = text from AcctFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; adapted from STRef2LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// use generic handler
	gsSTObjFocus = "RefFld2"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef2LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef2LostFocus	6/24/20
'/**/
'// STRef3LostFocus.bas
'//---------------------------------------------------------------
'// STRef3LostFocus - Handle AcctFldn LostFocus event.
'//		6/24/20.	wmk.	10:30
'//---------------------------------------------------------------

public sub STRef3LostFocus()

'//	Usage.	macro call or
'//			call STRef3LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsRefs(0) = text from AcctFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// use generic handler
	gsSTObjFocus = "RefFld3"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef3LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef3LostFocus	6/24/20
'/**/
'// STRef4LostFocus.bas
'//---------------------------------------------------------------
'// STRef4LostFocus - Handle RefFldn LostFocus event.
'//		6/25/20.	wmk.	13:30
'//---------------------------------------------------------------

public sub STRef4LostFocus()

'//	Usage.	macro call or
'//			call STRef4LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsRefs(0) = text from RefFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// use generic handler
	gsSTObjFocus = "RefFld4"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef4LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef4LostFocus	6/25/20
'/**/
'// STRef5LostFocus.bas
'//---------------------------------------------------------------
'// STRef5LostFocus - Handle RefFldn LostFocus event.
'//		6/25/20.	wmk.	13:30
'//---------------------------------------------------------------

public sub STRef5LostFocus()

'//	Usage.	macro call or
'//			call STRef5LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsRefs(0) = text from RefFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// use generic handler
	gsSTObjFocus = "RefFld5"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef5LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef5LostFocus	6/25/20
'/**/
'// STRef6LostFocus.bas
'//---------------------------------------------------------------
'// STRef6LostFocus - Handle RefFldn LostFocus event.
'//		6/25/20.	wmk.	13:30
'//---------------------------------------------------------------

public sub STRef6LostFocus()

'//	Usage.	macro call or
'//			call STRef6LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsRefs(0) = text from RefFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// use generic handler
	gsSTObjFocus = "RefFld6"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef6LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef6LostFocus	6/25/20
'/**/
'// STRef7LostFocus.bas
'//---------------------------------------------------------------
'// STRef7LostFocus - Handle RefFldn LostFocus event.
'//		6/25/20.	wmk.	13:30
'//---------------------------------------------------------------

public sub STRef7LostFocus()

'//	Usage.	macro call or
'//			call STRef7LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsRefs(0) = text from RefFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "RefFld7"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef7LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef7LostFocus	6/25/20
'/**/
'// STRef8LostFocus.bas
'//---------------------------------------------------------------
'// STRef8LostFocus - Handle RefFldn LostFocus event.
'//		6/25/20.	wmk.	13:30
'//---------------------------------------------------------------

public sub STRef8LostFocus()

'//	Usage.	macro call or
'//			call STRef8LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsRefs(0) = text from RefFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "RefFld8"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef8LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef8LostFocus	6/25/20
'/**/
'// STRef9LostFocus.bas
'//---------------------------------------------------------------
'// STRef9LostFocus - Handle RefFldn LostFocus event.
'//		6/25/20.	wmk.	13:30
'//---------------------------------------------------------------

public sub STRef9LostFocus()

'//	Usage.	macro call or
'//			call STRef9LostFocus()
'//
'// Entry.	RefFld1..RefFld10 lost focus
'//			gsSTObjFocus = control name with focus (e.g. "RefFld1")
'//
'//	Exit.	gsRefs(0) = text from RefFld1..10
'//
'// Calls.	STRefLostFocus
'//
'//	Modification history.
'//	---------------------
'//	6/25/20.	wmk.	original code; adapted from STRef1LostFocus
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "RefFld9"					'// clear object focus tracker
	iStatus = STRefLostFocus(gsSTObjFocus)
	if iStatus < 0 then
		GoTo ErrorHandler
	endif

NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STRef9LostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STRef9LostFocus	6/25/20
'/**/
'// STRefLostFocus.bas
'//---------------------------------------------------------------
'// STRefLostFocus - Handle any RefFld LostFocus event.
'//		8/26/22.	wmk.	16:04
'//---------------------------------------------------------------

public function STRefLostFocus(psFldName As String) As Integer

'//	Usage.	macro call or
'//			call STRefLostFocus(sFieldName)
'//
'//	Parameters.	sFldName = name of dialog object lost focus
'//
'// Entry.	RerFld1..RefFld10 lost focus
'//			gsSTObjFocus = name of object with focus
'//			gdSTAccumTot = accumulated split total
'//			gdSTTotalAmt = total amount to split
'//
'//	Exit.	gsSTRefs(n) = text from RefFld1..10
'//
'// Calls.	STFld6ToIndex.
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; cloned from STRef1LostFocus.
'// 8/26/22.	wmk.	comments corrected.
'//
'//	Notes. gsSTObjFocus will be used by the [...] button Execute
'//	event to determine which Ref array string to store the selected
'// Ref into.


'//	constants.

'//	local variables.
dim iRetValue 		As Integer			'// returned status
dim iStatus			As Integer			'// general status
dim iRowIx			As Integer			'// index of present row
dim oDlgControl		As Object			'// field ptr
dim bRowComplete	As Boolean			'// row complete flag
dim bDlgComplete	As Boolean			'// dialog complete flag
dim sFldName 		As String			'// local copy of field name

	'// code.
	iRetValue = -1						'// set bad return
	ON ERROR GOTO ErrorHandler
'// Note: the callout to STUpdateRefArray was mysteriously getting messed up
'// and not exiting properly. ODDLY, this may have been because of an unhandled
'// error in the AcctFld handlers where STAcctLostFocus was removed from the code
'// inadvertantly, but still called and not caught in the runtime error handlers...

	sFldName = psFldName				'// use local copy of field name
	oDlgControl = puoSTDialog.getControl(sFldName)
	iRowIx = STFld6ToIndex(sFldName)
	gsSTRefs(iRowIx) = oDlgControl.Text
	
	iRetValue = 0

NormalExit:
	STRefLostFocus = iRetValue
	exit function
	
ErrorHandler:
	msgBox("In STRefLostFocus - unprocessed error")
	GoTo NormalExit
	
end function		'// end STRefLostFocus	8/26/22
'/**/
'// STRowComplete.bas
'//---------------------------------------------------------------
'// STRowComplete - Check if entry row has minimum information.
'//		6/20/20.	wmk.	20:00
'//---------------------------------------------------------------

public function STRowComplete(piRowIx As Integer) As Boolean

'//	Usage.	bVal = STRowComplete(iRowIx)
'//
'//	Parameters.	iRowIx = row to check for minimum information
'//
'//	Entry	gdSTAmts(iRowIx) has stored value
'//			gsSTCOAs(iRowIx) has 4 digits
'// Entry.	giSTLastIx = index of last stored row from dialog
'//			gdSTAmts(0..giSTLastIx0 = stored amounts from dialog
'//			gsSTCOAs(0..giSTLastIx) = stored account #s from dialog
'//
'//	Exit.	bVal = true if ALL of the following:
'//					gdSTAmts(0..STLastIx) = non-empty values
'//					gsSTCOAs(0..STLastIx) = non-empty values
'//					gdSTTotalAmt = gdSTAccumTot
'//				   false otherwise
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'// 6/20/20.	wmk.	bug fix with return value not being set
'//
'//	Notes. <Insert notes here>
'//

'//	constants.

'//	local variables.
dim bRetValue As Boolean		'// returned flag
dim bRowComplete	As Boolean	'// interim tests flag

	'// code.
	bRowComplete = (Len(gsSTCOAs(piRowIx)) = 4)
	bRowComplete = bRowComplete AND (gdSTAmts(piRowix) >= 0)	
	STRowComplete = bRowComplete
'STRowComplete = true	'//$
end function 	'// end STRowComplete	6/20/20
'/**/
'// STSetAccumTot.bas
'//---------------------------------------------------------------
'// STSetAccumTot - Get accumulated splits total.
'//		6/18/20.	wmk.	18:00
'//---------------------------------------------------------------

public function STSetAccumTot(pdTotal As Double) As Integer

'//	Usage.	iVar = STSetAccumTot(dTotal)
'//
'//	Paramters.	dTotal = new total to set in AccumTot field
'//
'//	Entry.	gdAccumTot = allocated for accumulated split total
'//			puoSTDialog	= loaded ST dialog object
'//
'//	Exit.	iVar = 0 - normal return
'//				   -1 - error setting accumulated total
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.		wmk.	original code
'//
'//	Notes. gdAccumTot is initialized to 0, then updated as user enters
'//	split entries
'//

'//	constants.

'//	local variables.
dim iRetValue	As Integer		'// returned value
dim oSTAccumTotFld	As Object	'// accumulated total text field

	'// code.
	ON ERROR GOTO ErrorHandler
	iStatus = 0					'// normal return
	gdAccumTot = pdTotal	

	'// update the AccumTot text box in the ST dialog
	oSTAccumTotFld = puoSTDialog.getControl("AccumTot")
	oSTAccumTotFld.Value = pdAccumTotal
	
NormalExit:
	STSetAccumTot = iStatus
	exit function
	
ErrorHandler:	
	iRetValue = -1
	GoTo NormalExit
end function		'// end STSetAccumTot		6/18/20
'/**/
'// STSetObjFocus.bas
'//---------------------------------------------------------------
'// STSetObjFocus - Set control name with focus inside object.
'//		6/22/20.	wmk.	10:00
'//---------------------------------------------------------------

public function STSetObjFocus(psControl As String) As Integer

'//	Usage.	iVar = STSetObjFocus( sControl )
'//
'//		sControl = dialog control name with focus
'//					(e.g. "AmtFld1")
'//
'// Entry.	puoSTDialog = ST dialog object
'//			puoSTDialog.<HasFocusObject>.Text = available
'//
'//	Exit.	iVar = 0 - control name stored in dialog
'//				 < 0 - error attempting to store active control namne
'//			puoSTDialog.<HasFocus-Object>.Text = psControl
'//
'// Calls.	<other subs/functions called>
'//
'//	Modification history.
'//	---------------------
'//	6/18/20.		wmk.	original code
'//
'//	Notes. STSetObjFocus should be called by any <control>Listener
'//	event when the <control> gets focus. This will provide an easy
'// way for a control to get its own name during processing. The
'// STClrObjFocus function should be called when the <control> loses
'// focus so other controls may reuse the dialog field "HasFocus".
'// A third method STGetObjFocus will retrieve the current "HasFocus"
'// .Text field for use during the control processing.
'//

'//	constants.

'//	local variables.
dim iRetValue As Any
dim oDlgControl		As Object		'// dialog control
dim sFocusName		As String		'// HasFocus text control field name

	'// code.
	
	iRetValue = -1
	ON ERROR GOTO ErrorHandler
msgBox("In STSetObjFocus.. calling .getControl(..")
	oDlgControl = puoSTDialog.getControl("HasFocusFld")
msgBox("In STSetObjFocus back from .getControl on entry.." + CHR(13)+CHR(10) _
+	"stored name = '" + oDlgControl.Text + "'")
	sFocusName = oDlgControl.Text

'	if Len(sFocusName) > 0 then
'		iRetValue = -2		'// set busy flag
'		GoTo ErrorHandler
'	endif		'// end HasFocus in use conditional

	oDlgControl.Text = psControl		'// set control name
msgBox("In STSetObjFocus sFocusName = '" + psControl + "'")
	
	iRetValue = 0			'// set normal return
	
NormalExit:	
	STSetObjFocus = iRetValue
	exit function

ErrorHandler:
	Select Case iRetValue
	Case -2
		msgBox("In STSetObjFocus - 'HasFocus' field busy")
	Case else
		msgBox("In STSetObjFocus - unprocessed error")
	End Select
	
	GoTo NormalExit
	
end function 	'// end STSetObjFocus	6/22/20
'/**/
'// STTotalAcctLostFocus.bas
'//---------------------------------------------------------------
'// STTotalAcctLostFocus - Handle TotalAcctFld LostFocus event.
'//		6/27/20.	wmk.	11:15
'//---------------------------------------------------------------

public sub STTotalAcctLostFocus()

'//	Usage.	macro call or
'//			call STTotalAcctLostFocus()
'//
'// Entry.	TotalAcctFld lost focus
'//			gbSTSplitCredits = true if Debit is parent account
'//							   false if Credit is parent acccount
'//			puoSTDialog = ST dialog object
'//
'//	Exit.	if gbSTSplitCredits = true; gsETAcct1 = user input from field
'//								false; gsETAcct2 = user input from field
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/27/20.	wmk.	original code


'//	constants.

'//	local variables.
dim iStatus			As Integer			'// general status
dim oDlgCtrl		As Object			'// dialog control

	'// code.
	ON ERROR GOTO ErrorHandler
	'// attempt to use generic handler
	gsSTObjFocus = "TotalAcctFld"	
	oDlgCtrl = puoSTDialog.getControl(gsSTObjFocus)
	if gbSTSplitCredits then
		gsETAcct1 = oDlgCtrl.Text
	else
		gsETAcct2 = oDlgCtrl.Text
	endif
	
NormalExit:
	exit sub
	
ErrorHandler:
	msgBox("In STTotalAcctLostFocus - unprocessed error")
	GoTo NormalExit
	
end sub		'// end STTotalAcctLostFocus	6/27/20
'/**/
'// STUpdateAccumTot.bas
'//---------------------------------------------------------------
'// STUpdateAccumTot - Update accumulated split total.
'//		7/7/20.	wmk.	12:00
'//---------------------------------------------------------------

public function STUpdateAccumTot(psAmtFldName As String) As Integer

'//	Usage.	iVar = STUpdateAccumTot( sAmtFldName )
'//
'// Parameters.	sAmtFldName dialog amount field name (e.g. "AmtFld1")
'//
'// Entry.	This routine should be called on the "LostFocus" event
'//			on any AmtFld in the dialog, with the fieldname passed
'//			as the parameter.	
'//			puoSTDialog = Split Transaction dialog object
'//
'//	Exit.	iVar = 0 - totals updated; limit not reached
'//				   1 - totals updated; limit reached
'//				 < 0 - error updating totals
'//			gdStAmts(i) = matching index i updated with new value from dialog
'//			gdSTAccumTot = updated with any old value from array
'//					subtracted and new value from dialog added
'//			puoSTDialog.UnsplitTot.Value = updated with any old value
'//					from array added and new value from dialog subtracted
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/19/20.	wmk.	original code
'//	6/21/20.	wmk.	change to check split total met and return 1
'//	7/7/20.		wmk.	change to string comparison to allow for values
'//						equal except fractional cents; properly update
'//						totals without exception; fix bug where storing string
'//						into gdSTAmts array instead of value
'//
'//	Notes. The Accumulated total field is initialized to 0 at dialog
'//	startup, then is updated whenever the user adds a split amount to
'// any of the AmtFld1..AmtFld10 fields. STUpdateAccumTot should be
'// called whenever an AmtFldn object changes; The dialog UnsplitTot
'// field is also updated to hold the tracking
'//iAmtFldIx = index of ST dialog AmtFld to add to
'//			accumulator. If the amount field in the dialog has
'//			been cleared, any current value in the psSTAmts array
'//			will be subtracted from the accumulated total, the
'//			new value from the dialog box will be both stored in 
'//			the gsSTAmts array and added to the accumulated total.
'//			Then the Accumulated Splits Total box will be updated
'//			in the dialog objects.

'//	constants.

'//	local variables.
dim iRetValue 	As Integer		'// returned value
dim iAmtFldIx	As Integer		'// amounts array index
dim sAmtFld		As String		'// amount field local name
dim oAmtFld		As Object		'// amount field from dialog
dim oAccumTot	As Object		'// accumulated total field
dim dUnsplit	As Double		'// unsplit current amount

	'// code.

	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sAmtFld = psAmtFldName		'// get local copy of name
	
	'// convert name to array index
	iAmtFldIx = STFld6ToIndex(sAmtFld)
	if iAmtFldIx < 0 then
		GoTo ErrorHandler
	endif
	
	'//	adjust both accumulators by old value
	gdSTAccumTot = gdSTAccumTot - gdSTAmts(iAmtFldIx)
	oAccumTot = puoSTDialog.getControl("UnsplitTot")
	dUnsplit = oAccumTot.Value
	oAccumTot.Value = dUnsplit + gdSTAmts(iAmtFldIx)
	
	
	'// get new value from split transaction dialog and add
	'// into accumulator
	oAmtFld = puoSTDialog.getControl(sAmtFld)
	gdSTAccumTot = gdSTAccumTot + oAmtFld.Value
	
	if gdSTAccumTot > gdSTTotalAmt then
		if StrComp(Str(gdSTAccumTot),Str(gdSTTotalAmt))= 0 then
			gdSTAccumTot = gdSTTotalAmt
		else
			msgBox("Splits with new amount exceeds Total to split")
			iRetValue = 0
			GoTo NormalExit		'// bail out without updating array or control
		endif
	endif
	
	'// store new value in gdSTAmts array for this index
	gdStAmts(iAmtFldIx) = oAmtFld.Value
	
	'// update transaction dialog Accumulated Splits and Unsplit field
	oAccumTot = puoSTDialog.getControl("AccumTot")
	oAccumTot.Value = gdSTAccumTot
	oAccumTot = puoSTDialog.getControl("UnsplitTot")
	dUnsplit = oAccumTot.Value
	oAccumTot.Value = dUnsplit - oAmtFld.Value
		
'	if gdSTAccumTot = gdSTTotalAmt then
	if StrComp(Str(gdSTAccumTot),Str(gdSTTotalAmt))=0 then
		iRetValue = 1
	else
		iRetValue = 0
	endif	'// end split total met conditional
	
NormalExit:
	STUpdateAccumTot = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STUpdateAccumTot - unprocessed error.")
	GoTo NormalExit
	
end function		'// end STUpdateAccumTot	7/7/20
'/**/
'// STUpdateCOAArray.bas
'//---------------------------------------------------------------
'// STUpdateCOAArray - Update public COAs array.
'//		6/24/20.	wmk.	07:15
'//---------------------------------------------------------------

public function STUpdateCOAArray(psFldName As String) As Integer

'//	Usage.	iVar = STUpdateCOAArray( sFldName )
'//
'// 		sFldName dialog amount field name (e.g. "DescFld1")
'//
'// Entry.	This routine should be called on the "LostFocus" event
'//			on any DescFldn in the dialog, with the fieldname passed
'//			as the parameter.	
'//			puoSTDialog = Split Transaction dialog object
'//
'//	Exit.	iVar = 0 - Desc array updated ok
'//				 < 0 - error updating Desc Array
'//			gdSTCOAs(i) = matching index i updated with new string
'//					 from dialog
'//
'// Calls.	STFld6ToIndex
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; adapted from STUpdateDescArray
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue 	As Integer		'// returned value
dim iDlgFldIx	As Integer		'// amounts array index
dim sDlgFld		As String		'// amount field local name
dim oDlgFld		As Object		'// amount field from dialog

	'// code.

	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sDlgFld = psFldName		'// get local copy of name
	
	'// convert name to array index
	iDlgFldIx = STFld6ToIndex(sDlgFld)
	if iDlgFldIx < 0 then
		GoTo ErrorHandler
	endif
	
	'// get new field dialog and store in array
	oDlgFld = puoSTDialog.getControl(sDlgFld)
	gdSTCOAs(iDlgFldIx) = oDlgFld.Text
	
	iRetValue = 0			'// normal return
	
NormalExit:
	STUpdateCOAArrray = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STUpdateCOAArray - unprocessed error.")
	GoTo NormalExit
	
end function		'// end STUpdateCOAArray	6/24/20
'/**/
'// STUpdateDescArray.bas
'//---------------------------------------------------------------
'// STUpdateDescArray - Update public Descriptors array.
'//		6/24/20.	wmk.	07:15
'//---------------------------------------------------------------

public function STUpdateDescArray(psFldName As String) As Integer

'//	Usage.	iVar = STUpdateDescArray( sFldName )
'//
'// 		sFldName dialog amount field name (e.g. "DescFld1")
'//
'// Entry.	This routine should be called on the "LostFocus" event
'//			on any DescFldn in the dialog, with the fieldname passed
'//			as the parameter.	
'//			puoSTDialog = Split Transaction dialog object
'//
'//	Exit.	iVar = 0 - Desc array updated ok
'//				 < 0 - error updating Desc array
'//			gdSTDescs(i) = matching index i updated with new string
'//					 from dialog
'//
'// Calls.	STFld7ToIndex
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; adapted from STUpdateAccumTot
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue 	As Integer		'// returned value
dim iDlgFldIx	As Integer		'// amounts array index
dim sDlgFld		As String		'// amount field local name
dim oDlgFld		As Object		'// amount field from dialog

	'// code.

	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sDlgFld = psFldName		'// get local copy of name
	
	'// convert name to array index
	iDlgFldIx = STFld7ToIndex(sDlgFld)
	if iDlgFldIx < 0 then
		GoTo ErrorHandler
	endif
	
	'// get new field dialog and store in array
	oDlgFld = puoSTDialog.getControl(sDlgFld)
	gdSTDescs(iDlgFldIx) = oDlgFld.Text
	
	iRetValue = 0			'// normal return
	
NormalExit:
	STUpdateDescArrray = iRetValue
	exit function
	
ErrorHandler:
	msgBox("STUpdateDescArray - unprocessed error.")
	GoTo NormalExit
	
end function		'// end STUpdateDescArray	6/24/20
'/**/
'// STUpdateRefArray.bas
'//---------------------------------------------------------------
'// STUpdateRefArray - Update public Reference array.
'//		6/24/20.	wmk.	06:45
'//---------------------------------------------------------------

public function STUpdateRefArray(psFldName As String) As Integer

'//	Usage.	iVar = STUpdateRefArray( sFldName )
'//
'// 		sFldName dialog amount field name (e.g. "RefFld1")
'//
'// Entry.	This routine should be called on the "LostFocus" event
'//			on any REfFldn in the dialog, with the fieldname passed
'//			as the parameter.	
'//			puoSTDialog = Split Transaction dialog object
'//
'//	Exit.	iVar = 0 - Ref array updated ok
'//				 < 0 - error updating Ref Array
'//			gsSTRefs(i) = matching index i updated with new string
'//					 from dialog
'//
'// Calls.	STFld7ToIndex
'//
'//	Modification history.
'//	---------------------
'//	6/24/20.	wmk.	original code; adapted from STUpdateDescArray
'//
'//	Notes.

'//	constants.

'//	local variables.
dim iRetValue 	As Integer		'// returned value
dim iDlgFldIx	As Integer		'// amounts array index
dim sDlgFld		As String		'// amount field local name
dim oDlgFld		As Object		'// amount field from dialog

	'// code.
'//WARNING: IT APPEARS THAT THE STACK GETS MESSED UP WHEN GOING 2 LEVELS
'// DOWN IN AN EVENT HANDLER... THIS ROUTINE WILL REMAIN IN PLACE, BUT
'// CANNOT BE CALLED DOWNSTREAM BY "LOSTFOCUS" EVENT HANDLING...
'// Experiment with making this a parameterless function and see if the
'// situation improves...

	iRetValue = -1
	ON ERROR GOTO ErrorHandler
	sDlgFld = psFldName		'// get local copy of name
msgBox("In STUpdateRefArray.. psFldName = '" + psFldName + "'" + CHR(13)+CHR(10) _
+	"sDlgFld = '" + sDlgFld + "'")	
	'// convert name to array index
	iDlgFldIx = STFld6ToIndex(sDlgFld)
msgBox(".. iDlgFldIx = " + iDlgFldIx + "'")	
	if iDlgFldIx < 0 then
		GoTo ErrorHandler
	endif
	
	'// get new field dialog and store in array

msgBox(".. invoking .getControl")	
	
	oDlgFld = puoSTDialog.getControl(sDlgFld)
msgBox(".. back from .getControl")	
	
	gsSTRefs(iDlgFldIx) = oDlgFld.Text
msgBox(".. gsSTRefs() value stored")	
	
	iRetValue = 0			'// normal return
	GoTo NormalExit
	
ErrorHandler:
	msgBox("STUpdateRefArray - unprocessed error.")

NormalExit:
msgBox("Exiting STUpdateRefArray; iRetValue = " + iRetValue)
	STUpdateRefArrray = iRetValue
	
end function		'// end STUpdateRefArray	6/24/20
'/**/
'// VSDoneBtnRun.bas
'//---------------------------------------------------------------------
'// VSDoneBtnRun - Event handler <Done> from View Split.
'//		6/27/20.	wmk.	17:00
'//---------------------------------------------------------------------

public sub VSDoneBtnRun()

'//	Usage.	macro call or
'//			call VSDoneBtnRun()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked the [Record & Finish] button in ET dialog
'//			puoETDialog = Enter Transaction dialog object
'//
'//	Exit.	Transaction data recorded in GL sheet at user selection
'//			Transaction data in public vars cleared
'//			ET dialog ended with flag = 2 (user closed)
'//
'// Calls.	ETDialogRecord, ETPubVarsReset
'//
'//	Modification history.
'//	---------------------
'//	6/27/20.	wmk.	original code
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Done> cmd button in the View Split  dialog.

'//	constants.

'//	local variables.
dim oVSDoneBtn	As Object		'// Record & Finish button
dim iStatus 		As Integer		'// general status

	'// code.
	ON ERROR GoTo ErrorHandler

NormalExit:
	puoVSDialog.endDialog(0)		'// end dialog; same as Cancel
	exit sub
	
ErrorHandler:
	msgBox("VSDoneBtnRun - unprocessed error.")
	GoTo NormalExit
	
end sub		'// end VSDoneBtnRun	6/17/20
'/**/
'// VSEditSplitBtnRun.bas
'//---------------------------------------------------------------------
'// VSEditSplitBtnRun - Event handler <Done> from View Split.
'//		6/27/20.	wmk.	17:00
'//---------------------------------------------------------------------

public sub VSEditSplitBtnRun()

'//	Usage.	macro call or
'//			call VSEditSplitBtnRun()
'//
'// Entry.	normal entry from event handler for dialogue where
'//			user clicked the <Edit Split> button in VS dialog
'//			puoVSDialog = View Transaction dialog object
'//
'//	Exit.	VS dialog ended with flag = 2 (Edit Split)
'//
'// Calls.
'//
'//	Modification history.
'//	---------------------
'//	6/27/20.	wmk.	original code
'//
'//	Notes. This sub is the linked macro to the changed status event linked
'// to the <Edit Split> cmd button in the View Split  dialog.

'//	constants.

'//	local variables.

	'// code.
	ON ERROR GoTo ErrorHandler
	gbSTEditMode = true				'// set ST edit mode flag
	
NormalExit:
	puoVSDialog.endDialog(2)		'// end dialog; same as Cancel
	exit sub
	
ErrorHandler:
	msgBox("VSEditSplitBtnRun - unprocessed error.")
	GoTo NormalExit
	
end sub		'// end VSEditSplitBtnRun	6/27/20
'/**/
