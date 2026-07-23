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
