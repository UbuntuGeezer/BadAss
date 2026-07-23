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
