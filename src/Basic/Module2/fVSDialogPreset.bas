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
