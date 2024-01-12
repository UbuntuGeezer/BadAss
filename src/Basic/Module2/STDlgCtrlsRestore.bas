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
