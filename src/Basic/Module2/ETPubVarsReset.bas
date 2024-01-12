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
