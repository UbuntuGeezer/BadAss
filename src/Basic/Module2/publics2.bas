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
