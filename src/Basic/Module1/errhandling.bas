'// errhandling.bas
'//---------------------------------------------------------------
'// errhandling.bas - Definitions for development error handling.
'//		wmk. 5/26/20.	17:30
'//---------------------------------------------------------------
'//
'// Header Description.
'// -------------------
'//	errhandling.bas contains definitions that may be inserted in any
'// OOoBasic Module for handing error conditions. This is intended for
'// use with development modules, and should be removed for Production
'// spreadsheets. This will ensure clean code released into the
'// Production environment.
'//
'// error handling interface methods.
'// ErrLogSetup - set up error logging (gbErrLogging, goCellRangeAddress, gsErrModule)
'// ErrLogDisable() - disable error logging by setting gbErrLogging = false
'// ErrLogGetCellInfo(lColumn, lRow) - get error focus cell column, row
'// ErrLogSetCellInfo(lColumn, lRow) - set error focus cell column, row
'//	ErrLogGetCellName() - get error focus cell alphanumeric name
'//	ErrLogGetDisplay() - get error message box enabled status
'//	ErrLogSetDisplay(bDisplayOn) - set error message box enabled status
'//	ErrLogGetModule() - get current module name gsErrModule
'// ErrLogSetModule(sName) - set current module name gsErrModule
'// ErrLogGetRecording - Get error recording status from error log globals
'// ErrLogSetRecording - Set error recording status in error log globals
'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
'// LogError - make error log entry([goCellRangeAddress], psErrName, psErrMsg)

'// global constants.

'// public constants.
public const ERRLOGSHEET="Error Log"			'// error logging sheet name
public const ERRLOGINSROW=5						'// insertion row for log messages
'// global variables.
'// global variables.
global gbErrDisplay	As Boolean					'// msgBox on/off flag
global gbErrLogging as Boolean					'// logging on/off flag; LogSetup sets/clears
global gbErrRecording As Boolean				'// log sheet recording on/off flag
global gsErrSheet As String						'// name of sheet error thrown on
global gsErrModule As String					'// sub/function throwing error
global gsErrName As String						'// name of error code
global gsErrMsg As String						'// error message
global goErrRangeAddress As Object				'// cell range address

'// public variables.
'//

'// error handling interface methods.
'// ErrLogSetup - set up error logging (gbErrLogging, goCellRangeAddress, gsErrModule)
'// ErrLogDisable() - disable error logging by setting gbErrLogging = false
'// ErrLogGetCellInfo(lColumn, lRow) - get error focus cell column, row
'// ErrLogSetCellInfo(lColumn, lRow) - set error focus cell column, row
'//	ErrLogGetCellName() - get error focus cell alphanumeric name
'//	ErrLogGetDisplay() - get error message box enabled status
'//	ErrLogSetDisplay(bDisplayOn) - set error message box enabled status
'//	ErrLogGetModule() - get current module name gsErrModule
'// ErrLogSetModule(sName) - set current module name gsErrModule
'// ErrLogGetRecording - Get error recording status from error log globals
'// ErrLogSetRecording - Set error recording status in error log globals
'// ErrLogGetSheet - Get sheet index from error log globals.
'// ErrLogSetSheet - Set sheet index in error log globals.
'// LogError - make error log entry([goCellRangeAddress], psErrName, psErrMsg)
'//-----------------end errhandling.bas header--------------------------
'/**/

