#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: WebDriver_Extended
	Description...........: Extended functions for WebDriver_Core.au3	
    Global dependencies...: WebDriver_Core.au3;	WebDriver_Globals.au3; webdriver and browser version compatibility

    Author................: exorcistas@github.com
    Modified..............: 2021-02-09
    Version...............: v0.7.1.1rc
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once
#include <WebDriver_Core.au3>

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
	_WD_Ext_TableElementToArray($_sSessionId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH, $_sStartNode = '')
	_WD_Ext_WaitElement($_sSessionId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH, $_sStartNode = '', $_iTimeout = 30000)
	_WD_Ext_Attach($_sSessionId, $_sText, $_sMode = 'title')
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region EXTENDED_FUNCTIONS

	Func _WD_Ext_TableElementToArray($_sSessionId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH, $_sStartNode = '')
		Local $_iTableElement = _WD_FindElement($_sSessionId, $_sSelector, $_sLocatorStrategy, $_sStartNode)
			If @error Then 
				Return SetError(@error, @extended, "")
			ElseIf (_WD_GetElementTagName($_sSessionId, $_iTableElement) <> "table") Then
				Return SetError($_WD_ERROR_ID_NotFound, 0, "")
			EndIf

		Local $_sHTML = _WD_GetElementProperty($_sSessionId, $_iTableElement, "outerHTML")
			If @error Then Return SetError(@error, @extended, "")
	
		Local $_aRegExTable = StringRegExp($_sHTML, "(?isU)(?|<(/)tr>\s*|<t[dh].*>(.*)</t[dh]>)", 3)
		Local $_aHTMLTable[ UBound($_aRegExTable) ] [ UBound($_aRegExTable) ]
		Local $_iRow = 0, $_iCol = 0, $_iMaxRow = 0

			For $i = 0 To UBound($_aRegExTable) - 1
				If $_aRegExTable[$i] = "/" Then
					$_iRow += 1
					$_iCol = 0
				Else
					$_aHTMLTable[$_iRow][$_iCol] = $_aRegExTable[$i]
					$_iCol += 1
					If $_iCol > $_iMaxRow Then $_iMaxRow = $_iCol
				EndIf
			Next
		Redim $_aHTMLTable[$_iRow][$_iMaxRow]
			;~ _ArrayDisplay($_aHTMLTable)

		Return SetError($_WD_ERROR_ID_Success, $_WD_HTTPRESULT, $_aHTMLTable)
	EndFunc

	Func _WD_Ext_WaitElement($_sSessionId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH, $_sStartNode = '', $_iTimeout = 30000)
		Local $_sElementId = ""
		Local $hWaitTimer = TimerInit()

		While 1
			$_sElementId = _WD_FindElement($_sSessionId, $_sSelector, $_sLocatorStrategy, $_sStartNode)
				If @error Then 
					If (TimerDiff($hWaitTimer) > $_iTimeout) Then
						$_iErrorCode = $_WD_ERROR_ID_Timeout
						ExitLoop
					EndIf
				Else
					If $_sElementId <> "" Then
						$_iErrorCode = $_WD_ERROR_ID_Success
						ExitLoop
					EndIf
				EndIf
			Sleep(100)
		WEnd

		Return SetError($_iErrorCode, $_WD_HTTPRESULT, $_sElementId)
	EndFunc

	Func _WD_Ext_Attach($_sSessionId, $_sText, $_sMode = 'title')
		If (($_sText = "") OR ($_sMode = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")
		
		Local $_aWinHandles = _WD_GetWindowHandles($_sSessionId)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sHandle
		For $_sHandle in $_aWinHandles

			_WD_SwitchToWindow($_sSessionId, $_sHandle)
				If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sMode = StringLower($_sMode)
			Switch $_sMode
				Case "title"
					Local $_sTitle = _WD_GetTitle($_sSessionId)
					If StringUpper($_sTitle) = StringUpper($_sText) Then Return SetError($_WD_ERROR_ID_Success, $_WD_HTTPRESULT, $_sHandle)

				Case "url"
					Local $_sUrl = _WD_GetCurrentURL($_sSessionId)
					If StringUpper($_sUrl) = StringUpper($_sText) Then Return SetError($_WD_ERROR_ID_Success, $_WD_HTTPRESULT, $_sHandle)

				Case Else
					ContinueLoop
			EndSwitch
		Next

		Return SetError($_WD_ERROR_ID_NotFound, $_WD_HTTPRESULT, 0)
	EndFunc

#EndRegion EXTENDED_FUNCTIONS