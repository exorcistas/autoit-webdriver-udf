#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: WebDriver_JS_Scripts
	Description...........: JavaScript functions for WebDriver_Core.au3

    Global dependencies...: WebDriver_Core.au3;	WebDriver_Globals.au3; webdriver and browser version compatibility
	Documentation.........: https://w3c.github.io/webdriver/
							https://www.w3schools.com/js/DEFAULT.asp
							https://developer.mozilla.org/en-US/docs/Web/API/

    Author................: exorcistas@github.com
    Modified..............: 2021-07-14
    Version...............: v0.2rc
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once
#include <WebDriver_Core.au3>

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
	_WD_JS_HighlightElement($_sSessionId, $_sElementId)
	_WD_JS_ScrollElementIntoView($_sSessionId, $_sElementId)
	_WD_JS_LoadWait($_sSessionId, $_iTimeout = 30000, $_iExtraTimeout = "")
	_WD_JS_GetElementParent($_sSessionId, $_sElementId)
	_WD_JS_SetFocusOnElement($_sSessionId, $_sElementId)
	_WD_JS_SetCheckboxValue($_sSessionId, $_sElementId, $_bCheck = True)
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region JAVASCRIPT_FUNCTIONS

	Func _WD_JS_HighlightElement($_sSessionId, $_sElementId)
		Local $_sMethod = "background: #FFFF66; border-radius: 5px; padding-left: 3px;"
		Local $_sScript = "arguments[0].style='" & $_sMethod & "'; return true;"
		Local $_sJSON_Element = '{"' & $_WD_ELEMENT_ID & '":"' & $_sElementId & '"}'

		Local $_sResponse = _WD_ExecuteScript($_sSessionId, $_sScript, $_sJSON_Element)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;# https://www.w3schools.com/jsref/met_element_scrollintoview.asp
	Func _WD_JS_ScrollElementIntoView($_sSessionId, $_sElementId)
		Local $sJsonElement = '{"' & $_WD_ELEMENT_ID & '":"' & $_sElementId & '"}'
		$_sResponse = _WD_ExecuteScript($_sSessionId, "return arguments[0].scrollIntoView", $sJsonElement)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;# https://www.w3schools.com/jsref/prop_doc_readystate.asp
	Func _WD_JS_LoadWait($_sSessionId, $_iTimeout = 30000, $_iExtraTimeout = "")
		Local $hWaitTimer = TimerInit()

		While 1
			$_sResponse = _WD_ExecuteScript($_sSessionId, 'return document.readyState')
			$_sResponse = __WD_GetJSONValue($_sResponse, "[value]")

			If $_sResponse <> 'complete' Then
				If (TimerDiff($hWaitTimer) > $_iTimeout) Then
					$_iErrorCode = $_WD_ERROR_ID_Timeout
					ExitLoop
				EndIf
			Else
				$_iErrorCode = $_WD_ERROR_ID_Success
				ExitLoop
			EndIf
			Sleep(100)
		WEnd

		If $_iExtraTimeout <> "" Then Sleep($_iExtraTimeout)

		Return SetError($_iErrorCode, $_WD_HTTPRESULT)
	EndFunc

	;# https://www.w3schools.com/jsref/prop_node_parentelement.asp
	Func _WD_JS_GetElementParent($_sSessionId, $_sElementId)
		Local $sJsonElement = '{"' & $_WD_ELEMENT_ID & '":"' & $_sElementId & '"}'
		$_sResponse = _WD_ExecuteScript($_sSessionId, "return arguments[0].parentElement", $sJsonElement)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;# https://www.w3schools.com/jsref/met_html_focus.asp
	Func _WD_JS_SetFocusOnElement($_sSessionId, $_sElementId)
		Local $sJsonElement = '{"' & $_WD_ELEMENT_ID & '":"' & $_sElementId & '"}'
		$_sResponse = _WD_ExecuteScript($_sSessionId, "return arguments[0].focus()", $sJsonElement)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;# https://www.w3schools.com/jsref/prop_checkbox_checked.asp
	Func _WD_JS_SetCheckboxValue($_sSessionId, $_sElementId, $_bCheck = True)
		Local $sJsonElement = '{"' & $_WD_ELEMENT_ID & '":"' & $_sElementId & '"}'
		$_sResponse = _WD_ExecuteScript($_sSessionId, "return arguments[0].checked = " & $_bCheck, $sJsonElement)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

#EndRegion JAVASCRIPT_FUNCTIONS
