#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: WebDriver_Core
	Description...........: W3C WebDriver is a remote control interface that enables introspection and control of user agents.
							It provides a platform- and language-neutral wire protocol as a way for out-of-process programs
							to remotely instruct the behavior of web browsers (Chrome, MS Edge, Firefox)

	Dependencies..........: WebDriver_Globals.au3; Json.au3; WinHttp.au3
	                        webdriver and browser version compatibility

	Documentation.........: https://w3c.github.io/webdriver/
							https://www.w3.org/TR/webdriver/#endpoints

    Author................: exorcistas@github.com
    Modified..............: 2022-01-04
    Version...............: v0.9.7rc
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once
#include <Array.au3>
#include <File.au3>
#include <WinAPIProc.au3>
#include <Json.au3>
#include <WinHttp.au3>
#include <WebDriver_Globals.au3>

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
%% BASE	%%
	_WD_Configure($_sBrowserName = $_WD_DEFAULT_BROWSER, $_sWorkingDir = $_WD_DRIVER_DIR)
	_WD_StartupDriver($_bShowConsole = $_WD_DEBUG_CONSOLE)
	_WD_ShutdownDriver($_sDriver = $_WD_DRIVER)
	_WD_GetCurrentBrowserVersion($_sBrowserName = $_WD_CURRENT_BROWSER)
	_WD_GetCurrentDriverVersion($_sDriverExe = $_WD_DRIVER, $_sDriverInstallPath = $_WD_DRIVER_DIR)
	_WD_CheckUpdateRequired($_sCurrentBrowserVersion, $_sCurrentDriverVersion)
	_WD_CheckLatestAvailableDriverVersion($_sCurrentBrowserVersion, $_sBrowserName = $_WD_CURRENT_BROWSER)
	_WD_GetLatestDriverRelease($_sDriverVersion, $_sDriverExe = $_WD_DRIVER, $_sDriverInstallPath = $_WD_DRIVER_DIR)
	_WD_DownloadFile($_sURL, $_sDestinationPath)
	_WD_DriverInstallRelease($_sDriverReleaseZipPath, $_sDriverExe = $_WD_DRIVER, $_sDriverInstallPath = $_WD_DRIVER_DIR)

%% SESSIONS %%
	_WD_NewSession()
	_WD_DeleteSession($_sSessionId)
	_WD_Status()
%% TIMEOUTS %%
	_WD_SetTimeouts($_sSessionId, $_iPageLoadTimeout = 30000, $_iScriptTimeout = 30000, $_iImplicitTimeout = 0)
	_WD_GetTimeouts($_sSessionId)
%% NAVIGATION %%
	_WD_NavigateTo($_sSessionId, $_sURL)
	_WD_GetCurrentURL($_sSessionId)
	_WD_GetTitle($_sSessionId)
	_WD_Refresh($_sSessionId)
	_WD_Back($_sSessionId)
	_WD_Forward($_sSessionId)
%% CONTEXTS %%
	_WD_GetWindowHandle($_sSessionId)
	_WD_GetWindowHandles($_sSessionId)
	_WD_CloseWindow($_sSessionId)
	_WD_MaximizeWindow($_sSessionId)
	_WD_MinimizeWindow($_sSessionId)
	_WD_FullscreenWindow($_sSessionId)
	_WD_SwitchToWindow($_sSessionId, $_sWindowHandle)
	_WD_NewWindow($_sSessionId, $_bOpenTab = True)
	_WD_SwitchToFrame($_sSessionId, $_sFrameIndexOrId)
	_WD_SwitchToParentFrame($_sSessionId)
%% ELEMENTS %%
	_WD_FindElement($_sSessionId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH, $_sStartNode = '')
	_WD_FindElements($_sSessionId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH, $_sStartNode = '')
	_WD_FindElementFromShadowRoot($_sSessionId, $_sShadowId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH)
	_WD_FindElementsFromShadowRoot($_sSessionId, $_sShadowId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH)
	_WD_GetActiveElement($_sSessionId)
	_WD_GetElementShadowRoot($_sSessionId, $_sElementId)
	_WD_GetElementAttribute($_sSessionId, $_sElementId, $_sAttributeName)
	_WD_GetElementProperty($_sSessionId, $_sElementId, $_sPropertyName)
	_WD_GetElementText($_sSessionId, $_sElementId)
	_WD_GetElementTagName($_sSessionId, $_sElementId)
	_WD_GetElementRect($_sSessionId, $_sElementId)
	_WD_ElementClick($_sSessionId, $_sElementId)
	_WD_ElementClear($_sSessionId, $_sElementId)
	_WD_ElementSendKeys($_sSessionId, $_sElementId, $_sKeys)
%% DOCUMENT %%
	_WD_GetPageSource($_sSessionId)
	_WD_ExecuteScript($_sSessionId, $_sScript, $_sJSON_Arguments = "[]", $_bAsync = False)
%% COOKIES %%
	_WD_GetAllCookies($_sSessionId)
	_WD_GetNamedCookie($_sSessionId, $_sCookieName)
	_WD_AddCookie($_sSessionId, $_sCookieName, $_sCookieValue)
	_WD_DeleteCookie($_sSessionId, $_sCookieName)
	_WD_DeleteAllCookies($_sSessionId)
%% ACTIONS %%
	_WD_PerformActions($_sSessionId, $_sJSON_Actions)
	_WD_ReleaseActions($_sSessionId)
%% USER PROMPTS %%
	_WD_AcceptAlert($_sSessionId)
	_WD_GetAlertText($_sSessionId)
	_WD_SendAlertText($_sSessionId, $_sAlertText)
	_WD_DismissAlertText($_sSessionId)
%% SCREEN CAPTURE %%
	_WD_TakeScreenshot($_sSessionId, $_sElementId = '')
%% INTERNAL %%
	__WD_GET($_sURL)
	__WD_POST($_sURL, $_sData)
	__WD_DELETE($_sURL)
	__WD_GetJSONValue($_sResponse, $_sReturnValue = '')
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region BASE

	Func _WD_Configure($_sBrowserName = $_WD_DEFAULT_BROWSER, $_sWorkingDir = $_WD_DRIVER_DIR)
		DirCreate($_sWorkingDir)
		$_WD_CURRENT_BROWSER = StringUpper($_sBrowserName)

		Switch $_WD_CURRENT_BROWSER
			Case "MSEDGE", "EDGE"
				$_WD_CURRENT_BROWSER = "msedge"
				$_WD_DRIVER = $_WD_DRIVER_EDGE
				$_WD_CAPABILITIES = $_WD_CAPABILITIES_JSON_EDGE
				$_WD_DRIVER_PARAMS = $_WD_DRIVER_PARAMS_EDGE
				$_WD_PORT = $_WD_PORT_EDGE

			Case "CHROME"
				$_WD_CURRENT_BROWSER = "chrome"
				$_WD_DRIVER = $_WD_DRIVER_CHROME
				$_WD_CAPABILITIES = $_WD_CAPABILITIES_JSON_CHROME
				$_WD_DRIVER_PARAMS = $_WD_DRIVER_PARAMS_CHROME
				$_WD_PORT = $_WD_PORT_CHROME

			Case Else
				Return SetError($_WD_ERROR_ID_InvalidRequest, 0, False)
		EndSwitch

		$_WD_DRIVER_DIR = $_sWorkingDir
		If $_WD_DEBUG_CONSOLE Then ConsoleWrite(@CRLF & ">> [_WD_Configure] Driver full path set: " & $_WD_DRIVER_DIR & $_WD_DRIVER & @CRLF)

		Return SetError($_WD_ERROR_ID_Success, 0, True)
	EndFunc

	Func _WD_StartupDriver($_bShowConsole = $_WD_DEBUG_CONSOLE)
		If $_WD_DRIVER = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, 0)
		$_WD_DRIVER = $_WD_DRIVER_DIR & $_WD_DRIVER

		Local $_sCommand = StringFormat('"%s" %s ', $_WD_DRIVER, $_WD_DRIVER_PARAMS)
		If $_WD_DEBUG_CONSOLE Then
			ConsoleWrite(@CRLF & ">> [_WD_StartupDriver] OS:" & @TAB & @OSVersion & " " & @OSType & " " & @OSBuild & " " & @OSServicePack & @CRLF & _
					">> Driver:" & @TAB & $_WD_DRIVER & @CRLF & _
					">> Capabilities:" & @TAB & $_WD_CAPABILITIES & @CRLF & _
					">> Parameters:" & @TAB & $_WD_DRIVER_PARAMS & @CRLF & _
					">> Port:" & @TAB & $_WD_PORT & @CRLF)
		EndIf

		Local $_flagConsole = $_bShowConsole ? @SW_SHOW : @SW_HIDE
		Local $_iProcID = Run($_sCommand, "", $_flagConsole)
			If @error Then Return SetError($_WD_ERROR_ID_GeneralError, 0, 0)

		Sleep($_WD_DELAY_TIME)
		Return SetError($_WD_ERROR_ID_Success, 0, $_iProcID)
	EndFunc

	Func _WD_ShutdownDriver($_sDriver = $_WD_DRIVER)
		Local $_iProcID = 0, $_aData

		$_sDriver = StringRegExpReplace($_sDriver, "^.*\\(.*)$", "$1")
		Do
			$_iProcID = ProcessExists($_sDriver)
			If $_iProcID Then

				$_aData = _WinAPI_EnumChildProcess($_iProcID)
				If IsArray($_aData) Then
					For $i = 0 To UBound($_aData) - 1
						If $_aData[$i][1] = 'conhost.exe' Then ProcessClose($_aData[$i][0])
					Next
				EndIf
				ProcessClose($_iProcID)
			EndIf
		Until Not $_iProcID
		If $_WD_DEBUG_CONSOLE Then ConsoleWrite(@CRLF & ">> [_WD_ShutdownDriver] driver <" & $_sDriver & "> exited with " & $_WD_ERROR_COUNTER & " error(s) encountered" & @CRLF)
	EndFunc

	Func _WD_GetCurrentBrowserVersion($_sBrowserName = $_WD_CURRENT_BROWSER)
		If $_sBrowserName = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, False)

		Local Const $_reg_BrowserPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'

		Local $_sBrowserInstallPath = RegRead($_reg_BrowserPath & StringLower($_sBrowserName) & ".exe", "")
		Local $_sBrowserVersion = FileGetVersion($_sBrowserInstallPath, "FileVersion")

		If $_WD_DEBUG_CONSOLE Then ConsoleWrite(@CRLF & ">> [_WD_GetCurrentBrowserVersion]: " & $_sBrowserVersion & @CRLF & _
					@TAB & "BrowserInstallPath:" & @TAB & $_sBrowserInstallPath & @CRLF)

		Return SetError($_WD_ERROR_ID_Success, 0, $_sBrowserVersion)
	EndFunc

	Func _WD_GetCurrentDriverVersion($_sDriverExe = $_WD_DRIVER, $_sDriverInstallPath = $_WD_DRIVER_DIR)
		Local $_sDriverPath = $_sDriverInstallPath & $_sDriverExe
		Local $_iProcID = Run($_sDriverPath & " --version", $_sDriverInstallPath, @SW_HIDE, $STDOUT_CHILD)
			If NOT $_iProcID Then Return SetError($_WD_ERROR_ID_Exception, 1, "")
			If NOT ProcessWaitClose($_iProcID, 30) Then Return SetError($_WD_ERROR_ID_Exception, 2, "")

		Local $_sDriverVersion = StringRegExp(StdoutRead($_iProcID), "\s+([^\s]+)", 1)[0]
			If $_WD_DEBUG_CONSOLE Then ConsoleWrite(@CRLF & ">> [_WD_Ext_GetCurrentDriverVersion]: " & $_sDriverVersion & @CRLF & _
				@TAB & "DriverPath:" & @TAB & $_sDriverPath & @CRLF)

		Return SetError($_WD_ERROR_ID_Success, 0, $_sDriverVersion)
	EndFunc

	Func _WD_CheckUpdateRequired($_sCurrentBrowserVersion, $_sCurrentDriverVersion)
		$_sMajorBrowserVersion = StringLeft($_sCurrentBrowserVersion, StringInStr($_sCurrentBrowserVersion, ".")-1)
		$_sMajorDriverVersion = StringLeft($_sCurrentDriverVersion, StringInStr($_sCurrentDriverVersion, ".")-1)

		Local $_bUpdateRequire = ($_sMajorBrowserVersion <= $_sMajorDriverVersion) ? False : True
			If $_WD_DEBUG_CONSOLE Then ConsoleWrite(@CRLF & ">> [_WD_CheckUpdateRequired]: " & $_bUpdateRequire & @CRLF)

		Return $_bUpdateRequire
	EndFunc

	Func _WD_CheckLatestAvailableDriverVersion($_sCurrentBrowserVersion, $_sBrowserName = $_WD_CURRENT_BROWSER)
		Local $_sMajorBrowserVersion = StringLeft($_sCurrentBrowserVersion, StringInStr($_sCurrentBrowserVersion, ".")-1)
		Local $_sUrl = ""

		Switch $_sBrowserName
			Case "MSEDGE"
				$_sUrl = "https://msedgedriver.azureedge.net/LATEST_RELEASE_" & $_sMajorBrowserVersion & "_WINDOWS"

			Case "CHROME"
				$_sUrl = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & $_sMajorBrowserVersion

			Case Else
				Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")
		EndSwitch

		Local $_aTemp = StringRegExp(StringStripWS(BinaryToString(InetRead($_sUrl, 8+2), $_WD_BFORMAT), 8), "\d+.\d+.\d+.\d+", 3)
			If @error Then Return SetError($_WD_ERROR_ID_GeneralError, @error, False)
		Local $_sLatestVersion = $_aTemp[0]

			If $_WD_DEBUG_CONSOLE Then ConsoleWrite(@CRLF & ">> [_WD_GetLatestAvailableDriverVersion]: " & $_sLatestVersion & @CRLF & _
				@TAB & "BrowserName:" & @TAB & $_sBrowserName & @CRLF & _
				@TAB & "Url:" & @TAB & $_sUrl & @CRLF)

		Return SetError($_WD_ERROR_ID_Success, 0, $_sLatestVersion)
	EndFunc

	Func _WD_GetLatestDriverRelease($_sDriverVersion, $_sDriverExe = $_WD_DRIVER, $_sDriverInstallPath = $_WD_DRIVER_DIR)
		Local $_sDriverReleaseUrl = ""

		Switch $_sDriverExe
			Case $_WD_DRIVER_CHROME	;-- https://chromedriver.chromium.org/downloads/version-selection
				$_sDriverReleaseUrl = "https://chromedriver.storage.googleapis.com/" & $_sDriverVersion & "/chromedriver_win32.zip"

			Case $_WD_DRIVER_EDGE	;-- https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
			;	$_sDriverReleaseUrl = "https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/" & $_sDriverVersion & "/edgedriver_win32.zip"
				$_sDriverReleaseUrl = "https://msedgedriver.azureedge.net/" & $_sDriverVersion & "/edgedriver_win32.zip"

			Case Else
				Return SetError($_WD_ERROR_ID_InvalidRequest, 0, False)

		EndSwitch

		;-- download webdriver
		Local $_sTempFile = _TempFile($_sDriverInstallPath, "webdriver_v" & $_sDriverVersion & "_", ".zip")
		Local $_bSuccess = _WD_DownloadFile($_sDriverReleaseUrl, $_sTempFile)
			If $_WD_DEBUG_CONSOLE Then ConsoleWrite(@CRLF & ">> [_WD_GetLatestDriverRelease]: " & $_bSuccess & @CRLF & _
				@TAB & "DriverReleaseUrl:" & @TAB & $_sDriverReleaseUrl & @CRLF & _
				@TAB & "SavedFile:" & @TAB & $_sTempFile & @CRLF)

		Return SetError(@error, @extended, $_sTempFile)
	EndFunc

	Func _WD_DownloadFile($_sURL, $_sDestinationPath)
		Local $_binaryData = InetRead($_sURL)
			If @error Then Return SetError($_WD_ERROR_ID_SendReceiveError, 0, False)

		Local $hFile = FileOpen($_sDestinationPath, 18)
			If $hFile = -1 Then Return SetError($_WD_ERROR_ID_GeneralError, 0, False)
		$_iResult = FileWrite($hFile, $_binaryData)
			If $_iResult = 0 Then Return SetError($_WD_ERROR_ID_ActionIntercepted, 0, False)
		FileClose($hFile)

		Return SetError($_WD_ERROR_ID_Success, 0, True)
	EndFunc

	Func _WD_DriverInstallRelease($_sDriverReleaseZipPath, $_sDriverExe = $_WD_DRIVER, $_sDriverInstallPath = $_WD_DRIVER_DIR)
		If NOT FileExists($_sDriverReleaseZipPath) Then SetError($_WD_ERROR_ID_InvalidRequest, 0, False)

		;-- close and remove existing webdriver
		_WD_ShutdownDriver()
		Sleep($_WD_DELAY_TIME)
		FileDelete($_sDriverInstallPath & $_sDriverExe)
		DirRemove($_sDriverInstallPath & "Driver_Notes", 1)	;-- extra for msedgedriver

		; extract new downloaded webdriver
		Local $_oShell = ObjCreate("Shell.Application")
		Local $_ZipFiles = $_oShell.NameSpace($_sDriverReleaseZipPath).items
		$_oShell.NameSpace($_sDriverInstallPath).CopyHere($_ZipFiles, 20)
		Sleep($_WD_DELAY_TIME)

		;-- clean up
		FileDelete($_sDriverReleaseZipPath)

		$_WD_DRIVER_DIR = $_sDriverInstallPath
		$_WD_DRIVER = $_sDriverExe

			If $_WD_DEBUG_CONSOLE Then ConsoleWrite(@CRLF & ">> [_WD_DriverInstallRelease]: " & True & @CRLF & _
				@TAB & "InstallPath:" & @TAB & $_sDriverInstallPath & $_sDriverExe & @CRLF)

		Return SetError($_WD_ERROR_ID_Success, 0, True)
	EndFunc

#EndRegion BASE

#Region SESSIONS
	;-- https://www.w3.org/TR/webdriver/#sessions

	;-- https://www.w3.org/TR/webdriver/#new-session-0
	;-- https://www.w3.org/TR/webdriver/#capabilities
	Func _WD_NewSession()
		If $_WD_CAPABILITIES = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session", $_WD_CAPABILITIES)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, "")

		Local $_sSessionId = __WD_GetJSONValue($_sResponse, "[value][sessionId]")

		Return SetError(@error, @extended, $_sSessionId)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#delete-session
	Func _WD_DeleteSession($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_DELETE($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, "")
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#status
	Func _WD_Status()
		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/status")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, "")

		Local $_sStatus = __WD_GetJSONValue($_sResponse, "[value][ready]")

		Return SetError(@error, @extended, $_sStatus)
	EndFunc
#EndRegion SESSIONS

#Region TIMEOUTS
	;-- https://www.w3.org/TR/webdriver/#timeouts

	;-- https://www.w3.org/TR/webdriver/#set-timeouts
	Func _WD_SetTimeouts($_sSessionId, $_iPageLoadTimeout = 30000, $_iScriptTimeout = 30000, $_iImplicitTimeout = 0)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		;-- Implicit timeout:
		Local $_sJSON_Timeouts = '{"implicit":' & $_iImplicitTimeout & '}'
		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/timeouts", $_sJSON_Timeouts)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 1)

		;-- PageLoad timeout:
		$_sJSON_Timeouts = '{"pageLoad":' & $_iPageLoadTimeout & '}'
		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/timeouts", $_sJSON_Timeouts)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 2)

		;-- Script timeout:
		$_sJSON_Timeouts = '{"script":' & $_iScriptTimeout & '}'
		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/timeouts", $_sJSON_Timeouts)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 3)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-timeouts
	Func _WD_GetTimeouts($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")
		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/timeouts")

			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc
#EndRegion TIMEOUTS

#Region NAVIGATION
	;-- https://www.w3.org/TR/webdriver/#navigation

	;-- https://www.w3.org/TR/webdriver/#navigate-to
	Func _WD_NavigateTo($_sSessionId, $_sURL = 'about:blank')
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")
		Local $_sJSON_URL = '{"url":"' & $_sURL & '"}'

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/url", $_sJSON_URL)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-current-url
	Func _WD_GetCurrentURL($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/url")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sURL = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sURL)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-title
	Func _WD_GetTitle($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/title")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sTitle = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sTitle)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#refresh
	Func _WD_Refresh($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/refresh", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#back
	Func _WD_Back($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/back", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#forward
	Func _WD_Forward($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/forward", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc
#EndRegion NAVIGATION

#Region CONTEXTS
	;-- https://www.w3.org/TR/webdriver/#contexts

	;-- https://www.w3.org/TR/webdriver/#get-window-handle
	Func _WD_GetWindowHandle($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/window")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sHandle = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sHandle)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-window-handles
	Func _WD_GetWindowHandles($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/window/handles")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_aHandles = __WD_GetJSONValue($_sResponse, "[value]")

		If UBound($_aHandles) > 0 Then
			$_iError = $_WD_ERROR_ID_Success
		Else
			$_iError = $_WD_ERROR_ID_NotFound
		EndIf

		Return SetError(@error, @extended, $_aHandles)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#close-window
	Func _WD_CloseWindow($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_DELETE($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/window")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#maximize-window
	Func _WD_MaximizeWindow($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/window/maximize", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#minimize-window
	Func _WD_MinimizeWindow($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/window/minimize", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#fullscreen-window
	Func _WD_FullscreenWindow($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/window/fullscreen", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#switch-to-window
	Func _WD_SwitchToWindow($_sSessionId, $_sWindowHandle)
		If (($_sSessionId = "") OR ($_sWindowHandle = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sJSON_Window = '{"handle":"' & $_sWindowHandle & '"}'
		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/window", $_sJSON_Window)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#new-window
	Func _WD_NewWindow($_sSessionId, $_bOpenTab = True)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sType = $_bOpenTab ? 'tab' : 'window'
		Local $_sJSON_Window = '{"type":"' & $_sType & '"}'
		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/window/new", $_sJSON_Window)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sHandle = __WD_GetJSONValue($_sResponse, "[value][handle]")

		Return SetError(@error, @extended, $_sHandle)
	EndFunc

	;-- https://w3c.github.io/webdriver/#switch-to-frame
	Func _WD_SwitchToFrame($_sSessionId, $_sFrameIndexOrId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sJSON_Window = $_WD_JSON_EMPTY
		If IsInt($_sFrameIndexOrId) Then
			$_sJSON_Window = '{"id":"' & $_sFrameIndexOrId & '"}'
		Else
			$_sJSON_Window = '{"id":{"' & $_WD_ELEMENT_ID & '":"' & $_sFrameIndexOrId & '"}}'
		EndIf

			Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/frame", $_sJSON_Window)
				If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://w3c.github.io/webdriver/#switch-to-parent-frame
	Func _WD_SwitchToParentFrame($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sJSON_Window = $_WD_JSON_EMPTY
		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/frame/parent", $_sJSON_Window)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
		$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc
#EndRegion CONTEXTS

#Region ELEMENTS
	;-- https://www.w3.org/TR/webdriver/#elements
	;-- https://www.w3.org/TR/webdriver/#element-retrieval

	;-- https://www.w3.org/TR/webdriver/#find-element
	;-- https://www.w3.org/TR/webdriver/#find-element-from-element
	Func _WD_FindElement($_sSessionId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH, $_sStartNode = '')
		If (($_sSessionId = "") OR ($_sSelector = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		If $_sStartNode <> "" Then $_sStartNode = "/" & $_sStartNode & "/element"
		$_sSelector = StringRegExpReplace($_sSelector, "([" & $_WD_ESCAPE_CHARS & "])", "\\$1")
		Local $_sJSON_Element = '{"using":"' & $_sLocatorStrategy & '","value":"' & $_sSelector & '"}'

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element" & $_sStartNode, $_sJSON_Element)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sElementId = __WD_GetJSONValue($_sResponse, "[value][" & $_WD_ELEMENT_ID & "]")

		Return SetError(@error, @extended, $_sElementId)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#find-elements
	;-- https://www.w3.org/TR/webdriver/#find-elements-from-element
	;-- https://www.w3schools.com/xml/xpath_syntax.asp
	;-- http://www.whitebeam.org/library/guide/TechNotes/xpathtestbed.rhtm
	Func _WD_FindElements($_sSessionId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH, $_sStartNode = '')
		If (($_sSessionId = "") OR ($_sSelector = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		$_sStartNode = ($_sStartNode <> "") ? "/element/" & $_sStartNode & "/elements" : "/elements"
		Local $_sJSON_Element = '{"using":"' & $_sLocatorStrategy & '","value":"' & $_sSelector & '"}'

		$_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & $_sStartNode, $_sJSON_Element)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		;-- get elements array from JSON
		Local $_sJSON_Elements = __WD_GetJSONValue($_sResponse, "[value]")
			If @error Then Return SetError(@error, @extended, $_sJSON_Elements)

		Local $_aElements[0], $_oValue, $_iError
		Local $_iRow = 0
		;-- get element id values to output array
		If UBound($_sJSON_Elements) > 0 Then
			Local $_sKey = "[" & $_WD_ELEMENT_ID & "]"
			Dim $_aElements[UBound($_sJSON_Elements)]	;-- redeclare array with new dimensions
				For $_oValue In $_sJSON_Elements
					$_aElements[$_iRow] = Json_Get($_oValue, $_sKey)
					$_iRow += 1
				Next
			$_iError = $_WD_ERROR_ID_Success
		Else
			$_iError = $_WD_ERROR_ID_NotFound
		EndIf

		Return SetError($_iError, $_WD_HTTPRESULT, $_aElements)
	EndFunc

    ;-- https://w3c.github.io/webdriver/#find-element-from-shadow-root
	Func _WD_FindElementFromShadowRoot($_sSessionId, $_sShadowId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH)
		If (($_sSessionId = "") OR ($_sSelector = "") OR ($_sShadowId = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sJSON_Element = '{"using":"' & $_sLocatorStrategy & '","value":"' & $_sSelector & '"}'

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/shadow/" & $_sShadowId & "/element", $_sJSON_Element)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

        Local $_sElementId = __WD_GetJSONValue($_sResponse, "[value][" & $_WD_ELEMENT_ID & "]")

		Return SetError(@error, @extended, $_sElementId)
	EndFunc

    ;-- https://w3c.github.io/webdriver/#find-elements-from-shadow-root
	Func _WD_FindElementsFromShadowRoot($_sSessionId, $_sShadowId, $_sSelector, $_sLocatorStrategy = $_WD_LOCATOR_XPATH)
		If (($_sSessionId = "") OR ($_sSelector = "") OR ($_sShadowId = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sJSON_Element = '{"using":"' & $_sLocatorStrategy & '","value":"' & $_sSelector & '"}'

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/shadow/" & $_sShadowId & "/elements", $_sJSON_Element)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		;-- get elements array from JSON
		Local $_sJSON_Elements = __WD_GetJSONValue($_sResponse, "[value]")
			If @error Then Return SetError(@error, @extended, $_sJSON_Elements)

		Local $_aElements[0], $_oValue, $_iError
		Local $_iRow = 0
		;-- get element id values to output array
		If UBound($_sJSON_Elements) > 0 Then
			Local $_sKey = "[" & $_WD_ELEMENT_ID & "]"
			Dim $_aElements[UBound($_sJSON_Elements)]	;-- redeclare array with new dimensions
				For $_oValue In $_sJSON_Elements
					$_aElements[$_iRow] = Json_Get($_oValue, $_sKey)
					$_iRow += 1
				Next
			$_iError = $_WD_ERROR_ID_Success
		Else
			$_iError = $_WD_ERROR_ID_NotFound
		EndIf

		Return SetError($_iError, $_WD_HTTPRESULT, $_aElements)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-active-element
	Func _WD_GetActiveElement($_sSessionId)
		If $_sSessionId = "" Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/active")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sElementId = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sElementId)
	EndFunc

	;-- https://w3c.github.io/webdriver/#get-element-shadow-root
	Func _WD_GetElementShadowRoot($_sSessionId, $_sElementId)
		If (($_sSessionId = "") OR ($_sElementId = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/" & $_sElementId & "/shadow")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sShadowId = __WD_GetJSONValue($_sResponse, "[value][" & $_WD_SHADOW_ID & "]")

		Return SetError(@error, @extended, $_sShadowId)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-element-attribute
	Func _WD_GetElementAttribute($_sSessionId, $_sElementId, $_sAttributeName)
		If (($_sSessionId = "") OR ($_sElementId = "") OR ($_sAttributeName = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/" & $_sElementId & "/attribute/" & $_sAttributeName)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sAttribute = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sAttribute)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-element-property
	Func _WD_GetElementProperty($_sSessionId, $_sElementId, $_sPropertyName)
		If (($_sSessionId = "") OR ($_sElementId = "") OR ($_sPropertyName = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/" & $_sElementId & "/property/" & $_sPropertyName)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sProperty = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sProperty)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-element-text
	Func _WD_GetElementText($_sSessionId, $_sElementId)
		If (($_sSessionId = "") OR ($_sElementId = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/" & $_sElementId & "/text")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sText = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sText)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-element-tag-name
	Func _WD_GetElementTagName($_sSessionId, $_sElementId)
		If (($_sSessionId = "") OR ($_sElementId = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/" & $_sElementId & "/name")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sTagName = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sTagName)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-element-rect
	Func _WD_GetElementRect($_sSessionId, $_sElementId)
		If (($_sSessionId = "") OR ($_sElementId = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/" & $_sElementId & "/rect")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sJSON_Rectangle = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sJSON_Rectangle)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#element-click
	Func _WD_ElementClick($_sSessionId, $_sElementId)
		If (($_sSessionId = "") OR ($_sElementId = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/" & $_sElementId & "/click", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#element-clear
	Func _WD_ElementClear($_sSessionId, $_sElementId)
		If (($_sSessionId = "") OR ($_sElementId = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/" & $_sElementId & "/clear", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#element-send-keys
	;-- https://www.w3.org/TR/webdriver/#keyboard-actions
	Func _WD_ElementSendKeys($_sSessionId, $_sElementId, $_sKeys)
		If (($_sSessionId = "") OR ($_sElementId = "") OR ($_sKeys = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		$_sKeys = StringRegExpReplace($_sKeys, "([" & $_WD_ESCAPE_CHARS & "])", "\\$1")
		Local $_sJSON_Keys = '{"id":"' & $_sElementId & '", "text":"' & $_sKeys & '"}'

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/element/" & $_sElementId & "/value", $_sJSON_Keys)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc
#EndRegion ELEMENTS

#Region DOCUMENT
	;-- https://www.w3.org/TR/webdriver/#document

	;-- https://www.w3.org/TR/webdriver/#get-page-source
	Func _WD_GetPageSource($_sSessionId)
		If ($_sSessionId = "") Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/source")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sSource = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sSource)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#executing-script
	;-- https://www.w3.org/TR/webdriver/#execute-script
	;-- https://www.w3.org/TR/webdriver/#execute-async-script
	Func _WD_ExecuteScript($_sSessionId, $_sScript, $_sJSON_Arguments = "[]", $_bAsync = False)
		If (($_sSessionId = "") OR ($_sScript = "") OR ($_sJSON_Arguments = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sSync = ($_bAsync) ? 'async' : 'sync'
		$_sScript = StringRegExpReplace($_sScript, "([" & $_WD_ESCAPE_CHARS & "])", "\\$1")
		Local $_sJSON_Script = '{"script":"' & $_sScript & '", "args":[' & $_sJSON_Arguments & ']}'

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/execute/" & $_sSync, $_sJSON_Script)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc
#EndRegion DOCUMENT

#Region COOKIES
	;-- https://www.w3.org/TR/webdriver/#cookies

	;-- https://www.w3.org/TR/webdriver/#get-all-cookies
	Func _WD_GetAllCookies($_sSessionId)
		If ($_sSessionId = "") Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/cookie")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sJSON_Cookies = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sJSON_Cookies)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-named-cookie
	;-- https://www.w3.org/TR/webdriver/#dfn-cookie-name
	Func _WD_GetNamedCookie($_sSessionId, $_sCookieName)
		If (($_sSessionId = "") OR ($_sCookieName = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/cookie/" & $_sCookieName)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sJSON_Cookie = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sJSON_Cookie)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#add-cookie
	Func _WD_AddCookie($_sSessionId, $_sCookieName, $_sCookieValue)
		If (($_sSessionId = "") OR ($_sCookieName = "") OR ($_sCookieValue = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sJSON_Cookie = '{"cookie": {"name":"' & $_sCookieName & '","value":"' & $_sCookieValue & '"}}'

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/cookie/", $_sJSON_Cookie)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#delete-cookie
	Func _WD_DeleteCookie($_sSessionId, $_sCookieName)
		If (($_sSessionId = "") OR ($_sCookieName = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_DELETE($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/cookie/" & $_sCookieName)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#delete-all-cookies
	Func _WD_DeleteAllCookies($_sSessionId)
		If ($_sSessionId = "") Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_DELETE($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/cookie/")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc
#EndRegion COOKIES

#Region ACTIONS
	;-- https://www.w3.org/TR/webdriver/#actions

	;-- https://www.w3.org/TR/webdriver/#perform-actions
	Func _WD_PerformActions($_sSessionId, $_sJSON_Actions)
		If (($_sSessionId = "") OR ($_sJSON_Actions = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/actions", $_sJSON_Actions)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#release-actions
	Func _WD_ReleaseActions($_sSessionId)
		If ($_sSessionId = "") Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_DELETE($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/actions")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc
#EndRegion ACTIONS

#Region USER_PROMPTS
	;-- https://www.w3.org/TR/webdriver/#user-prompts

	;-- https://www.w3.org/TR/webdriver/#dfn-accept-alert
	Func _WD_AcceptAlert($_sSessionId)
		If ($_sSessionId = "") Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/alert/accept", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#get-alert-text
	Func _WD_GetAlertText($_sSessionId)
		If ($_sSessionId = "") Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/alert/text")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		Local $_sJSON_AlertText = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sJSON_AlertText)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#send-alert-text
	Func _WD_SendAlertText($_sSessionId, $_sAlertText)
		If (($_sSessionId = "") OR ($_sAlertText = "")) Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sJSON_AlertText = '{"text":"' & $_sAlertText & '"}'
		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/alert/text", $_sJSON_AlertText)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc

	;-- https://www.w3.org/TR/webdriver/#dismiss-alert
	Func _WD_DismissAlertText($_sSessionId)
		If ($_sSessionId = "") Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		Local $_sResponse = __WD_POST($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & "/alert/dismiss", $_WD_JSON_EMPTY)
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)
			$_sResponse = __WD_GetJSONValue($_sResponse)

		Return SetError(@error, @extended, $_sResponse)
	EndFunc
#EndRegion USER_PROMPTS

#Region SCREEN_CAPTURE
	;-- https://www.w3.org/TR/webdriver/#screen-capture

	;-- https://www.w3.org/TR/webdriver/#take-screenshot
	;-- https://www.w3.org/TR/webdriver/#take-element-screenshot
	Func _WD_TakeScreenshot($_sSessionId, $_sElementId = '')
		If ($_sSessionId = "") Then Return SetError($_WD_ERROR_ID_InvalidRequest, 0, "")

		If $_sElementId <> "" Then $_sElementId = "/element/" & $_sElementId
		Local $_sResponse = __WD_GET($_WD_BASE_URL & ":" & $_WD_PORT & "/session/" & $_sSessionId & $_sElementId & "/screenshot")
			If @error Then Return SetError(@error, $_WD_HTTPRESULT, 0)

		$_sJSON_Capture = __WD_GetJSONValue($_sResponse, "[value]")

		Return SetError(@error, @extended, $_sJSON_Capture)
	EndFunc
#EndRegion SCREEN_CAPTURE

#Region INTERNAL_FUNCTIONS

	Func __WD_GET($_sURL)
		Local $_iError = 0, $_sResponseText
		$_WD_HTTPRESULT = 0

		Local $_aURL = _WinHttpCrackUrl($_sURL)
		If IsArray($_aURL) Then
			;~ _ArrayDisplay($_aURL)
			Local $hOpen = _WinHttpOpen()   ;-- initialize and get session handle
			Local $hConnect = _WinHttpConnect($hOpen, $_aURL[2], $_aURL[3])   ;-- get connection handle
			If @error Then
				$_iError = $_WD_ERROR_ID_SocketError
				$_WD_ERROR_COUNTER += 1
			Else
				Switch $_aURL[1]
					Case $INTERNET_SCHEME_HTTP
						$_sResponseText = _WinHttpSimpleRequest($hConnect, "GET", $_aURL[6] & $_aURL[7])
					Case $INTERNET_SCHEME_HTTPS
						$_sResponseText = _WinHttpSimpleSSLRequest($hConnect, "GET", $_aURL[6] & $_aURL[7])
					Case Else
						SetError($_WD_ERROR_ID_InvalidRequest)
						$_WD_ERROR_COUNTER += 1
				EndSwitch

				$_iError = @error
				$_WD_HTTPRESULT = @extended
				If $_iError Then
					$_iError = $_WD_ERROR_ID_SendReceiveError
					$_WD_ERROR_COUNTER += 1
				ElseIf $_WD_HTTPRESULT =  $HTTP_STATUS_SERVER_ERROR Or $_WD_HTTPRESULT = $HTTP_STATUS_REQUEST_TIMEOUT Then
					$_iError = $_WD_ERROR_ID_Timeout
					$_WD_ERROR_COUNTER += 1
				EndIf
			EndIf

			_WinHttpCloseHandle($hConnect)
			_WinHttpCloseHandle($hOpen)
		Else
			$_iError = $_WD_ERROR_ID_InvalidRequest
			$_WD_ERROR_COUNTER += 1
		EndIf

		$_WD_SESSION_DETAILS = $_sResponseText
		If $_WD_DEBUG_CONSOLE Then
			ConsoleWrite(@CRLF & ">> [__WD_GET](" & $_sURL & "): " & @TAB & @CRLF & _
					">> Response: " & $_sResponseText & @CRLF & _
					">> Error code: " & $_iError & @TAB & " | $_WD_HTTPRESULT: " & $_WD_HTTPRESULT & @CRLF)
		EndIf
		Return SetError($_iError, $_WD_HTTPRESULT, $_sResponseText)
	EndFunc

	Func __WD_POST($_sURL, $_sData)
		Local $_iError = 0, $_sResponseText
		$_WD_HTTPRESULT = 0

		Local $_aURL = _WinHttpCrackUrl($_sURL)
		If IsArray($_aURL) Then
			;_ArrayDisplay($_aURL)
			Local $hOpen = _WinHttpOpen()	 ;-- initialize and get session handle
			Local $hConnect = _WinHttpConnect($hOpen, $_aURL[2], $_aURL[3])	;-- get connection handle
				If @error Then
					$_iError = $_WD_ERROR_ID_SocketError
					$_WD_ERROR_COUNTER += 1
				Else
					Switch $_aURL[1]
						Case $INTERNET_SCHEME_HTTP
							$_sResponseText = _WinHttpSimpleRequest($hConnect, "POST", $_aURL[6] & $_aURL[7], Default, StringToBinary($_sData, $_WD_BFORMAT))
						Case $INTERNET_SCHEME_HTTPS
							$_sResponseText = _WinHttpSimpleSSLRequest($hConnect, "POST", $_aURL[6] & $_aURL[7], Default, StringToBinary($_sData, $_WD_BFORMAT))
						Case Else
							SetError($_WD_ERROR_ID_InvalidRequest)
							$_WD_ERROR_COUNTER += 1
					EndSwitch

					$_iError = @error
					$_WD_HTTPRESULT = @extended
					If $_iError Then
						$_iError = $_WD_ERROR_ID_SendReceiveError
						$_WD_ERROR_COUNTER += 1
					ElseIf $_WD_HTTPRESULT =  $HTTP_STATUS_SERVER_ERROR Or $_WD_HTTPRESULT = $HTTP_STATUS_REQUEST_TIMEOUT Then
						$_iError = $_WD_ERROR_ID_Timeout
						$_WD_ERROR_COUNTER += 1
					EndIf
				EndIf

			_WinHttpCloseHandle($hConnect)
			_WinHttpCloseHandle($hOpen)
		Else
			$_iError = $_WD_ERROR_ID_InvalidRequest
			$_WD_ERROR_COUNTER += 1
		EndIf

		$_WD_SESSION_DETAILS = $_sResponseText
		If $_WD_DEBUG_CONSOLE Then
			ConsoleWrite(@CRLF & ">> [__WD_POST](" & $_sURL & "): " & @TAB & @CRLF & _
					">> Request: " & $_sData & @CRLF & _
					">> Response: " & $_sResponseText & @CRLF & _
					">> Error code: " & $_iError & @TAB & " | $_WD_HTTPRESULT: " & $_WD_HTTPRESULT & @CRLF)
		EndIf
		Return SetError($_iError, $_WD_HTTPRESULT, $_sResponseText)
	EndFunc

	Func __WD_DELETE($_sURL)
		Local $_iError = 0, $_sResponseText
		$_WD_HTTPRESULT = 0

		Local $_aURL = _WinHttpCrackUrl($_sURL)
		If IsArray($_aURL) Then
			;_ArrayDisplay($_aURL)
			Local $hOpen = _WinHttpOpen()	 ;-- initialize and get session handle
			Local $hConnect = _WinHttpConnect($hOpen, $_aURL[2], $_aURL[3])	;-- get connection handle
			If @error Then
				$_iError = $_WD_ERROR_ID_SocketError
				$_WD_ERROR_COUNTER += 1
			Else
				Switch $_aURL[1]
					Case $INTERNET_SCHEME_HTTP
						$_sResponseText = _WinHttpSimpleRequest($hConnect, "DELETE", $_aURL[6] & $_aURL[7])
					Case $INTERNET_SCHEME_HTTPS
						$_sResponseText = _WinHttpSimpleSSLRequest($hConnect, "DELETE", $_aURL[6] & $_aURL[7])
					Case Else
						SetError($_WD_ERROR_ID_InvalidRequest)
						$_WD_ERROR_COUNTER += 1
				EndSwitch

				$_iError = @error
				$_WD_HTTPRESULT = @extended
				If $_iError Then
					$_iError = $_WD_ERROR_ID_SendReceiveError
					$_WD_ERROR_COUNTER += 1
				ElseIf $_WD_HTTPRESULT =  $HTTP_STATUS_SERVER_ERROR Or $_WD_HTTPRESULT = $HTTP_STATUS_REQUEST_TIMEOUT Then
					$_iError = $_WD_ERROR_ID_Timeout
					$_WD_ERROR_COUNTER += 1
				EndIf
			EndIf

			_WinHttpCloseHandle($hConnect)
			_WinHttpCloseHandle($hOpen)
		Else
			$_iError = $_WD_ERROR_ID_SocketError
			$_WD_ERROR_COUNTER += 1
		EndIf

		$_WD_SESSION_DETAILS = $_sResponseText
		If $_WD_DEBUG_CONSOLE Then
			ConsoleWrite(@CRLF & ">> [__WD_DELETE](" & $_sURL & "): " & @TAB & @CRLF & _
						">> Response: " & $_sResponseText & @CRLF & _
						">> Error code: " & $_iError & @TAB & " | $_WD_HTTPRESULT: " & $_WD_HTTPRESULT & @CRLF)
		EndIf
		Return SetError($_iError, $_WD_HTTPRESULT, $_sResponseText)
	EndFunc

	Func __WD_GetJSONValue($_sResponse, $_sReturnValue = '')
		Local $_oJson = Json_Decode($_sResponse)
		Local $_iErrorCode

		If $_WD_HTTPRESULT = $HTTP_STATUS_OK Then
			If $_sReturnValue <> '' Then
				$_sReturnValue = Json_Get($_oJson, $_sReturnValue)
			Else
				$_sReturnValue = $_sResponse
			EndIf
			$_iErrorCode = $_WD_ERROR_ID_Success

		Else
			$_WD_ERROR_COUNTER += 1
			Local $_sError = Json_Get($_oJson, "[value][error]")
			Local $_sMessage = Json_Get($_oJson, "[value][message]")
			$_sReturnValue = $_sError
			Switch $_sError
				Case $WD_ERROR_InsecureCertificate, $WD_ERROR_InvalidArgument, $WD_ERROR_InvalidCookieDomain, $WD_ERROR_InvalidSelector, $WD_ERROR_InvalidSessionId, $WD_ERROR_UnsupportedOperation
					 $_iErrorCode = $_WD_ERROR_ID_InvalidRequest

				Case $WD_ERROR_AlertNotFound, $WD_ERROR_CookieNotFound, $WD_ERROR_ElementNotFound, $WD_ERROR_FrameNotFound, $WD_ERROR_WindowNotFound, $WD_ERROR_StaleElementRef
					$_iErrorCode = $_WD_ERROR_ID_NotFound

				Case $WD_ERROR_ScriptTimeout, $WD_ERROR_Timeout
					$_iErrorCode = $_WD_ERROR_ID_Timeout

				Case $WD_ERROR_ElementClickIntercept, $WD_ERROR_UnableSetCookie, $WD_ERROR_UnableCaptureScreen, $WD_ERROR_InvalidElementState, $WD_ERROR_UnexpectedAlertOpen, $WD_ERROR_TargetOutOfBounds
					$_iErrorCode = $_WD_ERROR_ID_ActionIntercepted

				Case $WD_ERROR_UnknownCommand, $WD_ERROR_UnknownError, $WD_ERROR_UnknownMethod
					$_iErrorCode = $_WD_ERROR_ID_Unknown

				Case Else
					$_iErrorCode = $_WD_ERROR_ID_Exception
			EndSwitch
			If $_WD_DEBUG_CONSOLE Then ConsoleWrite(">> JSON Error: " & $_sError & @TAB & "| Message: " & $_sMessage & @TAB & "| Code: " & $_iErrorCode & @CRLF)
		EndIf

		Return SetError($_iErrorCode, $_WD_HTTPRESULT, $_sReturnValue)
	EndFunc

#EndRegion INTERNAL_FUNCTIONS