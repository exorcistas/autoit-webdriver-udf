#include-once

#Region GENERAL_GLOBAL_VARS
	Global Const $_WD_DEFAULT_BROWSER = "msedge"
	Global $_WD_CURRENT_BROWSER = ""
	Global $_WD_DRIVER = "" ; Path to web driver executable
	Global $_WD_DRIVER_DIR = @ScriptDir & "\"
	Global $_WD_DRIVER_PARAMS = ""  ; Parameters to pass to web driver executable
	Global $_WD_BASE_URL = "http://127.0.0.1"
	Global $_WD_PORT = 0    ; Port used for web driver communication
	;Global Const $_WD_PROXY = "https://someproxy.com:8080"	; replace with relevant proxy uri/port
	Global $_WD_CAPABILITIES = ""    ; Port used for web driver communication
	Global $_WD_HTTPRESULT  ; Last response from WinHTTP request
	Global $_WD_SESSION_DETAILS = ""    ; Last response from _WD_CreateSession
	Global $_WD_BFORMAT = 4  ; Binary format (UTF8 == 4) // see <StringConstants.au3>
	Global $_WD_ESCAPE_CHARS = '\\"'    ; Characters to escape
	Global Const $_WD_ELEMENT_ID = "element-6066-11e4-a52e-4f735466cecf"    ; web element identifier CLSID
	Global Const $_WD_SHADOW_ID = "shadow-6066-11e4-a52e-4f735466cecf"
	Global Const $_WD_JSON_EMPTY  = "{}"
	Global $_WD_DEBUG_CONSOLE = TRUE
	Global $_WD_DELAY_TIME = 3000
#EndRegion GENERAL_GLOBAL_VARS

#Region CHROME_DRIVER
	;-- https://chromedriver.chromium.org/
	Global Const $_WD_DRIVER_CHROME = 'chromedriver.exe'
	Global Const $_WD_PORT_CHROME = 9515
	Global Const $_WD_CAPABILITIES_JSON_CHROME = '{"capabilities":{"alwaysMatch":{"goog:chromeOptions":{"w3c":true,"args":["--start-maximized","--disable-infobars","--disable-notifications","--disable-web-security","--disable-default-apps","--force-renderer-accessibility","--enable-features=DataReductionProxyDecidesTransform,NetworkServiceInProcess","--enable-experimental-ui-automation"]},"pageLoadStrategy":"normal","timeouts":{"pageLoad":600000, "implicit":0, "script":30000},"proxy": {"proxyType":"manual","autodetect": "true"}}}}'
	Global Const $_WD_DRIVER_PARAMS_CHROME = '--verbose --log-path="' & @ScriptDir & '\chromedriver-log.log"'
#EndRegion CHROME_DRIVER

#Region EDGE_DRIVER
	;-- https://docs.microsoft.com/en-us/microsoft-edge/webdriver
	Global Const $_WD_DRIVER_EDGE = 'msedgedriver.exe'
	Global Const $_WD_PORT_EDGE = 9515
	Global Const $_WD_CAPABILITIES_JSON_EDGE = '{"capabilities":{"alwaysMatch":{"browserName":"msedge","ms:edgeChromium": true,"ms:edgeOptions": {"w3c":true,"args":["--start-maximized","--disable-infobars","--disable-notifications","--disable-default-apps","--disable-web-security","--force-renderer-accessibility","--enable-features=DataReductionProxyDecidesTransform,NetworkServiceInProcess","--enable-experimental-ui-automation"]},"pageLoadStrategy":"normal","timeouts":{"pageLoad":600000, "implicit":0, "script":30000},"proxy": {"proxyType":"manual","autodetect": "true"}}}}'
	Global Const $_WD_DRIVER_PARAMS_EDGE = '--verbose --log-path="' & @ScriptDir & '\edgedriver-log.log"'
#EndRegion EDGE_DRIVER

#Region FIREFOX_DRIVER
;-- https://developer.mozilla.org/en-US/docs/Web/WebDriver
	Global Const $_WD_DRIVER_FIREFOX = 'geckodriver.exe'
	Global Const $_WD_PORT_FIREFOX = 4444
	Global Const $_WD_CAPABILITIES_JSON_FIREFOX = '{"desiredCapabilities":{"javascriptEnabled":true,"nativeEvents":true,"acceptInsecureCerts":true}}'
	Global Const $_WD_DRIVER_PARAMS_FIREFOX = '--log trace'
#EndRegion FIREFOX_DRIVER

#Region LOCATOR_STRATEGIES
    ;-- https://www.w3.org/TR/webdriver/#dfn-strategy
	Global Const $_WD_LOCATOR_CSS = "css selector"
	Global Const $_WD_LOCATOR_LINKTEXT = "link text"
	Global Const $_WD_LOCATOR_PARTIAL_LINKTEXT = "partial link text"
	Global Const $_WD_LOCATOR_TAG_NAME = "tag name"
	Global Const $_WD_LOCATOR_XPATH = "xpath"
#EndRegion LOCATOR_STRATEGIES

#Region ERROR_DEFINITIONS
	Global Enum _
		$_WD_ERROR_ID_Success = 0, _ ; No error
		$_WD_ERROR_ID_GeneralError, _ ; General error
		$_WD_ERROR_ID_SocketError, _ ; No socket
		$_WD_ERROR_ID_SendReceiveError, _ ; Send / Receive HTTP request error
		$_WD_ERROR_ID_Timeout, _ ;
		$_WD_ERROR_ID_InvalidRequest, _ ; Invalid input values provided
		$_WD_ERROR_ID_ActionIntercepted, _ ;
		$_WD_ERROR_ID_Unknown, _ ; Unknown command, method or error
		$_WD_ERROR_ID_Exception, _ ; Exception from web driver
		$_WD_ERROR_ID_NotFound

	;-- https://www.w3.org/TR/webdriver/#errors
	Global Const $WD_ERROR_ElementClickIntercept = "element click intercepted"
	Global Const $WD_ERROR_ElementNotInteractable = "element not interactable"
	Global Const $WD_ERROR_InsecureCertificate = "insecure certificate"
	Global Const $WD_ERROR_InvalidArgument = "invalid argument"
	Global Const $WD_ERROR_InvalidCookieDomain = "invalid cookie domain"
	Global Const $WD_ERROR_InvalidElementState = "invalid element state"
	Global Const $WD_ERROR_InvalidSelector = "invalid selector"
	Global Const $WD_ERROR_InvalidSessionId = "invalid session id"
	Global Const $WD_ERROR_JavascriptError = "javascript error"
	Global Const $WD_ERROR_TargetOutOfBounds = "move target out of bounds"
	Global Const $WD_ERROR_AlertNotFound = "no such alert"
	Global Const $WD_ERROR_CookieNotFound = "no such cookie"
	Global Const $WD_ERROR_ElementNotFound = "no such element"
	Global Const $WD_ERROR_FrameNotFound = "no such frame"
	Global Const $WD_ERROR_WindowNotFound = "no such window"
	Global Const $WD_ERROR_ScriptTimeout = "script timeout"
	Global Const $WD_ERROR_SessionNotCreated = "session not created"
	Global Const $WD_ERROR_StaleElementRef = "stale element reference"
	Global Const $WD_ERROR_Timeout = "timeout"
	Global Const $WD_ERROR_UnableSetCookie = "unable to set cookie"
	Global Const $WD_ERROR_UnableCaptureScreen = "unable to capture screen"
	Global Const $WD_ERROR_UnexpectedAlertOpen = "unexpected alert open"
	Global Const $WD_ERROR_UnknownCommand = "unknown command"
	Global Const $WD_ERROR_UnknownError = "unknown error"
	Global Const $WD_ERROR_UnknownMethod = "unknown method"
	Global Const $WD_ERROR_UnsupportedOperation = "unsupported operation"

	Global $_WD_ERROR_COUNTER = 0
#EndRegion ERROR_DEFINITIONS