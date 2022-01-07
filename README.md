# W3C WebDriver UDF

## Table of Contents
+ [About](#about)
+ [Prerequisites](#prerequisites)
+ [Installation](#installation)
+ [Usage](#usage)
+ [Documentation](#documentation)

## About <a name = "about"></a>
This UDF is an AutoIt wrapper for 'WebDriver' automation.   
It allows to interact with modern browsers like ```Google Chrome``` or ```MS Edge``` via web requests channel provided by Webdriver.  
```Note, that not all functionality might be mapped or updated in current published version.```

## Prerequisites <a name = "prerequisites"></a>
Supported browser version must be installed.

## Installation <a name = "installation"></a>
Simply copy ```.au3``` files to your development directory and use ```#include``` to point to these files in the source code.  
UDF requires ```WinHttp.au3``` as dependency, which is included as well.

## Usage <a name = "usage"></a>
* Use ```WebDriver_*.au3``` for core functionality to Chrome or Edge browsers.   
* UDF uses Webdriver .exe dedicated to certain browser to startup and engage with it. ```_WD_Configure``` and ```_WD_StartupDriver``` functions responsible for intialization.
* Use ```WebDriver_Core.au3``` to download and updated required Webdriver .exe file

## Documentation <a name = "documentation"></a>
* [W3C WebDriver documentation](https://w3c.github.io/webdriver/)
* [W3C WebDriver Endpoints](https://www.w3.org/TR/webdriver/#endpoints)
* [ChromeDriver](https://chromedriver.chromium.org/)
* [EdgeDriver](https://docs.microsoft.com/en-us/microsoft-edge/webdriver)
* [GeckoDriver](https://developer.mozilla.org/en-US/docs/Web/WebDriver)
* [MSDN WinHttp functions](https://docs.microsoft.com/en-us/windows/win32/winhttp/winhttp-functions)
* [AutoIt JSON UDF](https://www.autoitscript.com/forum/topic/148114-a-non-strict-json-udf-jsmn/)