' Launch Google Chrome
BrowserExecutable =  "Chrome.exe"
SystemUtil.Run BrowserExecutable,"","","",3

Set AppContext=Browser("CreationTime:=0")

AppContext.Navigate "https://sapui5.hana.ondemand.com/test-resources/sap/m/demokit/cart/webapp/index.html"
AppContext.Maximize
AIUtil.SetContext AppContext

AIUtil("search").Search Parameter("searchValue")

AIUtil.FindTextBlock("No products found").CheckExists True

' Close browser
AppContext.Close
