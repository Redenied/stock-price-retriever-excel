Sub StockRetrieve2()

'Declare variables
Dim bot As New WebDriver
'Dim stockPrice As Double
Dim stockPrice As WebElement
Dim i As Integer
Dim link As String

'Open chrome in background
bot.AddArgument "--headless"

'Extend response time tolerance
bot.SetPreference "pageLoadStrategy", "normal"
bot.Timeouts.PageLoad = 100000

'Init new Chrome instance
bot.Start "chrome"

'Init variables
i = 0
'stockPrice = 0
link = ""

'For loop to get each stock price
For i = 4 To 23
    'Skip to the next iteration if an error occurs
    On Error Resume Next
    
    'Link to stock
    link = Sheets("Ventas").Range("T" & i).Value

    'Open up a new tab in chrome (javascript)
    bot.ExecuteScript "window.open(arguments[0])", link

    'Switch to new tab
    bot.SwitchToNextWindow

    'Store stock price (same position for all stocks)
    Set stockPrice = bot.FindElementByXPath("//*[@id='quote-header-info']/div[3]/div[1]/div/fin-streamer[1]")

    'Write stock price to worksheet
    Sheets("Ventas").Range("L" & i).Value = stockPrice.Text
    
    'Close current tab and switch back to previous tab
    bot.ExecuteScript "window.close()"
    bot.SwitchToPreviousWindow
    
Next i

bot.Quit ' Close the browser window

MsgBox "Stock prices updated successfully uwu"

End Sub

