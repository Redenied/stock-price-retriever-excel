Attribute VB_Name = "Module1"
Sub StockRetrieve()
    ' Declare variables
    Dim bot As New WebDriver
    Dim stockPrice As WebElement
    Dim i As Integer
    Dim link As String

    ' Open Chrome in headless mode (in the background)
    bot.AddArgument "--headless"

    ' Extend response time tolerance
    bot.SetPreference "pageLoadStrategy", "normal"
    bot.Timeouts.PageLoad = 100000

    ' Start a new Chrome instance
    bot.Start "chrome"

    ' Initialize variables
    i = 0
    link = ""

    ' For loop to get each stock price
    For i = 3 To 36
        ' Skip to the next iteration if an error occurs
        On Error Resume Next
        
        ' Get link to stock
        link = Sheets("Portfolio").Range("M" & i).Value
    
        ' Open a new tab in chrome (javascript)
        bot.ExecuteScript "window.open(arguments[0])", link
    
        ' Switch to new tab
        bot.SwitchToNextWindow
    
        ' Store stock price (same position for all stocks)
        Set stockPrice = bot.FindElementByXPath("//*[@id='quote-header-info']/div[3]/div[1]/div/fin-streamer[1]")
    
        ' Write stock price to worksheet
        Sheets("Portfolio").Range("O" & i).Value = stockPrice.Text
        
        ' Close current tab and switch back to previous tab
        bot.ExecuteScript "window.close()"
        bot.SwitchToPreviousWindow
    Next i

' Quit the browser
bot.Quit

MsgBox "Stock prices updated successfully!"

End Sub
