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
    i = 2
    link = ""

    ' While loop to get each stock price
    Do While Not IsEmpty(Sheets("Portfolio").Range("K" & i).Value)
        ' Skip to the next iteration if an error occurs
        On Error Resume Next
        
        ' Get link to stock
        link = Sheets("Portfolio").Range("K" & i).Value
    
        ' Open a new tab in chrome (javascript)
        bot.ExecuteScript "window.open(arguments[0])", link
    
        ' Switch to new tab
        bot.SwitchToNextWindow
    
        ' Store stock price (same position for all stocks)
        Set stockPrice = bot.FindElementByXPath("//*[@id='quote-header-info']/div[3]/div[1]/div/fin-streamer[1]")
    
        ' Write stock price to worksheet
        Sheets("Portfolio").Range("G" & i).Value = stockPrice.Text
        
        ' Close current tab and switch back to previous tab
        bot.ExecuteScript "window.close()"
        bot.SwitchToPreviousWindow
        
        i = i + 1
    Loop

' Quit the browser
bot.Quit

MsgBox "Stock prices updated successfully!"

End Sub

