Sub StockRetrieve2()
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
    For i = 4 To 23
        ' Skip to the next iteration if an error occurs
        On Error Resume Next

        ' Get the link to the stock
        link = Sheets("Ventas").Range("T" & i).Value

        ' Open a new tab in Chrome using JavaScript
        bot.ExecuteScript "window.open(arguments[0])", link

        ' Switch to the new tab
        bot.SwitchToNextWindow

        ' Find the stock price element (same position for all stocks)
        Set stockPrice = bot.FindElementByXPath("//*[@id='quote-header-info']/div[3]/div[1]/div/fin-streamer[1]")

        ' Write the stock price to the worksheet
        Sheets("Ventas").Range("L" & i).Value = stockPrice.Text

        ' Close the current tab and switch back to the previous tab
        bot.ExecuteScript "window.close()"
        bot.SwitchToPreviousWindow
    Next i

    ' Quit the browser
    bot.Quit

    MsgBox "Stock prices updated successfully!"

End Sub
