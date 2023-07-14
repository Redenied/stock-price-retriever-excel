Attribute VB_Name = "Módulo2"
Sub DayBeforeValues()
    'Copies cash value for day before
    Range("D32").Copy
    Range("D33").PasteSpecial xlPasteValues

    'Copies stocks ytd yield for day before
    Range("I29").Copy
    Range("I30").PasteSpecial xlPasteValues

    'Copies total portfolio value for day before
    Range("H32").Copy
    Range("H33").PasteSpecial xlPasteValues

    'Deselect las used cell
    Application.SendKeys "{ESC}"
End Sub
