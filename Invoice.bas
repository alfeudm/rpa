Attribute VB_Name = "Módulo1"
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10


Sub GetPDF()

Application.Wait Now + TimeValue("00:00:05")

 Dim objShell
        
        Set objShell = CreateObject("shell.application")
            objShell.ToggleDesktop
        Set objShell = Nothing
        
'Open PDF Files

Dim strPDF_File_Name As String
strPDF_File_Name = "Invoice1.pdf"
ActiveWorkbook.FollowHyperlink strPDF_File_Name

Application.Wait Now + TimeValue("00:00:05")

SetCursorPos 700, 400 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0



Call GetInvoice


Range("A1") = "Invoice 1"
Range("m1").Select
Range("m1").PasteSpecial
Range("m7").Select
Range("m7").Copy
Range("A2").Select
Range("A2").PasteSpecial


Range("m8").Select
Range("m8").Copy
Range("A3").Select
Range("A3").PasteSpecial

Range("M:M").Clear


strPDF_File_Name = "Invoice2.pdf"
ActiveWorkbook.FollowHyperlink strPDF_File_Name

Call GetInvoice


Range("b1") = "Invoice 2"
Range("m1").Select
Range("m1").PasteSpecial
Range("m7").Select
Range("m7").Copy
Range("b2").Select
Range("b2").PasteSpecial


Range("m8").Select
Range("m8").Copy
Range("b3").Select
Range("b3").PasteSpecial

Range("M:M").Clear


strPDF_File_Name = "Invoice3.pdf"
ActiveWorkbook.FollowHyperlink strPDF_File_Name

Call GetInvoice


Range("c1") = "Invoice 3"
Range("m1").Select
Range("m1").PasteSpecial
Range("m7").Select
Range("m7").Copy
Range("c2").Select
Range("c2").PasteSpecial


Range("m8").Select
Range("m8").Copy
Range("c3").Select
Range("c3").PasteSpecial

Range("M:M").Clear

strPDF_File_Name = "Invoice4.pdf"
ActiveWorkbook.FollowHyperlink strPDF_File_Name

Call GetInvoice


Range("d1") = "Invoice 4"
Range("m1").Select
Range("m1").PasteSpecial
Range("m7").Select
Range("m7").Copy
Range("d2").Select
Range("d2").PasteSpecial


Range("m8").Select
Range("m8").Copy
Range("d3").Select
Range("d3").PasteSpecial

Range("M:M").Clear

Cells.Select
Cells.EntireColumn.AutoFit

Range("e1").Select

End Sub

Public Sub GetInvoice()

Application.Wait Now + TimeValue("00:00:03")
SendKeys "^a"
Application.Wait Now + TimeValue("00:00:02")
SendKeys "^c"
Application.Wait Now + TimeValue("00:00:02")
SendKeys "^q"
Application.Wait Now + TimeValue("00:00:02")

End Sub
