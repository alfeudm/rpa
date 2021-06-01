Attribute VB_Name = "pdfill"
Sub pdfill()


For i = 2 To 3
If Planilha3.Range("a" & i).Text = "" Then
                Exit For
                
            End If

Shell ("D:\Program Files (x86)\PlotSoft\PDFill\PDFill_PDF_Tools.exe")

Application.Wait Now + TimeValue("00:00:03")

SetCursorPos 710, 1000 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

Application.Wait Now + TimeValue("00:00:01")
SendKeys "{tab}"
Application.Wait Now + TimeValue("00:00:01")

SendKeys "{up}"
Application.Wait Now + TimeValue("00:00:04")

SendKeys "~"
SendKeys "~"

Application.Wait Now + TimeValue("00:00:04")

SendKeys "{tab 5}"

Application.Wait Now + TimeValue("00:00:01")

 SendKeys "~"

Application.Wait Now + TimeValue("00:00:01")

Application.SendKeys (Planilha3.Range("e1").Text) 'local onde está o pdf 1

Application.Wait Now + TimeValue("00:00:02")

SendKeys "~"

SendKeys "~"


Application.Wait Now + TimeValue("00:00:01")
SendKeys "~"
SendKeys "{tab 6}"

Application.SendKeys (Planilha3.Range("a2").Text)

Application.Wait Now + TimeValue("00:00:02")

SendKeys "{tab 7}"
                                        

Application.Wait Now + TimeValue("00:00:02")

SendKeys "~"

Application.Wait Now + TimeValue("00:00:03")

SendKeys "{tab 5}" 'adicionar outro pdf

SendKeys "~"

Application.Wait Now + TimeValue("00:00:02")

Application.SendKeys (Planilha3.Range("a3").Text) 'segundo pdf

Application.Wait Now + TimeValue("00:00:01")

SendKeys "~"

Application.Wait Now + TimeValue("00:00:02")

SendKeys "{tab 4}" 'clicar em salvar

Application.Wait Now + TimeValue("00:00:01")

SendKeys "~"

Application.Wait Now + TimeValue("00:00:02")

SendKeys "{tab 6}"
SendKeys "~"
Application.Wait Now + TimeValue("00:00:02")

Application.SendKeys (Planilha3.Range("f1").Text) 'local onde vai o pdf

Application.Wait Now + TimeValue("00:00:02")
SendKeys "~"
                    
Application.Wait Now + TimeValue("00:00:02")

SendKeys "{tab 6}"

Application.Wait Now + TimeValue("00:00:01")

Application.SendKeys (Planilha3.Range("a2").Text)

Application.Wait Now + TimeValue("00:00:03")

SendKeys "~"



SendKeys "%{f4}"
Application.Wait Now + TimeValue("00:00:01")
SendKeys "%{f4}"

    Application.Wait Now + TimeValue("00:00:03")

 'Planilha3.Activate
  'Rows("2:3").Select
    'Selection.Delete Shift:=xlUp
    'Range("C2").Select
    



Next

End Sub
