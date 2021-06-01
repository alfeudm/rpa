Attribute VB_Name = "pdf24"
Sub pdf24()


Shell ("D:\Program Files\PDF24\pdf24-Ocr.exe")


Application.Wait Now + TimeValue("00:00:02")

For i = 53 To 503
If Planilha2.Range("a" & i).Text = "" Then
                Exit For
                
            End If

SetCursorPos 250, 190 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
 Application.Wait Now + TimeValue("00:00:01")
 
 Application.SendKeys (Planilha2.Range("a" & i).Text)

Application.Wait Now + TimeValue("00:00:01")

SendKeys "~"

Application.Wait Now + TimeValue("00:00:01")

SetCursorPos 250, 190 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
  
 Application.Wait Now + TimeValue("00:00:01")
 
 Application.SendKeys (Planilha2.Range("b" & i).Text)

Application.Wait Now + TimeValue("00:00:01")

SendKeys "~"

Application.Wait Now + TimeValue("00:00:01")

SetCursorPos 250, 190 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
  Application.Wait Now + TimeValue("00:00:01")
  
 
 Application.SendKeys (Planilha2.Range("c" & i).Text)

Application.Wait Now + TimeValue("00:00:01")

SendKeys "~"

Application.Wait Now + TimeValue("00:00:01")

SetCursorPos 250, 190 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
  Application.Wait Now + TimeValue("00:00:01")
  
 
 Application.SendKeys (Planilha2.Range("d" & i).Text)

Application.Wait Now + TimeValue("00:00:01")

SendKeys "~"

Application.Wait Now + TimeValue("00:00:01")
  
  SetCursorPos 250, 190 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
 Application.Wait Now + TimeValue("00:00:01")
 
 Application.SendKeys (Planilha2.Range("e" & i).Text)

Application.Wait Now + TimeValue("00:00:01")

SendKeys "~"

Application.Wait Now + TimeValue("00:00:01")

 SetCursorPos 650, 190 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

Application.Wait Now + TimeValue("00:0:33")


 SetCursorPos 370, 190 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
  Application.Wait Now + TimeValue("00:00:02")
  
  'SendKeys "%{f4}"


Next


End Sub
