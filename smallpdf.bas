Attribute VB_Name = "smallpdf"
Sub juntar_smallpdf()

For i = 1 To 3
If Planilha3.Range("a" & i).Text = "" Then
                Exit For
                
            End If

ActiveWorkbook.FollowHyperlink "https://smallpdf.com/pt/juntar-pdf", NewWindow:=False


Application.Wait Now + TimeValue("00:00:05")

SetCursorPos 590, 490 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
 Application.Wait Now + TimeValue("00:00:02")
 
 Application.SendKeys (Planilha3.Range("a" & i).Text)
 
 Application.Wait Now + TimeValue("00:00:01")
 
 SendKeys "~"
 
 Application.Wait Now + TimeValue("00:00:09")
 
 SetCursorPos 590, 490 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
  Application.Wait Now + TimeValue("00:00:02")
  
  SetCursorPos 590, 585 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
Application.Wait Now + TimeValue("00:00:02")

SetCursorPos 290, 620 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
  Application.Wait Now + TimeValue("00:00:02")
  
  SetCursorPos 210, 490 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
  Application.Wait Now + TimeValue("00:00:02")
  
  Application.SendKeys (Planilha3.Range("b" & i).Text)
 
 Application.Wait Now + TimeValue("00:00:01")
 
 SendKeys "~"
 
 Application.Wait Now + TimeValue("00:00:10")
 
 SetCursorPos 1010, 605 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
  Application.Wait Now + TimeValue("00:00:15")
  
 
  SetCursorPos 1010, 272 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
  Application.Wait Now + TimeValue("00:00:08")
  
  SendKeys "^w"
  
  'Application.Wait Now + TimeValue("00:00:05")
  
  'Planilha3.Activate
  'Rows("1:2").Select
   ' Selection.Delete Shift:=xlUp
    'Range("C2").Select
 
Next


End Sub
