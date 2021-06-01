Attribute VB_Name = "pdf2go"
Sub juntar_pdf()

For i = 1 To 1000
If Planilha3.Range("a" & i).Text = "" Then
                Exit For
                
            End If

ActiveWorkbook.FollowHyperlink "https://www.pdf2go.com/merge-pdf", NewWindow:=False


Application.Wait Now + TimeValue("00:00:05")

SetCursorPos 590, 490 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
 Application.Wait Now + TimeValue("00:00:03")
 
 Application.SendKeys (Planilha3.Range("a1").Text)
 
 Application.Wait Now + TimeValue("00:00:02")
 
 SendKeys "~"
 
 Application.Wait Now + TimeValue("00:00:02")
 
 
 
 Application.Wait Now + TimeValue("00:00:08")

SendKeys "{down 5}"

Application.Wait Now + TimeValue("00:00:02")

SetCursorPos 290, 350 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
 Application.Wait Now + TimeValue("00:00:01")
 
 Application.SendKeys (Planilha3.Range("a2").Text)
 
 Application.Wait Now + TimeValue("00:00:02")
 
 SendKeys "~"
 
 Application.Wait Now + TimeValue("00:00:02")
 
 SendKeys "{up 5}"
 
 Application.Wait Now + TimeValue("00:00:03")
 
 SetCursorPos 1100, 190 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  
 Application.Wait Now + TimeValue("00:00:02")
 
 SetCursorPos 1100, 290 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
 
  Application.Wait Now + TimeValue("00:00:15")
  
  
  
  SendKeys "^w"
  
  Application.Wait Now + TimeValue("00:00:05")
  
  'Planilha3.Activate
    'Range("A1:A2").Select
    'Application.CutCopyMode = False
    'Selection.ClearContents
    'Range("A3").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'ActiveWindow.SmallScroll Down:=-210
    'Selection.Cut Destination:=Range("A1:A216")
    'Range("B1").Select
  
  'Planilha3.Activate
  'Rows("1:2").Select
    'Selection.Delete Shift:=xlUp
    'Range("C2").Select
 
Next

End Sub
