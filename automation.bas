Attribute VB_Name = "Módulo1"
Sub Criar_coluna()
    
    
    Planilha1.Activate               'cria coluna Status
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Status"
    Range("G1").Select               'cria coluna Data
    ActiveCell.FormulaR1C1 = "Data"
    
               
    For i = 2 To 1000 'buscar valores abaixo de 100R$
            
            If Planilha1.Range("d" & i).value < 100 And Planilha1.Range("a" & i).value <> "" Then
    
             Range("f" & i).value = "Nao enviar"
             
             End If
             
        If Planilha1.Range("a" & i).value <> "" Then 'colocar a data atual
             
        Range("g" & i).value = Date
                        
       
        End If
    Next
  
End Sub


Sub verificar_email() 'verificar se o campo email está vazio

For e = 2 To 1000
If Planilha1.Range("e" & e).value = "" And Planilha1.Range("a" & e).value <> "" Then

Range("f" & e).value = "FALHA - Sem e-mail"

End If

Next

End Sub
Function IsValidEmail(sEmailAddress As String) As Boolean

   
    Dim sEmailPattern As String
    Dim oRegEx As Object
    Dim bReturn As Boolean

    
    'Use the below regular expressions
   
    sEmailPattern = "^[a-z0-9_.-]+@([a-z]{4,}\.)+(?:[com]{3,})+(\.[br]{2}|)$"
    'sEmailPattern = "^[a-z0-9_.-]+@[a-z]{2,}\.[com]{3,}$"
            
    
    'Create Regular Expression Object
    Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Global = True
    oRegEx.IgnoreCase = True
    oRegEx.Pattern = sEmailPattern
    
        
        
    bReturn = False
    
    'Check if Email match regex pattern
    If oRegEx.Test(sEmailAddress) Then
        'Debug.Print "Valid Email ('" & sEmailAddress & "')"
        bReturn = True
    Else
        'Debug.Print "Invalid Email('" & sEmailAddress & "')"
        bReturn = False
    End If

    'Return validation result
    IsValidEmail = bReturn
    
    
End Function

Sub validar_email()

For m = 2 To 1000
If Planilha1.Range("e" & m).value <> "" And Planilha1.Range("a" & m).value <> "" And Planilha1.Range("d" & m).value >= 100 Then
 
 verifica = IsValidEmail(Planilha1.Range("e" & m).value)
  If verifica = False Then
    Planilha1.Range("f" & m).value = "FALHA - E-mail invalido"
    
    Else
    
    Planilha1.Range("f" & m).value = "SUCESSO"
    
    
End If
End If
On Error Resume Next
Next

End Sub


Public Function fncSplitNome(VarNome As String)

'Pegar apenas o primeiro nome

Dim VarSplit

VarSplit = Split(VarNome, " ")

fncSplitNome = VarSplit(0)



End Function

Sub pegarNome()

Planilha1.Range("h1").value = "FirstName"


For n = 2 To 1000
If Planilha1.Range("f" & n).value = "SUCESSO" And Planilha1.Range("d" & n).value >= 100 Then

nome = fncSplitNome(Planilha1.Range("b" & n).value)

Planilha1.Range("h" & n).value = nome


End If
Next


End Sub

