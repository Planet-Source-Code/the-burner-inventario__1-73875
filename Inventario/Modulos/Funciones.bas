Attribute VB_Name = "Funciones"
Public Type USER_INFO
    USER_PK As String
    USER_NAME As String
    USER_ISADMIN As String
End Type
Public CurrUser                     As USER_INFO


'bind_dc "SELECT * FROM tbl_IC_Products", "ProductCode", dcProd, "PK", True


Public Sub bind_dc(ByVal srcSQL As String, ByVal srcBindField As String, ByRef SrcDC As DataCombo, Optional srcColBound As String, Optional ShowFirstRec As Boolean)
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open srcSQL, PrimeData, adOpenStatic, adLockOptimistic
    
    With SrcDC
        .ListField = srcBindField
        .BoundColumn = srcColBound
        Set .RowSource = rs
        'Display the first record
        If ShowFirstRec = True Then
            If Not rs.RecordCount < 1 Then
                .BoundText = rs.Fields(srcColBound)
                .Tag = rs.RecordCount & "*~~~~~*" & rs.Fields(srcColBound)
            Else
                .Tag = "0*~~~~~*0"
            End If
        End If
    End With
    Set rs = Nothing
End Sub

Public Function SQLDate(ConvertDate As Date) As String
SQLDate = Format(ConvertDate, "dd/mm/yyyy")
End Function

Public Function getValorCampo(ByVal srcSQL As String, ByVal whichField As String) As String
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open srcSQL, PrimeData, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then getValorCampo = rs.Fields(whichField)
    
    Set rs = Nothing
End Function

Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(srcCurr, "#,##0.00")
End Function

Public Sub LimpiarTexto(ByRef sForm As Form)
    Dim CONTROL As CONTROL
    For Each CONTROL In sForm.Controls
        If (TypeOf CONTROL Is TextBox) Then CONTROL = vbNullString
    Next CONTROL
    Set CONTROL = Nothing
End Sub

Public Sub LoadForm(ByRef srcForm As Form)
    srcForm.Show
    'srcForm.WindowState = vbMaximized
    srcForm.SetFocus
End Sub

Public Function Generar(ByVal srcNo As String, ByVal src1stStr As String, ByVal src2ndStr As String) As String
    If Len(src2ndStr) <= Len(srcNo) Then
        GenerateID = src1stStr & srcNo
    Else
        Generar = src1stStr & Left$(src2ndStr, Len(src2ndStr) - Len(srcNo)) & srcNo
    End If
    
End Function

Public Function GetTxtVal(ByVal sTxt As String) As Double

    Dim sNew As String
    Dim sC As String
    Dim i As Integer
    
    'default
    GetTxtVal = 0
        
    sTxt = Trim(sTxt)
    
    If Len(sTxt) > 0 Then
        For i = 1 To Len(sTxt)
            sC = Mid(sTxt, i, 1)
            If sC = "-" Or sC = "." Or sC = "1" Or sC = "2" Or sC = "3" Or sC = "4" Or sC = "5" Or sC = "6" Or sC = "7" Or sC = "8" Or sC = "9" Or sC = "0" Then
                sNew = sNew & sC
            End If
        Next
    
        If Len(sNew) > 0 Then
            GetTxtVal = Val(sNew)
        End If
    End If
    
    
End Function

Public Function HLTxt(ByRef txt As Object)
On Error Resume Next
    txt.SelStart = 0
    txt.SelLength = Len(txt)
    txt.SetFocus
End Function





Public Function EncryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i As Integer
 
 
 i = 0
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 
 Do
  i = i + 1
  OldChar = Asc(Mid(InString, i, 1))
  CryptChar = Asc(Mid(TempKey, i, 1))
  Select Case i Mod 2
   Case 0
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
   Case Else
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
  End Select
  If NewChar < 35 Then
   OutString = OutString + "!" + Chr(NewChar + 40)
  Else
   OutString = OutString + Chr(NewChar)
  End If
 Loop Until i = Len(InString)
 
 EncryptString = OutString

End Function

Public Function DecryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i, c As Integer
 
 
 c = 0
 i = 0
 
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 Do
 
  i = i + 1
  c = c + 1
  OldChar = Asc(Mid(InString, c, 1))
  If OldChar = 33 Then
   c = c + 1
   OldChar = Asc(Mid(InString, c, 1))
   OldChar = OldChar - 40
  End If
  CryptChar = Asc(Mid(TempKey, i, 1))
  Select Case i Mod 2
   Case 0
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
   Case Else
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
  End Select
  OutString = OutString + Chr(NewChar)
 Loop Until c = Len(InString)
 
 DecryptString = OutString

End Function

Public Function Esta_Vacio(ByRef sText As Variant, Optional UseTagValue As Boolean) As Boolean
    On Error Resume Next
    If sText.Text = "" Then
        Esta_Vacio = True
        If UseTagValue = True Then
            MsgBox "El Campo '" & sText.Tag & "' Es requerido!", vbExclamation
        Else
            MsgBox "El campo es requerido.", vbExclamation
        End If
        sText.SetFocus
    Else
        Esta_Vacio = False
    End If
End Function



