Attribute VB_Name = "Conexion_1"
Option Explicit

Public Const DBFileName = "Datos.mdb"
Public PrimeData As New ADODB.Connection

Public DBPathFileName As String
Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Sub Main()
    FrmMain.Show
End Sub

Public Sub Main_AfterSD()
    If InitDB = False Then
        Exit Sub
    End If
    If OpenDB = False Then
        Exit Sub
    End If
End Sub

Private Sub TestUnit()

     
End Sub


Public Function InitDB() As Boolean
    
    Dim fso As New FileSystemObject
    
    InitDB = False
    
    
    If fso.FileExists(App.Path & "\" & DBFileName) = False Then
       End
       DBPathFileName = App.Path & "\" & DBFileName
       GoTo RAE
    End If
    
    
    DBPathFileName = App.Path & "\" & DBFileName
    
    InitDB = True
    
RAE:
    Set fso = Nothing
End Function

Public Function OpenDB() As Boolean

    OpenDB = False
    
    
    If ConnectDB(PrimeData, DBPathFileName) = False Then
       GoTo RAE
    End If
    
    OpenDB = True
    
RAE:
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






