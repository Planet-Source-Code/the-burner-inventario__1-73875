VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Sonia 
   Caption         =   "Form1"
   ClientHeight    =   11670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   ScaleHeight     =   11670
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   585
      Left            =   11040
      TabIndex        =   4
      Top             =   10290
      Width           =   2115
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   420
      TabIndex        =   2
      Top             =   270
      Width           =   3195
   End
   Begin MSComctlLib.ProgressBar progreso 
      Height          =   255
      Left            =   270
      TabIndex        =   1
      Top             =   8370
      Width           =   12945
      _ExtentX        =   22834
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   585
      Left            =   270
      TabIndex        =   0
      Top             =   10410
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   3690
      Width           =   6285
   End
End
Attribute VB_Name = "Sonia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim sql As String
Dim cont As Integer
Dim i As Long
sql = "select * from Quitar"
Dim ssql As String
If rs.State = adStateOpen Then rs.Close
rs.Open sql, PrimeData, adOpenStatic, adLockOptimistic

MsgBox rs.RecordCount
rs.MoveFirst
i = 0
ssql = ""
While Not rs.EOF
         Me.Refresh
         
         'ssql = "Select * from msmanaure where"
         'ssql = ssql + " msmanaure.campo5='" & rs.Fields("campo2") & "' and msmanaure.campo10='" & rs.Fields("campo7") & "' and msmanaure.campo11='" & rs.Fields("campo8") & "'"
         'ssql = ssql + " and msmanaure.campo12='" & rs.Fields("campo9") & "' and msmanaure.campo14='" & rs.Fields("campo11") & "'"
         'If rs1.State = adStateOpen Then rs1.Close
         '   rs1.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
            
         'If rs1.RecordCount >= 1 Then
            ssql = "Delete FROM msmanaure WHERE"
            ssql = ssql + " msmanaure.campo5='" & rs.Fields("campo5") & "' and msmanaure.campo6='" & rs.Fields("campo6") & "' and msmanaure.campo7='" & rs.Fields("campo7") & "'"
            ssql = ssql + " and msmanaure.campo8='" & rs.Fields("campo8") & "' and msmanaure.campo9='" & rs.Fields("campo9") & "'"
            ssql = ssql + " and msmanaure.campo10='" & rs.Fields("campo10") & "'"
            ssql = ssql + " and msmanaure.campo11='" & rs.Fields("campo11") & "'"
            If rs1.State = adStateOpen Then rs1.Close
            rs1.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
               
         'End If
         
            progreso.Value = (i * 100) / rs.RecordCount
            Me.Refresh
            'Label1.Caption = "Voy X :" + Str(i)
            i = i + 1
            rs.MoveNext
            Me.Refresh
    
Wend

MsgBox "Se recorrieron: " + Str(i)

End Sub


Private Sub Command2_Click()


Dim rs As New Recordset
Dim rs1 As New Recordset
Dim sql As String
Dim cont As Integer
Dim i As Long
sql = "select * from Archivo2"
Dim ssql As String
If rs.State = adStateOpen Then rs.Close
rs.Open sql, PrimeData, adOpenStatic, adLockOptimistic

MsgBox rs.RecordCount
rs.MoveFirst
i = 0
ssql = ""
For j = 1 To 40968
         Me.Refresh
         
         'ssql = "Select * from msmanaure where"
         'ssql = ssql + " msmanaure.campo5='" & rs.Fields("campo2") & "' and msmanaure.campo10='" & rs.Fields("campo7") & "' and msmanaure.campo11='" & rs.Fields("campo8") & "'"
         'ssql = ssql + " and msmanaure.campo12='" & rs.Fields("campo9") & "' and msmanaure.campo14='" & rs.Fields("campo11") & "'"
         'If rs1.State = adStateOpen Then rs1.Close
         '   rs1.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
            
         'If rs1.RecordCount >= 1 Then
            ssql = "Delete FROM msmanaure WHERE"
            ssql = ssql + " msmanaure.campo5= Archivo1.campo2 and msmanaure.campo10=Archivo1.campo7 and msmanaure.campo11=Archivo1.campo8"
            ssql = ssql + " and msmanaure.campo12=Archivo1.campo9 and msmanaure.campo14=Archivo1.campo11"
            
            If rs1.State = adStateOpen Then rs1.Close
            rs1.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
               
         'End If
         
            progreso.Value = (i * 100) / rs.RecordCount
            Me.Refresh
            Label1.Caption = "Voy X :" + Str(i)
            i = i + 1
      '      rs.MoveNext
            Me.Refresh
    
Next j

MsgBox "Se recorrieron: " + Str(i)

End Sub

Private Sub Command3_Click()
Dim rs As New Recordset
Dim rs1 As New Recordset
Dim sql As String
Dim cont As Integer
Dim i As Long
sql = "select * from Archivo1"
Dim ssql As String
If rs.State = adStateOpen Then rs.Close
rs.Open sql, PrimeData, adOpenStatic, adLockOptimistic

MsgBox rs.RecordCount
End Sub

Private Sub Form_Load()
Call Conexion_1.Main_AfterSD


End Sub

