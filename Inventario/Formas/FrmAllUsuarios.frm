VERSION 5.00
Begin VB.Form FrmAllUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizacion de Usuarios"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   5730
   Begin Proyecto1.OsenXPButton CmdAgrega 
      Height          =   525
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Agregar Productos..."
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "FrmAllUsuarios.frx":0000
      PICN            =   "FrmAllUsuarios.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.OsenXPButton OsenXPButton1 
      Height          =   525
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Agregar Productos..."
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "&Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "FrmAllUsuarios.frx":01A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.LynxGrid3 GrillaUsuarios 
      Height          =   4215
      Left            =   30
      TabIndex        =   0
      Top             =   1560
      Width           =   5625
      _ExtentX        =   9499
      _ExtentY        =   8493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorBkg    =   16777215
      BackColorFixed  =   33023
      BackColorSel    =   15849673
      ThemeStyle      =   3
      EditTrigger     =   0
      SBackColor1     =   0
      SBackColor2     =   0
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   150
      Picture         =   "FrmAllUsuarios.frx":01BC
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Usuarios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   495
      Index           =   0
      Left            =   900
      TabIndex        =   5
      Top             =   30
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DISTRIBUIDORA POMPI LTDA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   930
      TabIndex        =   4
      Top             =   480
      Width           =   2745
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kra. 14 No. 14-40 Tel. (035) - 7267710 Maicao - Guajira"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   930
      TabIndex        =   3
      Top             =   690
      Width           =   4020
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   -30
      Width           =   12375
   End
End
Attribute VB_Name = "FrmAllUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAgrega_Click()
With FrmActUsuarios
     .Modo = "Nuevo"
     .Show 1
     
End With
End Sub

Private Sub Form_Load()
With GrillaUsuarios
        .Redraw = False
        .AddColumn "ID", 30        '0
        .AddColumn "Codigo", 95        '
        .AddColumn "Nombre", 150
        .AddColumn "Privilegios", 95        '
        .Redraw = True
        .Refresh

End With
Me.TraerUsuarios
End Sub

Public Sub TraerUsuarios()
   
    Dim vRs As New ADODB.Recordset
    Dim sSql As String
    Dim il As Long
                
    'clear list
    Me.GrillaUsuarios.Redraw = False
    Me.GrillaUsuarios.Clear
    
    sSql = "SELECT * from TBL_Usuario"
    If ConnectRS(PrimeData, vRs, sSql) = False Then
       MsgBox Me.Name & "," & "Usuarios" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
        GoTo RAE
    End If
    
    If vRs.RecordCount = 0 Then Exit Sub
    'Me.LblTotal.Caption = "Total Usuarios: " + Str(vRs.RecordCount)
    vRs.MoveFirst
    While vRs.EOF = False
        With GrillaUsuarios
            il = .AddItem(vRs.Fields("ID"))
            .ItemImage(il) = 2
            .CellText(il, 1) = vRs.Fields("Codigo_Usuario")
            .CellText(il, 2) = vRs.Fields("Nombre_Usuario")
            .CellText(il, 3) = vRs.Fields("Privilegio")
        End With
        
        vRs.MoveNext
    Wend
    
RAE:
    Set vRs = Nothing
    Me.GrillaUsuarios.Redraw = True
    Me.GrillaUsuarios.Refresh
End Sub

Private Sub GrillaUsuarios_DblClick()
  If Me.GrillaUsuarios.RowCount = 0 Then Exit Sub
  With FrmActUsuarios
       .lblID.Caption = Me.GrillaUsuarios.CellText(Me.GrillaUsuarios.Row, 0)
       .TxtCodigo.Text = Me.GrillaUsuarios.CellText(Me.GrillaUsuarios.Row, 1)
       .TxtNombre.Text = Me.GrillaUsuarios.CellText(Me.GrillaUsuarios.Row, 2)
       .TxtPrivilegio.Text = Me.GrillaUsuarios.CellText(Me.GrillaUsuarios.Row, 3)
       .Modo = "Editar"
       
       .Show 1
  End With
End Sub

Private Sub OsenXPButton1_Click()
Unload Me
End Sub
