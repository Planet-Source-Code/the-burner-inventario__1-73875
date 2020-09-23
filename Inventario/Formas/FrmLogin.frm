VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login de Usuarios"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.OsenXPButton CmdLogin 
      Default         =   -1  'True
      Height          =   525
      Left            =   3390
      TabIndex        =   6
      Top             =   2070
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "FrmLogin.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtContraseña 
      BackColor       =   &H00EAFDFF&
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3390
      PasswordChar    =   "X"
      TabIndex        =   2
      Tag             =   "Constraseña"
      Top             =   1320
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DUsuarios 
      Height          =   315
      Left            =   3390
      TabIndex        =   0
      Top             =   900
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin Proyecto1.OsenXPButton OsenXPButton2 
      Height          =   525
      Left            =   5100
      TabIndex        =   7
      Top             =   2070
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "&Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "FrmLogin.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   2100
      TabIndex        =   8
      Top             =   540
      Width           =   4020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   240
      Index           =   1
      Left            =   2010
      TabIndex        =   5
      Top             =   1350
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   240
      Index           =   0
      Left            =   1980
      TabIndex        =   4
      Top             =   930
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   0
      Picture         =   "FrmLogin.frx":0038
      Top             =   900
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login de Usuario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   240
      Index           =   3
      Left            =   2070
      TabIndex        =   3
      Top             =   30
      Width           =   1590
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
      Left            =   2100
      TabIndex        =   1
      Top             =   330
      Width           =   2745
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   0
      Picture         =   "FrmLogin.frx":3EED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6840
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdLogin_Click()
If DUsuarios.Text = "" Then DUsuarios.SetFocus: Exit Sub
If Me.TxtContraseña.Text = "" Then Me.TxtContraseña.SetFocus: Exit Sub


    Dim strPass As String
    
    strPass = getValorCampo("SELECT ID,Clave_Usuario FROM tbl_Usuario WHERE Codigo_Usuario= '" & DUsuarios.BoundText & "'", "Clave_Usuario")
    strPass = Funciones.DecryptString(strPass, "POMPI")
    
    If LCase(Me.TxtContraseña.Text) = LCase(strPass) Then
        With CurrUser
             .USER_NAME = DUsuarios.Text
             .USER_PK = Funciones.getValorCampo("SELECT Codigo_Usuario,Nombre_Usuario FROM tbl_Usuario WHERE Codigo_Usuario= '" & DUsuarios.BoundText & "'", "Nombre_Usuario")
            .USER_ISADMIN = Funciones.getValorCampo("SELECT Codigo_Usuario,Privilegio FROM tbl_Usuario WHERE Codigo_Usuario= '" & DUsuarios.BoundText & "'", "Privilegio")
             FrmMain.LblCodigoUsuario.Caption = .USER_NAME
             FrmMain.LblNombre.Caption = .USER_PK
             FrmMain.LblPerfil.Caption = .USER_ISADMIN
            'MsgBox .USER_PK
        End With
        
        FrmMain.CloseMe = True
        Unload Me
    Else
        MsgBox "Contraseña Invalida. Intente de nuevo!", vbExclamation
        Me.TxtContraseña.SetFocus
    End If
    strPass = vbNullString

End Sub

Private Sub Form_Load()
Call Conexion_1.Main_AfterSD

'bind_dc "SELECT * FROM TBL_Usuario", "", dcUser, "PK"
bind_dc "SELECT * FROM TBL_Usuario", "Codigo_Usuario", DUsuarios, "Codigo_Usuario", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If FrmMain.CloseMe = False Then End
End Sub

Private Sub OsenXPButton2_Click()
End
End Sub
