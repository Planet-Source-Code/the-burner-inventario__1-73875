VERSION 5.00
Begin VB.Form FrmActUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualización de Usuarios"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.OsenXPButton CmdGuardar 
      Height          =   525
      Left            =   2430
      TabIndex        =   10
      Top             =   2940
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "&Actualizar"
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
      MICON           =   "FrmActUsuarios.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox TxtPrivilegio 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FrmActUsuarios.frx":001C
      Left            =   1290
      List            =   "FrmActUsuarios.frx":0026
      TabIndex        =   9
      Tag             =   "Perfil"
      Top             =   2520
      Width           =   2085
   End
   Begin VB.TextBox TxtRepita 
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
      Left            =   1290
      PasswordChar    =   "X"
      TabIndex        =   6
      Tag             =   "Contraseña"
      Top             =   2130
      Width           =   2055
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
      Left            =   1290
      PasswordChar    =   "X"
      TabIndex        =   5
      Tag             =   "Constraseña"
      Top             =   1740
      Width           =   2055
   End
   Begin VB.TextBox TxtNombre 
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
      Left            =   1290
      TabIndex        =   4
      Tag             =   "Nombre"
      Top             =   1350
      Width           =   4845
   End
   Begin VB.TextBox TxtCodigo 
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
      Left            =   1290
      TabIndex        =   2
      Tag             =   "Codigo"
      Top             =   990
      Width           =   2055
   End
   Begin Proyecto1.OsenXPButton CmdSalir 
      Height          =   525
      Left            =   4410
      TabIndex        =   11
      Top             =   2940
      Width           =   1905
      _ExtentX        =   3360
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
      MICON           =   "FrmActUsuarios.frx":0043
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
      Index           =   7
      Left            =   60
      TabIndex        =   15
      Top             =   330
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
      Index           =   4
      Left            =   60
      TabIndex        =   14
      Top             =   540
      Width           =   4020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Privilegio"
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
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   2550
      Width           =   870
   End
   Begin VB.Label LblID 
      Caption         =   "_"
      Height          =   285
      Left            =   5340
      TabIndex        =   12
      Top             =   930
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repitala"
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
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   2130
      Width           =   795
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1770
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualización de Usuarios"
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
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "FrmActUsuarios.frx":005F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6570
   End
End
Attribute VB_Name = "FrmActUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modo As String

Private Sub CmdGuardar_Click()
Dim vRsUsuarios As New Recordset
Dim sSql As String

'Si va Agregar uno nuevo
If Modo = "Nuevo" Then

   Me.TxtCodigo.Enabled = True
   Me.TxtContraseña.Enabled = True
   Me.TxtRepita.Enabled = True
   
   If Funciones.Esta_Vacio(Me.TxtCodigo, True) = True Then Exit Sub
   If Funciones.Esta_Vacio(Me.TxtNombre, True) = True Then Exit Sub
   If Funciones.Esta_Vacio(Me.TxtContraseña, True) = True Then Exit Sub
   If Funciones.Esta_Vacio(Me.TxtRepita, True) = True Then Exit Sub
   If Funciones.Esta_Vacio(Me.TxtPrivilegio, True) = True Then Exit Sub
   
   Contraseña = Funciones.EncryptString(Me.TxtContraseña.Text, "POMPI")
   
   If Me.TxtContraseña.Text <> Me.TxtRepita.Text Then MsgBox "Las Claves no son Iguales, Favor Verifiqueles", vbInformation: Exit Sub
   sSql = "Select * From TBL_Usuario Where Codigo_Usuario = '" & Me.TxtCodigo & "'"
   If ConnectRS(PrimeData, vRsUsuarios, sSql) = False Then
       MsgBox Me.Name & "," & "Act. Usuarios" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
       GoTo RAE
   End If
   If vRsUsuarios.RecordCount >= 1 Then
      MsgBox "el Usuario ya Existe en el Sistema", vbInformation
      Funciones.HLTxt Me.TxtCodigo
      Exit Sub
   End If
   sSql = ""
   sSql = "INSERT INTO TBL_Usuario(Codigo_Usuario,Nombre_Usuario,Clave_Usuario,Privilegio) "
   sSql = sSql + " VALUES('" & Me.TxtCodigo.Text & "', '" & Me.TxtNombre.Text & "',"
   sSql = sSql + " '" & Contraseña & "', '" & Me.TxtPrivilegio.Text & "')"
   If ConnectRS(PrimeData, vRsUsuarios, sSql) = False Then
       MsgBox Me.Name & "," & "Act. Usuarios" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
       GoTo RAE
   End If
   
End If
'Si va Editar un usuario
If Modo = "Editar" Then
   Me.TxtContraseña.Enabled = False
   Me.TxtCodigo.Enabled = False
   Me.TxtRepita.Enabled = False
   
   sSql = "UPDATE TBL_Usuario SET Nombre_Usuario = '" & Me.TxtNombre.Text & "',"
   sSql = sSql + "Privilegio = '" & Me.TxtPrivilegio.Text & "'"
   sSql = sSql + " WHERE ID = " & Me.lblID.Caption
   If ConnectRS(PrimeData, vRsUsuarios, sSql) = False Then
       MsgBox Me.Name & "," & "Act. Usuarios" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
       GoTo RAE
   End If
End If
FrmAllUsuarios.TraerUsuarios
MsgBox "Datos Actualizados Correctamente", vbInformation
Unload Me
RAE:
Set vRsUsuarios = Nothing
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.TxtPrivilegio.ListIndex = 0
End Sub

Private Sub TxtCodigo_GotFocus()
'Me.TxtCodigo.Text = Me.TxtCodigo.Text
End Sub

