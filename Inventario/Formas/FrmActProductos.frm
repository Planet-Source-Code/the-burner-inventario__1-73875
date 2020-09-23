VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmActProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualización de Productos"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtId 
      Height          =   375
      Left            =   1050
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2820
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtGrupo 
      BackColor       =   &H00EAFDFF&
      Enabled         =   0   'False
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
      Left            =   3540
      TabIndex        =   16
      Tag             =   "Codigo"
      Top             =   1740
      Width           =   3105
   End
   Begin Proyecto1.OsenXPButton CmdGuardar 
      Height          =   525
      Left            =   3990
      TabIndex        =   11
      Top             =   2640
      Width           =   1365
      _ExtentX        =   2408
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
      MICON           =   "FrmActProductos.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtCosto 
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
      Left            =   1080
      TabIndex        =   10
      Tag             =   "Codigo"
      Top             =   2190
      Width           =   1755
   End
   Begin VB.TextBox TxtPrecio 
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
      Left            =   1080
      TabIndex        =   8
      Tag             =   "Codigo"
      Top             =   1740
      Width           =   1755
   End
   Begin VB.TextBox TxtDescripcion 
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
      Left            =   1080
      TabIndex        =   6
      Tag             =   "Codigo"
      Top             =   1320
      Width           =   5565
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
      Left            =   1080
      TabIndex        =   4
      Tag             =   "Codigo"
      Top             =   900
      Width           =   1095
   End
   Begin Proyecto1.OsenXPButton CmdSalir 
      Height          =   525
      Left            =   5400
      TabIndex        =   12
      Top             =   2640
      Width           =   1365
      _ExtentX        =   2408
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
      MICON           =   "FrmActProductos.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo dcProd 
      Height          =   315
      Left            =   3180
      TabIndex        =   14
      Top             =   900
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo"
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
      Index           =   7
      Left            =   2970
      TabIndex        =   17
      Top             =   1770
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2250
      TabIndex        =   15
      Top             =   930
      Width           =   855
   End
   Begin VB.Label lblID 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2910
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo"
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
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   2310
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
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
      Index           =   5
      Left            =   90
      TabIndex        =   7
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
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
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   1350
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
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
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   960
      Width           =   495
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
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   4020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualización de Productos"
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
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   2640
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
      Left            =   60
      TabIndex        =   0
      Top             =   330
      Width           =   2745
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   0
      Picture         =   "FrmActProductos.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12420
   End
End
Attribute VB_Name = "FrmActProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Modos As String

Private Sub CmdGuardar_Click()
Dim vRsProductos As New Recordset
Dim ssql As String

If Me.TxtCosto.Text > Me.TxtPrecio.Text Then
   MsgBox "El Precio de Costo no puede ser mayor al Precio de Venta.", vbCritical
   Exit Sub
End If
If Modos = "Nuevo" Then
   ssql = "Select * From TBL_Producto Where Codigo_Producto = '" & Me.TxtCodigo.Text & "'"
   If ConnectRS(PrimeData, vRsProductos, ssql) = False Then
       MsgBox Me.Name & "," & "Act. Productos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
   End If
   If vRsProductos.RecordCount >= 1 Then
    MsgBox "El Codigo del Producto ya Existe", vbCritical
    Exit Sub
   End If
   ssql = "INSERT INTO TBL_Producto(Codigo_Producto,Descripcion_Producto,Cod_grupo,Precio_Costo,Precio_Venta) "
   ssql = ssql + " VALUES('" & Me.TxtCodigo.Text & "', '" & Me.TxtDescripcion & "',"
   ssql = ssql + " '" & Me.lblID.Caption & "', '" & Me.TxtCosto.Text & "',"
   ssql = ssql + " '" & Me.TxtPrecio.Text & "')"
   
   If ConnectRS(PrimeData, vRsProductos, ssql) = False Then
       MsgBox Me.Name & "," & "Act. Productos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
   End If
   ssql = "INSERT INTO TBL_InventarioDiario(Cod_Producto) VALUES('" & Me.TxtCodigo.Text & "')"
   If ConnectRS(PrimeData, vRsProductos, ssql) = False Then
       MsgBox Me.Name & "," & "Act. Productos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
   End If
   
End If

If Modos = "Editar" Then
  Me.TxtCodigo.Enabled = False
  ssql = "UPDATE TBL_Producto SET Descripcion_Producto = '" & Me.TxtDescripcion.Text & "',"
  ssql = ssql + " Precio_Venta = '" & Me.TxtPrecio.Text & "',"
  ssql = ssql + " Precio_Costo = '" & Me.TxtCosto.Text & "',"
  ssql = ssql + " Cod_Grupo = " & dcProd.BoundText & ""
  ssql = ssql + " Where ID = " & Me.TxtId.Text
  If ConnectRS(PrimeData, vRsProductos, ssql) = False Then
       MsgBox Me.Name & "," & "Act. Productos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
   End If
End If
FrmAllProductos.TraerProductos (Me.lblID.Caption)
MsgBox "Datos Actualizados Correctamente", vbInformation
Unload Me
RAE:
Set vRsProductos = Nothing

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub dcProd_Change()
Me.lblID.Caption = dcProd.BoundText
End Sub

Private Sub Form_Load()
bind_dc "SELECT * FROM tbl_Grupo", "Descripcion", dcProd, "ID", True
End Sub

