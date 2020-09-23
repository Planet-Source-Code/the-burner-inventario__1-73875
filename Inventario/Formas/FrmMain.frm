VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D28F8786-0BB9-402B-92DC-F32DE23A324E}#3.0#0"; "OutlookBar.ocx"
Begin VB.MDIForm FrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ANGEL - INVENTARIO. Distribuidora POMPI LTDA"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17025
   LinkTopic       =   "MDIForm1"
   Picture         =   "FrmMain.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   16965
      TabIndex        =   5
      Top             =   9735
      Width           =   17025
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e-Mail : carlosj.gamez@gmail.com Riohacha, La Guajira - Colombia"
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
         Index           =   8
         Left            =   9780
         TabIndex        =   17
         Top             =   510
         Width           =   4740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. 301 614 0094"
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
         Index           =   9
         Left            =   9780
         TabIndex        =   15
         Top             =   750
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANGEL.NET"
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
         Left            =   9780
         TabIndex        =   14
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Software de Inventario"
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
         Index           =   5
         Left            =   9780
         TabIndex        =   11
         Top             =   30
         Width           =   2280
      End
      Begin VB.Label LblPerfil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "_"
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
         Left            =   1800
         TabIndex        =   10
         Top             =   810
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil"
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
         Index           =   2
         Left            =   1800
         TabIndex        =   9
         Top             =   570
         Width           =   495
      End
      Begin VB.Label LblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "_"
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
         Left            =   1800
         TabIndex        =   8
         Top             =   300
         Width           =   120
      End
      Begin VB.Label LblCodigoUsuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "_"
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
         Left            =   3330
         TabIndex        =   7
         Top             =   60
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Actual"
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
         Left            =   1800
         TabIndex        =   6
         Top             =   60
         Width           =   1410
      End
      Begin VB.Image Image2 
         Height          =   1260
         Left            =   60
         Picture         =   "FrmMain.frx":1F69
         Top             =   30
         Width           =   1335
      End
   End
   Begin VB.PictureBox CR 
      Align           =   1  'Align Top
      Height          =   1185
      Left            =   0
      ScaleHeight     =   1125
      ScaleWidth      =   16965
      TabIndex        =   1
      Top             =   0
      Width           =   17025
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DISTRIBUIDORA MRNETWORK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   10
         Left            =   5370
         TabIndex        =   16
         Top             =   60
         Width           =   6390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DISTRIBUIDORA MRNETWORK"
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
         Left            =   5400
         TabIndex        =   13
         Top             =   540
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calle 11 No. 8-06 Cel 301 614 0094"
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
         Left            =   5400
         TabIndex        =   12
         Top             =   750
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Software de Inventario"
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
         TabIndex        =   4
         Top             =   420
         Width           =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANGEL-  INVENTARIO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   30
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   1455
         Index           =   1
         Left            =   0
         Picture         =   "FrmMain.frx":779B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   19320
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8550
      Left            =   0
      ScaleHeight     =   8550
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   1185
      Width           =   2700
      Begin OutlookBar.ctxOutlookBar ctxOutlookBar1 
         Height          =   6045
         Left            =   30
         TabIndex        =   2
         Top             =   180
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   10663
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormatControl   =   "FrmMain.frx":9541
         FormatGroup     =   "FrmMain.frx":96B5
         FormatGroupHover=   "FrmMain.frx":97A5
         FormatGroupPressed=   "FrmMain.frx":9881
         FormatGroupSelected=   "FrmMain.frx":9955
         FormatItem      =   "FrmMain.frx":9A15
         FormatItemLargeIcons=   "FrmMain.frx":9B11
         FormatItemHover =   "FrmMain.frx":9C0D
         FormatItemPressed=   "FrmMain.frx":9CE9
         FormatItemSelected=   "FrmMain.frx":9D95
         FormatSmallIcon =   "FrmMain.frx":9E41
         FormatSmallIconHover=   "FrmMain.frx":9F3D
         FormatSmallIconPressed=   "FrmMain.frx":A039
         FormatSmallIconSelected=   "FrmMain.frx":A135
         FormatLargeIcon =   "FrmMain.frx":A231
         FormatLargeIconHover=   "FrmMain.frx":A319
         FormatLargeIconPressed=   "FrmMain.frx":A415
         FormatLargeIconSelected=   "FrmMain.frx":A511
         Groups          =   "FrmMain.frx":A60D
      End
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   9540
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":15ED9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":168EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":172FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":17697
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":17A31
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":17DCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":18165
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":18B77
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":19589
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":19F9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1A9AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1B3BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1BDD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1C7E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1CD7F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CloseMe  As Boolean


Private Sub ctxOutlookBar1_ButtonClick(ByVal oBtn As OutlookBar.cButton)
On Error GoTo Error_No_Conexion
Select Case oBtn.Caption
   
   Case "Productos"
       If CurrUser.USER_ISADMIN = "Operador" Then
          MsgBox "Usted no tiene privilegio para Accesar a Entradas y Salidas de Productos." + vbNewLine + _
          "Comuniquese con el Administrador del Aplicativo", vbCritical
          Exit Sub
       End If

         'Me.Enabled = False
         FrmAllProductos.Show vbModal
   
   Case "Usuarios"
         If CurrUser.USER_ISADMIN = "Operador" Then
            MsgBox "Usted no tiene privilegio para Accesar a Entradas y Salidas de Productos." + vbNewLine + _
            "Comuniquese con el Administrador del Aplicativo", vbCritical
            Exit Sub
        End If

         'Me.Enabled = False
         FrmAllUsuarios.Show
   
   Case "E/S Bodega"
         'Me.Enabled = False
         If CurrUser.USER_ISADMIN = "Operador" Then
            MsgBox "Usted no tiene privilegio para Accesar a Entradas y Salidas de Productos." + vbNewLine + _
            "Comuniquese con el Administrador del Aplicativo", vbCritical
            Exit Sub
         End If
 
         FrmEntradas.Show
         
   Case "Pre-Factura"
         'Me.Enabled = False
         FrmPrefactura.Show
 
   Case "Entrega Diaria"
   
         'Me.Enabled = False
         FrmQueryFacturas.Show
   
   Case "Recargue"
   
         'Me.Enabled = False
         FrmRecargues.Show
         
   Case "Inventario Diario"
   
         'Me.Enabled = False
         FrmQueryInv.Show
         
   Case "Inventario Actual"
         QueryProductosInv.Show 1
         
   Case "Ventas Total x Producto"
         QueryProdxRangoFechas.Show
    
   Case "Inv. Precio Total"
         QueryProductosPrecio.Show
  
End Select
Error_No_Conexion:
End Sub


Private Sub MDIForm_Load()

Me.Show
Image1(1).Width = CR.Width
Image1(1).Height = CR.Height

FrmLogin.Show 1
'MsgBox "El Sistema ha detectado que su configuración en Memoria Es de 256MG." + vbNewLine + _
'       "Se recomienda Minimo 512MG Para un mejor desempeño", vbCritical
       
'frmShortcuts.Show
'Set lvWin.SmallIcons = i16x16
'    Set lvWin.Icons = i16x16
'lvWin.ListItems.Add(, "frmShortcuts", "@Shortcuts", 1, 1).Bold = True
End Sub

Public Sub AddToWin(ByVal srcDName As String, ByVal srcFormName As String)
    On Error Resume Next
    Dim xItem As ListItem
    
    Set xItem = lvWin.ListItems.Add(, srcFormName, srcDName, 1, 1)
    xItem.ToolTipText = srcDName
    xItem.SubItems(1) = "***" & srcDName & "***"
    xItem.Selected = True
    
    Set xItem = Nothing
End Sub

