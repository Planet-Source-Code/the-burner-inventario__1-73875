VERSION 5.00
Begin VB.Form QueryProductosPrecio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Productos y Valor Total en Inventario"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Opcionet"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   3735
         TabIndex        =   3
         Top             =   240
         Width           =   3735
         Begin VB.OptionButton optbarkod 
            Caption         =   "Ordenados por Codigo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Value           =   -1  'True
            Width           =   3495
         End
         Begin VB.OptionButton optDescrip 
            Caption         =   "Ordenados por Nombre"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   3495
         End
      End
   End
   Begin Proyecto1.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   90
      TabIndex        =   0
      Top             =   2430
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   53
   End
   Begin Proyecto1.OsenXPButton OsenXPButton1 
      Height          =   585
      Left            =   2670
      TabIndex        =   1
      Top             =   2550
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1032
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
      MICON           =   "QueryProductosPrecio.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Productos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   240
      Picture         =   "QueryProductosPrecio.frx":001C
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00AE692B&
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "QueryProductosPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OsenXPButton1_Click()
InvActual
End Sub

Private Sub InvActual()

Dim VrsProductosActual As New Recordset
Dim sSql As String

If optbarkod.Value = True Then
   sSql = "Select Codigo_Producto,Descripcion_Producto,Qty,Precio_Venta, (Precio_Venta * qty) as total from TBL_Producto Order By Codigo_Producto"
Else
  sSql = "Select * from TBL_Producto Order By Descripcion_Producto"
End If

If ConnectRS(PrimeData, VrsProductosActual, sSql) = False Then
   MsgBox Me.Name & "," & "Productos Actuales" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
End If
If VrsProductosActual.RecordCount = 0 Then Exit Sub

Set RptProductosPrecio.DataSource = VrsProductosActual
RptProductosPrecio.Sections("S1").Controls("Label15").Caption = Funciones.CurrUser.USER_NAME
RptProductosPrecio.Show 1
RAE:
Set VrsProductosActual = Nothing
Unload Me

End Sub
