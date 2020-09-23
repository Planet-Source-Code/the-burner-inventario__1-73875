VERSION 5.00
Begin VB.Form FrmQueryAllPrpoductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Todos los Productos"
   ClientHeight    =   9825
   ClientLeft      =   11415
   ClientTop       =   3270
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   5040
   Begin Proyecto1.LynxGrid3 GrillaProductos 
      Height          =   8715
      Left            =   60
      TabIndex        =   0
      Top             =   990
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   15372
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Todos los Productos"
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
      TabIndex        =   3
      Top             =   60
      Width           =   1950
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
      Left            =   0
      TabIndex        =   2
      Top             =   360
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
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   4020
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   0
      Picture         =   "FrmQueryAllPrpoductos.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12420
   End
End
Attribute VB_Name = "FrmQueryAllPrpoductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SIBgColor1 = &HF4FFFF
Private Const SIBgColor2 = &HE3F9FB
Private Sub Form_Load()
With GrillaProductos
        .Redraw = False
        .AddColumn "Codigo", 50        '1
        .AddColumn "Descripcion", 250 '2
        '.AddColumn "", 0
        .RowHeightMin = 21
        '.ImageList = ilList
        
        .Redraw = True
        .Refresh

End With
TraerProductos
End Sub

Private Sub TraerProductos()
Dim li As Long
Dim sCat As String
Dim vRsProductos As New Recordset
Dim sSql As String
sSql = "SELECT * FROM TBL_Producto"
If ConnectRS(PrimeData, vRsProductos, sSql) = False Then
   MsgBox Me.Name & "," & "Productos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
End If
If vRsProductos.RecordCount = 0 Then Exit Sub
Me.GrillaProductos.Clear
vRsProductos.MoveFirst
GrillaProductos.Redraw = False
While Not vRsProductos.EOF
   With Me.GrillaProductos
        li = .AddItem(vRsProductos.Fields("Codigo_producto"))
        '.ItemImage(li) = 1
        .CellText(li, 1) = vRsProductos.Fields("Descripcion_Producto")
        If sCat <> .CellText(li, 0) Then
            'change bgcolor
            If lBgColor = SIBgColor1 Then
                lBgColor = SIBgColor2
            Else
                lBgColor = SIBgColor1
            End If
        End If
        .ItemBackColor(li) = lBgColor
        sCat = .CellText(li, 0)
        
        End With
        vRsProductos.MoveNext
Wend
GrillaProductos.Redraw = True
RAE:

Set vRsProductos = Nothing
End Sub
