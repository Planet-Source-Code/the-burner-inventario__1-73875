VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form QueryProdxRangoFechas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Ventas de Productos por Fechas"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11025
   Begin Proyecto1.OsenXPButton CmdConsultar 
      Height          =   375
      Left            =   5310
      TabIndex        =   8
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "[ &Consultar ]"
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
      MICON           =   "QueryProdxRangoFechas.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSMask.MaskEdBox TxtInicial 
      Height          =   315
      Left            =   690
      TabIndex        =   0
      Top             =   1200
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Proyecto1.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   90
      TabIndex        =   4
      Top             =   1620
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   53
   End
   Begin MSMask.MaskEdBox TxtFinal 
      Height          =   315
      Left            =   3300
      TabIndex        =   1
      Top             =   1200
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Proyecto1.LynxGrid3 GrillaQueryProd 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   1740
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9763
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2670
      TabIndex        =   7
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   1260
      Width           =   525
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de Ventas de Productos por Fechas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   120
      Picture         =   "QueryProdxRangoFechas.frx":001C
      Top             =   150
      Width           =   690
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "QueryProdxRangoFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConsultar_Click()
Dim vRsRango As New Recordset
Dim Criterio As String
Dim li As Long
Me.GrillaQueryProd.Redraw = False
Me.GrillaQueryProd.Clear
'Criterio = "Select * From TBL_Factura Where (Fecha >= #" & SQLDate(Me.TxtInicial.Text) & "#) And (Fecha <= #" & SQLDate(Me.TxtFinal.Text) & "#)"
Criterio = "SELECT A1.Codigo_Producto,  A1.Descripcion_Producto,A2.Valor_Unitario as ValorUni, SUM(A2.Venta) as Venta, Sum(A2.Valor_Unitario * A2.Venta) as Total_Venta"
Criterio = Criterio + " FROM  TBL_PRoducto A1,TBL_DetFactura A2,TBL_Factura A3"
Criterio = Criterio + " Where A1.Codigo_Producto = A2.Codigo_Producto AND "
Criterio = Criterio + " Fecha >= #" & SQLDate(Me.TxtInicial.Text) & "# And Fecha <= #" & SQLDate(Me.TxtFinal.Text) & "# AND A3.Consecutivo = A2.Consecutivo"
Criterio = Criterio + " GROUP BY A1.Codigo_Producto,A1.Descripcion_Producto,A2.Valor_Unitario"
If ConnectRS(PrimeData, vRsRango, Criterio) = False Then
   MsgBox Me.Name & "," & "Productos Axtuales" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
End If
If vRsRango.RecordCount = 0 Then MsgBox "No hay": Exit Sub
'MsgBox vRsRango.RecordCount
If vRsRango.RecordCount = 0 Then Exit Sub
vRsRango.MoveFirst
While Not vRsRango.EOF
   With Me.GrillaQueryProd
        li = .AddItem(vRsRango.Fields("Codigo_producto"))
        '.ItemImage(li) = 1
        .CellText(li, 1) = vRsRango.Fields("Descripcion_Producto")
        .CellText(li, 2) = vRsRango.Fields("Venta")
        .CellText(li, 3) = FormatNumber(vRsRango.Fields("ValorUni"), 2)
        .CellText(li, 4) = FormatNumber(vRsRango.Fields("Total_Venta"), 2)
        .CellFontBold(li, 4) = True
        End With
        vRsRango.MoveNext
Wend
Set RptQueryFact.DataSource = vRsRango
RptQueryFact.Sections("S1").Controls("Label15").Caption = Funciones.CurrUser.USER_NAME
RptQueryFact.Sections("Section4").Controls("lblInicial").Caption = Me.TxtInicial.Text
RptQueryFact.Sections("Section4").Controls("lblFinal").Caption = Me.TxtFinal.Text
RptQueryFact.Show 1
RAE:
Me.GrillaQueryProd.Redraw = True
Me.GrillaQueryProd.Refresh
Set vRsRango = Nothing
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim sDate As String
TxtInicial = Format(Date, "dd/mm/yyyy")
TxtFinal = Format(Date, "dd/mm/yyyy")
Me.TxtInicial.SetFocus
With GrillaQueryProd
        .Redraw = False
        .AddColumn "Codigo", 50        '1
        .AddColumn "Descripcion", 330 '2
        .AddColumn "Cant", 70, lgAlignRightBottom, lgNumeric '3
        .AddColumn "P. Venta", 100, lgAlignRightBottom, lgNumeric '4
        .AddColumn "Total", 140, lgAlignRightBottom, lgNumeric '5
        
        '.AddColumn "", 0
        .RowHeightMin = 21
        '.ImageList = ilList
        
        .Redraw = True
        .Refresh

End With


End Sub

