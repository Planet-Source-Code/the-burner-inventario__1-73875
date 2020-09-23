VERSION 5.00
Begin VB.Form FrmRecargues 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recargue de Productos"
   ClientHeight    =   7815
   ClientLeft      =   3555
   ClientTop       =   2865
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   11415
   Begin Proyecto1.OsenXPButton OsenXPButton1 
      Height          =   615
      Left            =   10680
      TabIndex        =   21
      Top             =   150
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "OsenXPButton1"
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
      MICON           =   "FrmRecargues.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox lnTxtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   9330
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6930
      Width           =   1995
   End
   Begin VB.TextBox TxtNombre 
      BackColor       =   &H00EAFDFF&
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   30
      Width           =   5505
   End
   Begin VB.TextBox TxtZona 
      BackColor       =   &H00F4FFFF&
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
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   420
      Width           =   7275
   End
   Begin VB.TextBox TxtCodigo 
      BackColor       =   &H00F4FFFF&
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
      Left            =   750
      TabIndex        =   4
      Top             =   30
      Width           =   1725
   End
   Begin VB.TextBox TxtNrofact 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1110
      Width           =   2385
   End
   Begin VB.TextBox TxtFecha 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4170
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1110
      Width           =   1245
   End
   Begin VB.TextBox TxtValor 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6030
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1110
      Width           =   1245
   End
   Begin VB.TextBox TxtEstado 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1110
      Width           =   2745
   End
   Begin Proyecto1.LynxGrid3 GrillaInvDiario 
      Height          =   5385
      Left            =   30
      TabIndex        =   7
      Top             =   1500
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9499
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
      SBackColor1     =   0
      SBackColor2     =   0
   End
   Begin Proyecto1.OsenXPButton CmdAceptar 
      Height          =   465
      Left            =   9810
      TabIndex        =   8
      Top             =   7290
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "&Registrar"
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
      MICON           =   "FrmRecargues.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.OsenXPButton OsenXPButton2 
      Height          =   465
      Left            =   30
      TabIndex        =   22
      Top             =   6960
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
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
      MICON           =   "FrmRecargues.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.OsenXPButton CmdEliminar 
      Height          =   465
      Left            =   690
      TabIndex        =   23
      Top             =   6960
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
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
      MICON           =   "FrmRecargues.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Total"
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
      Index           =   2
      Left            =   8190
      TabIndex        =   20
      Top             =   6990
      Width           =   1095
   End
   Begin VB.Label TxtConsecutivo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F4FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   8520
      TabIndex        =   18
      Top             =   420
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Consecutivo"
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
      Index           =   1
      Left            =   8520
      TabIndex        =   17
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label LblID 
      AutoSize        =   -1  'True
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8220
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Zona"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   14
      Top             =   90
      Width           =   570
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   0
      Top             =   810
      Width           =   11325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de la Factura"
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
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   810
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consecutivo"
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
      Index           =   2
      Left            =   30
      TabIndex        =   12
      Top             =   1170
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Index           =   3
      Left            =   3600
      TabIndex        =   11
      Top             =   1170
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Index           =   4
      Left            =   5520
      TabIndex        =   10
      Top             =   1170
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Index           =   5
      Left            =   7380
      TabIndex        =   9
      Top             =   1170
      Width           =   645
   End
End
Attribute VB_Name = "FrmRecargues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdEliminar_Click()
Dim iItem As Long
If Me.GrillaInvDiario.RowCount = 0 Then Exit Sub
If Me.GrillaInvDiario.CellText(Me.GrillaInvDiario.Row, 0) = "000" Then
    Me.GrillaInvDiario.RemoveItem Me.GrillaInvDiario.Row
End If
End Sub

Private Sub Form_Load()

With GrillaInvDiario
        .Redraw = False
        .AddColumn "ID", 50        '0
        .AddColumn "Codigo", 50        '1
        .AddColumn "Descripcion", 385 '545 '2
        .AddColumn "Cant.", 50, lgAlignRightBottom, lgNumeric '3
        .AddColumn "Cargue", 50, lgAlignRightBottom, lgNumeric '
        .AddColumn "P. Unit", 80, lgAlignRightBottom, lgNumeric '
        .AddColumn "Total", 80, lgAlignRightBottom, lgNumeric '
        
        '.CellText(li, 6) = vRs.Fields("Valor_Unitario")
        '.AddColumn "", 0
        .RowHeightMin = 21
        '.ImageList = ilList
        
        .Redraw = True
        .Refresh
    End With

Consecutivos
End Sub

Private Sub Consecutivos()
Dim RsConsecutivo As New Recordset
Dim sSql As String
sSql = "Select * from TBL_Consecutivo"
If RsConsecutivo.State = adStateOpen Then RsConsecutivo.Close
   RsConsecutivo.Open sSql, PrimeData, adOpenStatic, adLockOptimistic
'Me.Text1.Text = Generar(1101, "FACT-", "00000000")
Me.TxtConsecutivo.Caption = Generar(RsConsecutivo.Fields("ID"), "MOV-", "00000000")
End Sub


Private Sub TraerEmpleados(nCodigo As String)

Dim vRs As New ADODB.Recordset
Dim i As Long
Dim sSql As String
GrillaInvDiario.Redraw = True
GrillaInvDiario.Clear

 
 sSql = "Select * FROM TBL_Empleado Where Codigo = '" & nCodigo & "'"
 'MsgBox sSql
 If ConnectRS(PrimeData, vRs, sSql) = False Then
       MsgBox Me.Name & "," & "Empleados" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
       GoTo RAE
 End If
 If vRs.RecordCount = 0 Then
    'removerItems
    MsgBox "Este Codigo no Existe", vbInformation
    Exit Sub
 End If
 
 With Me
      .TxtNombre.Text = vRs.Fields("Nombre")
      .TxtZona.Text = vRs.Fields("Zona")
      .LblID.Caption = vRs.Fields("ID")
 End With
 TraerFacturaCliente (Me.TxtCodigo.Text)
RAE:

End Sub
Private Sub TraerFacturaCliente(cCodigoEmpleado As String)
Dim vRs As New ADODB.Recordset
Dim i As Long
Dim sSql As String
GrillaInvDiario.Redraw = False
GrillaInvDiario.Clear
sSql = "Select * From TBL_Factura Where Codigo_empleado = '" & cCodigoEmpleado & "'"
sSql = sSql + " AND Estado = 'N'"
If ConnectRS(PrimeData, vRs, sSql) = False Then
   MsgBox Me.Name & "," & "Facturas" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
End If
If vRs.RecordCount = 0 Then MsgBox "El Cliente no tiene una Prefactura en Tramite", vbInformation: Exit Sub
Me.TxtNrofact.Text = vRs.Fields("Consecutivo")
Me.TxtFecha.Text = vRs.Fields("Fecha")
Me.TxtValor.Text = FormatNumber(vRs.Fields("Valor_Total"), 2)
If vRs.Fields("Estado") = "N" Then
   Me.TxtEstado.Text = "En Tramite"
End If
TraerDetalleFactura (Me.TxtNrofact.Text)
Form_CalTotales

'ID,Consecutivo,Codigo_Producto,Descripcion_producto,Qty,Dev,Venta,Valor_Unitario

RAE:

End Sub
Public Sub Form_CalTotales()
    Dim li As Long
    Dim dTA As Double
    Me.lnTxtTotal.Text = "0.00"
    dTA = 0
    For li = 0 To Me.GrillaInvDiario.RowCount - 1
        dTA = dTA + GetTxtVal(Me.GrillaInvDiario.CellText(li, 6))
    Next
    Me.lnTxtTotal.Text = FormatNumber(dTA, 2)
End Sub


Private Sub TraerDetalleFactura(cConsecutivo As String)
Dim vRs As New ADODB.Recordset
Dim i As Long
Dim sSql As String
Dim li As Long
GrillaInvDiario.Redraw = False
GrillaInvDiario.Clear
sSql = "Select d.ID,d.Consecutivo,d.Codigo_Producto,d.Qty as Cant,d.Dev,d.Venta,d.Valor_unitario,d.Total,"
sSql = sSql + "p.Codigo_Producto as codigo, p.Descripcion_producto, f.Consecutivo"
sSql = sSql + " FROM TBL_Producto p,TBL_DetFactura d, TBL_Factura f"
sSql = sSql + " WHERE p.Codigo_Producto = d.Codigo_Producto"
sSql = sSql + " AND d.Consecutivo = f.Consecutivo "
sSql = sSql + " AND d.Consecutivo = '" & cConsecutivo & "'"
sSql = sSql + " Order by p.Codigo_Producto"
'MsgBox sSql
If ConnectRS(PrimeData, vRs, sSql) = False Then
   MsgBox Me.Name & "," & "Detalle de Factura" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
End If
If vRs.RecordCount = 0 Then Exit Sub
vRs.MoveFirst
While Not vRs.EOF
   With Me.GrillaInvDiario
        li = .AddItem(vRs.Fields("ID"))
        '.ItemImage(li) = 1
        .CellText(li, 1) = vRs.Fields("Codigo")
        .CellText(li, 2) = vRs.Fields("Descripcion_Producto")
        .CellText(li, 3) = vRs.Fields("Cant")
        .CellText(li, 4) = 0
        .CellText(li, 5) = FormatNumber(vRs.Fields("Valor_Unitario"))
        .CellText(li, 6) = FormatNumber(.CellText(li, 3) * .CellText(li, 5))
        .CellFontBold(li, 6) = True
        'If IsNull(vRs.Fields("Dev").Value) Then
        '   .CellText(li, 4) = "0"
        'End If
        
        '.CellText(li, 5) = ReadField(vRs.Fields("ProdDescription"))
        '.CellText(li, 6) = ReadField(vRs.Fields("UnitPrice"))
        '.CellText(li, 7) = ReadField(vRs.Fields("Amount"))
        End With
        
        vRs.MoveNext
Wend
RAE:
Me.GrillaInvDiario.Redraw = True
Me.GrillaInvDiario.Refresh
Set vRs = Nothing
End Sub

Private Sub GrillaInvDiario_DblClick()
If Me.GrillaInvDiario.RowCount = 0 Then Exit Sub
   With FrmCargue
        .TxtCodigo.Text = Me.GrillaInvDiario.CellText(Me.GrillaInvDiario.Row, 1)
        .TxtDescripcion = Me.GrillaInvDiario.CellText(Me.GrillaInvDiario.Row, 2)
        .TxtQty.Text = Me.GrillaInvDiario.CellText(Me.GrillaInvDiario.Row, 3)
        .TxtCargue.Text = Me.GrillaInvDiario.CellText(Me.GrillaInvDiario.Row, 4)
        .TxtPrecio.Text = Me.GrillaInvDiario.CellText(Me.GrillaInvDiario.Row, 5)
        .cFila = Me.GrillaInvDiario.Row
        .Show vbModal
   End With

End Sub

Private Sub OsenXPButton1_Click()
'Dim calc As String
'calc = Shell("calc.exe", vbMaximizedFocus)
End Sub

Private Sub OsenXPButton2_Click()
If Me.GrillaInvDiario.RowCount = 0 Then Exit Sub
   With FrmCargue
        .TxtCodigo.Text = ""
        .TxtDescripcion = ""
        .TxtQty.Text = "0"
        .TxtCargue.Text = "0"
        .TxtPrecio.Text = "0.00"
        .sModo = "Agregar"
        '.cFila = Me.GrillaInvDiario.Row
        .Show vbModal
   End With

End Sub


Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TraerEmpleados (Me.TxtCodigo.Text)
 End If
End Sub


Private Sub CmdAceptar_Click()
Dim vRs As New Recordset
Dim vRsKardex As New Recordset
Dim vRsInventarioDiario As New Recordset
Dim vRsProductos As New Recordset
Dim RsConsecutivo As New Recordset
Dim vRsDetFactura As New Recordset
Dim vRsFactura As New Recordset
Dim sSql As String
Dim Fecha As String
Dim Vendedor  As String
If Me.GrillaInvDiario.RowCount = 0 Then MsgBox "No hay registros para Recargar", vbCritical: Exit Sub

Vendedor = "V" + LTrim(Str(Me.LblID.Caption))
Fecha = Date
If Me.TxtCodigo.Text = "" Then
   MsgBox "No hay empleados para Realizar la Transaccion", vbCritical
   Exit Sub
End If
'Actualizar Inventario Diario
PrimeData.BeginTrans
For li = 0 To Me.GrillaInvDiario.RowCount - 1
    'Pregunto si devolvieron algo par ingrearlo a Kardex, Inventario diario y Producto
    If Me.GrillaInvDiario.CellText(Val(li), 4) > 0 Then
       'Update la tabla de TBL_InventarioDiario
       'MsgBox Vendedor
       sSql = "UPDATE TBL_InventarioDiario SET " & Vendedor & " = " & Vendedor & "  + " & Me.GrillaInvDiario.CellText(Val(li), 4) & ""
       sSql = sSql + " WHERE Cod_Producto = '" & Me.GrillaInvDiario.CellText(Val(li), 1) & "'"
       'MsgBox sSql
       If ConnectRS(PrimeData, vRsInventarioDiario, sSql) = False Then
          MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
          GoTo RAE
       End If
       sSql = ""
       'Update a la tabla de Productos
       sSql = "UPDATE TBL_Producto SET QTY = QTY - " & Me.GrillaInvDiario.CellText(Val(li), 4) & ""
       sSql = sSql + " WHERE Codigo_Producto = '" & Me.GrillaInvDiario.CellText(Val(li), 1) & "'"
       If ConnectRS(PrimeData, vRsProductos, sSql) = False Then
          MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
          GoTo RAE
       End If
       'Insert A la tabla de kardex
       sSql = ""
       sSql = "Insert into TBL_Kardex(Consecutivo,Fecha_movimiento,Codigo_producto,Qty,venta_producto,total,Tipo_transaccion)"
       sSql = sSql + " VALUES('" & Me.TxtConsecutivo.Caption & "', '" & Fecha & "',"
       sSql = sSql + " '" & Me.GrillaInvDiario.CellText(Val(li), 1) & "',"
       sSql = sSql + " '" & Me.GrillaInvDiario.CellText(Val(li), 4) & "',"
       sSql = sSql + " '" & Me.GrillaInvDiario.CellText(Val(li), 5) & "',"
       sSql = sSql + " '" & Me.GrillaInvDiario.CellText(Val(li), 4) * Me.GrillaInvDiario.CellText(Val(li), 5) & "'," 'VlrUn*Cantrecargue
       sSql = sSql + " 'FV')"
       If ConnectRS(PrimeData, vRsKardex, sSql) = False Then
          MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
          GoTo RAE
       End If
       'Aca le pregunto si el producto ya existe en el detalle de la factura de lo contrario lo agregue
       sSql = ""
       sSql = "Select * From TBL_Detfactura Where Consecutivo = '" & Me.TxtNrofact.Text & "'"
       sSql = sSql + " AND Codigo_Producto = '" & Me.GrillaInvDiario.CellText(Val(li), 1) & "'"
       'MsgBox sSql
       If ConnectRS(PrimeData, vRsDetFactura, sSql) = False Then
          MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
          GoTo RAE
       End If
       If vRsDetFactura.RecordCount = 0 Then
        '  MsgBox "Uno nuevo "
          sSql = "INSERT INTO TBL_DetFactura(Consecutivo,Codigo_Producto,Qty,Valor_Unitario,Total)"
          sSql = sSql + " VALUES('" & Me.TxtNrofact.Text & "', '" & Me.GrillaInvDiario.CellText(Val(li), 1) & "',"
          sSql = sSql + " '" & Me.GrillaInvDiario.CellText(Val(li), 4) & "',"
          sSql = sSql + " '" & Me.GrillaInvDiario.CellText(Val(li), 5) & "',"
          sSql = sSql + " '" & Me.GrillaInvDiario.CellText(Val(li), 4) * Me.GrillaInvDiario.CellText(Val(li), 5) & "')" 'VlrUn*Cantrecargue
        '  MsgBox sSql
          If ConnectRS(PrimeData, vRsDetFactura, sSql) = False Then
           MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
           GoTo RAE
        End If
          
       'Update TBL_DetFactura WHERE ID = Celltext(li,0)
       Else
       'MsgBox "UPDATE"
       sSql = "UPDATE TBL_DetFactura SET Qty =  " & Me.GrillaInvDiario.CellText(Val(li), 3) & " + " & Me.GrillaInvDiario.CellText(Val(li), 4) & " ,"
       sSql = sSql + "Total = '" & Me.GrillaInvDiario.CellText(Val(li), 4) * Me.GrillaInvDiario.CellText(Val(li), 5) & "'"
       sSql = sSql + " WHERE ID = " & Val(Me.GrillaInvDiario.CellText(Val(li), 0)) & ""
       
        If ConnectRS(PrimeData, vRsDetFactura, sSql) = False Then
           MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
           GoTo RAE
        End If
       End If
    End If
    
Next li

'Update el Total de la Factura ..
sSql = ""
sSql = "UPDATE TBL_Factura SET Valor_Total = '" & Me.lnTxtTotal.Text & "' WHERE Consecutivo =  '" & Me.TxtNrofact.Text & "'"
'MsgBox sSql
If ConnectRS(PrimeData, vRsFactura, sSql) = False Then
   MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
End If

sSql = ""
sSql = "UPDATE TBL_Consecutivo SET ID = ID + 1 "
If RsConsecutivo.State = adStateOpen Then RsConsecutivo.Close
RsConsecutivo.Open sSql, PrimeData, adOpenStatic, adLockOptimistic

PrimeData.CommitTrans
MsgBox "Transacción Realizada Correctamente", vbInformation
Unload Me
RAE:


End Sub





