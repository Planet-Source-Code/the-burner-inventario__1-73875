VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrefactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PreFactura Vendedores y Bussers"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10770
   Begin VB.CheckBox Check1 
      Caption         =   "No Genera Prefactura"
      Height          =   285
      Left            =   4170
      TabIndex        =   30
      Top             =   2160
      Width           =   1875
   End
   Begin VB.TextBox TTotal 
      Alignment       =   1  'Right Justify
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
      Left            =   9300
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   6840
      Width           =   1365
   End
   Begin VB.TextBox TxtNombre 
      BackColor       =   &H00EAFDFF&
      Height          =   315
      Left            =   2580
      TabIndex        =   25
      Top             =   210
      Width           =   5505
   End
   Begin VB.TextBox TxtCodProducto 
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
      Left            =   840
      TabIndex        =   14
      Top             =   1350
      Width           =   705
   End
   Begin VB.TextBox TxtDescripcion 
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
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1350
      Width           =   3765
   End
   Begin VB.TextBox TxtGrupo 
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
      Left            =   5970
      TabIndex        =   12
      Top             =   1350
      Width           =   4095
   End
   Begin VB.TextBox TxtQty 
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
      Left            =   840
      TabIndex        =   11
      Top             =   1710
      Width           =   705
   End
   Begin VB.TextBox TxtPrecio 
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1740
      Width           =   1305
   End
   Begin VB.TextBox TxtTotal 
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
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1740
      Width           =   1935
   End
   Begin VB.TextBox TxtQtyInv 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Top             =   2130
      Width           =   705
   End
   Begin VB.TextBox TxtCosto 
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
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2100
      Width           =   1365
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
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
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
      Left            =   780
      TabIndex        =   0
      Top             =   210
      Width           =   1725
   End
   Begin Proyecto1.LynxGrid3 GrillaInvDiario 
      Height          =   4185
      Left            =   60
      TabIndex        =   5
      Top             =   2610
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   7382
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
   Begin Proyecto1.OsenXPButton CmdEliminar 
      Height          =   525
      Left            =   9210
      TabIndex        =   6
      Top             =   2040
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
      MICON           =   "FrmPrefactura.frx":0000
      PICN            =   "FrmPrefactura.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.OsenXPButton CmdAgrega 
      Height          =   525
      Left            =   8460
      TabIndex        =   15
      ToolTipText     =   "Agregar Productos..."
      Top             =   2040
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
      MICON           =   "FrmPrefactura.frx":013D
      PICN            =   "FrmPrefactura.frx":0159
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.OsenXPButton CmdGuardar 
      Height          =   525
      Left            =   9960
      TabIndex        =   16
      Top             =   2040
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
      MICON           =   "FrmPrefactura.frx":02DD
      PICN            =   "FrmPrefactura.frx":02F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   10260
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrefactura.frx":0693
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrefactura.frx":0C2D
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   8220
      TabIndex        =   29
      Top             =   300
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
      Left            =   8160
      TabIndex        =   28
      Top             =   600
      Width           =   1605
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   8550
      TabIndex        =   27
      Top             =   6900
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Productos"
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
      Left            =   360
      TabIndex        =   24
      Top             =   1050
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   23
      Top             =   1410
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   1
      Left            =   5430
      TabIndex        =   22
      Top             =   1410
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   21
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   3
      Left            =   1650
      TabIndex        =   20
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   3540
      TabIndex        =   19
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "En Inventario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2280
      TabIndex        =   18
      Top             =   2190
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   17
      Top             =   2160
      Width           =   420
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
      Left            =   10230
      TabIndex        =   4
      Top             =   330
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
      Left            =   120
      TabIndex        =   3
      Top             =   660
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
      Left            =   120
      TabIndex        =   1
      Top             =   270
      Width           =   570
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   90
      Top             =   1020
      Width           =   9975
   End
End
Attribute VB_Name = "FrmPrefactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TraerEmpleados(nCodigo As String)

Dim vRs As New ADODB.Recordset
Dim i As Long
Dim ssql As String
GrillaInvDiario.Redraw = False
GrillaInvDiario.Clear
 Rem : Rutina para saber si El Empleado tiene una factura en tramite
 ssql = "Select * from TBL_Factura Where Codigo_empleado = '" & nCodigo & "'"
 ssql = ssql + "AND Estado = 'N'"
 If ConnectRS(PrimeData, vRs, ssql) = False Then
       MsgBox Me.Name & "," & "Empleados" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
 End If
 If vRs.RecordCount >= 1 Then
    MsgBox "El Codigo tiene una Prefactura en Tramite", vbCritical
    Funciones.HLTxt Me.TxtCodigo
    Exit Sub
 End If
 Rem : Busco los Empleados
 ssql = "Select * FROM TBL_Empleado Where Codigo = '" & nCodigo & "'"
 'MsgBox sSql
 If ConnectRS(PrimeData, vRs, ssql) = False Then
       MsgBox Me.Name & "," & "Empleados" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
 End If
 If vRs.RecordCount = 0 Then MsgBox "Este Codigo no Existe", vbInformation: Exit Sub
 With Me
      .TxtNombre.Text = vRs.Fields("Nombre")
      .TxtZona.Text = vRs.Fields("Zona")
      .lblID.Caption = vRs.Fields("ID")
 End With
 Me.TxtCodProducto.SetFocus
RAE:

End Sub

Private Sub QueryGrupos(id)

Dim vRs As New ADODB.Recordset
Dim ssql As String
 
 ssql = "Select * from TBL_Grupo Where ID = " & id
 If ConnectRS(PrimeData, vRs, ssql) = False Then
       MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
 End If
 If vRs.RecordCount = 0 Then Exit Sub
 Me.TxtGrupo.Text = vRs.Fields("Descripcion")

RAE:
Set vRs = Nothing

End Sub

Private Sub QueryProductos(cCodProducto)
Dim vRs As New ADODB.Recordset
Dim ssql As String
 
 
 ssql = "Select * from TBL_Producto Where Codigo_Producto = '" & cCodProducto & "'"
 If ConnectRS(PrimeData, vRs, ssql) = False Then
       MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
 End If
 If vRs.RecordCount = 0 Then MsgBox "Producto no Existe", vbInformation: Me.TxtCodProducto.SetFocus: Me.TxtDescripcion.Text = "": Me.TxtGrupo.Text = "": Exit Sub
    If TipoTrans = "S" Then
      If vRs.Fields("Qty") = 0 Then
         MsgBox "El Producto : " + vRs.Fields("Descripcion_producto") + " No tiene Saldo En Inventario"
         Exit Sub
      End If
    End If
    
    Me.TxtDescripcion.Text = vRs.Fields("Descripcion_producto")
    Me.TxtQtyInv.Text = vRs.Fields("Qty")
    If vRs.Fields("Precio_Venta") = "" Then
       MsgBox "El Articulo no tiene precio de Venta", vbInformation
       Me.TxtPrecio.Text = 0
    Else
       Me.TxtPrecio.Text = vRs.Fields("Precio_Venta")
    End If
    
    If vRs.Fields("Precio_Costo") = "" Then
       MsgBox "El Articulo no tiene precio de Costo", vbInformation
       Me.TxtCosto.Text = 0
    Else
       Me.TxtCosto.Text = vRs.Fields("Precio_Costo")
    End If
    
    Me.TxtQty.Text = 1
    QueryGrupos (vRs.Fields("Cod_Grupo"))
    
    
    Me.TxtPrecio.Text = FormatNumber(Me.TxtPrecio.Text, 2)
    Me.TxtCosto.Text = FormatNumber(Me.TxtCosto.Text, 2)
    Me.TxtTotal.Text = FormatNumber(Me.TxtPrecio.Text, 2) * Me.TxtQty
    Me.TxtTotal.Text = FormatNumber(Me.TxtTotal.Text, 2)
    Me.TxtQty.SetFocus
    
    
RAE:
Set vRs = Nothing

End Sub


Private Sub CmdAgrega_Click()
Dim li As Long
Dim lProdID As Long
Dim dupli As Long
On Error GoTo err
    'validate
    If Not (GetTxtVal(TxtQty.Text) > 0) Then
        MsgBox "Por Favor Digite un Valor Valido", vbExclamation
        HLTxt TxtQty
        Exit Sub
    End If
        
    If Not (GetTxtVal(Me.TxtPrecio.Text) > 0) Then
        MsgBox "El Producto no tiene precio de Venta", vbExclamation
        HLTxt Me.TxtPrecio
        Exit Sub
    End If
    
    dupli = Me.GrillaInvDiario.FindItem(Me.TxtCodProducto.Text, 0, lgSMEqual, False)
    'MsgBox dupli
    If dupli >= 0 Then
        If MsgBox("El Producto ya se encuentra en la Lista" & vbNewLine & vbNewLine & _
            "Usted Desea Reemplazar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            Me.GrillaInvDiario.RemoveItem dupli
        Else
            'the answer is NO
            Exit Sub
        End If
    End If

With Me.GrillaInvDiario
        .Redraw = False
        li = .AddItem(CStr(Me.TxtCodProducto.Text))
        .ItemImage(li) = 1
        .CellText(li, 1) = Me.TxtDescripcion.Text
        .CellText(li, 2) = Val(Me.TxtQty.Text)
        .CellText(li, 3) = FormatNumber(Me.TxtPrecio.Text, 2)
        .CellText(li, 4) = FormatNumber(Me.TxtTotal.Text, 2)
        '.CellText(li, 4) = Me.TxtPrecio.Text * Me.TxtQty.Text
        
        .Redraw = True
        .Refresh
    End With
Me.TxtCodProducto.SetFocus
Me.TxtCodProducto.Text = ""
Me.TxtDescripcion.Text = ""
Me.TxtPrecio.Text = "0.00"
Me.TxtQty.Text = "0"
Me.TxtCosto.Text = "0.00"
Me.TxtGrupo.Text = ""
Me.TxtQtyInv.Text = "0"
Me.TxtTotal.Text = "0.00"

Funciones.HLTxt Me.TxtCodProducto
Call Form_CalTotalAmount
err:
'MsgBox "Error : " + err.Description
End Sub

Private Sub CmdEliminar_Click()
If Me.GrillaInvDiario.RowCount > 0 Then
        GrillaInvDiario.RemoveItem GrillaInvDiario.Row
    
        'calculate total amount
        Call Form_CalTotalAmount
    
    End If
End Sub
Private Sub Form_CalTotalAmount()

    Dim li As Long
    Dim dTA As Double
    
    'clear
    TTotal.Text = "0.00"
    
    dTA = 0
    For li = 0 To Me.GrillaInvDiario.RowCount - 1
        dTA = dTA + GetTxtVal(Me.GrillaInvDiario.CellText(li, 4))
    Next
    
    TTotal.Text = FormatNumber(dTA, 2)
    
    'If GetTxtVal(txtPayAmtOnDate.Text) < 0 Then
    '    Exit Sub
    'End If
    
    'txtSIBalance.Text = FormatNumber(GetTxtVal(txtTotalAmount.Text) - GetTxtVal(txtPayAmtOnDate.Text), 2)
    
End Sub

Private Sub Consecutivos()
Dim RsConsecutivo As New Recordset
Dim ssql As String
ssql = "Select * from TBL_Consecutivo"
If RsConsecutivo.State = adStateOpen Then RsConsecutivo.Close
   RsConsecutivo.Open ssql, PrimeData, adOpenStatic, adLockOptimistic
'Me.Text1.Text = Generar(1101, "FACT-", "00000000")
Me.TxtConsecutivo.Caption = Generar(RsConsecutivo.Fields("IDFacturas"), "FACT-", "00000000")
End Sub

Private Sub CmdGuardar_Click()
Dim vRsKardex As New Recordset
Dim vRsFactura As New Recordset
Dim vRsDetFactura As New Recordset
Dim vRsInventarioDiario As New Recordset
Dim vRsProducto As New Recordset
Dim RsConsecutivo As New Recordset
Dim ssql As String
Dim li As Long
Dim FV As String
Dim Estado As String
Dim Fecha As String
Dim Vendedor  As String
Vendedor = "V" + LTrim(Str(Me.lblID.Caption))
Fecha = Date
FV = "FV"
Estado = ""
'Estado = "N" 'No Facturado aun
'Estado = IIf(Me.Check1.Value = 1, "S", "N")
If Me.Check1.Value = 0 Then
   Estado = "N"
Else
   Estado = "S"
End If

PrimeData.BeginTrans
For li = 0 To Me.GrillaInvDiario.RowCount - 1
'Actualizo Kardex
 ssql = "INSERT INTO TBL_Kardex(Consecutivo,Fecha_movimiento,Codigo_Producto,Qty,Venta_Producto,"
 ssql = ssql + "Total,Tipo_Transaccion)"
 ssql = ssql + " Values('" & Me.TxtConsecutivo.Caption & "', '" & Fecha & "',"
 ssql = ssql + " '" & Me.GrillaInvDiario.CellText(li, 0) & "',"
 ssql = ssql + " '" & Me.GrillaInvDiario.CellText(li, 2) & "',"
 ssql = ssql + " '" & Me.GrillaInvDiario.CellText(li, 3) & "',"
 ssql = ssql + " '" & Me.GrillaInvDiario.CellText(li, 4) & "',"
 ssql = ssql + " '" & FV & "')"
 If ConnectRS(PrimeData, vRsKardex, ssql) = False Then
       MsgBox Me.Name & "," & "kardex" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
   '    GoTo RAE
 End If
 'Actualizo Detalle de Factura
 ssql = ""
 ssql = "INSERT INTO TBL_DetFactura(Consecutivo,Codigo_Producto,Qty,"
 ssql = ssql + "Valor_Unitario,Total)"
 ssql = ssql + " Values('" & Me.TxtConsecutivo.Caption & "','" & Me.GrillaInvDiario.CellText(li, 0) & "',"
 ssql = ssql + " '" & Me.GrillaInvDiario.CellText(li, 2) & "',"
 ssql = ssql + " '" & Me.GrillaInvDiario.CellText(li, 3) & "',"
 ssql = ssql + " '" & Me.GrillaInvDiario.CellText(li, 4) & "')"
 If ConnectRS(PrimeData, vRsKardex, ssql) = False Then
       MsgBox Me.Name & "," & "Detalle de Factura" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
     '  GoTo RAE
 End If
 'Actualizo Inventario Diario
 ssql = ""
 ssql = "Select Cod_Producto From TBL_InventarioDiario where Cod_Producto = '" & Me.GrillaInvDiario.CellText(li, 0) & "'"
    If vRsInventarioDiario.State = adStateOpen Then vRsInventarioDiario.Close
    vRsInventarioDiario.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
    If vRsInventarioDiario.RecordCount = 0 Then
       'Inserto sino Existe El Codigo de Producto y el Concepto
       ssql = "INSERT INTO TBL_InventarioDiario(Cod_Producto, " & Vendedor & ")"
       ssql = ssql + " VALUES('" & Me.GrillaInvDiario.CellText(li, 0) & "',"
       ssql = ssql + " '" & Me.GrillaInvDiario.CellText(li, 2) & "')"
       If ConnectRS(PrimeData, vRsInventarioDiario, ssql) = False Then
          MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
     '     GoTo RAE
       End If
    Else
       'Si Existe Actualizo
      ' MsgBox "Entro"
       
       ssql = "Update TBL_InventarioDiario SET  " & Vendedor & " = " & Vendedor & " + Val('" & Me.GrillaInvDiario.CellText(li, 2) & "')"
       ssql = ssql + " Where Cod_Producto = '" & Me.GrillaInvDiario.CellText(li, 0) & "'"
       MsgBox ssql
       If ConnectRS(PrimeData, vRsInventarioDiario, ssql) = False Then
          MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       '   GoTo RAE
       End If
    End If
 
    'Actualizo Productos
    ssql = ""
    ssql = "Update TBL_Producto Set Qty = Qty  - Val('" & Me.GrillaInvDiario.CellText(li, 2) & "')"
    ssql = ssql + " " + "Where Codigo_Producto = '" & Me.GrillaInvDiario.CellText(li, 0) & "'"
    If ConnectRS(PrimeData, vRsProducto, ssql) = False Then
          MsgBox Me.Name & "," & "Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
        '  GoTo RAE
    End If
Next

'Consecutivo,Codigo_empleado,Zona,Fecha,Valor_Total,Estado
'MsgBox Estado
    ssql = ""
    ssql = "INSERT INTO TBL_Factura(Consecutivo,Codigo_empleado,Zona,"
    ssql = ssql + "Fecha,Valor_Total,Estado)"
    ssql = ssql + " VALUES('" & Me.TxtConsecutivo.Caption & "',"
    ssql = ssql + "'" & Me.TxtCodigo.Text & "', '" & Me.TxtZona.Text & "',"
    ssql = ssql + "'" & Fecha & "', '" & Me.TTotal.Text & "', '" & Estado & "')"
    'MsgBox sSql
    If ConnectRS(PrimeData, vRsFactura, ssql) = False Then
          MsgBox Me.Name & "," & "Facturas" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
        '  GoTo RAE
    End If
    
'Actualizo el Consecutivo de las facturas
ssql = ""
ssql = "UPDATE TBL_Consecutivo SET IDFacturas = IDFacturas + 1 "
If RsConsecutivo.State = adStateOpen Then RsConsecutivo.Close
RsConsecutivo.Open ssql, PrimeData, adOpenStatic, adLockOptimistic
PrimeData.CommitTrans
MsgBox "Transacción Realizada Correctamente", vbInformation
Unload Me
'Funciones.LimpiarTexto Me
'form_refresh
'removerItems

'Consecutivos
'Me.TxtCodigo.SetFocus
'GrillaInvDiario.Redraw = False
'Me.GrillaInvDiario.Clear
'Me.Refresh
'RAE:

'PrimeData.RollbackTrans


End Sub
Private Sub form_refresh()
Me.GrillaInvDiario.Clear
End Sub
Private Sub removerItems()
Dim li As Long
    For li = 0 To Me.GrillaInvDiario.RowCount
        Me.GrillaInvDiario.RemoveItem (li)
    Next
    Me.GrillaInvDiario.SetFocus
End Sub
Private Sub Form_Load()


With GrillaInvDiario
        .Redraw = False
        .AddColumn "Codigo", 50        '0
        .AddColumn "Descripcion", 350 '1
        .AddColumn "Cant.", 100  '2
        .AddColumn "V/Unit.", 100   '3
        .AddColumn "Total", 100   '3
        .AddColumn "", 0
        .RowHeightMin = 21
        .ImageList = ilList
        
        .Redraw = True
        .Refresh
    End With
Consecutivos
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TraerEmpleados (Me.TxtCodigo.Text)
End If
End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   QueryProductos (Me.TxtCodProducto.Text)
End If

End Sub

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Val(Me.TxtQty) > Val(Me.TxtQtyInv.Text) Then
      MsgBox "la Cantidad es mayor que la que hay en Inventario", vbInformation
      HLTxt TxtQty
      Exit Sub
    End If
    Me.TxtTotal.Text = Me.TxtQty.Text * Me.TxtPrecio.Text
    Me.TxtTotal.Text = FormatNumber(Me.TxtTotal.Text)
    Me.CmdAgrega.SetFocus
End If
End Sub
