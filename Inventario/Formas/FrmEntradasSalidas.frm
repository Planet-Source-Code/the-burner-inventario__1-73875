VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmEntradas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entradas y Salidas de Productos"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10215
   Begin Proyecto1.OsenXPButton CmdHelp 
      Height          =   675
      Left            =   9300
      TabIndex        =   33
      Top             =   60
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&Buscar"
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
      MICON           =   "FrmEntradasSalidas.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.OsenXPButton OsenXPButton1 
      Height          =   525
      Left            =   8550
      TabIndex        =   28
      Top             =   2250
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
      MICON           =   "FrmEntradasSalidas.frx":001C
      PICN            =   "FrmEntradasSalidas.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      TabIndex        =   25
      Top             =   2400
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2670
      TabIndex        =   24
      Top             =   7380
      Visible         =   0   'False
      Width           =   1305
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
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2430
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2730
      TabIndex        =   20
      Top             =   7770
      Visible         =   0   'False
      Width           =   1065
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
      TabIndex        =   18
      Top             =   2040
      Width           =   1935
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
      TabIndex        =   17
      Top             =   2040
      Width           =   1305
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
      TabIndex        =   16
      Top             =   2010
      Width           =   705
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
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1650
      Width           =   4095
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
      TabIndex        =   14
      Top             =   1650
      Width           =   3765
   End
   Begin VB.TextBox TxtCodProducto 
      BackColor       =   &H00BB5900&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   840
      TabIndex        =   13
      Top             =   1650
      Width           =   705
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   60
      TabIndex        =   11
      Top             =   2820
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Inv."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "E/S"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "V. Unit."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Item"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.OptionButton OptSalida 
      Caption         =   "Salidas"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.OptionButton OptEntrada 
      Caption         =   "Entradas"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dcProd 
      Height          =   315
      Left            =   1050
      TabIndex        =   19
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin Proyecto1.OsenXPButton CmdAgrega 
      Height          =   525
      Left            =   7800
      TabIndex        =   29
      ToolTipText     =   "Agregar Productos..."
      Top             =   2250
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
      MICON           =   "FrmEntradasSalidas.frx":0159
      PICN            =   "FrmEntradasSalidas.frx":0175
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
      Left            =   9330
      TabIndex        =   30
      Top             =   2250
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
      MICON           =   "FrmEntradasSalidas.frx":02F9
      PICN            =   "FrmEntradasSalidas.frx":0315
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
      BackStyle       =   0  'Transparent
      Caption         =   "Que deseas Realizar ?"
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
      Index           =   0
      Left            =   450
      TabIndex        =   0
      Top             =   570
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   30
      Picture         =   "FrmEntradasSalidas.frx":06AF
      Top             =   -30
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada y Salida de Productos"
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
      Left            =   690
      TabIndex        =   32
      Top             =   120
      Width           =   2925
   End
   Begin VB.Label TxtTotales 
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
      Left            =   7950
      TabIndex        =   31
      Top             =   6510
      Width           =   1695
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
      Left            =   8070
      TabIndex        =   27
      Top             =   810
      Width           =   1935
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
      TabIndex        =   26
      Top             =   2460
      Width           =   420
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
      TabIndex        =   23
      Top             =   2490
      Width           =   1005
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   5145
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   30
      Top             =   7050
      Width           =   10065
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   12
      Top             =   6540
      Width           =   525
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
      TabIndex        =   10
      Top             =   2100
      Width           =   360
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
      TabIndex        =   9
      Top             =   2100
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
      TabIndex        =   8
      Top             =   2100
      Width           =   645
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
      TabIndex        =   7
      Top             =   1710
      Width           =   435
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
      TabIndex        =   6
      Top             =   1710
      Width           =   495
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
      TabIndex        =   5
      Top             =   1350
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   14  'Copy Pen
      X1              =   240
      X2              =   8880
      Y1              =   1230
      Y2              =   1230
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
      Left            =   6930
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Concepto"
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
      Left            =   120
      TabIndex        =   3
      Top             =   870
      Width           =   855
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   90
      Top             =   1350
      Width           =   9975
   End
End
Attribute VB_Name = "FrmEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













Public TipoTrans As String

Private Sub CmdAgrega_Click()
'> <
Dim Suma As Long

If Me.TxtCodProducto.Text = "" Then MsgBox "El Codigo no puede Estar Vacio", vbCritical: Exit Sub
If Val(Me.TxtPrecio.Text) <= 0 Then MsgBox "El precio de Venta esta Vacio", vbCritical: Me.TxtPrecio.SetFocus: Me.TxtPrecio.Text = FormatNumber(Me.TxtPrecio.Text, 2) * -1: Exit Sub
If Val(Me.TxtQty.Text) <= 0 Then MsgBox "La Cantidad no puede ser Cero (0)", vbCritical: Me.TxtQty.SetFocus: Me.TxtQty.Text = 1: Exit Sub

For i = 1 To Me.ListView1.ListItems.Count
      If Me.ListView1.ListItems.Item(i).Text = Me.TxtCodProducto.Text Then
         MsgBox "Este Producto ya esta Ingresado...", vbInformation, "INVENTARIO"
         Exit Sub
      End If
 Next i
 
 Me.ListView1.ListItems.Add , , Me.TxtCodProducto.Text
 '
 With Me.ListView1.ListItems.Item(Me.ListView1.ListItems.Count)
      .SubItems(1) = Me.TxtDescripcion.Text
      .SubItems(2) = Me.TxtQtyInv.Text
      .SubItems(3) = Me.TxtQty.Text
      If TipoTrans = "E" Then
        .SubItems(4) = Me.TxtCosto.Text
      End If
      If TipoTrans = "S" Then
        .SubItems(4) = Me.TxtPrecio.Text
      End If
      .SubItems(5) = Me.TxtTotal.Text
      .SubItems(6) = Me.ListView1.ListItems.Count
 End With
 LimpiarTexto Me
 
 For i = 1 To Me.ListView1.ListItems.Count
      Suma = Suma + Me.ListView1.ListItems.Item(i).SubItems(5)
 Next i
 
 Me.TxtTotales.Caption = FormatNumber(Suma, 2)
 Me.TxtCodProducto.SetFocus
 

End Sub


Private Sub CmdGuardar_Click()
Dim vRs As New Recordset
Dim vRsProductos As New Recordset
Dim RsConsecutivo As New Recordset
Dim vRsKardex As New Recordset
Dim Fechas As String
Fechas = Date
If Me.ListView1.ListItems.Count = 0 Then
   MsgBox "No hay Item para la Transaccion", vbCritical
   Exit Sub
End If
On Error GoTo err

PrimeData.BeginTrans

For i = 1 To Me.ListView1.ListItems.Count
    ssql = "Select Cod_Producto From TBL_InventarioDiario where Cod_Producto = '" & Me.ListView1.ListItems(i).Text & "'"
    If vRs.State = adStateOpen Then vRs.Close
    vRs.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
    If vRs.RecordCount = 0 Then
       'Inserto sino Existe El Codigo de Producto y el Concepto
       ssql = "INSERT INTO TBL_InventarioDiario(Cod_Producto, " & dcProd.BoundText & ")"
       ssql = ssql + " VALUES('" & Me.ListView1.ListItems(i).Text & "',"
       ssql = ssql + " '" & Me.ListView1.ListItems.Item(i).ListSubItems(3).Text & "')"
       If vRs.State = adStateOpen Then vRs.Close
          vRs.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
          ssql = ""
    Else
       'Si Existe Actualizo
       ssql = "Update TBL_InventarioDiario SET  " & dcProd.BoundText & " = " & dcProd.BoundText & " + Val('" & Me.ListView1.ListItems.Item(i).ListSubItems(3).Text & "')"
       ssql = ssql + " Where Cod_Producto = '" & Me.ListView1.ListItems(i).Text & "'"
       'MsgBox sSql
       If vRs.State = adStateOpen Then vRs.Close
          vRs.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
          ssql = ""
    End If
    'Si el Tipo de Transaccion es una Entrada Sumo a la tabla de Productos
    If TipoTrans = "E" Then
        ssql = ""
        ssql = "Update TBL_Producto Set Qty = Qty  + " & Me.ListView1.ListItems.Item(i).ListSubItems(3).Text & ""
        ssql = ssql + " " + "Where Codigo_Producto = '" & Me.ListView1.ListItems(i).Text & "'"
        If vRsProductos.State = adStateOpen Then vRs.Close
           vRsProductos.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
    End If
    'Si el Tipo de Transaccion es una Salida Resto
    If TipoTrans = "S" Then
        ssql = ""
        ssql = "Update TBL_Producto Set Qty = Qty  - " & Me.ListView1.ListItems.Item(i).ListSubItems(3).Text & ""
        ssql = ssql + " " + "Where Codigo_Producto = '" & Me.ListView1.ListItems(i).Text & "'"
        If vRsProductos.State = adStateOpen Then vRsProductos.Close
           vRsProductos.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
    End If
    'Ingreso en kardex Fecha_movimiento, Codigo_producto,Qty,Venta_producto,total,Tipo_transaccion
    'TxtConsecutivo
    
    ssql = ""
    ssql = "Insert into TBL_Kardex(Consecutivo,Fecha_movimiento,Codigo_producto,Qty,Venta_producto,total,Tipo_transaccion)"
    ssql = ssql + " " + "VALUES('" & Me.TxtConsecutivo.Caption & "', '" & Fechas & "','" & Me.ListView1.ListItems(i).Text & "',"
    ssql = ssql + "'" & Me.ListView1.ListItems.Item(i).ListSubItems(3).Text & "',"
    ssql = ssql + "'" & Me.ListView1.ListItems.Item(i).ListSubItems(4).Text & "',"
    ssql = ssql + "'" & Me.ListView1.ListItems.Item(i).ListSubItems(5).Text & "',"
    ssql = ssql + "'" & dcProd.BoundText & "')"
    
    If vRsKardex.State = adStateOpen Then vRsKardex.Close
       vRsProductos.Open ssql, PrimeData, adOpenStatic, adLockPessimistic
    
       
    status(2).Caption = "Actualizando Registros" + " " + Me.ListView1.ListItems(i).Text
Next i
'Aca Termino
ssql = ""
ssql = "UPDATE TBL_Consecutivo SET ID = ID + 1 "
If RsConsecutivo.State = adStateOpen Then RsConsecutivo.Close
RsConsecutivo.Open ssql, PrimeData, adOpenStatic, adLockOptimistic
PrimeData.CommitTrans
MsgBox "Registros Actualizados Correctamente", vbInformation
Unload Me
err:

'MsgBox err.Description
'PrimeData.RollbackTrans
End Sub

Private Sub CmdHelp_Click()
FrmQueryAllPrpoductos.Show 1
End Sub

Private Sub Command1_Click()
MsgBox dcProd.BoundText
End Sub

Private Sub Command2_Click()
Dim vRs As New Recordset
Dim ssql As String
Dim UNO
UNO = 1
'sSql = "alter table TBL_Producto ADD COLUMN 01 text(25)"
ssql = "insert into TBL_InventarioDiario(" & UNO & ") values(10)"
 If ConnectRS(PrimeData, vRs, ssql) = False Then
       MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
 End If
RAE:
Set vRs = Nothing
End Sub

Private Sub Form_Load()
 'Call Conexion_1.Main_AfterSD

 TipoTrans = "E"
 FrmMain.AddToWin Me.Caption, Name
 bind_dc "SELECT * FROM tbl_concepto where Codigo_Concepto <> 'ES' AND Tipo = '" & TipoTrans & "'", "Descripcion", dcProd, "Codigo_Concepto", True
 Consecutivos
 
End Sub
Private Sub Consecutivos()
Dim RsConsecutivo As New Recordset
Dim ssql As String
ssql = "Select * from TBL_Consecutivo"
If RsConsecutivo.State = adStateOpen Then RsConsecutivo.Close
   RsConsecutivo.Open ssql, PrimeData, adOpenStatic, adLockOptimistic
'Me.Text1.Text = Generar(1101, "FACT-", "00000000")
Me.TxtConsecutivo.Caption = Generar(RsConsecutivo.Fields("ID"), "MOV-", "00000000")
End Sub


Private Sub QueryProductos(cCodProducto)
Dim vRs As New ADODB.Recordset
Dim ssql As String
 
 ssql = "Select * from TBL_Producto Where Codigo_Producto = '" & cCodProducto & "'"
 If ConnectRS(PrimeData, vRs, ssql) = False Then
       MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
       GoTo RAE
 End If
 If vRs.RecordCount = 0 Then Me.status(2).Caption = "No Existe el Producto": Me.TxtCodProducto.SetFocus: Me.TxtDescripcion.Text = "": Me.TxtGrupo.Text = "": Exit Sub
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



Private Sub OptEntrada_Click()
TipoTrans = "E"
bind_dc "SELECT * FROM tbl_concepto where Codigo_Concepto <> 'ES' AND Tipo = '" & TipoTrans & "'", "Descripcion", dcProd, "Codigo_Concepto", True
Me.ListView1.ListItems.Clear
End Sub

Private Sub OptSalida_Click()
TipoTrans = "S"
bind_dc "SELECT * FROM tbl_concepto where Codigo_Concepto <> 'ES' AND Tipo = '" & TipoTrans & "'", "Descripcion", dcProd, "Codigo_Concepto", True
Me.ListView1.ListItems.Clear
End Sub

Private Sub OsenXPButton1_Click()
Dim Suma1 As Long
Suma1 = 0
If Not ListView1.SelectedItem Is Nothing Then
   
       If MsgBox("Desea eliminar el Producto ?", vbQuestion + vbYesNo) = vbYes Then
          ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
          For i = 1 To Me.ListView1.ListItems.Count
              Suma1 = Suma1 + Me.ListView1.ListItems.Item(i).SubItems(5)
          Next i
          Me.TxtTotales.Caption = FormatNumber(Suma1, 2)
       End If
       
End If

End Sub

Private Sub TxtCodProducto_GotFocus()
'Me.TxtCodProducto.Text = "": Me.TxtDescripcion.Text = "": Me.TxtGrupo.Text = "": Me.TxtCodProducto.SetFocus
LimpiarTexto Me

End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   QueryProductos (Me.TxtCodProducto.Text)
   
End If

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

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TxtTotal.Text = FormatNumber(Me.TxtPrecio.Text, 2) * Val(Me.TxtQty.Text)
   Me.TxtTotal.Text = FormatNumber(Me.TxtTotal.Text, 2)
   Me.CmdAgrega.SetFocus
End If
End Sub
