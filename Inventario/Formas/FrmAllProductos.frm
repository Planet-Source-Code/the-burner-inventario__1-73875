VERSION 5.00
Begin VB.Form FrmAllProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualización de Productos"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   12480
   Begin Proyecto1.OsenXPButton CmdNuevo 
      Height          =   465
      Left            =   30
      TabIndex        =   5
      Top             =   900
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "&Agregar"
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
      MICON           =   "FrmAllProductos.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.LynxGrid3 GrillaProductos 
      Height          =   7155
      Left            =   2460
      TabIndex        =   0
      Top             =   900
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   11033
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
   Begin Proyecto1.LynxGrid3 GrillaGrupos 
      Height          =   6075
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   11033
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
   Begin Proyecto1.OsenXPButton OsenXPButton3 
      Height          =   465
      Left            =   0
      TabIndex        =   6
      Top             =   1410
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "&Salir"
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
      MICON           =   "FrmAllProductos.frx":001C
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
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   330
      Width           =   2745
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
      TabIndex        =   3
      Top             =   30
      Width           =   2640
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
   Begin VB.Image Image2 
      Height          =   855
      Left            =   0
      Picture         =   "FrmAllProductos.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12420
   End
End
Attribute VB_Name = "FrmAllProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SIBgColor1 = &HF4FFFF
Private Const SIBgColor2 = &HE3F9FB
Private lcGrupo As Long
Private lnGrupo As String


Private Sub CmdNuevo_Click()
With FrmActProductos
     .Modos = "Nuevo"
     .Show 1
End With
End Sub

Private Sub Form_Load()
'Codigo_producto, Descripcion_producto,Precio_Costo,Precio_Venta,Qty,observaciones
With GrillaProductos
        .Redraw = False
        .AddColumn "ID", 30        '0
        .AddColumn "Codigo", 50        '1
        .AddColumn "Descripcion", 350 '2
        .AddColumn "P. Costo", 70, lgAlignRightBottom, lgNumeric '3
        .AddColumn "P. Venta", 70, lgAlignRightBottom, lgNumeric '4
        .AddColumn "Inv.", 70, lgAlignRightBottom, lgNumeric '5
        
        '.AddColumn "", 0
        .RowHeightMin = 21
        '.ImageList = ilList
        
        .Redraw = True
        .Refresh

End With

With GrillaGrupos
        .Redraw = False
        .AddColumn "ID", 30        '0
        .AddColumn "Grupo", 125        '
        .Redraw = True
        .Refresh

End With
TraerGrupos
End Sub

Private Sub TraerGrupos()
Dim li As Long
Dim vRsGrupos As New Recordset
Dim sSql As String
sSql = "Select * from TBL_Grupo"

If ConnectRS(PrimeData, vRsGrupos, sSql) = False Then
   MsgBox Me.Name & "," & "Productos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
End If
If vRsGrupos.RecordCount = 0 Then Exit Sub
vRsGrupos.MoveFirst
While Not vRsGrupos.EOF
   With Me.GrillaGrupos
        li = .AddItem(vRsGrupos.Fields("ID"))
        '.ItemImage(li) = 1
        .CellText(li, 1) = vRsGrupos.Fields("Descripcion")
        End With
        
        vRsGrupos.MoveNext
Wend
RAE:
Me.GrillaGrupos.Redraw = True
Me.GrillaGrupos.Refresh
Set vRs = Nothing

End Sub

Public Sub TraerProductos(cID)
Dim li As Long
Dim sCat As String
Dim vRsProductos As New Recordset
Dim sSql As String
sSql = "SELECT * FROM TBL_Producto where Cod_grupo = " & cID & ""
If ConnectRS(PrimeData, vRsProductos, sSql) = False Then
   MsgBox Me.Name & "," & "Productos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
End If
If vRsProductos.RecordCount = 0 Then Exit Sub
Me.GrillaProductos.Clear
GrillaProductos.Redraw = False
vRsProductos.MoveFirst
While Not vRsProductos.EOF
   With Me.GrillaProductos
        li = .AddItem(vRsProductos.Fields("ID"))
        '.ItemImage(li) = 1
        .CellText(li, 1) = vRsProductos.Fields("Codigo_producto")
        .CellText(li, 2) = vRsProductos.Fields("Descripcion_Producto")
        .CellText(li, 3) = FormatNumber(vRsProductos.Fields("Precio_Costo"))
        .CellText(li, 4) = FormatNumber(vRsProductos.Fields("Precio_Venta"))
        .CellText(li, 5) = vRsProductos.Fields("Qty")
        .CellFontBold(li, 5) = True
        dTSRP = dTSRP + GetTxtVal(.CellText(li, 5))
        
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
Me.GrillaProductos.AddItem ""
li = GrillaProductos.AddItem("")
GrillaProductos.CellText(li, 2) = "Totales"
GrillaProductos.CellText(li, 5) = dTSRP
Me.GrillaProductos.ItemBackColor(li) = &H80C0FF
Me.GrillaProductos.CellFontBold(li, 5) = True

RAE:
Me.GrillaProductos.Redraw = True
Me.GrillaProductos.Refresh
Set vRs = Nothing

End Sub

Private Sub GrillaGrupos_DblClick()
GrillaProductos.Redraw = False
TraerProductos (Me.GrillaGrupos.CellText(GrillaGrupos.Row, 0))
GrillaProductos.Redraw = True
lcGrupo = Me.GrillaGrupos.CellText(Me.GrillaGrupos.Row, 0)
lnGrupo = Me.GrillaGrupos.CellText(Me.GrillaGrupos.Row, 1)
'MsgBox lcGrupo
GrillaProductos.Refresh

End Sub

Private Sub GrillaProductos_DblClick()
If Me.GrillaProductos.CellText(Me.GrillaProductos.Row, 0) = "" Then Exit Sub
  With FrmActProductos
       .TxtId.Text = Me.GrillaProductos.CellText(Me.GrillaProductos.Row, 0)
       .TxtCodigo.Text = Me.GrillaProductos.CellText(Me.GrillaProductos.Row, 1)
       .TxtDescripcion.Text = Me.GrillaProductos.CellText(Me.GrillaProductos.Row, 2)
       .TxtPrecio.Text = Me.GrillaProductos.CellText(Me.GrillaProductos.Row, 4)
       .TxtCosto = Me.GrillaProductos.CellText(Me.GrillaProductos.Row, 3)
       .lblID.Caption = lcGrupo
       .TxtGrupo.Text = lnGrupo
       .Modos = "Editar"
       .Show 1
  End With
End Sub


Private Sub OsenXPButton3_Click()
Unload Me
End Sub
