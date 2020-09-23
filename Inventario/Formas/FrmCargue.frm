VERSION 5.00
Begin VB.Form FrmCargue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recargue de Productos"
   ClientHeight    =   2220
   ClientLeft      =   6000
   ClientTop       =   4605
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6555
   Begin VB.TextBox TxtPrecio 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1170
      Width           =   1785
   End
   Begin VB.TextBox TxtInv 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   5610
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   330
      Width           =   885
   End
   Begin VB.TextBox TxtDescripcion 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   750
      Width           =   5355
   End
   Begin VB.TextBox TxtCodigo 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   885
   End
   Begin VB.TextBox TxtQty 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1140
      Width           =   855
   End
   Begin VB.TextBox TxtCargue 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   2820
      TabIndex        =   0
      Top             =   1140
      Width           =   885
   End
   Begin Proyecto1.OsenXPButton CmdGuardar 
      Height          =   525
      Left            =   4890
      TabIndex        =   1
      Top             =   1620
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   926
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
      MICON           =   "FrmCargue.frx":0000
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
      Height          =   195
      Index           =   6
      Left            =   3810
      TabIndex        =   13
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "En Inventario"
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
      Index           =   5
      Left            =   4260
      TabIndex        =   11
      Top             =   390
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de la Transaccion"
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
      Left            =   90
      TabIndex        =   9
      Top             =   30
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
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
      TabIndex        =   8
      Top             =   810
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   2
      Left            =   30
      TabIndex        =   7
      Top             =   420
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
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
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   1230
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargue"
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
      Index           =   4
      Left            =   2130
      TabIndex        =   5
      Top             =   1170
      Width           =   600
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   -30
      Top             =   30
      Width           =   11325
   End
End
Attribute VB_Name = "FrmCargue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cFila As Integer
Public sModo As String
Private Sub CmdGuardar_Click()
Dim li As Long
If Me.TxtCargue.Text = "" Then MsgBox "Falta la Cantidad de cargue", vbCritical: Me.TxtCargue.SetFocus: Exit Sub
  If Val(Me.TxtCargue.Text) > Val(Me.TxtInv.Text) Then
     MsgBox "La Cantidad a recargar supera la Cantidad en Inventario", vbCritical
     Me.TxtCargue.Text = 0
     Funciones.HLTxt Me.TxtCargue
     Exit Sub
  End If
  If sModo = "Agregar" Then
    dupli = FrmRecargues.GrillaInvDiario.FindItem(Me.TxtCodigo.Text, 1, lgSMEqual, False)
    If dupli >= 0 Then
        MsgBox "El producto ya se encuentra en la Lsita", vbInformation
            'Me.GrillaInvDiario.RemoveItem dupli
        Exit Sub
    End If
     li = FrmRecargues.GrillaInvDiario.AddItem("000")
     FrmRecargues.GrillaInvDiario.CellText(li, 1) = Me.TxtCodigo.Text
     FrmRecargues.GrillaInvDiario.CellText(li, 2) = Me.TxtDescripcion.Text
     FrmRecargues.GrillaInvDiario.CellText(li, 3) = 0
     FrmRecargues.GrillaInvDiario.CellText(li, 4) = Me.TxtCargue.Text
     FrmRecargues.GrillaInvDiario.CellText(li, 5) = FormatNumber(Me.TxtPrecio.Text, 2)
     FrmRecargues.GrillaInvDiario.CellText(li, 6) = Val(FrmRecargues.GrillaInvDiario.CellText(li, 3) + _
     Val(FrmRecargues.GrillaInvDiario.CellText(li, 4))) * FrmRecargues.GrillaInvDiario.CellText(li, 5)
     FrmRecargues.GrillaInvDiario.CellText(li, 6) = FormatNumber(FrmRecargues.GrillaInvDiario.CellText(li, 6))
     FrmRecargues.GrillaInvDiario.CellFontBold(li, 6) = True
     FrmRecargues.Form_CalTotales
     Unload Me
     Exit Sub
  End If
  
  FrmRecargues.GrillaInvDiario.CellText(Str(cFila), 4) = Me.TxtCargue.Text
  FrmRecargues.GrillaInvDiario.CellText(Str(cFila), 6) = Val(FrmRecargues.GrillaInvDiario.CellText(Str(cFila), 3) + _
  Val(FrmRecargues.GrillaInvDiario.CellText(Str(cFila), 4))) * FrmRecargues.GrillaInvDiario.CellText(Str(cFila), 5)
  FrmRecargues.GrillaInvDiario.CellText(Str(cFila), 6) = FormatNumber(FrmRecargues.GrillaInvDiario.CellText(Str(cFila), 6))
  FrmRecargues.Form_CalTotales
  Unload Me
  
End Sub

Private Sub TxtDev_KeyPress(KeyAscii As Integer)
'Valido que sean solo numeros.
If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8 Or KeyAscii = 13 Then
   If KeyAscii = 13 Then
      Me.CmdGuardar.SetFocus
   End If
Else
    KeyAscii = 0
End If

End Sub

Private Sub Form_Activate()
 Dim Cantidad As Long
 Dim Valor As Long
 If sModo = "Agregar" Then
  Me.TxtCodigo.Locked = False
  Exit Sub
 End If
 
 Cantidad = getValorCampo("SELECT Codigo_Producto,Qty FROM TBL_Producto WHERE Codigo_Producto = '" & Me.TxtCodigo.Text & "'", "Qty")
 Me.TxtInv.Text = Cantidad
 Me.TxtCodigo.Locked = True
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.sModo = ""
End Sub

Private Sub TxtCargue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.CmdGuardar.SetFocus
End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
Dim dupli As Long

If KeyAscii = 13 Then
    QueryProductos (Me.TxtCodigo.Text)
    Me.TxtCargue.Text = ""
    Me.TxtCargue.SetFocus
End If



End Sub

Private Sub QueryProductos(cCodProducto)
Dim vRs As New ADODB.Recordset
Dim sSql As String
 
 sSql = "Select * from TBL_Producto Where Codigo_Producto = '" & cCodProducto & "'"
 If ConnectRS(PrimeData, vRs, sSql) = False Then
       MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
       GoTo RAE
 End If
 If vRs.RecordCount = 0 Then MsgBox "Producto no Existe", vbInformation:  Exit Sub
      If vRs.Fields("Qty") = 0 Then
         MsgBox "El Producto : " + vRs.Fields("Descripcion_producto") + " No tiene Saldo En Inventario"
         Exit Sub
    End If
    Me.TxtDescripcion.Text = vRs.Fields("Descripcion_Producto")
    Me.TxtInv.Text = vRs.Fields("Qty")
    Me.TxtPrecio.Text = FormatNumber(vRs.Fields("Precio_Venta"))
    
RAE:
Set vRs = Nothing
End Sub

