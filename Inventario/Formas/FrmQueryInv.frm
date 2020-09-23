VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQueryInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Diaria de Inventario"
   ClientHeight    =   10035
   ClientLeft      =   3750
   ClientTop       =   4140
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   14580
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.OsenXPButton CmdCerrar 
      Height          =   375
      Left            =   11490
      TabIndex        =   4
      Top             =   8850
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cerrar Dia"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      MICON           =   "FrmQueryInv.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.OsenXPButton CmdQuery 
      Height          =   375
      Left            =   11490
      TabIndex        =   2
      Top             =   8370
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Consultar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      MICON           =   "FrmQueryInv.frx":001C
      PICN            =   "FrmQueryInv.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.LynxGrid3 GrillaInvDiario 
      Height          =   7905
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13944
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
   Begin MSComctlLib.ImageList ilList 
      Left            =   9330
      Top             =   0
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
            Picture         =   "FrmQueryInv.frx":048A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQueryInv.frx":0A24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Proyecto1.LynxGrid3 GrillaTotales 
      Height          =   1605
      Left            =   90
      TabIndex        =   1
      Top             =   8340
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   2831
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
   Begin Proyecto1.OsenXPButton CmdSalir 
      Height          =   375
      Left            =   11490
      TabIndex        =   3
      Top             =   9330
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      MICON           =   "FrmQueryInv.frx":1436
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
      Caption         =   "SubTotales por Grupos"
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
      Left            =   330
      TabIndex        =   5
      Top             =   8100
      Width           =   3585
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   90
      Top             =   8070
      Width           =   14445
   End
End
Attribute VB_Name = "FrmQueryInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vRs As New ADODB.Recordset
Dim vRsProductos As New ADODB.Recordset

Private Sub CmdCerrar_Click()
Dim vRsFacturas As New Recordset
Dim vRsInv As New Recordset
Dim sSql As String
Dim Respuesta As VbMsgBoxResult
Dim kk As Long
sSql = "Select * FROM TBL_Factura Where Estado = 'N'"
If ConnectRS(PrimeData, vRsFacturas, sSql) = False Then
   MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
End If
If vRsFacturas.RecordCount >= 1 Then
   MsgBox "Hay facturas en Tramite. Tramite las Facturas Antes de hacer cierre del Día", vbCritical
   Exit Sub
End If
'MsgBox "OK"
PrimeData.BeginTrans
sSql = ""
Respuesta = MsgBox("Esta Seguro que Desea hacer el Cierre Diario ?. Esta Operación es Irreversible." + vbNewLine + "Los Saldos se Transladarán al día siguiente", vbYesNo + vbCritical)
If Respuesta = vbYes Then
   'MsgBox "Uy datos Borrados."
   sSql = "Delete From TBL_InventarioDiario"
   If ConnectRS(PrimeData, vRsFacturas, sSql) = False Then
      MsgBox Me.Name & "," & "Borrando Todos los Datos de Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
   End If
   sSql = ""
   For kk = 0 To Me.GrillaInvDiario.RowCount - 1
       sSql = "Insert into TBL_InventarioDiario(Cod_Producto,ES) Values('" & Me.GrillaInvDiario.CellText(kk, 2) & "', '" & Val(Me.GrillaInvDiario.CellText(kk, 14)) & "')"
       'MsgBox sSql
       If ConnectRS(PrimeData, vRsInv, sSql) = False Then
          MsgBox Me.Name & "," & "Ingresando Saldos Nuevos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
          GoTo RAE
       End If
   Next kk
   
End If

PrimeData.CommitTrans
MsgBox "Operación Realizada Correctamente", vbInformation
Unload Me

'ssql = "Delete From Rep_Inventariodiario"
'If ConnectRS(PrimeData, vRsFacturas, ssql) = False Then
'   MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
'   GoTo RAE
'End If
'ssql = "INSERT INTO Rep_Inventariodiario Select * from vRsProductos"


RAE:
Set vRsFacturas = Nothing

End Sub

Private Sub CmdQuery_Click()
Dim vRs As New ADODB.Recordset
Dim vRsProductos As New ADODB.Recordset
Dim i As Long
Dim il As Long
Dim TotalEntradas As Long
Dim TotalSalidas As Long
Dim SaldoTotal As Long


Dim sSql As String
sSql = "Select * from TBL_Grupo"
If ConnectRS(PrimeData, vRs, sSql) = False Then
       MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
       'GoTo RAE
End If
If vRs.RecordCount = 0 Then Exit Sub

j = 0
Dim Codigo_G As String
vRs.MoveFirst
GrillaInvDiario.Redraw = False
Me.GrillaTotales.Redraw = False
Me.GrillaInvDiario.Clear
Me.GrillaTotales.Clear
While Not vRs.EOF
        GrillaInvDiario.AddItem ""
        Me.GrillaInvDiario.CellText(Val(j), 0) = (vRs.Fields("ID"))
        Me.GrillaInvDiario.CellText(Val(j), 1) = vRs.Fields("Descripcion")
        Codigo_G = vRs.Fields("ID")
        Descripcion_Grupo = vRs.Fields("Descripcion")
        
        sSql = ""
        sSql = "Select TBL_Producto.Codigo_producto, TBL_Producto.Descripcion_producto,TBL_Producto.Cod_Grupo,"
        sSql = sSql + " TBL_InventarioDiario.* FROM TBL_Producto,TBL_InventarioDiario "
        sSql = sSql + " Where TBL_Producto.Cod_Grupo = " & Codigo_G & ""
        sSql = sSql + " AND TBL_Producto.Codigo_Producto = TBL_InventarioDiario.Cod_Producto order by TBL_Producto.Codigo_producto"
       ' MsgBox sSql
        
        If ConnectRS(PrimeData, vRsProductos, sSql) = False Then
           MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
           'GoTo RAE
        End If
        Reccount = vRsProductos.RecordCount
        If vRsProductos.RecordCount = 0 Then Exit Sub
           i = 0
            While Not vRsProductos.EOF
               'MsgBox "OK"
               
               Me.GrillaInvDiario.ItemImage(Val(j)) = 1
               GrillaInvDiario.AddItem ""
               Me.GrillaInvDiario.CellText(Val(j), 2) = (vRsProductos.Fields("Codigo_Producto"))
               Me.GrillaInvDiario.CellText(Val(j), 3) = vRsProductos.Fields("Descripcion_Producto")
               Me.GrillaInvDiario.CellText(Val(j), 4) = vRsProductos.Fields("ES")
               Me.GrillaInvDiario.CellText(Val(j), 5) = vRsProductos.Fields("EC")
               Me.GrillaInvDiario.CellText(Val(j), 6) = vRsProductos.Fields("ED")
               
               
               TotalEntradas = vRsProductos.Fields("ES") + vRsProductos.Fields("EC") + _
                                                           vRsProductos.Fields("ED")
                                                           
               Me.GrillaInvDiario.CellText(Val(j), 7) = vRsProductos.Fields("V1") 'Salida Busser1
               Me.GrillaInvDiario.CellText(Val(j), 8) = vRsProductos.Fields("V2") 'Salida Busser2
               Me.GrillaInvDiario.CellText(Val(j), 9) = vRsProductos.Fields("V3") 'SALIDA VENDEDOR SERGIO
               Me.GrillaInvDiario.CellText(Val(j), 10) = vRsProductos.Fields("V4") 'SALIDA VENDEDOR RAFAEL
               Me.GrillaInvDiario.CellText(Val(j), 11) = vRsProductos.Fields("V5") 'SALIDA VENDEDOR FERNANDO
               Me.GrillaInvDiario.CellText(Val(j), 12) = vRsProductos.Fields("SC") 'SALIDA POR CAMBIO
               Me.GrillaInvDiario.CellText(Val(j), 13) = vRsProductos.Fields("SO") 'SALIDA POR OBSEQUIO
               
               TotalSalidas = vRsProductos.Fields("V1") + vRsProductos.Fields("V2") + vRsProductos.Fields("V3") + _
                              vRsProductos.Fields("V4") + vRsProductos.Fields("V5") + vRsProductos.Fields("SC") + _
                              vRsProductos.Fields("SO")
               SaldoTotal = TotalEntradas - TotalSalidas
               Me.GrillaInvDiario.CellText(Val(j), 14) = SaldoTotal
               Me.GrillaInvDiario.CellFontBold(Val(j), 14) = True
               
               SumaES = SumaES + vRsProductos.Fields("ES")
               SumaEC = SumaEC + vRsProductos.Fields("EC")
               SumaED = SumaED + vRsProductos.Fields("ED")
               SumaS1 = SumaS1 + vRsProductos.Fields("V1")
               SumaS2 = SumaS2 + vRsProductos.Fields("V2")
               SumaS3 = SumaS3 + vRsProductos.Fields("V3")
               SumaS4 = SumaS4 + vRsProductos.Fields("V4")
               SumaS5 = SumaS5 + vRsProductos.Fields("V5")
               SumaSC = SumaSC + vRsProductos.Fields("SC")
               SumaSO = SumaSO + vRsProductos.Fields("SO")
               
               
               j = j + 1
               vRsProductos.MoveNext
               
            Wend
                il = Me.GrillaTotales.AddItem
                Me.GrillaTotales.CellText(il, 0) = Descripcion_Grupo
                Me.GrillaTotales.CellText(il, 1) = SumaES
                Me.GrillaTotales.CellText(il, 2) = SumaEC
                Me.GrillaTotales.CellText(il, 3) = SumaED
                Me.GrillaTotales.CellText(il, 4) = SumaS1
                Me.GrillaTotales.CellText(il, 5) = SumaS2
                Me.GrillaTotales.CellText(il, 6) = SumaS3
                Me.GrillaTotales.CellText(il, 7) = SumaS4
                Me.GrillaTotales.CellText(il, 8) = SumaS5
                Me.GrillaTotales.CellText(il, 9) = SumaSC
                Me.GrillaTotales.CellText(il, 10) = SumaSO
                TotalSumaEntradas = SumaES + SumaEC + SumaED
                TotalSumaSalidas = SumaS1 + SumaS2 + SumaS3 + SumaS4 + SumaS5 + SumaSC + SumaSO
                Me.GrillaTotales.CellText(il, 11) = TotalSumaEntradas - TotalSumaSalidas
                Me.GrillaTotales.CellFontBold(il, 11) = True
                
                
                
        vRs.MoveNext
        
        SumaES = 0
        SumaEC = 0
        SumaED = 0
        SumaS1 = 0
        SumaS2 = 0
        SumaS3 = 0
        SumaS4 = 0
        SumaS5 = 0
        SumaSC = 0
        SumaSO = 0
Wend



' Falta hacer Aqui hago una copia de Seguridad de los datos del Dia actual


'ssql = "Delete From Rep_InventarioDiario"
'If ConnectRS(PrimeData, vRs, ssql) = False Then
'    MsgBox Me.Name & "," & "Borrando Todos los Datos de Inventario Diario" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
'   GoTo RAE
'End If

'

'ssql = "INSERT INTO Rep_InventarioDiario Select * from TBL_InventarioDiario"
'ssql = "INSERT INTO  TBL_InventarioDiario Select * from Rep_InventarioDiario"
'If ConnectRS(PrimeData, vRsProductos, ssql) = False Then
'   MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & ssql & "'", vbExclamation
'   GoTo RAE
'End If

For kk = 0 To Me.GrillaInvDiario.RowCount - 1
    If Me.GrillaInvDiario.CellText(Val(kk), 3) = "" Then
       Me.GrillaInvDiario.RemoveItem (Val(kk))
    End If
Next kk

sSql = ""
sSql = "Select TBL_Producto.Codigo_producto, TBL_Producto.Descripcion_producto,TBL_Producto.Cod_Grupo,TBL_InventarioDiario.*,"
sSql = sSql + " ((TBL_InventarioDiario.ES + TBL_InventarioDiario.EC + TBL_InventarioDiario.ED) -"
sSql = sSql + " (TBL_InventarioDiario.V1 + TBL_InventarioDiario.V2 + TBL_InventarioDiario.V3 +"
sSql = sSql + "  TBL_InventarioDiario.V4 + TBL_InventarioDiario.V5 + TBL_InventarioDiario.SC +"
sSql = sSql + "  TBL_InventarioDiario.SO)) AS TOTAL"
sSql = sSql + " FROM TBL_Producto,TBL_InventarioDiario"
sSql = sSql + " WHERE TBL_Producto.Codigo_Producto = TBL_InventarioDiario.Cod_Producto Order By  TBL_Producto.Codigo_producto"
If ConnectRS(PrimeData, vRsProductos, sSql) = False Then
   MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
   GoTo RAE
   End If
' = 0 Then Exit Sub
Me.CmdCerrar.Enabled = IIf(vRsProductos.RecordCount = 0, False, True)
Set Inventario.DataSource = vRsProductos
Inventario.Show 1

    
RAE:
    Set vRs = Nothing
    GrillaInvDiario.Redraw = True
    Me.GrillaTotales.Redraw = True
    GrillaInvDiario.Refresh
    Me.GrillaTotales.Refresh
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
With GrillaInvDiario
    
        .Redraw = False
        .AddColumn "C.Grupo", 50        '0
        .AddColumn "D.Grupo", 80 '1
        .AddColumn "C. Prod.", 60  '2
        .AddColumn "Descr. Prod.", 250   '3
        .AddColumn "ES", 40   '3
        .AddColumn "EC", 40   '3
        .AddColumn "Dev", 40 '4
        .AddColumn "Busser 1", 47, lgAlignCenterCenter   '5
        .AddColumn "Busser 2", 47, lgAlignCenterCenter         '6
        .AddColumn "SER", 40, lgAlignCenterCenter         '7
        .AddColumn "RAF", 40, lgAlignCenterCenter         '8
        .AddColumn "FER", 40, lgAlignCenterCenter         '9
        .AddColumn "SC", 40, lgAlignCenterCenter         '8
        .AddColumn "SO", 40, lgAlignCenterCenter         '9
        .AddColumn "Total Inv.", 85, lgAlignCenterCenter         '8
        .AddColumn "", 0
        .RowHeightMin = 21
        .ImageList = ilList
        
        .Redraw = True
        .Refresh
    End With
    
    With Me.GrillaTotales
    
        .Redraw = False
        .AddColumn "Grupo", 200        '0
        .AddColumn "ES", 50   '3
        .AddColumn "EC", 50   '3
        .AddColumn "ED", 50 '4
        .AddColumn "Busser 1", 50, lgAlignCenterCenter   '5
        .AddColumn "Busser 2", 50, lgAlignCenterCenter         '6
        .AddColumn "03", 40, lgAlignCenterCenter         '7
        .AddColumn "04", 40, lgAlignCenterCenter         '8
        .AddColumn "05", 40, lgAlignCenterCenter         '9
        .AddColumn "SC", 40, lgAlignCenterCenter         '8
        .AddColumn "SO", 40, lgAlignCenterCenter         '9
        .AddColumn "Total Inv.", 85, lgAlignCenterCenter         '8
        .AddColumn "", 0
        .RowHeightMin = 21
        .ImageList = ilList
        
        .Redraw = True
        .Refresh
    End With
End Sub

