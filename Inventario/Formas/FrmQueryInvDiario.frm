VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQueryInvDiario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12495
   Begin VB.TextBox TotalEC 
      Height          =   315
      Left            =   1530
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6690
      Width           =   1095
   End
   Begin VB.TextBox TotalES 
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6690
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   7230
      TabIndex        =   4
      Top             =   5820
      Width           =   2025
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   9990
      TabIndex        =   3
      Top             =   5370
      Width           =   1605
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4500
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
   Begin Proyecto1.LynxGrid3 GrillaInvDiario 
      Height          =   4875
      Left            =   30
      TabIndex        =   1
      Top             =   360
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8599
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   1875
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   9360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmQueryInvDiario.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmQueryInvDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim vRs As New ADODB.Recordset
Dim i As Long
    GrillaInvDiario.Redraw = False
    GrillaInvDiario.Clear
Dim sSql As String
 
 sSql = "Select TBL_Producto.Codigo_producto,TBL_Producto.Descripcion_producto, TBL_Producto.Cod_Grupo,"
 sSql = sSql + "TBL_GRUPO.ID as IdGrupo,TBL_Grupo.Descripcion as DGrupo,TBL_InventarioDiario.*"
 sSql = sSql + " FROM TBL_Producto,TBL_Grupo,TBL_InventarioDiario"
 sSql = sSql + " WHERE Tbl_Producto.Cod_Grupo = Tbl_Grupo.ID"
 sSql = sSql + " AND TBL_Producto.Codigo_Producto = TBL_InventarioDiario.Cod_Producto"
 
 
 If ConnectRS(PrimeData, vRs, sSql) = False Then
       MsgBox Me.Name & "," & "Movimientos" & "," & "No se puede conectar a la BD. SQL Expresion: '" & sSql & "'", vbExclamation
       GoTo RAE
 End If
 
i = 0
Suma = 0
yFact = 1000000
vRs.MoveFirst
While Not vRs.EOF
'yFact = vRs.Fields("IDGrupo")

 '    If xFact = yFact Then
  
        GrillaInvDiario.AddItem (vRs.Fields("IDGrupo"))
        Me.GrillaInvDiario.CellText(i, 0) = vRs.Fields("IDGrupo")
        Me.GrillaInvDiario.CellText(i, 1) = vRs.Fields("DGrupo")
        Me.GrillaInvDiario.CellText(i, 2) = vRs.Fields("Codigo_Producto")
        Me.GrillaInvDiario.CellText(i, 3) = vRs.Fields("Descripcion_Producto")
        Me.GrillaInvDiario.CellText(i, 4) = vRs.Fields("ES")
        Me.GrillaInvDiario.CellText(i, 5) = vRs.Fields("EC")
        Me.GrillaInvDiario.CellText(i, 6) = vRs.Fields("ED")
        
        'Me.GrillaInvDiario.CellText(i, 6) = Val(vRs.Fields("ES") + vRs.Fields("EC"))
        i = i + 1
   '   xFact = vRs.Fields("IDGrupo").Value
       SumaES = SumaES + Val(vRs.Fields("ES"))
       SumaEC = SumaEC + Val(vRs.Fields("EC"))
       SumaED = SumaED + Val(vRs.Fields("ED"))
       
       vRs.MoveNext
  
 Wend
 
       'MsgBox i
       GrillaInvDiario.AddItem ""
       Me.GrillaInvDiario.CellText(i, 3) = "Totales:"
       Me.GrillaInvDiario.CellText(i, 4) = SumaES
       Me.GrillaInvDiario.CellText(i, 5) = SumaEC
       
       
       
RAE:
    Set vRs = Nothing
    GrillaInvDiario.Redraw = True
    GrillaInvDiario.Refresh
End Sub

Private Sub Command2_Click()
Dim itemq As String
Dim Sw As Boolean
itemq = Val(GrillaInvDiario.CellText(Val(j), 0))
NumRows = Me.GrillaInvDiario.RowCount
For j = 0 To Me.GrillaInvDiario.RowCount - 1
    Suma = 0
    For H = 4 To Me.GrillaInvDiario.Cols
        
     If Me.GrillaInvDiario.CellText(Val(j), 0) = itemq Then
       Suma = Suma + Val(GrillaInvDiario.CellText(Val(j), Val(H)))
       
    Else
       Sw = True
       Columna = H
       
    End If
        
    Next H
    If Sw = True Then
       itemq = GrillaInvDiario.CellText(Val(j), 0)
       Me.GrillaInvDiario.AddItem "Total", Val(j)
       Me.GrillaInvDiario.CellText(Val(j), Val(Columna)) = Suma
       'List1.AddItem suma
       'NumRows = NumRows + 1
       
       Suma = 0
    End If
Next j
For k = 0 To Me.GrillaInvDiario.RowCount - 1
 'MsgBox "Este es el Valor de item " + Str(itemq)
  If Val(Me.GrillaInvDiario.CellText(Val(k), 0)) = 4 Then
     total = total + Val(Me.GrillaInvDiario.CellText(Val(k), 4))
     'MsgBox k
  End If
  
Next k
Me.GrillaInvDiario.AddItem "Total", Me.GrillaInvDiario.RowCount + 1
Me.GrillaInvDiario.CellText(Val(k), 4) = total

 
'Me.GrillaInvDiario.AddItem "Total"
'Me.GrillaInvDiario.CellText(Val(j), 4) = suma
End Sub

Private Sub Form_Load()

With GrillaInvDiario
    
        .Redraw = False
        .AddColumn "C.Grupo", 50        '0
        .AddColumn "D.Grupo", 80 '1
        .AddColumn "C. Prod.", 60  '2
        .AddColumn "Descr. Prod.", 250   '3
        .AddColumn "ES", 50   '3
        .AddColumn "EC", 50   '3
        .AddColumn "ED.", 50   '3
        '.AddColumn "Dev", 40 '4
        '.AddColumn "Busser 1", 47, lgAlignCenterCenter   '5
        '.AddColumn "Busser 2", 47, lgAlignCenterCenter         '6
        '.AddColumn "03", 40, lgAlignCenterCenter         '7
        '.AddColumn "04", 40, lgAlignCenterCenter         '8
        '.AddColumn "05", 40, lgAlignCenterCenter         '9
        '.AddColumn "SC", 40, lgAlignCenterCenter         '8
        '.AddColumn "SO", 40, lgAlignCenterCenter         '9
        '.AddColumn "Total Inv.", 85, lgAlignCenterCenter         '8
        '.AddColumn "", 0
        .RowHeightMin = 21
        .ImageList = ilList
        
        .Redraw = True
        .Refresh
    End With
End Sub

Private Sub GrillaInvDiario_DblClick()
IDItem = Me.GrillaInvDiario.CellText(GrillaInvDiario.Row, 0)
MsgBox IDItem
End Sub
