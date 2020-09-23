VERSION 5.00
Begin VB.Form FrmTransaccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transacción de Tramite de Factura"
   ClientHeight    =   2700
   ClientLeft      =   5940
   ClientTop       =   4830
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtVenta 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1230
      Width           =   885
   End
   Begin Proyecto1.OsenXPButton CmdGuardar 
      Height          =   525
      Left            =   4950
      TabIndex        =   13
      Top             =   2100
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
      MICON           =   "FrmTransaccion.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtTotal 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1650
      Width           =   2325
   End
   Begin VB.TextBox TxtValor 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1620
      Width           =   1485
   End
   Begin VB.TextBox TxtDev 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   3210
      TabIndex        =   0
      Top             =   1230
      Width           =   885
   End
   Begin VB.TextBox TxtQty 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1230
      Width           =   855
   End
   Begin VB.TextBox TxtCodigo 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   450
      Width           =   885
   End
   Begin VB.TextBox TxtDescripcion 
      BackColor       =   &H00EAFDFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   5445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venta"
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
      Index           =   7
      Left            =   4200
      TabIndex        =   15
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   6
      Left            =   2760
      TabIndex        =   12
      Top             =   1710
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Uni."
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
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Devolución"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   1260
      Width           =   930
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
      Left            =   90
      TabIndex        =   7
      Top             =   1320
      Width           =   750
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
      Left            =   60
      TabIndex        =   5
      Top             =   510
      Width           =   570
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
      Left            =   90
      TabIndex        =   3
      Top             =   900
      Width           =   975
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   0
      Top             =   120
      Width           =   11325
   End
End
Attribute VB_Name = "FrmTransaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fila As Integer

Private Sub CmdGuardar_Click()
If Val(Me.TxtDev.Text) > Val(Me.TxtQty.Text) Then
         Me.TxtDev.Text = 0
         Funciones.HLTxt Me.TxtDev
         MsgBox "El Valor De la devolución no puede ser Mayor que la Entregada.", vbCritical
         
         Exit Sub
      End If
      Me.TxtVenta.Text = Val(Me.TxtVenta.Text)
      Me.TxtVenta.Text = Val(Me.TxtQty.Text) - Val(Me.TxtDev.Text)
      Me.TxtTotal.Text = Me.TxtVenta.Text * Me.TxtValor.Text
      Me.TxtTotal.Text = FormatNumber(Me.TxtTotal.Text)
      FrmQueryFacturas.GrillaInvDiario.CellText(Str(Fila), 4) = Val(Me.TxtDev.Text)
      FrmQueryFacturas.GrillaInvDiario.CellText(Str(Fila), 5) = Val(Me.TxtVenta.Text)
      FrmQueryFacturas.GrillaInvDiario.CellText(Str(Fila), 7) = Me.TxtTotal.Text
      FrmQueryFacturas.Form_CalTotal
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
