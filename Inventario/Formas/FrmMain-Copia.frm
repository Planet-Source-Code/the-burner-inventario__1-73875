VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picAdvisory 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00464646&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   10230
      TabIndex        =   5
      Top             =   7500
      Width           =   10230
      Begin VB.PictureBox picAd 
         Appearance      =   0  'Flat
         BackColor       =   &H00464646&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         Picture         =   "FrmMain-Copia.frx":0000
         ScaleHeight     =   495
         ScaleWidth      =   2895
         TabIndex        =   6
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox CR 
      Align           =   1  'Align Top
      Height          =   1185
      Left            =   0
      ScaleHeight     =   1125
      ScaleWidth      =   10170
      TabIndex        =   4
      Top             =   0
      Width           =   10230
      Begin VB.Image Image1 
         Height          =   1455
         Index           =   1
         Left            =   0
         Picture         =   "FrmMain-Copia.frx":053B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9390
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6315
      Left            =   0
      ScaleHeight     =   6315
      ScaleWidth      =   2310
      TabIndex        =   0
      Top             =   1185
      Width           =   2310
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   1
         Top             =   -75
         Width           =   2250
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Formularios Abiertos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   375
            TabIndex        =   2
            Top             =   195
            Width           =   1740
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   75
            Picture         =   "FrmMain-Copia.frx":22E1
            Top             =   150
            Width           =   240
         End
      End
      Begin MSComctlLib.ListView lvWin 
         Height          =   6390
         Left            =   0
         TabIndex        =   3
         Top             =   570
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   11271
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmMain-Copia.frx":2CE3
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Form Name"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   4230
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":39BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":43CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":4DE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":517B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":5515
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":58AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":5C49
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":665B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":706D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":7A7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":8491
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":8EA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":98B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":A2C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain-Copia.frx":A863
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvWin_Click()
If lvWin.ListItems.Count < 1 Then Exit Sub
    
    Select Case lvWin.SelectedItem.Key
        Case "frmShortcuts": frmShortcuts.Show: frmShortcuts.WindowState = vbMaximized: frmShortcuts.SetFocus
    
        'FrmEntradas
        Case "FrmEntradas": LoadForm FrmEntradas
    End Select
End Sub

Private Sub MDIForm_Load()
Me.Show
Image1(1).Width = CR.Width
Image1(1).Height = CR.Height
'frmShortcuts.Show
'Set lvWin.SmallIcons = i16x16
'    Set lvWin.Icons = i16x16
'lvWin.ListItems.Add(, "frmShortcuts", "@Shortcuts", 1, 1).Bold = True
End Sub

Public Sub AddToWin(ByVal srcDName As String, ByVal srcFormName As String)
    On Error Resume Next
    Dim xItem As ListItem
    
    Set xItem = lvWin.ListItems.Add(, srcFormName, srcDName, 1, 1)
    xItem.ToolTipText = srcDName
    xItem.SubItems(1) = "***" & srcDName & "***"
    xItem.Selected = True
    
    Set xItem = Nothing
End Sub
