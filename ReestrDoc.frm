VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ReestrDoc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Реестр документов "
   ClientHeight    =   7500
   ClientLeft      =   168
   ClientTop       =   552
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReestrDoc.frx":0000
   LinkTopic       =   "Form7"
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin KvPay.xpcmdbutton xpcmdbutton4 
      Height          =   492
      Left            =   7440
      TabIndex        =   29
      Top             =   480
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   868
      Caption         =   "Отчеты"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton6 
      Height          =   252
      Left            =   5520
      TabIndex        =   27
      Top             =   7200
      Width           =   2172
      _ExtentX        =   3831
      _ExtentY        =   445
      Caption         =   "Посмотреть"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   252
      Left            =   5520
      TabIndex        =   17
      Top             =   6840
      Width           =   2172
      _ExtentX        =   3831
      _ExtentY        =   445
      Caption         =   "Оплата"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   11775
      _cx             =   20770
      _cy             =   9128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ReestrDoc.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   0   'False
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4605
      Top             =   3735
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReestrDoc.frx":042E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReestrDoc.frx":0540
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReestrDoc.frx":0652
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   508
      ButtonWidth     =   910
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Ins"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin KvPay.xpcmdbutton xpcmdbutton3 
      Height          =   492
      Left            =   3960
      TabIndex        =   18
      Top             =   480
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   868
      Caption         =   "Проверка"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton2 
      Height          =   492
      Left            =   5040
      TabIndex        =   28
      Top             =   480
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   868
      Caption         =   "Загрузить документы"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label23 
      Caption         =   "По пл.поручениям"
      Height          =   252
      Left            =   3600
      TabIndex        =   26
      Top             =   7200
      Width           =   1812
   End
   Begin VB.Label Label22 
      Caption         =   "По периодам оплаты"
      Height          =   252
      Left            =   3600
      TabIndex        =   25
      Top             =   6840
      Width           =   1932
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Аванс"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8520
      TabIndex        =   24
      ToolTipText     =   $"ReestrDoc.frx":0764
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Текущий период"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4440
      TabIndex        =   23
      ToolTipText     =   $"ReestrDoc.frx":0853
      Top             =   1080
      Width           =   2040
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Прошлый период"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      ToolTipText     =   "Сумма платежей за прошлый период"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9360
      TabIndex        =   20
      ToolTipText     =   "Сумма платежей за будущие  периоды"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      ToolTipText     =   "Сумма платежей за текущий период"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   3840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   3840
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line7 
      X1              =   3840
      X2              =   3840
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   3840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Начисления"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Оплата"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2640
      TabIndex        =   15
      Top             =   6960
      Width           =   852
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Субсидии"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   14
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   6960
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Left            =   9240
      TabIndex        =   12
      Top             =   6960
      Width           =   75
   End
   Begin VB.Label Label9 
      Caption         =   "Начисления"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Оплата"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Субсидии"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   8880
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "НА СУММУ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Кол-во документов:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Menu Меню 
      Caption         =   "Меню"
      Begin VB.Menu Новый_документ 
         Caption         =   "Новый документ"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Удалить 
         Caption         =   "Удалить"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Выход 
         Caption         =   "Выход"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "ReestrDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_kat As ADODB.Recordset
Dim Dl As ADODB.Recordset
Dim DatT As ADODB.Recordset
'Dim mconn As ADODB.Connection
'Dim Prov As ADODB.Recordset
Dim ProvA As ADODB.Recordset
Dim SumRS As ADODB.Recordset
Public R As Integer

Private Sub BtnEnh1_1_Click()

End Sub

Private Sub BtnEnh1_Click()
BankImport.Show
End Sub

Private Sub BtnEnh2_1_Click()

End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub FG_AfterDataRefresh()
цвет
End Sub

Private Sub FG_Click()
'Label1 = FG.TextMatrix(FG.Row, 10)
'Label1.Refresh
'MsgBox (FG.TextMatrix(2, 2))

End Sub

Private Sub FG_DblClick()
R = FG.Row
ReestrDoc.Hide
'ReestrDoc.Enabled = False
Doc.Show
End Sub




Private Sub FG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then FG_DblClick
If KeyAscii = 27 Then Выход_Click
End Sub

Private Sub FG_RowColChange()
'On Error Resume Next
End Sub







Private Sub Form_Activate()
Form_Load

End Sub

Private Sub Form_Initialize()
FSize Me
FG.Width = Me.Width - 250
End Sub

Private Sub Form_Resize()
FG.Width = Me.Width - 250
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True
MainMenu.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
'    On Error Resume Next
    Select Case Button.KEY
        Case "New"
            'ToDo: Add 'New' button code.
            Новый_документ_Click
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            Удалить_Click
        Case "Save"
            'ToDo: Add 'Save' button code.
            Выход_Click
    End Select
End Sub


Private Sub Form_Load()



'Set mconn = New ADODB.Connection
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Jet OLEDB:Database Password=" + MainForm.Pas + ";"
'mconn.Open
'"data/Kvartplata.mdb"

Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
 
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockBatchOptimistic


Set Dl = New ADODB.Recordset
Set Dl.ActiveConnection = Mconn
 
Dl.CursorType = adOpenStatic
'adOpenForwardOnly
Dl.LockType = adLockBatchOptimistic



FG.AutoResize = False




'Rs_kat.Open "Reestrdoc"
rs_kat.Open ("SELECT ReestrDoc.Cod, ReestrDoc.Data, ReestrDoc.Nach, ReestrDoc.Coment, ReestrDoc.Summa, ReestrDoc.Status, ReestrDoc.Tip, ReestrDoc.NachCod, ReestrDoc.KodDom, ReestrDoc.Adres FROM ReestrDoc")
Set FG.DataSource = rs_kat
Статус
FG.ColHidden(10) = True









'FG.ColDataType(1) = flexDTBoolean

' Cвойства, свойства необходимые для сортировки
    FG.AllowUserResizing = flexResizeBoth
    FG.ExtendLastCol = True
    FG.ExplorerBar = flexExSortShowAndMove
    FG.AutoSearch = flexSearchFromCursor

'Label1 = "Адрес"
'Label1.Enabled = False
Итог
Итог1
If Arhiv = True Then xpcmdbutton2.Enabled = False


End Sub



Private Sub xpcmdbutton1_Click()

Analizlgot.Titl = "Суммы оплаты по периодам начисления.  Оплата произведена в " + MainMenu.Command13.Caption
'Str(MainForm.DR)




Analizlgot.G = 3
'Reports.sq = "SELECT Year([RealData]) AS Год, Right('0'+Trim(Str(Month([RealData])))+'.'+Str(Year([RealData])),8) AS [Оплата за], Doc.DataR AS [Дата оплаты], Doc.NameN AS Наименование, Sum(Doc.Summa) AS Сумма From Doc GROUP BY Year([RealData]), Right('0'+Trim(Str(Month([RealData])))+'.'+Str(Year([RealData])),8), Doc.DataR, Doc.NameN ORDER BY Year([RealData]), Right('0'+Trim(Str(Month([RealData])))+'.'+Str(Year([RealData])),8)"
'Reports.sq = "SELECT Right('0'+Trim(Str(Month([RealData])))+'.'+Str(Year([RealData])),8) as [За период], Sum(Doc.Summa) AS [Сумма оплаты] From Doc GROUP BY Doc.RealData"
Reports.sq = "SELECT Right('0'+Trim(Str(Month([RealData])))+'.'+Str(Year([RealData])),8) AS [За период], Sum(Doc.Summa) AS [Сумма оплаты] From Doc GROUP BY Right('0'+Trim(Str(Month([RealData])))+'.'+Str(Year([RealData])),8)"
'Unload Me
Analizlgot.Об 0

Analizlgot.fg1.MergeCells = flexMergeRestrictAll
Analizlgot.fg1.MergeCol(-1) = True
'AnalizLgot.FG1.MergeCol(FG1.Cols - 1) = False
Analizlgot.Show

End Sub

Private Sub xpcmdbutton2_Click()
Me.Hide
MenuZ.Show
End Sub

Private Sub xpcmdbutton3_Click()
Proverka.Show
End Sub

Private Sub xpcmdbutton4_Click()
Analizlgot.Titl = "Оплата на  " + Str(MainForm.DR)
Analizlgot.G = 7
Reports.sq = "SELECT [MainOccupant]![OLDNUM] AS Лиц_сч, doc.NameN AS Учлуга, doc.DataR AS Дата, doc.Summa AS Сумма, doc.PLNOM AS Пл_поручение, doc.Com AS Коментарий FROM doc INNER JOIN MainOccupant ON doc.KodKv = MainOccupant.Numer ORDER BY [MainOccupant]![OLDNUM]"
Unload Me
Analizlgot.Об 0
'Analizlgot.fg1.MergeCells = flexMergeRestrictAll
'Analizlgot.fg1.MergeCol(-1) = True
'AnalizLgot.FG1.MergeCol(FG1.Cols - 1) = False
Analizlgot.Show


End Sub

Private Sub xpcmdbutton5_Click()
BankTXTimpott.Show 1
End Sub

Private Sub xpcmdbutton6_Click()
Analizlgot.Titl = "Суммы оплаты по периодам начисления.  Оплата произведена в " + MainMenu.Command13.Caption
'Str(MainForm.DR)




Analizlgot.G = 3
'Reports.sq = "SELECT Year([RealData]) AS Год, Right('0'+Trim(Str(Month([RealData])))+'.'+Str(Year([RealData])),8) AS [Оплата за], Doc.DataR AS [Дата оплаты], Doc.NameN AS Наименование, Sum(Doc.Summa) AS Сумма From Doc GROUP BY Year([RealData]), Right('0'+Trim(Str(Month([RealData])))+'.'+Str(Year([RealData])),8), Doc.DataR, Doc.NameN ORDER BY Year([RealData]), Right('0'+Trim(Str(Month([RealData])))+'.'+Str(Year([RealData])),8)"
Reports.sq = "SELECT Doc.PLNOM AS [Номер п/п], Sum(Doc.Summa) AS [Сумма оплаты] From Doc GROUP BY Doc.PLNOM"
'Unload Me
Analizlgot.Об 0

Analizlgot.fg1.MergeCells = flexMergeRestrictAll
Analizlgot.fg1.MergeCol(-1) = True
'AnalizLgot.FG1.MergeCol(FG1.Cols - 1) = False
Analizlgot.Show

End Sub

Private Sub xpcmdbutton7_Click()
BankSocGarimpott.Show 1
End Sub

Private Sub Выход_Click()
'ReestrDoc.Hide
Unload Me
MainMenu.Enabled = True
MainMenu.Show
End Sub

Private Sub Новый_документ_Click()
If Arhiv = True Then Exit Sub

Me.Enabled = False
DocShapka.Show
End Sub

Private Sub Удалить_Click()
If Arhiv = True Then Exit Sub

Dim DelItem As String
With rs_kat
DelItem = FG.TextMatrix(FG.Row, 1)
If MsgBox("Вы хотите удалить " + FG.TextMatrix(FG.Row, 2) + " " + FG.TextMatrix(FG.Row, 3) + " " + FG.TextMatrix(FG.Row, 4) + "? Из лицевых счетов будут удалены все начисления находящиеся в этом документе!", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''

kd = FG.TextMatrix(FG.Row, 1)
'MsgBox (kd)

Dl.Open ("SELECT Doc.Cod, Doc.Key ,Doc.KodKv ,Doc.NameKv From Doc WHERE (((Doc.Cod)=" + kd + "))")




nkl = 1
If Not Dl.EOF Then Dl.MoveFirst
Do While Not Dl.EOF
nkl = nkl + 1
Dl.MoveNext
Loop

'Dim DelI(nkl)
Me.Enabled = False
MainMenu.Enabled = False



Pod.Show
Pod.ProgressBar1.Max = nkl + 1
Pod.ProgressBar1.Value = 1

On Error GoTo EndLop
Dl.MoveFirst
EndLop:

Do While Not Dl.EOF
Pod.Label1.Caption = "П О Д О Ж Д И Т Е!" + vbNewLine + "Удаляю из лиц.сч. №" + Str(Dl("Kodkv")) + " " + Dl("Namekv")
Pod.Label1.Refresh
Mconn.Execute ("DELETE Adding.*, Adding.KodDoc From Adding WHERE (((Adding.KodDoc)=" + Str(Dl.Fields("Key").Value) + "))")

Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
Dl.MoveNext
Loop

Pod.Show
Pod.ProgressBar1.Max = nkl + 1
Pod.ProgressBar1.Value = 1


If Not Dl.EOF Then Dl.MoveFirst
Do While Not Dl.EOF
Pod.Label1.Caption = "П О Д О Ж Д И Т Е, теперь" + vbNewLine + "пересчитываю лиц.сч. №" + Str(Dl("Kodkv")) + " " + Dl("Namekv")
Pod.Label1.Refresh
'Расчет сальдо и количества
MainForm.КоличествоСальдо Str(Dl("Kodkv"))
MainForm.RSaldoK Str(Dl("Kodkv"))


Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
Dl.MoveNext
Loop


Pod.ProgressBar1.Max = nkl + 1
Pod.ProgressBar1.Value = 1
On Error GoTo ermf
.MoveFirst
ermf:
Do While Not .EOF
If rs_kat("Cod") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop
Kod = FG.TextMatrix(FG.Row, 1)
Mconn.Execute ("DELETE Doc.*, Doc.Cod From Doc WHERE (((Doc.Cod)=" + Kod + "))")
'mconn.Execute ("DELETE Adding.*, Adding.KodDoc From Adding WHERE (((Adding.KodDoc)=" + Kod + "))")

.UpdateBatch
Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
FG.DataRefresh


If .EOF And FG.Rows > 1 Then .MoveLast
End If
End With
Unload Pod
Me.Enabled = True

FG.SetFocus
Dl.Close
End Sub
Public Sub Новый()
Dim n, N1 As Integer
If MsgBox("Добавить новую запись?", vbYesNo) = vbYes Then
n = 0
If Not rs_kat.EOF Then rs_kat.MoveFirst
Do While Not rs_kat.EOF
If rs_kat("Cod").Value = "" Then
rs_kat.Delete
rs_kat.MoveFirst
End If
N1 = rs_kat("Cod").Value
If N1 > n Then n = N1
rs_kat.MoveNext
Loop

rs_kat.AddNew
'Rs_kat("") = N + 1
rs_kat("Coment") = DocShapka.Text2.Text
rs_kat("Nach") = DocShapka.Combo2.Text
rs_kat("Status") = 0
rs_kat("Tip") = DocShapka.Combo1.Text
rs_kat("Data") = DocShapka.Text1
rs_kat("NachCod") = Val(DocShapka.Combo2.Text)
rs_kat("KodDom") = Val(DocShapka.Combo3.Text)
rs_kat("Adres") = DocShapka.Combo3.Text


rs_kat.UpdateBatch
FG.DataRefresh
If Not rs_kat.EOF Then rs_kat.MoveLast
End If

End Sub
Public Sub Статус()
' Пееребераем записи грида
'For Rw = 1 To FG.Rows - 1
'если статус 1 то левая колонка синяя
'If FG.TextMatrix(FG.Row, 6) = 1 Then
                'FG.Cell(flexcpFontBold, 1, 2, 2, 3) = True
 '
  '              FG.Cell(flexcpBackColor, Rw, 0) = vbBlue
   '             End If
               'FG.Cell(flexcpFontUnderline, FG.Rows - 1, 0) = True
               'FG.Cell(flexcpPicture, FG.Rows - 1, 0) = imgFolder

'Next Rw
End Sub
Private Sub цвет()
Dim rw As Integer
For rw = 1 To FG.Rows - 1
'MsgBox (FG.TextMatrix(Rw, 6))
If FG.TextMatrix(rw, 6) = "1" Then
'FG1.Cell(flexcpFontBold, Rw, 1, Rw, FG1.Cols) = True
'fg1.Cell(flexcpBackColor, Rw, 0) = vbCyan
'FG.Cell(flexcpFontBold, Rw, 6) = True
'FG.Cell(flexcpBackColor, Rw, 6, Rw, 6) = vbCyan
FG.Cell(flexcpForeColor, rw, 1, rw, 7) = vbBlue
FG.Cell(flexcpFontBold, rw, 1, rw, 7) = False
End If

'MsgBox (FG.TextMatrix(rw, 7))

'End If

'If fg1.TextMatrix(Rw, 23) = "+" Then fg1.Cell(flexcpForeColor, Rw, 18, Rw, 18) = vbBlue
'If fg1.TextMatrix(Rw, 23) = "-" Then fg1.Cell(flexcpForeColor, Rw, 18, Rw, 18) = vbRed
Next rw
'fg1.Refresh

End Sub
Private Sub Итог()
Dim s As Double
Dim Kol As Integer

s = 0
Kol = 0
For rw = 1 To FG.Rows - 1
If FG.TextMatrix(rw, 5) <> "" Then s = s + Round(FG.TextMatrix(rw, 5), 2)
Kol = Kol + 1
Next rw
Label5 = Str(Round(s, 2))
Label6 = Str(Kol)
End Sub

Private Sub Итог1()
Set SumRS = New ADODB.Recordset
Set DatT = New ADODB.Recordset
Set SumRS.ActiveConnection = Mconn

Dim prSumO As Double
Dim prSumS As Double
Dim prSumN As Double
Dim STec As Double
Dim SPro As Double
Dim SAva As Double
   
DatT.Open ("SELECT Year([TekData]) AS ГОД, Month([TekData]) AS МЕСЯЦ FROM Settings"), Mconn
   
On Error GoTo en1
SumRS.Open ("SELECT Doc.Tip, Sum(Doc.Summa) AS [Sum-Summa], Doc.RealData From Doc GROUP BY Doc.Tip, Doc.RealData")
SumRS.MoveFirst

Do While Not SumRS.EOF
If SumRS("Tip").Value = "+" Then
prSumN = prSumN + Round(SumRS("Sum-Summa").Value, 2)
Label7.Caption = Str(prSumN)
End If

If SumRS("Tip").Value = "-" Then
prSumO = prSumO + Round(SumRS("Sum-Summa").Value, 2)
xpcmdbutton1.Caption = Str(Round(prSumO, 2))

If Year(SumRS("RealData")) = DatT("ГОД") And Month(SumRS("RealData")) = DatT("МЕСЯЦ") Then STec = STec + Round(SumRS("Sum-Summa").Value, 2)

If Year(SumRS("RealData")) > DatT("ГОД") Then SAva = SAva + Round(SumRS("Sum-Summa").Value, 2)
If Year(SumRS("RealData")) = DatT("ГОД") And Month(SumRS("RealData")) > DatT("МЕСЯЦ") Then SAva = SAva + Round(SumRS("Sum-Summa").Value, 2)

If Year(SumRS("RealData")) < DatT("ГОД") Then SPro = SPro + Round(SumRS("Sum-Summa").Value, 2)
If Year(SumRS("RealData")) = DatT("ГОД") And Month(SumRS("RealData")) < DatT("МЕСЯЦ") Then SPro = SPro + Round(SumRS("Sum-Summa").Value, 2)

Label4.Caption = Str(STec)
Label17.Caption = Str(SAva)
Label18.Caption = Str(SPro)

Me.Label19.ToolTipText = "!!!ПРЕДУПРЕЖДЕНИЕ!!! Данный расчет произведен исключительно на основании данных банковских реестров, т.е. информация о периоде за который произведен платеж взята со слов плательщиков, и не может претендовать на исключительную точность!"

End If

If SumRS("Tip").Value = "s" Then
prSumS = prSumS + Round(SumRS("Sum-Summa").Value, 2)
Label3.Caption = Str(prSumS)
End If

SumRS.MoveNext
Loop
en1:

If SumRS.State = adStateClosed Then
Else
SumRS.Close
End If
DatT.Close
End Sub


