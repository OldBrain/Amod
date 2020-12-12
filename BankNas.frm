VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form BankNas 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5184
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   7584
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid Fg 
      Height          =   3132
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   7212
      _cx             =   12721
      _cy             =   5524
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
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
      FormatString    =   $"BankNas.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   4
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   2
      AutoSizeMouse   =   -1  'True
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
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Выберите начисления соответствующие колонкам реестра "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6612
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   480
      Picture         =   "BankNas.frx":010B
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   720
      Picture         =   "BankNas.frx":0855
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   240
      Picture         =   "BankNas.frx":0F9F
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleHelp 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      Height          =   156
      Left            =   1440
      Picture         =   "BankNas.frx":16E9
      Top             =   0
      Width           =   156
   End
End
Attribute VB_Name = "BankNas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCombo As ADODB.Recordset
Dim rsNas As ADODB.Recordset
Dim Cl As String


Private Sub Command1_Click()
'ДисКоннект
MenuNastr.Enabled = True
Unload Me
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
rsNas.MoveFirst
Do While Not rsNas.EOF
rsCombo.MoveFirst
Do While Not rsCombo.EOF
If rsCombo("КодKategor") = "" Or rsCombo("Kod") = "" Then Exit Do
If rsCombo("Kod") = rsNas("NachCod") Then
rsNas("KatCod") = rsCombo("КодKategor")

If rsCombo("Kod") = -1 Then rsNas("KatCod") = -1

rsNas.UpdateBatch
End If
rsCombo.MoveNext
Loop
rsNas.MoveNext
Loop

End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Col = 3 Then Fg.ComboData("NachCod") = Cl
If Col = 5 Then Cancel = True
If Fg.TextMatrix(Row, 2) = "unknow" Then Cancel = True

End Sub

Private Sub Form_Load()
'Коннект
lblTitle = "Настройка импорта оплаты из реестров банка"
MakeWindow Me, True
Set rsCombo = New ADODB.Recordset
Set rsNas = New ADODB.Recordset
rsNas.Open ("SELECT BankNastr.Код, BankNastr.ReestrPole, BankNastr.NachCod, BankNastr.NachName, BankNastr.KatCod, BankNastr.KodSG, BankNastr.ExpPole FROM BankNastr"), Mconn, adOpenStatic, adLockPessimistic


rsCombo.Open ("SELECT nachisleniy.Kod, nachisleniy.КодKategor, nachisleniy.Kategor, nachisleniy.Naim, nachisleniy.Tip From Nachisleniy WHERE (((nachisleniy.Tip)='-')) OR (((nachisleniy.Tip)='s')) ORDER BY nachisleniy.КодKategor"), Mconn

Set Fg.DataSource = rsNas
'Fg.ColComboList(2) = Cl

rsCombo.MoveFirst
Do While Not rsCombo.EOF
'Fg.ColComboList(2) = Str(rsCombo("Kod")) + rsCombo("Naim") + "|"
Cl = Cl + "#" + Str(rsCombo("Kod")) + ";" + rsCombo("Naim") + "|"
rsCombo.MoveNext
Loop
Fg.ColComboList(3) = Cl
'rsNas.Close
End Sub

