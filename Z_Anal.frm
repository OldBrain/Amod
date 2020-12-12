VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Z_Anal 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7050
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5850
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5535
      _cx             =   9763
      _cy             =   10186
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Z_Anal.frx":0000
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
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
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      Caption         =   "Сохранить"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   240
      Left            =   0
      Picture         =   "Z_Anal.frx":00FA
      ToolTipText     =   "О программе"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   960
      Picture         =   "Z_Anal.frx":063C
      Top             =   0
      Width           =   285
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "123"
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   5370
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1320
      Picture         =   "Z_Anal.frx":0D86
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   480
      Picture         =   "Z_Anal.frx":14D0
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Z_Anal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBestA As ADODB.Recordset
Dim rsAnal As ADODB.Recordset
Private Sub Form_Load()
MakeWindow Me, True
lblTitle.Caption = "Выбор аналитик" + " Сч. " + Z_Sootn.Shet

' Открываем аналитики Б4

Set rsAnal = New ADODB.Recordset
rsAnal.Open ("SELECT BestAnal.KeySh, BestAnal.vkl, BestAnal.Schet, BestAnal.Anal, BestAnal.AnslName From BestAnal WHERE (((BestAnal.KeySh)=" + Str(Z_Sootn.KeySh) + "))"), Mconn, adOpenKeyset, adLockOptimistic
' Открываем аналитики Квартплаты
Set rsBestA = New ADODB.Recordset
rsBestA.Open ("Select Schet, Name, Code from Analit WHERE Analit.Schet='" + Z_Sootn.Shet + "' ORDER BY Code"), BestConn, adOpenStatic, adLockReadOnly

Z_Sootn.Enabled = False
Me.Enabled = False
Pod.Show
Pod.Label1.Caption = "Пожалуйса подождите. Идет проверка соответствия аналитик."
Pod.ProgressBar1.Value = 1
Pod.ProgressBar1.Max = 5000






'Добавляем аналитики в квартплату если их нет
I = 0

rsBestA.MoveFirst




Do While Not rsBestA.EOF



DoEvents

If rsAnal.RecordCount > 0 Then rsAnal.MoveFirst
Do While Not rsAnal.EOF

If rsBestA("Schet") = rsAnal("Schet") And rsBestA("Code") = rsAnal("Anal") And rsAnal("KeySh") = Z_Sootn.KeySh Then
I = I + 1
Exit Do
End If
rsAnal.MoveNext
Loop


If I = 0 Then
rsAnal.AddNew
rsAnal("KeySh") = Z_Sootn.KeySh
rsAnal("Schet") = rsBestA("Schet")
rsAnal("Anal") = rsBestA("Code")
rsAnal("AnslName") = rsBestA("Name")
rsAnal("Vkl") = False
rsAnal.UpdateBatch
End If

rsBestA.MoveNext
I = 0

Pod.ProgressBar1.Value = Pod.ProgressBar1.Value + 1
Loop

Unload Pod
Z_Sootn.Enabled = True
Me.Enabled = True


Set FG.DataSource = rsAnal

End Sub

Private Sub xpcmdbutton1_Click()
Unload Me
Z_Sootn.Show

End Sub
