VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form ArhivLic 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7176
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11976
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   Picture         =   "ArhivLic.frx":0000
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   998
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      DownPicture     =   "ArhivLic.frx":391E
      Height          =   255
      Left            =   11640
      Picture         =   "ArhivLic.frx":3DCE
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Picture         =   "ArhivLic.frx":4282
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      Picture         =   "ArhivLic.frx":45D3
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000001&
      DownPicture     =   "ArhivLic.frx":4901
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      MaskColor       =   &H00FF0000&
      Picture         =   "ArhivLic.frx":4F42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1410
   End
   Begin VSFlex8Ctl.VSFlexGrid VSA 
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   11535
      _cx             =   20346
      _cy             =   4683
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483625
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483628
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
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ArhivLic.frx":5275
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   7
      MergeCompare    =   3
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
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Коментарий"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5400
      Width           =   8535
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   5640
      Width           =   8535
   End
   Begin VB.Line Line8 
      X1              =   584
      X2              =   24
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Line Line7 
      X1              =   24
      X2              =   24
      Y1              =   288
      Y2              =   352
   End
   Begin VB.Line Line6 
      X1              =   152
      X2              =   152
      Y1              =   288
      Y2              =   352
   End
   Begin VB.Line Line5 
      X1              =   280
      X2              =   280
      Y1              =   288
      Y2              =   352
   End
   Begin VB.Line Line4 
      X1              =   424
      X2              =   424
      Y1              =   288
      Y2              =   352
   End
   Begin VB.Line Line3 
      X1              =   584
      X2              =   584
      Y1              =   288
      Y2              =   352
   End
   Begin VB.Line Line2 
      X1              =   24
      X2              =   584
      Y1              =   352
      Y2              =   352
   End
   Begin VB.Line Line1 
      X1              =   24
      X2              =   584
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Сальдо на конец периода"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Оплата/субсидии"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Начислено"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Сальдо на начало"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "ArhivLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Lic.Command11.Visible = True
Lic.Clik = 0
Unload Me
End Sub

Private Sub Command2_Click()
Lic.Clik = Lic.Clik - 1
'MsgBox (Str(Lic.Clik))
Lic.ViewArhiv Lic.Clik
End Sub

Private Sub Command3_Click()
Lic.Clik = Lic.Clik + 1
'MsgBox (Str(Lic.Clik))
Lic.ViewArhiv Lic.Clik
End Sub

Private Sub Command4_Click()
Command1_Click
End Sub

Private Sub Form_Load()
Lic.Clik = 1
Label11 = VSA.TextMatrix(VSA.Row, 19)
If Label11.Caption = "" Then Label11.Caption = "Коментарий отсутствует"

'For Rw = 1 To VSA.Rows - 1
'Next
End Sub

Private Sub VSA_Click()
Label11 = VSA.TextMatrix(VSA.Row, 19)
If Label11.Caption = "" Then Label11.Caption = "Коментарий отсутствует"
End Sub
