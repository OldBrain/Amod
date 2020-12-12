VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Lgota 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Льготы"
   ClientHeight    =   8580
   ClientLeft      =   1788
   ClientTop       =   1632
   ClientWidth     =   11400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Lgota.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
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
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Удалить(F8)"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   9360
      TabIndex        =   17
      Text            =   "Combo5"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   9360
      TabIndex        =   16
      Text            =   "Combo4"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   5280
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   9360
      TabIndex        =   12
      Text            =   "Combo3"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   4200
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   9360
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "New (F4)"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   9360
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2040
      Width           =   1815
   End
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   6975
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   7935
      _cx             =   13996
      _cy             =   12303
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
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   255
      TreeColor       =   -2147483632
      FloodColor      =   65280
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Lgota.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      Ellipsis        =   1
      ExplorerBar     =   5
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
      ComboSearch     =   2
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
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   6000
         TabIndex        =   22
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   480
      Width           =   11175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5100
      Top             =   4050
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
            Picture         =   "Lgota.frx":04A5
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lgota.frx":05B7
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lgota.frx":06C9
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   11175
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      X1              =   11280
      X2              =   11280
      Y1              =   6000
      Y2              =   6720
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      X1              =   11280
      X2              =   10080
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      X1              =   8160
      X2              =   11280
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      X1              =   8160
      X2              =   8160
      Y1              =   6720
      Y2              =   6000
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      X1              =   9240
      X2              =   8160
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      X1              =   11280
      X2              =   11280
      Y1              =   4920
      Y2              =   5760
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      X1              =   11160
      X2              =   11280
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line18 
      BorderWidth     =   3
      X1              =   8280
      X2              =   8160
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line17 
      BorderWidth     =   3
      X1              =   8160
      X2              =   8160
      Y1              =   5760
      Y2              =   4920
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Мусор"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   9240
      TabIndex        =   19
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Коммунальные услуги"
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
      Left            =   8280
      TabIndex        =   18
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Line Line16 
      BorderWidth     =   3
      X1              =   8160
      X2              =   11280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      X1              =   11280
      X2              =   10680
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      X1              =   8160
      X2              =   9120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Отоплнние"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   9120
      TabIndex        =   13
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      X1              =   11280
      X2              =   11280
      Y1              =   3960
      Y2              =   4680
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      X1              =   11280
      X2              =   8160
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      X1              =   8160
      X2              =   8160
      Y1              =   3960
      Y2              =   4680
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00404000&
      BorderWidth     =   3
      X1              =   11280
      X2              =   11280
      Y1              =   2880
      Y2              =   3600
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   10920
      X2              =   11280
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404000&
      BorderWidth     =   3
      X1              =   11280
      X2              =   8160
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00404000&
      BorderWidth     =   3
      X1              =   8160
      X2              =   8160
      Y1              =   3600
      Y2              =   2880
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   8640
      X2              =   8160
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Техобслуживание"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Index           =   2
      Left            =   8640
      TabIndex        =   10
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   11280
      X2              =   11280
      Y1              =   1800
      Y2              =   2520
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   8160
      X2              =   11280
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   8160
      X2              =   8160
      Y1              =   2520
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   10560
      X2              =   11280
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   8160
      X2              =   9000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Index           =   1
      Left            =   9000
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Квартплата"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   0
      Left            =   8640
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Menu Меню_команд 
      Caption         =   "Menu"
      Index           =   101
      WindowList      =   -1  'True
      Begin VB.Menu New 
         Caption         =   "Новая запись"
         Index           =   102
         Shortcut        =   {F4}
      End
      Begin VB.Menu DelZap 
         Caption         =   "Удалить запись"
         Index           =   102
         Shortcut        =   {F8}
      End
      Begin VB.Menu Пусто 
         Caption         =   "Закрыть"
      End
   End
   Begin VB.Menu Инфо 
      Caption         =   "Инфо"
      Index           =   2
      Begin VB.Menu Список 
         Caption         =   "Список жильзов"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Реорганизация 
         Caption         =   "Реорганизация льгот"
      End
   End
End
Attribute VB_Name = "Lgota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim mconn As ADODB.Connection
Dim TheRS As ADODB.Recordset
Dim rsCombo As ADODB.Recordset
Dim Combo_rs1 As ADODB.Recordset
Dim rsProv As ADODB.Recordset
Dim fld$, lblvalue, lblOriginalValue
Dim SelRow As AffectEnum

Private Sub FG_Click()
Dim Cl As String
If FG.Col = 3 Then
FG.Editable = flexEDKbdMouse
If Combo_rs1.RecordCount > 0 Then Combo_rs1.MoveFirst
Do While Not Combo_rs1.EOF
'cl = cl + Combo_RS("Name_Kategor") + "|"
Cl = Cl + CStr(Combo_rs1("Tip")) & vbTab & Combo_rs1("Name") + "|"
Combo_rs1.MoveNext
Loop
FG.ComboList = Cl
Else
FG.ComboList = ""
FG.Editable = flexEDNone
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
   ' On Error Resume Next
    Select Case Button.KEY
        Case "New"
            Command2_Click
        Case "Delete"
           Command3_Click

        Case "Save"
            Command1_Click
    End Select
End Sub



Private Sub UpdateLabels()
    'On Error Resume Next
   ' Dim fld$, lblvalue, lblOriginalValue
    'fld = fg.TextMatrix(0, fg.Col)
    
    'lblvalue = TheRS(fld).Value
 '   lblvalue = TheRS(fld).Value
  '  lblOriginalValue = TheRS(fld).OriginalValue
'//////////////////////////////////////////////////
  '  If lblvalue <> lblOriginalValue Then
'fg.BackColorSel = vbMagenta
'fg.CellBackColor = vbCyan
'fg.CellFloodColor = vbCyan
'lblvalue.ForeColor = vbCyan
'Else
'fg.BackColorSel = -2147483635      ' default selection color
'lblvalue.ForeColor = vbBlack
'fg.CellBackColor = vbBlack
'End If
'/////////////////////////////////////////////////
End Sub

Private Sub Combo1_LostFocus()
TheRS.UpdateBatch
End Sub
Private Sub Combo2_LostFocus()
TheRS.UpdateBatch
End Sub
Private Sub Combo3_LostFocus()
TheRS.UpdateBatch
End Sub
Private Sub Combo4_LostFocus()
TheRS.UpdateBatch
End Sub
Private Sub Combo5_LostFocus()
TheRS.UpdateBatch
End Sub

Private Sub Command1_Click()
TheRS.UpdateBatch

Mconn.Execute ("UPDATE KLS_PRIV INNER JOIN Lgota ON KLS_PRIV.N_KLS = Lgota.Numer SET Lgota.NAME_KLS = [KLS_PRIV]![NAME_KLS], Lgota.LPKV = [KLS_PRIV]![LPKV], Lgota.LPTEH = [KLS_PRIV]![LPKV], Lgota.LPOTOPL = [KLS_PRIV]![LPOTOPL], Lgota.LPCOMM = [KLS_PRIV]![LPCOMM], Lgota.LPMUSOR = [KLS_PRIV]![LPMUSOR], Lgota.USEKV = [KLS_PRIV]![USEKV], Lgota.USETEH = [KLS_PRIV]![USETEH], Lgota.USEOTOPL = [KLS_PRIV]![USEOTOPL], Lgota.USECOMM = [KLS_PRIV]![USECOMM], Lgota.USEMUSOR = [KLS_PRIV]![USEMUSOR]")
ЗапЛьгот
Sprav.Show
Lgota.Hide
Unload Me
End Sub
Private Sub Command2_Click()
Dim n, N1 As Integer
If MsgBox("Добавить новую запись?", vbYesNo) = vbYes Then
n = 0
If TheRS.RecordCount > 0 Then TheRS.MoveFirst
Do While Not TheRS.EOF
If TheRS("N_KLS").Value = "" Then
TheRS.Delete
TheRS.MoveFirst
End If
N1 = TheRS("N_KLS").Value
If N1 > n Then n = N1
TheRS.MoveNext
Loop

'MsgBox (N + 1)
TheRS.AddNew
TheRS("N_KLS") = n + 1
TheRS("NAME_KLS") = "Новая льгота"
TheRS.UpdateBatch
TheRS.Requery
FG.DataRefresh
TheRS.MoveLast
End If
End Sub


Private Sub Command3_Click()







If MsgBox("Вы хотите удалить лготу КОД>" + FG.TextMatrix(FG.Row, 1) + vbNewLine + FG.TextMatrix(FG.Row, 2) + "?", vbYesNo) = vbYes Then

rsProv.Open ("SELECT Lgota.NomNum, MainOccupant.BanKN, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Lgota.Numer FROM Lgota INNER JOIN MainOccupant ON Lgota.NomNum = MainOccupant.Numer WHERE (((Lgota.Numer)=" + FG.TextMatrix(FG.Row, 1) + "))"), Mconn, adOpenKeyset, adLockPessimistic

If rsProv.RecordCount <> 0 Then

If MsgBox("Эта льгота используется в расчете! Удалять нельзя!", vbCritical, "Квартплата +") = vbOK Then
rsProv.Close
Exit Sub
End If

Else
rsProv.Close
TheRS.Delete (adAffectCurrent)
TheRS.MoveFirst
TheRS.UpdateBatch
FG.DataRefresh
End If
End If
End Sub

Private Sub DelZap_Click(Index As Integer)
'TheRS.Delete (adAffectCurrent)
'TheRS.MoveFirst
'TheRS.UpdateBatch
'fg.DataRefresh
Command3_Click
End Sub

'Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)

 '       Err.Raise 666, "IVSFlexDataSource", "This data is read-only."
'End Sub








Private Sub FG_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'fg.Refresh
Combo1.Refresh
Text1.Refresh
Text2.Refresh
'fg.DataRefresh

End Sub






Private Sub FG_LostFocus()
fld = FG.TextMatrix(0, FG.Col)
'SelRo = wadAffectCurrent
'MsgBox (fld)
End Sub

Private Sub Form_Load()

 
    
    ' открывать recordset для пакетной коррекции
    Set TheRS = New ADODB.Recordset
    Set TheRS.ActiveConnection = Mconn
    Set rsCombo = New ADODB.Recordset
    Set rsCombo.ActiveConnection = Mconn
'About.Show
    
    TheRS.CursorType = adOpenKeyset
    TheRS.LockType = adLockBatchOptimistic
    
    rsCombo.CursorType = adOpenStatic
    'adOpenDynamic
    'adOpenForwardOnly
    'adOpenStatic
   ' adOpenKeyset
   ' RsCombo.LockType = adLockBatchOptimistic
   
Set Combo_rs1 = New ADODB.Recordset
Set Combo_rs1.ActiveConnection = Mconn

Set rsProv = New ADODB.Recordset
Set rsProv.ActiveConnection = Mconn

Combo_rs1.CursorType = adOpenForwardOnly
Combo_rs1.LockType = adLockBatchOptimistic

Combo_rs1.Open "lgtip"
'TheRS.Open ("KLS_PRIV")
TheRS.Open ("SELECT KLS_PRIV.N_KLS, KLS_PRIV.NAME_KLS, KLS_PRIV.Tip, KLS_PRIV.LPKV, KLS_PRIV.LPTEH, KLS_PRIV.LPOTOPL, KLS_PRIV.LPCOMM, KLS_PRIV.LPMUSOR, KLS_PRIV.USEKV, KLS_PRIV.USETEH, KLS_PRIV.USEOTOPL, KLS_PRIV.USECOMM, KLS_PRIV.USEMUSOR From KLS_PRIV ORDER BY KLS_PRIV.N_KLS"), Mconn, adOpenKeyset, adLockPessimistic
rsCombo.Open "USE"
    
    
    '******************************************
    ' правопреемник recordset в сетку
    FG.FocusRect = 3
    'flexFocusSolid
    FG.Editable = False
    FG.DataMode = flexDMBoundImmediate
    
    
    Set FG.DataSource = TheRS
    FG.AutoSearch = flexSearchFromCursor
    FG.ExplorerBar = flexExSortShowAndMove
    
   
'************ Для текст1 **************
 Set Text1.DataSource = TheRS
 Text1.DataField = "NAME_KLS"
 '==============Квартплата ===============
Set Text2.DataSource = TheRS
Text2.DataField = "LPKV"
'Text2.Refresh
'---------------------------------------------------------------
Set Combo1.DataSource = TheRS
Combo1.DataField = "USEKV"
rsCombo.MoveFirst
Do While Not rsCombo.EOF
Combo1.AddItem rsCombo("USEKV")
rsCombo.MoveNext
Loop
'============= Техобслуживание =====================
Set Text3.DataSource = TheRS
Text3.DataField = "LPTEH"

'---------------------------------------------------------------
Set Combo2.DataSource = TheRS
Combo2.DataField = "USETEH"
rsCombo.MoveFirst
Do While Not rsCombo.EOF
Combo2.AddItem rsCombo("USEKV")
rsCombo.MoveNext
Loop

'============= Отопление =====================
Set Text4.DataSource = TheRS
Text4.DataField = "LPOTOPL"

'---------------------------------------------------------------
Set Combo3.DataSource = TheRS
Combo3.DataField = "USEOTOPL"
rsCombo.MoveFirst
Do While Not rsCombo.EOF
Combo3.AddItem rsCombo("USEKV")
rsCombo.MoveNext
Loop

'============= Коммунальные услуги =====================
Set Text5.DataSource = TheRS
Text5.DataField = "LPCOMM"

'---------------------------------------------------------------
Set Combo4.DataSource = TheRS
Combo4.DataField = "USECOMM"
rsCombo.MoveFirst
Do While Not rsCombo.EOF
Combo4.AddItem rsCombo("USEMUSOR")
rsCombo.MoveNext
Loop

'============= МУСОР =====================
Set Text6.DataSource = TheRS
Text6.DataField = "LPMUSOR"

'---------------------------------------------------------------
Set Combo5.DataSource = TheRS
Combo5.DataField = "USEMUSOR"
rsCombo.MoveFirst
Do While Not rsCombo.EOF
Combo5.AddItem rsCombo("USEMUSOR")
rsCombo.MoveNext
Loop


    
     '********* сортировка ******************
       
'  FG.Sort = flexSortGenericAscending
 ' FG.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove
    
    
              
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' закрытая связь (probly необязательно)
    'Mconn.Close
    Set TheRS = Nothing
    'Set TheConn = Nothing
End Sub

Private Sub New_Click(Index As Integer)
Command2_Click
End Sub

Private Sub Text1_LostFocus()
UpdateLabels
TheRS.UpdateBatch
'MsgBox (fg.SelectedRow)
End Sub

Private Sub Text2_LostFocus()
UpdateLabels
TheRS.UpdateBatch
'TheRS.Update
End Sub


Private Sub Text3_LostFocus()
TheRS.UpdateBatch
End Sub

Private Sub Text4_LostFocus()
TheRS.UpdateBatch
End Sub

Private Sub Text5_LostFocus()
TheRS.UpdateBatch
End Sub

Private Sub Text6_LostFocus()
TheRS.UpdateBatch
End Sub

Private Sub Пусто_Click()
Command1_Click
End Sub
Private Sub ЗапЛьгот()


'Добавляем  льготы для "квартплата" [Filter].[nm]
Mconn.Execute ("UPDATE KLS_PRIV INNER JOIN tmp_lgota ON KLS_PRIV.N_KLS = tmp_lgota.KodKls SET tmp_lgota.NAME_KLS = [KLS_PRIV]![NAME_KLS], tmp_lgota.Procent = [KLS_PRIV]![LPKV], tmp_lgota.Use = [KLS_PRIV]![USEKV] WHERE (((tmp_lgota.LgotaVid)=" + Chr(34) + "Квартплата" + Chr(34) + "))")

'Добавляем  льготы для "Отопление" [Filter].[nm]

Mconn.Execute ("UPDATE KLS_PRIV INNER JOIN tmp_lgota ON KLS_PRIV.N_KLS = tmp_lgota.KodKls SET tmp_lgota.NAME_KLS = [KLS_PRIV]![NAME_KLS], tmp_lgota.Procent = [KLS_PRIV]![LPotopl], tmp_lgota.Use = [KLS_PRIV]![USEotopl] WHERE (((tmp_lgota.LgotaVid)=" + Chr(34) + "Отопление" + Chr(34) + "))")
'Добавляем  льготы для "Техобслуживание" [Filter].[nm]

Mconn.Execute ("UPDATE KLS_PRIV INNER JOIN tmp_lgota ON KLS_PRIV.N_KLS = tmp_lgota.KodKls SET tmp_lgota.NAME_KLS = [KLS_PRIV]![NAME_KLS], tmp_lgota.Procent = [KLS_PRIV]![LPteh], tmp_lgota.Use = [KLS_PRIV]![USEteh] WHERE (((tmp_lgota.LgotaVid)=" + Chr(34) + "Техобслуживание" + Chr(34) + "))")
'Добавляем  льготы для "Мусор" [Filter].[nm]

Mconn.Execute ("UPDATE KLS_PRIV INNER JOIN tmp_lgota ON KLS_PRIV.N_KLS = tmp_lgota.KodKls SET tmp_lgota.NAME_KLS = [KLS_PRIV]![NAME_KLS], tmp_lgota.Procent = [KLS_PRIV]![LPmusor], tmp_lgota.Use = [KLS_PRIV]![USEmusor] WHERE (((tmp_lgota.LgotaVid)=" + Chr(34) + "Мусор" + Chr(34) + "))")
'Добавляем  льготы для "Коммунальные услуги" [Filter].[nm]

Mconn.Execute ("UPDATE KLS_PRIV INNER JOIN tmp_lgota ON KLS_PRIV.N_KLS = tmp_lgota.KodKls SET tmp_lgota.NAME_KLS = [KLS_PRIV]![NAME_KLS], tmp_lgota.Procent = [KLS_PRIV]![LPcomm], tmp_lgota.Use = [KLS_PRIV]![USEcomm] WHERE (((tmp_lgota.LgotaVid)=" + Chr(34) + "Коммунальные услуги" + Chr(34) + "))")
End Sub

Private Sub Реорганизация_Click()
Msg.Show
Msg.Label1.ForeColor = vbRed
Msg.sts = "Реорг"
Msg.Label1.Caption = "ВНИМАНИЕ" + vbNewLine + "Данный режим преднозначен для реорганизации кодоа льгот. Далее Вам будет предложено указать код реорганизуемой (изменяемой) льготы, и код льготы на который необходимо заменить старую льготу" + vbNewLine + "Перед выполнением реорганизации льгот ОБЯЗАТЕЛЬНО сделайте архивную копию базы данных, т.к. откат назад программой не предусмотрен, т.е. отменить Ваши действия будет невозможно" + vbNewLine + "Не пользуйтесь этим режимом без крайней необходимости"


End Sub

Private Sub Список_Click()
'rsProv.Open ("SELECT Lgota.NomNum, MainOccupant.BanKN, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Lgota.Numer FROM Lgota INNER JOIN MainOccupant ON Lgota.NomNum = MainOccupant.Numer WHERE (((Lgota.Numer)=" + FG.TextMatrix(FG.Row, 1) + "))"), Mconn, adOpenKeyset, adLockPessimistic
Analizlgot.StrSQL = "SELECT Lgota.Numer as [Код льготы], Lgota.name_kls as [Наименование льготы], Lgota.NomNum as [Код л сч], MainOccupant.OldNum as [Номер л сч], MainOccupant.BanKN as [Новый номер л сч], MainOccupant.FAM as [Фамилия], MainOccupant.IM as [Имя], MainOccupant.OT as [Отчество] FROM Lgota INNER JOIN MainOccupant ON Lgota.NomNum = MainOccupant.Numer WHERE (((Lgota.Numer)=" + FG.TextMatrix(FG.Row, 1) + "))"
Analizlgot.G = 9
Analizlgot.Show
Analizlgot.Об 0
Analizlgot.Label3 = "1111"
'rsProv.Close
End Sub
