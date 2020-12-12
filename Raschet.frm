VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form Filter 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Расчет"
   ClientHeight    =   7890
   ClientLeft      =   2595
   ClientTop       =   2595
   ClientWidth     =   11520
   HelpContextID   =   22
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10560
      TabIndex        =   10
      Top             =   0
      Width           =   240
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   741
      ButtonWidth     =   1191
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbarIcons(2)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F2"
            Key             =   "Семья"
            Object.ToolTipText     =   "Данные о жильцах"
            ImageKey        =   "Семья"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F3"
            Key             =   "Сырье. Материалы"
            Object.ToolTipText     =   "Постоянные начисления"
            ImageKey        =   "Сырье. Материалы"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F5"
            Key             =   "office0010"
            Object.ToolTipText     =   "Данные о квартире"
            ImageKey        =   "office0010"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F4"
            Key             =   "New"
            Object.ToolTipText     =   "Добавить новый лицевой счет"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F6"
            Key             =   "fil_end"
            Object.ToolTipText     =   "Снять фильтр"
            ImageKey        =   "fil_end"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F7"
            Key             =   "office0047"
            Object.ToolTipText     =   "Лицевой счет"
            ImageKey        =   "office0047"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F8"
            Key             =   "dell"
            Object.ToolTipText     =   "Удалить лиц.счет"
            ImageKey        =   "dell"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F12"
            Key             =   "exit"
            Object.ToolTipText     =   "Выход"
            ImageKey        =   "exit"
         EndProperty
      EndProperty
      Begin VB.CommandButton Command13 
         Caption         =   "A-"
         Height          =   375
         Left            =   9480
         TabIndex        =   14
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Caption         =   "A+"
         Height          =   375
         Left            =   9000
         TabIndex        =   13
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         Height          =   375
         Left            =   5400
         Picture         =   "Raschet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Расчитать отмеченных"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   11
         Top             =   0
         Width           =   2295
      End
      Begin VB.CommandButton Command7 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10800
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton Command8 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10320
         TabIndex        =   8
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Raschet.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Добавить<F4>"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Picture         =   "Raschet.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Квартиросъемщик <F5>"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Picture         =   "Raschet.frx":0AFE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Закрыть <F12>"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VSFlex8LCtl.VSFlexGrid FG 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11295
      _cx             =   19923
      _cy             =   12938
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   8388608
      GridColorFixed  =   255
      TreeColor       =   -2147483642
      FloodColor      =   192
      SheetBorder     =   -2147483643
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
      FormatString    =   $"Raschet.frx":0F40
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   2
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Picture         =   "Raschet.frx":101F
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Постоянные начисления <F3>"
      Top             =   7080
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Picture         =   "Raschet.frx":1329
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Члены семьи <F2>"
      Top             =   7080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Picture         =   "Raschet.frx":14F9
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Снять фильтр<F6>"
      Top             =   7080
      Width           =   495
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   0
      Left            =   8640
      Top             =   -360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":15FB
            Key             =   "office0010"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":1E6D
            Key             =   "Семья"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":21BF
            Key             =   "Сырье. Материалы"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":24D9
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":25EB
            Key             =   "office0047"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":287D
            Key             =   "dell"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":2A57
            Key             =   "fil_end"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   1
      Left            =   4950
      Top             =   3705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":2B69
            Key             =   "office0010"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":33DB
            Key             =   "Семья"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":372D
            Key             =   "Сырье. Материалы"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":3A47
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":3B59
            Key             =   "office0047"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":3DEB
            Key             =   "dell"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":3FC5
            Key             =   "fil_end"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":40D7
            Key             =   "Save2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   2
      Left            =   4950
      Top             =   3705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":4929
            Key             =   "Семья"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":4C7B
            Key             =   "Сырье. Материалы"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":4F95
            Key             =   "office0010"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":5807
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":5919
            Key             =   "fil_end"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":5A2B
            Key             =   "office0047"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":5CBD
            Key             =   "dell"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Raschet.frx":5E97
            Key             =   "exit"
         EndProperty
      EndProperty
   End
   Begin VB.Menu Меню 
      Caption         =   "Меню"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu Жильцы 
         Caption         =   "Жильцы"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu Удалить 
         Caption         =   "Постоянные начисления"
         Index           =   3
         Shortcut        =   {F3}
      End
      Begin VB.Menu Доб 
         Caption         =   "Добавить новый лиц.сч."
         Shortcut        =   {F4}
      End
      Begin VB.Menu Квартиросъемщик 
         Caption         =   "Квартиросъемщик"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Снять 
         Caption         =   "Снять фильтр"
         Shortcut        =   {F6}
      End
      Begin VB.Menu Расчет 
         Caption         =   "Лиц.счет"
         Shortcut        =   {F7}
      End
      Begin VB.Menu Удалить1 
         Caption         =   "Удалить"
         Index           =   8
         Shortcut        =   {F8}
      End
      Begin VB.Menu Закрыть 
         Caption         =   "Закрыть"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FSize As Double
Public Nm, FIO, ad
Option Explicit
Dim m_DS As FlexADO
Dim Conn_Add As ADODB.Connection
Dim Rs_Add As ADODB.Recordset
Dim r As Integer

Private Sub Command10_Click()
SposobR.Show

End Sub

Private Sub Command11_Click()
' PrintW.VP.ColorMode = cmMonochrome
'
FG.Font.Size = 8
FG.AutoResize = True
FG.AutoSize 0, FG.Cols - 1
FG.FontBold = False

PrintW.Show
        
        PrintW.VP.StartDoc
        PrintW.VP.RenderControl = FG.hwnd
        PrintW.VP.EndDoc
       'PrintW.VP.RenderControl = m_DS.m_RS
       PrintW.VP.FontSize = 8
       
       
       
               
        
End Sub

Private Sub Command12_Click()
FG.Font.Size = FSize
FG.Refresh

End Sub

Private Sub Command13_Click()

'FSize = FG.Font.Size
If FG.Font.Size >= 8 Then FG.Font.Size = FG.Font.Size - 1
'FG.AutoResize = True
'FG.AutoSize 0, FG.Cols - 1
FG.Refresh
End Sub

Private Sub Command7_Click()
For r = 2 To FG.Rows - 1
   'FG.TextMatrix(r, 6) = False
            FG.Cell(flexcpChecked, r, 6) = flexUnchecked
Conn_Add.Execute ("UPDATE MainOccupant SET MainOccupant.otm = False")
            
        Next

End Sub

Private Sub Command8_Click()
'flexNoCheckbox
For r = 2 To FG.Rows - 1
   
            FG.Cell(flexcpChecked, r, 6) = flexChecked
  Conn_Add.Execute ("UPDATE MainOccupant SET MainOccupant.otm = True")
        Next

End Sub

Private Sub Command9_Click()

For r = 2 To FG.Rows - 1

If FG.Cell(flexcpChecked, r, 6) = flexChecked Then
FG.Cell(flexcpChecked, r, 6) = flexUnchecked

GoTo N
End If

If FG.Cell(flexcpChecked, r, 6) = flexUnchecked Then
FG.Cell(flexcpChecked, r, 6) = flexChecked

End If
N:
Next

Conn_Add.Execute ("UPDATE MainOccupant SET MainOccupant.otm = IIf(MainOccupant!otm=True,False,True)")

End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If FG.Cell(flexcpChecked, FG.Row, FG.Col) = True Then FG.Editable = flexEDKbdMouse




End Sub

Private Sub FG_Click()
'If FG.Cell(flexcpChecked, FG.Row, FG.Col) = flexChecked Then FG.Cell(flexcpChecked, FG.Row, FG.Col) = flexNoCheckbox
'If FG.Cell(flexcpChecked, FG.Row, FG.Col) = flexNoCheckbox Then
'FG.Cell(flexcpChecked, FG.Row, FG.Col) = flexChecked
'MsgBox (FG.Cell(flexcpChecked, FG.Row, FG.Col))
'End If

End Sub

'Public f As String


Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.KEY
        Case "office0010"
           Command5_Click
        Case "Семья"
            Command3_Click
        Case "Сырье. Материалы"
         Command4_Click
        Case "New"
            Command6_Click
        Case "office0047"
            FG_DblClick
        Case "dell"
            'ToDo: Add 'dell' button code.
           Удалить1_Click (0)
        Case "fil_end"
            Command1_Click
        Case "exit"
            Command2_Click
    End Select
End Sub
'Dim nm As Integer
'Dim FIO As String
Private Sub ref()
Form_Load
End Sub

Private Sub Command1_Click()
    
    ' ясные данные фильтра, восстановления
    FG.Cell(flexcpText, 1, 0, 1, FG.Cols - 1) = ""
    FG.FlexDataSource = m_DS
    ОбнВыд
End Sub

Private Sub Command2_Click()
Unload Me
MainMenu.Enabled = True
MainMenu.Show
'Form_Unload (Filter)
Filter.Refresh
'Form_Unload() = True
Filter.Hide

End Sub

Private Sub Command3_Click()
If Filter.Nm = "" Then
MsgBox ("Вы не выбрали квартиросъемщика")
Else
OtheOwner.Show
FIO = FG.Cell(flexcpText, FG.Row, 1) + " " + FG.Cell(flexcpText, FG.Row, 2) + "  " + FG.Cell(flexcpText, FG.Row, 2)
OtheOwner.Caption = "Ответственный квартиросъемщик-> " + Filter.FIO
Filter.Nm = FG.Cell(flexcpText, FG.Row, 0)
End If
'Form_Load
End Sub

Private Sub Command4_Click()
If Filter.Nm = "" Then
MsgBox ("Вы не выбрали квартиросъемщика")
Else
Constant.Show
FIO = FG.Cell(flexcpText, FG.Row, 1) + " " + FG.Cell(flexcpText, FG.Row, 2) + "  " + FG.Cell(flexcpText, FG.Row, 2)
Constant.Caption = " " + Filter.FIO
Filter.Nm = FG.Cell(flexcpText, FG.Row, 0)
End If
End Sub



Private Sub Command5_Click()
FIO = FG.Cell(flexcpText, FG.Row, 1) + " " + FG.Cell(flexcpText, FG.Row, 2) + "  " + FG.Cell(flexcpText, FG.Row, 2)
Constant.Caption = " " + Filter.FIO
Filter.Nm = FG.Cell(flexcpText, FG.Row, 0)
'Filter.Hide
'Filter.FG.Clear
If Filter.Nm = "" Then
MsgBox ("Вы не выбрали квартиросъемщика")
Else
Kvart.Show
Filter.Hide
m_DS.m_RS.Close
m_DS.m_Conn.Close
End If

End Sub

Private Sub Command6_Click()

Dim N, N1 As Integer
If MsgBox("Добавить нового квартиросъемщика?", vbYesNo) = vbYes Then
N = 0
m_DS.m_RS.MoveFirst
Do While Not m_DS.m_RS.EOF
N1 = m_DS.m_RS("Номер").Value
If N1 > N Then N = N1
m_DS.m_RS.MoveNext
Loop






'Set Conn_Add = New ADODB.Connection
 ' Conn_Add.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'Conn_Add.Open "data/Kvartplata.mdb"
    
Set Rs_Add = New ADODB.Recordset
Set Rs_Add.ActiveConnection = Conn_Add
 
Rs_Add.CursorType = adOpenForwardOnly
Rs_Add.LockType = adLockBatchOptimistic
Rs_Add.Open "MainOccupant"

Rs_Add.AddNew
Rs_Add("Numer") = N + 1

Rs_Add.UpdateBatch
Rs_Add.Update
Rs_Add.Requery


Filter.ad = 1
Filter.Nm = N + 1
Kvart.Show
Filter.Hide
m_DS.m_RS.Close
m_DS.m_Conn.Close
Conn_Add.Close
'Rs_Add.Close
End If
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    ' новому фильтру, нужно восстановление
    If Row = 1 Then FG.FlexDataSource = m_DS
    
    
End Sub

Private Sub FG_DblClick()
If Filter.Nm = "" Then
MsgBox ("Вы не выбрали квартиросъемщика")
Else
FIO = FG.Cell(flexcpText, FG.Row, 1) + " " + FG.Cell(flexcpText, FG.Row, 2) + "  " + FG.Cell(flexcpText, FG.Row, 3)
Lic.Caption = " " + Filter.FIO
Filter.Nm = FG.Cell(flexcpText, FG.Row, 0)
Lic.Show
Filter.Enabled = False
End If
End Sub

Private Sub FG_EnterCell()

' переменная nm будет содержать текст текущей выбранной ячейки
  'nm = FG.Cell(flexcpText, FG.Row, FG.Col)
Filter.Caption = FG.Cell(flexcpText, FG.Row, 1) + " " + FG.Cell(flexcpText, FG.Row, 2) + "  " + FG.Cell(flexcpText, FG.Row, 2)
  ' nm присваевается значение номера выбранной ячейки
Nm = FG.Cell(flexcpText, FG.Row, 0)

'MsgBox (nm)

'FIO = m_DS.m_RS.Fields("Фамилия").Value + " " + m_DS.m_RS.Fields("Имя").Value + " " + m_DS.m_RS.Fields("Отчество").Value
'Filter.Caption = "Текущая запись-> " + FIO

End Sub


Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


    'для этого образца, мы только допускаем редактировать линию фильтра
   'If Row <> 1 Then
    'If Col <> 6 And Row <> 1 Then Cancel = True
    'End If
    FG.Editable = flexEDKbdMouse
    
    
    If FG.Cell(flexcpChecked, Row, 6) = flexChecked Then

FG.Cell(flexcpChecked, Row, 6) = flexUnchecked
Conn_Add.Execute ("UPDATE MainOccupant SET MainOccupant.otm = False WHERE (((MainOccupant.Numer)=" + FG.TextMatrix(Row, 0) + "))")
GoTo N1
       End If
       

If FG.Cell(flexcpChecked, Row, 6) = flexUnchecked Then
FG.Cell(flexcpChecked, Row, 6) = flexChecked
Conn_Add.Execute ("UPDATE MainOccupant SET MainOccupant.otm = True WHERE (((MainOccupant.Numer)=" + FG.TextMatrix(Row, 0) + "))")

End If
N1:

End Sub



Private Sub Form_Activate()
Form_Load

End Sub





Private Sub Form_Load()
FSize = FG.Font.Size
Set Conn_Add = New ADODB.Connection
  Conn_Add.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  Conn_Add.Open "data/Kvartplata.mdb"
    
 
Filter.ad = 0


'Filter.Caption = FIO
'////////////////////////////////////////////////////

FG.AutoSearch = flexSearchFromCursor
FG.ExplorerBar = flexExSortShowAndMove



    ' инициализируйте сетку (дополнительную)
    FG.FixedCols = 0
    FG.Editable = flexEDKbdMouse
    FG.BackColorFrozen = RGB(200, 255, 200)
    
    ' создайте исходный объект заказных данных
    Set m_DS = New FlexADO
 
        
    ' назначьте этим в сетку
    FG.FlexDataSource = m_DS
        
    FG.FrozenRows = 1
    FG.DataMode = flexDMBoundBatch
    
    
    'flexDMFree
    
        
    
    ' Cвойства, свойства необходимые для сортировки в этом гриде не работают
    ' из за строки поиска
    'FG.AllowUserResizing = flexResizeBoth
    'FG.ExtendLastCol = True
    'FG.ExplorerBar = flexExSortShowAndMove
    'FG.AutoSearch = flexSearchFromCursor
    
    'FG.Cell(flexcpChecked, 2, 6, FG.Rows, 6) = flexUnchecked
  
  
  ОбнВыд
  
    




End Sub
'////////// Сортировка

Private Sub Form_Resize()
    On Error Resume Next
    FG.Move FG.Left, FG.Top, ScaleWidth - FG.Left * 2, ScaleHeight - FG.Left - FG.Top
End Sub

Private Sub Доб_Click()
Command6_Click
End Sub

Private Sub Жильцы_Click(Index As Integer)
Command3_Click
End Sub

Private Sub Закрыть_Click()
Command2_Click
End Sub

Private Sub Квартиросъемщик_Click()
Command5_Click
End Sub

Private Sub Расчет_Click()
FG_DblClick
End Sub

Private Sub Снять_Click()
Command1_Click
End Sub

'Private Sub Удалить_Click()
'Command4_Click
'End Sub

 Private Sub Удалить1_Click(Index As Integer)


If MsgBox("Вы хотите удалить лицевой счет №" + FG.TextMatrix(FG.Row, 0) + " ответственный квартиросъемщик " + FG.TextMatrix(FG.Row, 1) + "  " + FG.TextMatrix(FG.Row, 2) + "  " + FG.TextMatrix(FG.Row, 3) + "? ", vbYesNo) = vbYes Then
'Filter.Hide
Unload Me
If MsgBox("Все данные связанные с этим лицевым счетом будут удалены без возможности восстановления. Вы уверены", vbYesNo) = vbYes Then


Conn_Add.Execute ("DELETE Adding.KodKv, Adding.* From Adding WHERE (((Adding.KodKv)=" + Filter.Nm + "))")
Conn_Add.Execute ("DELETE Constanta.Numer, Constanta.KodNach, Constanta.NameNach From Constanta WHERE (((Constanta.Numer)=" + Filter.Nm + "))")
Conn_Add.Execute ("DELETE Lgota.NomNum, Lgota.* From Lgota WHERE (((Lgota.NomNum)=" + Filter.Nm + "))")
Conn_Add.Execute ("DELETE OtheOwner.Numer, OtheOwner.* From OtheOwner WHERE (((OtheOwner.Numer)=" + Filter.Nm + "))")
Conn_Add.Execute ("DELETE tmp_lgota.KodKv, tmp_lgota.* From tmp_lgota WHERE (((tmp_lgota.KodKv)=" + Filter.Nm + "))")

Conn_Add.Execute ("DELETE MainOccupant.Numer, MainOccupant.* From MainOccupant WHERE (((MainOccupant.Numer)=" + Filter.Nm + "))")


'FG.Redraw = flexRDBuffered
'Command1_Click

'm_DS.m_RS.Requery
'Command1_Click
'FG.FlexDataSource = m_DS
'Unload Me

End If
End If


Filter.Show

End Sub
Private Sub Отметка()
Dim nu As Integer

For r = 2 To FG.Rows - 1
nu = Val(FG.TextMatrix(r, 4))

MsgBox (Str(nu))

If FG.Cell(flexcpChecked, r, 6) = flexChecked Then
MsgBox (FG.Cell(flexcpChecked, r, 6))
Conn_Add.Execute ("UPDATE MainOccupant SET MainOccupant.otm = True WHERE (((MainOccupant.Numer)=" + Nm + "))")
Else
If FG.Cell(flexcpChecked, r, 6) = flexUnchecked Then Conn_Add.Execute ("UPDATE MainOccupant SET MainOccupant.otm = False WHERE (((MainOccupant.Numer)=" + Str(nu) + "))")
End If
Next
End Sub
Private Sub ОбнВыд()
For r = 2 To FG.Rows - 1

'MsgBox (FG.Cell(flexcpChecked, r, 6))

If FG.TextMatrix(r, 7) = False Then
FG.Cell(flexcpChecked, r, 6) = flexUnchecked
GoTo N
End If

If FG.TextMatrix(r, 7) = True Then
FG.Cell(flexcpChecked, r, 6) = flexChecked
End If
N:
Next
End Sub
