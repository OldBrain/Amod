VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Filter 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������"
   ClientHeight    =   7896
   ClientLeft      =   2592
   ClientTop       =   2592
   ClientWidth     =   12816
   ControlBox      =   0   'False
   HelpContextID   =   22
   Icon            =   "Reestr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7896
   ScaleWidth      =   12816
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      DisabledPicture =   "Reestr.frx":030A
      DownPicture     =   "Reestr.frx":03B1
      Height          =   315
      Left            =   11520
      Picture         =   "Reestr.frx":045F
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "��������/��������  ���������� ����"
      Top             =   600
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "� ���/��"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "��������/��������  ���������� ����"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00BDC6BB&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10800
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00BDC6BB&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10560
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton Command9 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   240
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   348
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12816
      _ExtentX        =   22606
      _ExtentY        =   614
      ButtonWidth     =   1058
      ButtonHeight    =   487
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbarIcons(2)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F2"
            Key             =   "�����"
            Object.ToolTipText     =   "������ � �������"
            ImageKey        =   "�����"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F3"
            Key             =   "�����. ���������"
            Object.ToolTipText     =   "���������� ����������"
            ImageKey        =   "�����. ���������"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F5"
            Key             =   "office0010"
            Object.ToolTipText     =   "������ � ��������"
            ImageKey        =   "office0010"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F4"
            Key             =   "New"
            Object.ToolTipText     =   "�������� ����� ������� ����"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F6"
            Key             =   "fil_end"
            Object.ToolTipText     =   "����� ������"
            ImageKey        =   "fil_end"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F7"
            Key             =   "office0047"
            Object.ToolTipText     =   "������� ����"
            ImageKey        =   "office0047"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F12"
            Key             =   "exit"
            Object.ToolTipText     =   "�����"
            ImageKey        =   "exit"
         EndProperty
      EndProperty
      Begin VB.CommandButton Command18 
         BackColor       =   &H80000013&
         Caption         =   "��������� ������"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6000
         MaskColor       =   &H008080FF&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   1692
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00BDC6BB&
         Height          =   375
         Left            =   11040
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Reestr.frx":0506
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00BDC6BB&
         Height          =   300
         Left            =   4800
         Picture         =   "Reestr.frx":099C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "������ �� ������"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00BDC6BB&
         Caption         =   "A-"
         Height          =   375
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00BDC6BB&
         Caption         =   "A+"
         Height          =   300
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00BDC6BB&
         Height          =   300
         Left            =   5400
         Picture         =   "Reestr.frx":0A74
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "������ ����������"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00BDC6BB&
         Height          =   300
         Left            =   7680
         Picture         =   "Reestr.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00BDC6BB&
         Caption         =   ">"
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
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00BDC6BB&
         Caption         =   "<"
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
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Reestr.frx":10DA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "��������<F4>"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Picture         =   "Reestr.frx":151C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "��������������� <F5>"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Picture         =   "Reestr.frx":16A6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "������� <F12>"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VSFlex8LCtl.VSFlexGrid FG 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12615
      _cx             =   22251
      _cy             =   12726
      Appearance      =   3
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
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
      BackColorSel    =   12615680
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483641
      BackColorAlternate=   16777215
      GridColor       =   8388608
      GridColorFixed  =   0
      TreeColor       =   -2147483642
      FloodColor      =   0
      SheetBorder     =   255
      FocusRect       =   4
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1200
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Reestr.frx":1AE8
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   20
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
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
      CellButtonPicture=   "Reestr.frx":1C2C
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   16711680
      ForeColorFrozen =   255
      WallPaperAlignment=   0
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Image ImgNul 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         OLEDragMode     =   1  'Automatic
         Picture         =   "Reestr.frx":1F46
         Stretch         =   -1  'True
         Top             =   6600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image ImgcellDog 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3240
         OLEDragMode     =   1  'Automatic
         Picture         =   "Reestr.frx":2488
         Stretch         =   -1  'True
         Top             =   6600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Imgcell 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         OLEDragMode     =   1  'Automatic
         Picture         =   "Reestr.frx":25E9
         Stretch         =   -1  'True
         Top             =   6600
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Picture         =   "Reestr.frx":26AC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "���������� ���������� <F3>"
      Top             =   7080
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Picture         =   "Reestr.frx":29B6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "����� ����� <F2>"
      Top             =   7080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Picture         =   "Reestr.frx":2B86
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "����� ������<F6>"
      Top             =   7080
      Width           =   495
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   0
      Left            =   8640
      Top             =   -360
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":2C88
            Key             =   "office0010"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":34FA
            Key             =   "�����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":384C
            Key             =   "�����. ���������"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":3B66
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":3C78
            Key             =   "office0047"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":3F0A
            Key             =   "dell"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":40E4
            Key             =   "fil_end"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   1
      Left            =   4950
      Top             =   3705
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":41F6
            Key             =   "office0010"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":4A68
            Key             =   "�����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":4DBA
            Key             =   "�����. ���������"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":50D4
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":51E6
            Key             =   "office0047"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":5478
            Key             =   "dell"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":5652
            Key             =   "fil_end"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":5764
            Key             =   "Save2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   2
      Left            =   4950
      Top             =   3705
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":5FB6
            Key             =   "�����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":6308
            Key             =   "�����. ���������"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":6622
            Key             =   "office0010"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":6E94
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":6FA6
            Key             =   "fil_end"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":70B8
            Key             =   "office0047"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":734A
            Key             =   "dell"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reestr.frx":7524
            Key             =   "exit"
         EndProperty
      EndProperty
   End
   Begin VB.Menu ���� 
      Caption         =   "������ �/��"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu lsEdit 
         Caption         =   "�������� �/��."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu ����� 
         Caption         =   "����� �/��"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu ���� 
         Caption         =   "�������������� ������"
         Shortcut        =   {F9}
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu ������� 
         Caption         =   "���������� ����������"
         Index           =   3
         Shortcut        =   {F3}
      End
      Begin VB.Menu ��� 
         Caption         =   "�������� ����� ���.��."
         Shortcut        =   {F4}
      End
      Begin VB.Menu ��������������� 
         Caption         =   "���������������"
         Shortcut        =   {F5}
      End
      Begin VB.Menu �������1 
         Caption         =   "�������"
         Index           =   8
         Shortcut        =   +^{F8}
      End
      Begin VB.Menu ������ 
         Caption         =   "���.����"
         Shortcut        =   {F7}
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu ������_ 
      Caption         =   "������_"
      Index           =   2
      Begin VB.Menu ������� 
         Caption         =   "������ �� ������"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu ������ 
         Caption         =   "������ "
         Shortcut        =   +{F6}
      End
      Begin VB.Menu ����� 
         Caption         =   "����� ������"
      End
   End
   Begin VB.Menu �������� 
      Caption         =   "��������"
      Index           =   5
      Begin VB.Menu ������ 
         Caption         =   "��������� ������ "
      End
      Begin VB.Menu ��������2 
         Caption         =   "������� ������"
      End
      Begin VB.Menu ��������1 
         Caption         =   "�/��. ��� ����������"
      End
   End
   Begin VB.Menu �����1 
      Caption         =   "�������� ��������"
      Index           =   3
      Begin VB.Menu ����� 
         Caption         =   "���������� ��� ��� ����"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu �_��������� 
      Caption         =   "� ���������"
      Index           =   4
   End
End
Attribute VB_Name = "Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FSize As Double
Public Nm, FIO, ad, nNum
Public infRS As ADODB.Recordset
Public CL5 As Long
Public oldR As Long
Option Explicit
Public m_DS As FlexADO
'Dim mconn As ADODB.Connection
Dim Rs_Add As ADODB.Recordset
Dim R As Integer
'Dim SQ As String









Private Sub BtnEnh1_Click()
Unload Filter
MainMenu.Show
MainMenu.Enabled = True
'mconn.Close
End Sub

Private Sub Check1_Click()
If Fg.ColHidden(0) = True Then Fg.ColHidden(0) = False Else Fg.ColHidden(0) = True
If Fg.ColHidden(1) = True Then Fg.ColHidden(1) = False Else Fg.ColHidden(1) = True
Fg.Redraw = flexRDBuffered

End Sub

Private Sub Check2_Click()
 '���������������������������������������������������
' ��������

   'FG.ColComboList(12) = Imgcell
   ' ����.�����
   For R = 2 To Fg.Rows - 1
   
   If Fg.TextMatrix(R, 12) = 0 Then
   Fg.Cell(flexcpPicture, R, 12, R) = ImgNul
   
   
   Fg.Cell(flexcpBackColor, R, 12, R) = &H8080FF
   End If
   
   If Fg.TextMatrix(R, 12) = 1 Then
   Fg.Cell(flexcpPicture, R, 12, R) = Imgcell
   'FG.Cell(flexcpPictureAlignment, 2, 12, FG.Rows - 1) = flexPicAlignLeftCenter
   Fg.Cell(flexcpBackColor, R, 12, R) = &HE0E0E0
   End If
   
   '�������
   If Fg.TextMatrix(R, 12) = 2 Then
   Fg.Cell(flexcpPicture, R, 12, R) = ImgcellDog
   'FG.Cell(flexcpPictureAlignment, 2, 12, FG.Rows - 1) = flexPicAlignLeftCenter
   Fg.Cell(flexcpBackColor, R, 12, R) = &HE0E0E0
   End If
        

        Next
    
'����������������������������������������������������������

End Sub

Private Sub Command10_Click()
Filter.Enabled = False
'mconn.Execute ("DELETE Err_Ras.* FROM Err_Ras")

SposobR.Show
End Sub

Private Sub Command11_Click()
'FG.Font.Size = 8
'FG.AutoResize = False
'FG.AutoSize 0, FG.Cols - 1
'FG.FontBold = False
'PrintW.Show
 '       PrintW.VP.StartDoc
  '      PrintW.VP.RenderControl = FG.hwnd
   '     PrintW.VP.EndDoc
    '    PrintW.VP.FontSize = 8
       
       
 Analizlgot.Titl = "��������� ��������� �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))

Analizlgot.G = 9
Analizlgot.StrSQL = "SELECT KLS_PODR.NAIM_KLS AS �����, KLS_PODR.Num AS �����_����, MainOccupant.bankN as N, MainOccupant.Numer AS �����, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, MainOccupant.kv_num AS �_�� FROM MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� Where (((MainOccupant.Otm) = True)) ORDER BY MainOccupant.FAM"
'Analizlgot.�� 3
Unload Me
Analizlgot.Show
               
        
End Sub

Private Sub Command12_Click()
Dim tmp_Font As Double

tmp_Font = Fg.Font.Size
Fg.Font.Size = Fg.Font.Size + 1
If Fg.Font.Size = tmp_Font Then Fg.Font.Size = Fg.Font.Size + 2
Fg.Refresh
End Sub

Private Sub Command13_Click()

'FSize = FG.Font.Size
If Fg.Font.Size >= 8 Then Fg.Font.Size = Fg.Font.Size - 1
'FG.AutoResize = True
'FG.AutoSize 0, FG.Cols - 1
Fg.Refresh
End Sub

Private Sub Command14_Click()
oldR = Fg.Row
Fg.Row = 1
Otm.Show
End Sub

Private Sub Command15_Click()
For R = 2 To Fg.Rows - 1
   
            Fg.Cell(flexcpChecked, R, 7) = flexChecked
  Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = True  WHERE (((MainOccupant.Numer)=" + Fg.TextMatrix(R, 0) + "))")
        
        
        Next
End Sub

Private Sub Command16_Click()
Unload Filter
MainMenu.Show
MainMenu.Enabled = True
'mconn.Close
End Sub

Private Sub Command17_Click()
For R = 2 To Fg.Rows - 1
   
            Fg.Cell(flexcpChecked, R, 7) = flexUnchecked
  Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = False  WHERE (((MainOccupant.Numer)=" + Fg.TextMatrix(R, 0) + "))")
        
        
        Next
End Sub

Private Sub Command18_Click()
Arhiv_all.Show
End Sub

'Private Sub Command14_Click()
'For r = 2 To 20
'MainForm.��
'Next r
'End Sub

Private Sub Command7_Click()
Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = False")
For R = 2 To Fg.Rows - 1
   'FG.TextMatrix(r, 6) = False
            Fg.Cell(flexcpChecked, R, 7) = flexUnchecked
  
        Next

End Sub

Private Sub Command8_Click()
'flexNoCheckbox
Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = True")
For R = 2 To Fg.Rows - 1
   
            Fg.Cell(flexcpChecked, R, 7) = flexChecked
  'mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = True  WHERE (((MainOccupant.Numer)=" + FG.TextMatrix(r, 0) + "))")
        
        
        Next
'm_DS.m_RS.UpdateBatch
End Sub

Private Sub Command9_Click()

For R = 2 To Fg.Rows - 1

If Fg.Cell(flexcpChecked, R, 7) = flexChecked Then
Fg.Cell(flexcpChecked, R, 7) = flexUnchecked

GoTo n
End If

If Fg.Cell(flexcpChecked, R, 7) = flexUnchecked Then
Fg.Cell(flexcpChecked, R, 7) = flexChecked

End If
n:
Next

Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = IIf(MainOccupant!otm=True,False,True)")

End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If FG.Cell(flexcpChecked, FG.Row, FG.Col) = True Then FG.Editable = flexEDKbdMouse




End Sub



Private Sub FG_GotFocus()
' ����������� ���� ���������� ������������ �������� ����
PopUp.Show
PopUp.Height = Me.Top
PopUp.Width = Me.Width + Me.Left

PopUp.Enabled = False
PopUp.Refresh
MakeWindow PopUp, True
PopUp.Enabled = False
Fg.SetFocus

End Sub

Private Sub FG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Fg.Row <> 1 Then FG_DblClick
If KeyAscii = 32 And Fg.Row <> 1 Then
If MsgBox("���������� ������ �� " + Fg.TextMatrix(Fg.Row, Fg.Col), vbYesNo) = vbYes Then
Fg.TextMatrix(1, Fg.Col) = Fg.TextMatrix(Fg.Row, Fg.Col)
Command15.Visible = True
Command17.Visible = True
Fg.FlexDataSource = m_DS
End If
End If
End Sub

Private Sub Form_Activate()
Me.SetFocus
Filter.Fg.SetFocus
End Sub

'Public f As String







Private Sub Form_Unload(Cancel As Integer)
Dim rsLcKol As ADODB.Recordset

Set rsLcKol = New ADODB.Recordset


MainMenu.Command6.Caption = "������"
MainMenu.Command6.BackColor = &H80000018
MainMenu.Command6.Refresh
MainMenu.Enabled = True

rsLcKol.Open ("SELECT Count(MainOccupant.Numer) AS [Count-Numer] FROM MainOccupant"), Mconn
If rsLcKol.EOF = False And rsLcKol.BOF = False Then MainForm.LcKol = rsLcKol("Count-Numer")
rsLcKol.Close


rsLcKol.Open ("SELECT Count(MainOccupant.Numer) AS [Count-Numer], MainOccupant.Dog From MainOccupant GROUP BY MainOccupant.Dog HAVING (((MainOccupant.Dog)=2))"), Mconn
If rsLcKol.EOF = False And rsLcKol.BOF = False Then MainForm.LcKolD = rsLcKol("Count-Numer")
rsLcKol.Close



rsLcKol.Open ("SELECT Count(MainOccupant.Numer) AS [Count-Numer], MainOccupant.Dog From MainOccupant GROUP BY MainOccupant.Dog HAVING (((MainOccupant.Dog)=1))"), Mconn
If rsLcKol.EOF = False And rsLcKol.BOF = False Then MainForm.LcKolK = rsLcKol("Count-Numer")
rsLcKol.Close

MainMenu.lblTitle.ToolTipText = "���-�� ���/�� >" + Str(MainForm.LcKol) + ". � �.�. ���������-" + Str(MainForm.LcKolD) + ". ����.������-" + Str(MainForm.LcKolK)

End Sub

Private Sub lsEdit_Click()

nNum = InputBox("������ �����>>" + Fg.TextMatrix(Fg.Row, 1), Fg.TextMatrix(Fg.Row, 1) + "  " + Fg.TextMatrix(Fg.Row, 2) + "  " + Fg.TextMatrix(Fg.Row, 3), Fg.TextMatrix(Fg.Row, 1))

MsgBox (Len(nNum))

If nNum <> Fg.TextMatrix(Fg.Row, 1) And Len(nNum) <> 0 Then

'MsgBox (Fg.TextMatrix(Fg.Row, 0))

Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.OLDNUM = '" + nNum + "' WHERE (((MainOccupant.Numer)= " + Fg.TextMatrix(Fg.Row, 0) + "))")

Fg.TextMatrix(Fg.Row, 0) = nNum


End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    'On Error Resume Next
    Select Case Button.KEY
        Case "office0010"
           Command5_Click
        Case "�����"
            Command3_Click
        Case "�����. ���������"
         
        Case "New"
            Command6_Click
        Case "office0047"
            FG_DblClick
        Case "dell"
            
           �������1_Click (0)
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
    
    ' ����� ������ �������, ��������������
    Fg.Cell(flexcpText, 1, 0, 1, Fg.Cols - 1) = ""
    Fg.FlexDataSource = m_DS
    Command15.Visible = False
    Command17.Visible = False
    ������
End Sub

Private Sub Command2_Click()

Unload Filter

MainMenu.Show
MainMenu.Enabled = True
'mconn.Close
Unload Filter
End Sub

Private Sub Command3_Click()
If Filter.Nm = "" Then
MsgBox ("�� �� ������� ����������������")
Else
OtheOwner.Show
FIO = Fg.Cell(flexcpText, Fg.Row, 1) + " " + Fg.Cell(flexcpText, Fg.Row, 2) + "  " + Fg.Cell(flexcpText, Fg.Row, 2)
OtheOwner.lblTitle.Caption = "������������� ���������������-> " + Filter.FIO
Filter.Nm = Fg.Cell(flexcpText, Fg.Row, 0)
End If
'Form_Load
End Sub

Private Sub Command5_Click()
CL5 = Fg.Row
FIO = Fg.Cell(flexcpText, Fg.Row, 2) + " " + Fg.Cell(flexcpText, Fg.Row, 3) + "  " + Fg.Cell(flexcpText, Fg.Row, 4)

Filter.Nm = Fg.Cell(flexcpText, Fg.Row, 0)
'Filter.Hide
'Filter.FG.Clear
If Filter.Nm = "" Then
MsgBox ("�� �� ������� ����������������")
Else
Me.Enabled = False


Kvart.Show
'Filter.Hide
'm_DS.m_RS.Close
'm_DS.m_Conn.Close
End If

End Sub

Private Sub Command6_Click()

MsgBox "���������� ����� ������� ������, ������������ � ������ /���� �������/, ���� /���������/"
Exit Sub


Dim n, N1 As Double
nNum = 0
Set Rs_Add = New ADODB.Recordset
Set Rs_Add.ActiveConnection = Mconn
Rs_Add.CursorType = adOpenDynamic
Rs_Add.LockType = adLockPessimistic
Rs_Add.Open "MainOccupant"
If MsgBox("�������� ������ ����������������?", vbYesNo) = vbYes Then
n = 0
Rs_Add.MoveFirst
Do While Not Rs_Add.EOF
If Rs_Add.Fields("Oldnum").Value <> "" Then N1 = Val(Rs_Add.Fields("Oldnum").Value)
If N1 > n Then
n = N1
End If
Rs_Add.MoveNext
Loop
Rs_Add.AddNew
Rs_Add.Fields("oldnum").Value = n + 1
nNum = Rs_Add.Fields("Numer").Value
Rs_Add.Fields("Priv").Value = "��"
Rs_Add.Update
Filter.ad = 1
Kvart.Show
Kvart.Text4.Enabled = True
Filter.Hide
End If
End Sub



Private Sub fg_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim nr As Long, nc As Long      '��� ������ �������� ���� ����������� � ������ � �������
    
    On Error GoTo ex
    Static R As Long, c As Long     '��� �� ���������� ��� �������� ������� ������
    nr = Fg.MouseRow:    nc = Fg.MouseCol  ' get coordinates
    
    If nr < 1 Or nc = -1 Then
    Fg.ToolTipText = ""
    Exit Sub
    End If
    If c <> nc Or R <> nr Then                   ' update tooltip text
        
       If Fg.TextMatrix(nr, nc) <> "" Then
        Fg.ToolTipText = Fg.TextMatrix(nr, nc)
        End If
        R = nr:            c = nc
        DoEvents
    End If
ex:
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)


    ' ������ �������, ����� ��������������
    If Row = 1 Then
    Fg.FlexDataSource = m_DS
    Command15.Visible = True
    Command17.Visible = True
    End If
    
    'If Col = 12 Then
    'FG.FlexDataSource = m_DS
    'End If
    
End Sub

Private Sub FG_DblClick()



If Fg.Col = 12 Then
If Fg.Row = 1 Then Exit Sub
Filter.Enabled = False
Dogovor.Show
Exit Sub
End If




If Filter.Nm = "" Or Fg.Row = 0 Then
MsgBox ("�� �� ������� ����������������")
Else
Filter.Nm = Fg.Cell(flexcpText, Fg.Row, 0)
MainForm.Fnum = Filter.Nm

FIO = Fg.Cell(flexcpText, Fg.Row, 1) + " " + Fg.Cell(flexcpText, Fg.Row, 2) + "  " + Fg.Cell(flexcpText, Fg.Row, 3)
Lic.Caption = " " + Filter.FIO + " ��." + Fg.Cell(flexcpText, Fg.Row, 5) + " ��� �" + Fg.Cell(flexcpText, Fg.Row, 6) + " ��.� " + Fg.Cell(flexcpText, Fg.Row, 9)

Lic.Show
Filter.Enabled = False
End If
End Sub

Private Sub FG_EnterCell()

' ���������� nm ����� ��������� ����� ������� ��������� ������
  'nm = FG.Cell(flexcpText, FG.Row, FG.Col)
Filter.Caption = Fg.Cell(flexcpText, Fg.Row, 1) + " " + Fg.Cell(flexcpText, Fg.Row, 2) + "  " + Fg.Cell(flexcpText, Fg.Row, 3) + " � �/�� " + Fg.Cell(flexcpText, Fg.Row, 11)

PopUp.Label5.Caption = Fg.Cell(flexcpText, Fg.Row, 2) + "  " + Fg.Cell(flexcpText, Fg.Row, 3) + " " + Fg.Cell(flexcpText, Fg.Row, 4) + " � �/�� " + Fg.Cell(flexcpText, Fg.Row, 11)


'PopUp.Label6.Caption = " " + FG.Cell(flexcpText, FG.Row, 11)

  ' nm ������������� �������� ������ ��������� ������
Nm = Fg.Cell(flexcpText, Fg.Row, 0)

'MsgBox (nm)

'FIO = m_DS.m_RS.Fields("�������").Value + " " + m_DS.m_RS.Fields("���").Value + " " + m_DS.m_RS.Fields("��������").Value
'Filter.Caption = "������� ������-> " + FIO

End Sub


Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


If Row <> 1 Then Cancel = True


    '��� ����� �������, �� ������ ��������� ������������� ����� �������
   'If Row <> 1 Then
    'If Col <> 6 And Row <> 1 Then Cancel = True
    'End If
    Fg.Editable = flexEDKbdMouse
    
    
    If Fg.Cell(flexcpChecked, Row, 7) = flexChecked Then

Fg.Cell(flexcpChecked, Row, 7) = flexUnchecked
Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = False WHERE (((MainOccupant.Numer)=" + Fg.TextMatrix(Row, 0) + "))")
GoTo N1
       End If
       

If Fg.Cell(flexcpChecked, Row, 7) = flexUnchecked Then
Fg.Cell(flexcpChecked, Row, 7) = flexChecked
Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = True WHERE (((MainOccupant.Numer)=" + Fg.TextMatrix(Row, 0) + "))")

End If
N1:

End Sub









Private Sub Form_Load()
FSize = Fg.Font.Size
'Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
    
 
Filter.ad = 0

'Filter.Caption = FIO
'////////////////////////////////////////////////////

Fg.AutoSearch = flexSearchFromCursor
Fg.ExplorerBar = flexExSortShowAndMove



    ' ��������������� ����� (��������������)
    Fg.FixedCols = 0
    Fg.Editable = flexEDKbdMouse
    Fg.BackColorFrozen = RGB(200, 255, 200)
        
    'FG.Cell(flexcpFontSize, 1, 1, FG.Rows - 1, 1) = 8
    Fg.Cell(flexcpFontSize, 1, 0, Fg.Rows - 1, 1) = 8
    Fg.Cell(flexcpFontSize, 1, 11, Fg.Rows - 1, 11) = 8
    
    Fg.Cell(flexcpFontBold, 1, 11, Fg.Rows - 1, 11) = False
    Fg.Cell(flexcpFontBold, 1, 0, Fg.Rows - 1, 1) = False
    
    ' ����������� ������������� ������, ����� �������� ����������� �
    ' �������� ���� ���� � ������� 0:
    ' ��� �������: ��� ������������� ���������� ������ �������
    

    
    
    
   If MainForm.Pok = 1 Then Fg.ColHidden(0) = False
   If MainForm.Pok = 1 Then Fg.ColHidden(1) = True
   
   If MainForm.Pok = 0 Then Fg.ColHidden(1) = False
   If MainForm.Pok = 0 Then Fg.ColHidden(0) = True
   
' ���� �� ���� ���������� �������� � ���������

   If MainForm.Dog = 0 Then
   Fg.ColHidden(12) = True
   Check2.Visible = False
   End If
   
    ' �������� �������� ������ �������� ������
    Set m_DS = New FlexADO
 ' ��������� ���� � �����
    
    Fg.FlexDataSource = m_DS
    
    
    Fg.FrozenRows = 1
    
    Fg.DataMode = flexDMBoundBatch
    
    
    

   

    ' C�������, �������� ����������� ��� ���������� � ���� ����� �� ��������
    ' �� �� ������ ������
    'FG.AllowUserResizing = flexResizeBoth
    'FG.ExtendLastCol = True
    'FG.ExplorerBar = flexExSortShowAndMove
    'FG.AutoSearch = flexSearchFromCursor
    
    'FG.Cell(flexcpChecked, 2, 6, FG.Rows, 6) = flexUnchecked
  
  
  ������
  
    


Fg.Row = CL5
End Sub
'////////// ����������

Private Sub Form_Resize()
    On Error Resume Next
    Fg.Move Fg.Left, Fg.Top, ScaleWidth - Fg.Left * 2, ScaleHeight - Fg.Left - Fg.Top
End Sub

Private Sub �����_Click()
'Me.Enabled = False


Dni.Show
End Sub

Private Sub �����_Click()
Jdite.Show
Jdite.Label1.FontSize = 26
Jdite.Label1.Caption = Fg.TextMatrix(Fg.Row, 11)
End Sub

Private Sub ���_Click()
Command6_Click
End Sub

Private Sub ������_Click(Index As Integer)
Command3_Click
End Sub

Private Sub �������_Click()
Command2_Click
End Sub

Public Sub ����_Click()


Set infRS = New ADODB.Recordset
Set infRS.ActiveConnection = Mconn
infRS.Open ("SELECT MainOccupant.Numer, MainOccupant.OLDNUM, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.NLODGER, MainOccupant.NLODGERF, MainOccupant.NLODLIFT, MainOccupant.NROOM, MainOccupant.COMSPACE, MainOccupant.HABSPACE, MainOccupant.HABITATE, MainOccupant.BIRTHDAY, MainOccupant.NORDER, MainOccupant.KITCHSPACE, MainOccupant.BATHSPACE, MainOccupant.CORRSPACE, MainOccupant.TOILSPACE, MainOccupant.BALCSPACE, MainOccupant.DATARECEIV, MainOccupant.PASSPORT, MainOccupant.TELEPHONE, MainOccupant.LDOK, MainOccupant.LDATEBEG, MainOccupant.LDATEEND, MainOccupant.NAPARTMENT, MainOccupant.FLOOR, TipDom.Name_Dom, TipKv.Name_Kv FROM TipKv INNER JOIN (TipDom INNER JOIN MainOccupant ON TipDom.��� = MainOccupant.DomTip) ON TipKv.��� = MainOccupant.KV WHERE (((MainOccupant.Numer)=" + Filter.Nm + "))"), Mconn, adOpenKeyset, adLockPessimistic


Information.Show
End Sub

Private Sub ���������������_Click()
CL5 = Fg.Row
Fg.FlexDataSource = m_DS
Command5_Click
End Sub

Private Sub �_���������_Click(Index As Integer)
Dim AboutBox As New AboutBox
With AboutBox
    .Title = " ������ � ������ ������������ �������� ���������"
    .Version = "������: " + Str(App.Major) + "." + Str(App.Minor) + "." + Str(App.Revision)
    .Company = "���������� +  (C) Copyright, 2005, ���������" + vbNewLine
    .Copyright = " ������� ������ ������������"
    .Description = "����������� ������������� ��������"
    .License = "����� � ������� E-Mail:bestonline@list.ru ��������:+79881733600"
    .hWndOwner = Me.hwnd
    'Set .Icon = Me.Icon
    .AboutBox
End With
End Sub

Private Sub ��������1_Click()
Reports.sq = ""
Unload Reports
Analizlgot.Titl = "�/�� ��� ����������. �� " + MainMenu.Command13.Caption

Analizlgot.G = 7
Analizlgot.StrSQL = "SELECT MainOccupant.Numer AS �, KLS_PODR.NAIM_KLS AS �����, MainOccupant.kv_num AS ��, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS �������� FROM KLS_PODR INNER JOIN (MainOccupant LEFT JOIN AdNach ON MainOccupant.Numer = AdNach.KodKv) ON KLS_PODR.��� = MainOccupant.Dom WHERE (((AdNach.KodKv) Is Null))"
Analizlgot.�� 0

Unload Me
Analizlgot.Show
End Sub

Private Sub ��������2_Click()
Reports.sq = ""
Unload Reports
Analizlgot.Titl = "������� ������"
Analizlgot.G = 10
Analizlgot.StrSQL = "SELECT Adding.KodKv, MainOccupant.FAM AS �������, MainOccupant.IM AS ���, MainOccupant.OT AS ��������, MainOccupant.kv_num AS [�� �], KLS_PODR.NAIM_KLS AS �����, Adding.KodN AS [��� ���], Adding.NameN AS ����������, Adding.SummaI, Adding.Key AS ����� FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((Adding.KodKv) In (SELECT [KodKv] FROM [Adding] As Tmp GROUP BY [KodKv],[KodN],[KodKat],[SummaI] HAVING Count(*)>1  And [KodN] = [Adding].[KodN] And [KodKat] = [Adding].[KodKat] And [SummaI] = [Adding].[SummaI])) AND ((Adding.Tip)=" + Chr(34) + "+" + Chr(34) + ")) ORDER BY Adding.KodKv, Adding.KodN, Adding.NameN, Adding.SummaI"
Analizlgot.�� 0

Unload Me
Analizlgot.Show
End Sub

Private Sub ������_Click()
FG_DblClick
End Sub

Private Sub ������_Click()
Dim rsProv As ADODB.Recordset

Set rsProv = New ADODB.Recordset
rsProv.Open ("SELECT Adding.KodKv,Adding.key FROM Adding LEFT JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer WHERE (((MainOccupant.Numer) Is Null))"), Mconn, adOpenKeyset, adLockPessimistic
If rsProv.RecordCount > 0 Then
rsProv.MoveFirst
Do While Not rsProv.EOF
Mconn.Execute ("DELETE Adding.KodKv From Adding WHERE (((Adding.key)=" + Str(rsProv("key")) + "))")
rsProv.MoveNext
Loop
End If
'If Arhiv = True Then

Analizlgot.Titl = "��������� ��������� �� " + MainMenu.Command13.Caption
'+ " " + Str(Year(MainForm.DR))
'Else
'Analizlgot.Titl = "��������� ��������� �� " + MonthName(Month(MainForm.DR), False) + " " + Str(Year(MainForm.DR))
'End If
Analizlgot.Vid = "����"

Analizlgot.G = 8
'sq = "SELECT Adding.NameKat as [��������� ����������], KLS_PODR.NAIM_KLS as �����, IIf([Adding]![Kol]=0,0,round([Adding]![SaldoN]/[Adding]![Kol],2)) as [������ �� ������], IIf([Adding]![Tip]=" & Chr(34) & "+" & Chr(34) & ",round([Adding]![SummaI],2),0) AS ���������, IIf([Adding]![Tip]=" & Chr(34) & "-" & Chr(34) & ",round([Adding]![SummaI],2),0) AS ��������, IIf([Adding]![Tip]=" & Chr(34) & "s" & Chr(34) & ",round([Adding]![SummaI],2),0) AS ��������, IIf([Adding]![Kol]=0,0,round([Adding]![SaldoK]/[Adding]![Kol],2)) as [������ ��������] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.��� = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer ORDER BY Adding.NameKat"
'sq = "SELECT Adding.NameKat as [��������� ����������], ' ' as ' ',IIf([Adding]![Kol]=0,0,round([Adding]![SaldoN]/[Adding]![Kol],2)) as [������ �� ������], IIf([Adding]![Tip]=" & Chr(34) & "+" & Chr(34) & ",round([Adding]![SummaI],2),0) AS ���������, IIf([Adding]![Tip]=" & Chr(34) & "-" & Chr(34) & ",round([Adding]![SummaI],2),0) AS ��������, IIf([Adding]![Tip]=" & Chr(34) & "s" & Chr(34) & ",round([Adding]![SummaI],2),0) AS ��������, IIf([Adding]![Kol]=0,0,round([Adding]![SaldoK]/[Adding]![Kol],2)) as [������ ��������] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.��� = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer ORDER BY Adding.NameKat"
'sq = "SELECT Adding.NameKat as [��������� ����������], ' ' as _,IIf([Adding]![Kol]=0,0,[Adding]![SaldoN]/[Adding]![Kol]) as [������ �� ������], IIf([Adding]![Tip]=" & Chr(34) & "+" & Chr(34) & ",[Adding]![SummaI],0) AS ���������, IIf([Adding]![Tip]=" & Chr(34) & "-" & Chr(34) & ",[Adding]![SummaI],0) AS ��������, IIf([Adding]![Tip]=" & Chr(34) & "s" & Chr(34) & ",[Adding]![SummaI],0) AS ��������, IIf([Adding]![Kol]=0,0,[Adding]![SaldoK]/[Adding]![Kol]) as [������ ��������] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.��� = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer ORDER BY Adding.NameKat"
Reports.sq = "SELECT ' ' AS _, Adding.NameKat AS [��������� ����������], Sum(IIf([Adding]![Kol]=0,0,[Adding]![SaldoN]/[Adding]![Kol])) AS [������ �� ������], Sum(IIf([Adding]![Tip]='+',[Adding]![SummaI],0)) AS ���������, Sum(IIf([Adding]![Tip]='-',[Adding]![SummaI],0)) AS ��������, Sum(IIf([Adding]![Tip]='s',[Adding]![SummaI],0)) AS ��������, Sum(IIf([Adding]![Kol]=0,0,[Adding]![SaldoK]/[Adding]![Kol])) AS [������ ��������] FROM Adding LEFT JOIN (KLS_PODR RIGHT JOIN MainOccupant ON KLS_PODR.��� = MainOccupant.Dom) ON Adding.KodKv = MainOccupant.Numer GROUP BY Adding.NameKat, ' ' ORDER BY Adding.NameKat"
Analizlgot.�� 1

Analizlgot.Show


Analizlgot.Label1.Visible = True
'Label2.Visible = True

Analizlgot.Ok = Val(Str(Analizlgot.fg1.TextMatrix(1, 3))) + Val(Str(Analizlgot.fg1.TextMatrix(1, 4))) - Val(Str(Analizlgot.fg1.TextMatrix(1, 5))) - Val(Str(Analizlgot.fg1.TextMatrix(1, 6))) - Val(Str(Analizlgot.fg1.TextMatrix(1, 7)))

Analizlgot.Label1.Caption = Analizlgot.fg1.TextMatrix(1, 3) + " + " + Analizlgot.fg1.TextMatrix(1, 4) + " - " + Analizlgot.fg1.TextMatrix(1, 5) + " - " + Analizlgot.fg1.TextMatrix(1, 6) + " - " + Analizlgot.fg1.TextMatrix(1, 7) + " = " + Str(Round(Analizlgot.Ok, 2))
'Label2.Caption = Str(Ok)

If Round(Analizlgot.Ok, 2) <> 0 Then
Analizlgot.Command6.Visible = True
Analizlgot.Command7.Visible = True

'FG1.Cell(flexcpFontBold, 2, 1, 2, FG1.Cols - 1) = True
'FG1.Cell(flexcpBackColor, 2, 1, 2, FG1.Cols - 1) = vbRed

Else

osh:
If Err.Number <> 0 Then
MsgBox Err.Description
Err.Clear
End If




End If

End Sub

Private Sub �����_Click()
Command1_Click
End Sub

 Private Sub �������1_Click(Index As Integer)


If MsgBox("�� ������ ������� ������� ���� �" + Fg.TextMatrix(Fg.Row, 0) + " ������������� ��������������� " + Fg.TextMatrix(Fg.Row, 1) + "  " + Fg.TextMatrix(Fg.Row, 2) + "  " + Fg.TextMatrix(Fg.Row, 3) + "? ", vbYesNo) = vbYes Then
'Filter.Hide
Unload Me
If MsgBox("��� ������ ��������� � ���� ������� ������ ����� ������� ��� ����������� ��������������. �� �������", vbYesNo) = vbYes Then


Mconn.Execute ("DELETE Adding.KodKv, Adding.* From Adding WHERE (((Adding.KodKv)=" + Filter.Nm + "))")
Mconn.Execute ("DELETE Constanta.Numer, Constanta.KodNach, Constanta.NameNach From Constanta WHERE (((Constanta.Numer)=" + Filter.Nm + "))")
Mconn.Execute ("DELETE Lgota.NomNum, Lgota.* From Lgota WHERE (((Lgota.NomNum)=" + Filter.Nm + "))")
'Mconn.Execute ("DELETE OtheOwner.Numer, OtheOwner.* From OtheOwner WHERE (((OtheOwner.Numer)=" + Filter.Nm + "))")
Mconn.Execute ("DELETE tmp_lgota.KodKv, tmp_lgota.* From tmp_lgota WHERE (((tmp_lgota.KodKv)=" + Filter.Nm + "))")

Mconn.Execute ("DELETE MainOccupant.Numer, MainOccupant.* From MainOccupant WHERE (((MainOccupant.Numer)=" + Filter.Nm + "))")


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
Private Sub �������()
Dim nu As Integer

For R = 2 To Fg.Rows - 1
nu = Val(Fg.TextMatrix(R, 5))

MsgBox (Str(nu))

If Fg.Cell(flexcpChecked, R, 7) = flexChecked Then



MsgBox (Fg.Cell(flexcpChecked, R, 7))
Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = True WHERE (((MainOccupant.Numer)=" + Nm + "))")
Else
If Fg.Cell(flexcpChecked, R, 7) = flexUnchecked Then Mconn.Execute ("UPDATE MainOccupant SET MainOccupant.otm = False WHERE (((MainOccupant.Numer)=" + Str(nu) + "))")
End If
Next
End Sub
Private Sub ������()
For R = 2 To Fg.Rows - 1

'MsgBox (FG.Cell(flexcpChecked, r, 6))

If Fg.TextMatrix(R, 8) = False Then
Fg.Cell(flexcpChecked, R, 7) = flexUnchecked
GoTo n
End If

If Fg.TextMatrix(R, 8) = True Then
Fg.Cell(flexcpChecked, R, 7) = flexChecked
End If
n:
Next
End Sub

Private Sub ������_Click()
 'FG.Row = 1
 SendKeys "{Enter}"
End Sub

Private Sub �������_Click()
Command14_Click
End Sub
