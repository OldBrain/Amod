VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form LicArh 
   BackColor       =   &H00404080&
   Caption         =   "������� ����"
   ClientHeight    =   8028
   ClientLeft      =   132
   ClientTop       =   816
   ClientWidth     =   10752
   ControlBox      =   0   'False
   Icon            =   "LicArh.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8028
   ScaleWidth      =   10752
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000004&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   720
      Width           =   492
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000004&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   480
      Width           =   492
   End
   Begin VB.CommandButton Command14 
      Height          =   612
      Left            =   6600
      Picture         =   "LicArh.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5760
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H80000018&
      Caption         =   "�������"
      Height          =   615
      Left            =   7200
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "����������� �����"
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Height          =   735
      Left            =   9480
      Picture         =   "LicArh.frx":12038
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Height          =   615
      Left            =   6000
      Picture         =   "LicArh.frx":12152
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "����������� �����"
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Height          =   495
      Left            =   9960
      Picture         =   "LicArh.frx":1245C
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "�������� ������"
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H008080FF&
      Caption         =   "��������� ������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2160
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command9 
      Height          =   615
      Left            =   5520
      Picture         =   "LicArh.frx":1287D
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "���������� ����������"
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Height          =   615
      Left            =   5040
      Picture         =   "LicArh.frx":12B87
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "������ �������� ����� F5"
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Height          =   615
      Left            =   4560
      Picture         =   "LicArh.frx":12FC9
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "������ Ctrl-L"
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Height          =   735
      Left            =   9120
      Picture         =   "LicArh.frx":13405
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "�������������, ��������"
      Top             =   240
      Width           =   372
   End
   Begin VB.CommandButton Command6 
      Caption         =   "���."
      Height          =   735
      Left            =   8760
      Picture         =   "LicArh.frx":13847
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "���������� �� ����������"
      Top             =   240
      Width           =   372
   End
   Begin VSFlex8Ctl.VSFlexGrid VS 
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   6120
      Width           =   4335
      _cx             =   7646
      _cy             =   2355
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"LicArh.frx":13D9F
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
   Begin VB.CommandButton Command4 
      Caption         =   "��."
      Enabled         =   0   'False
      Height          =   735
      Left            =   7080
      Picture         =   "LicArh.frx":13E10
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "�������"
      Top             =   240
      Width           =   372
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���."
      Enabled         =   0   'False
      Height          =   735
      Left            =   6720
      Picture         =   "LicArh.frx":14252
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "�������� ����� ����������"
      Top             =   240
      Width           =   372
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   1
      EndProperty
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
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7440
      Picture         =   "LicArh.frx":148BC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "���������"
      Top             =   240
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   10200
      Picture         =   "LicArh.frx":14A11
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "������� F12"
      Top             =   240
      Width           =   495
   End
   Begin VSFlex8Ctl.VSFlexGrid fg1 
      Height          =   4695
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   10695
      _cx             =   18865
      _cy             =   8281
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
      BackColorSel    =   -2147483645
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483647
      GridColorFixed  =   4194432
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   4
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   45
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"LicArh.frx":14E53
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
      MergeCompare    =   1
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
      ShowComboButton =   0
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   3
      VirtualData     =   0   'False
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
      AccessibleRole  =   50
      Begin VSFlex8Ctl.VSFlexGrid V1 
         Height          =   1815
         Left            =   0
         TabIndex        =   31
         Top             =   2760
         Visible         =   0   'False
         Width           =   10455
         _cx             =   18441
         _cy             =   3201
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         FormatString    =   ""
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
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   8280
      TabIndex        =   41
      Top             =   240
      Width           =   492
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   732
      Left            =   7560
      TabIndex        =   34
      Top             =   5760
      Width           =   2412
   End
   Begin VB.Label Label18 
      BackColor       =   &H00400000&
      Caption         =   "����������� �����������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   7560
      Width           =   10335
   End
   Begin VB.Line Line22 
      X1              =   6840
      X2              =   7800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line19 
      X1              =   9120
      X2              =   9120
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line18 
      X1              =   10440
      X2              =   10440
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line17 
      X1              =   10440
      X2              =   4560
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line16 
      X1              =   4560
      X2              =   4560
      Y1              =   7080
      Y2              =   6720
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������ ���."
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
      Left            =   9120
      TabIndex        =   26
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���./��������"
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
      Left            =   7680
      TabIndex        =   25
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���������"
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
      Left            =   6000
      TabIndex        =   24
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������ ���."
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
      Left            =   4560
      TabIndex        =   23
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Line Line15 
      X1              =   9120
      X2              =   9120
      Y1              =   7440
      Y2              =   6840
   End
   Begin VB.Line Line14 
      X1              =   7680
      X2              =   7680
      Y1              =   7440
      Y2              =   6480
   End
   Begin VB.Line Line13 
      X1              =   6000
      X2              =   6000
      Y1              =   6480
      Y2              =   7440
   End
   Begin VB.Line Line12 
      X1              =   10440
      X2              =   10440
      Y1              =   6840
      Y2              =   7440
   End
   Begin VB.Line Line11 
      X1              =   4560
      X2              =   4560
      Y1              =   7440
      Y2              =   6480
   End
   Begin VB.Line Line10 
      X1              =   4560
      X2              =   10440
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line9 
      X1              =   10440
      X2              =   4560
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   9240
      TabIndex        =   22
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
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
      Left            =   7800
      TabIndex        =   21
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��������� ����������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5760
      Width           =   4335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "��������� ������"
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
      Left            =   9480
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line8 
      X1              =   5040
      X2              =   5040
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line Line7 
      X1              =   3240
      X2              =   3240
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line Line6 
      X1              =   1440
      X2              =   1440
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line Line4 
      X1              =   6720
      X2              =   6720
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6720
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6720
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������ ���."
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
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������ ���."
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
      Left            =   5040
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���������"
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
      Left            =   1440
      TabIndex        =   9
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���������"
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
      Left            =   1485
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������/��������"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��������� ����������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6735
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ��������������� 
         Caption         =   "������� ��������"
         Shortcut        =   {F9}
      End
      Begin VB.Menu ��������� 
         Caption         =   "���������"
         Shortcut        =   {F2}
      End
      Begin VB.Menu ���������_��������� 
         Caption         =   "���������"
         Shortcut        =   ^N
      End
      Begin VB.Menu ������_���_����� 
         Caption         =   "������ ��� �����"
         Shortcut        =   {F5}
      End
      Begin VB.Menu ����������_���������� 
         Caption         =   "���������� ����������"
         Shortcut        =   {F3}
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
         Shortcut        =   ^L
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Shortcut        =   ^O
      End
      Begin VB.Menu ����� 
         Caption         =   "�����"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu �������������� 
      Caption         =   "��������������"
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Shortcut        =   ^X
      End
      Begin VB.Menu ��������_���������� 
         Caption         =   "�������� ����������"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu ������_��_������ 
         Caption         =   "������ �� ������"
         Shortcut        =   ^S
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
         Shortcut        =   ^Q
      End
      Begin VB.Menu �������_���������� 
         Caption         =   "������� ����������"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu K���������� 
         Caption         =   "������ �����������"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ������������ 
         Caption         =   "������� � ����"
      End
      Begin VB.Menu ������������ 
         Caption         =   "�������� ����"
      End
      Begin VB.Menu ���������� 
         Caption         =   "����������"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu ������ 
      Caption         =   "������"
      Begin VB.Menu ������� 
         Caption         =   "��� ����"
      End
      Begin VB.Menu �������� 
         Caption         =   "��������� ��� ����� ������"
      End
      Begin VB.Menu ������ 
         Caption         =   "��������� ��� �����������"
      End
      Begin VB.Menu ��������� 
         Caption         =   "��������� ���"
      End
      Begin VB.Menu ������� 
         Caption         =   "������� � �������������"
      End
      Begin VB.Menu ���������� 
         Caption         =   "������ ��������� ���"
      End
   End
End
Attribute VB_Name = "LicArh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' � ������ � ��� ����� �������. ��������� ��� ������� ������� 1-�� �����

Option Explicit

Public Sal, Kat, Itog1, ops As Double
Public KODS, SPR As String
Public KODS_Kat, KODS_N, TMPI As Integer
'Public ������ As Label
Public InfoRS As ADODB.Recordset
Dim Zap1, Zap2
'Dim Ca As ADODB.connection
Dim Un As ADODB.Recordset
Dim FGS As Integer
Dim MainOc As ADODB.Recordset
Dim Perebor1 As ADODB.Recordset
Public Clik As Integer
Dim RS As ADODB.Recordset
Dim Inf As ADODB.Recordset
Dim Combo_RS, Dat, TMP_Lic, Perebor, GoodL, TMP As ADODB.Recordset
Dim Formula(100), Queri, Status As String
Dim NACH(999), OPL(999), Proc As Double
Dim Inf1 As ADODB.Recordset
'Dim rsSaldo As ADODB.Recordset
Dim F As String
Dim Q As String
Dim Qinf As String
Dim i As Integer
Dim Fg As Long
Dim rw As Long
Dim Cl As String
Dim nameRP As String
Dim K, j As Integer
Public Dolg As Currency
Dim FIO As String







Private Sub Check1_Click()

End Sub

Private Sub Command1_Click()

Unload Sch


For rw = 1 To fg1.Rows - 1
If fg1.TextMatrix(rw, 2) = "999" Then
MsgBox "��������� ��� ����������"
fg1.Row = rw
fg1.Col = 2
Exit Sub
End If
Next
����������
'Rs.UpdateBatch
MainForm.RSaldoK Filter.Nm


Unload Me
Filter.Enabled = True

End Sub

Private Sub Command10_Click()
Dim SaldoArh As ADODB.Recordset
Set SaldoArh = New ADODB.Recordset

Text1.Enabled = False
Status = "Text1"
If IsNumeric(Text1.Text) = False Then
If TMPI = 1 Then MsgBox ("��������� ����! ��� ���������� ����� � ������� ����� ����������� �������")
TMPI = TMPI + 1
Else
KODS_N = fg1.TextMatrix(fg1.Row, 2)
'Return_KODS_KAT
���������_������

SaldoArh.Open ("SELECT Saldo_Arh.KodKV, Saldo_Arh.KodKat, Saldo_Arh.SK From Saldo_Arh WHERE (((Saldo_Arh.KodKV)=" + Filter.Nm + ") AND ((Saldo_Arh.KodKat)=" + fg1.TextMatrix(FGS, 22) + "))"), Ca, adOpenKeyset, adLockPessimistic
              If SaldoArh.RecordCount = 0 Then
        

Ca.Execute ("INSERT INTO Saldo_Arh ( KodKV, KodKat, SK ) SELECT " + Filter.Nm + " AS ���������1, " + fg1.TextMatrix(FGS, 22) + " AS ���������2, " + Replace(Text1.Text, ",", ".") + " AS ���������3")
                          Else
Ca.Execute ("UPDATE Saldo_Arh SET Saldo_Arh.SK = " + Replace(Text1.Text, ",", ".") + " WHERE (((Saldo_Arh.KodKV)=" + Filter.Nm + ") AND ((Saldo_Arh.KodKat)=" + fg1.TextMatrix(FGS, 22) + "))")
                         End If
End If

 If IsNumeric(Text1.Text) = True Then SaldoArh.Close

Command10.Visible = False
fg1.Enabled = True
fg1.SetFocus

End Sub

Private Sub Command11_Click()
Dim Clik As Integer
Clik = 1
ViewArhiv Clik
Command11.Visible = False
End Sub

Private Sub Command12_Click()

Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long

nameRP = "lc"

'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = True

'������ ����� �������� � Word-e
'Set DocWord = WordApp.Documents.Add

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate

'��������� ��������� ��������
On Error GoTo est
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)
est:
 
If Err.Number = 5356 Then
Err.Clear
nameRP = Trim(Trim(nameRP) + Trim(Str(Int(Rnd() * 1000))))

DocWord.SaveAs (App.Path + "\Temp\" + nameRP + ".doc")
End If
'���������, ���� �� ��������� ��������� ��������� ��������� Saved � ���� ��������� �� ���� ��������� - ��������� ��;
'If DocWord.Saved = False Then DocWord.Save

WordApp.Options.CheckSpellingAsYouType = False


Set TableWord = DocWord.Tables(1)
'.Add(DocWord.Range(), 10, 2)


'�������� ����� � ������ � �������
'(�����_������, �����_�������)

TableWord.Cell(1, 2).Range.Text = MainForm.Label3
TableWord.Cell(2, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 5)
TableWord.Cell(2, 2).Range.Text = "�� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 9)

TableWord.Cell(1, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 11)
TableWord.Cell(2, 3).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 2) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 3) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 4)
TableWord.Cell(4, 1).Range.Text = "������ ���.��:" + MainForm.Label8 + "�."
TableWord.Cell(4, 2).Range.Text = Me.Label10

TableWord.Cell(5, 2).Range.Text = Me.Label11

' ������ �������� � ������
O9 = 0
S9 = 0
For rw = 1 To fg1.Rows - 1
If fg1.TextMatrix(rw, 23) = "-" Then O9 = O9 + fg1.TextMatrix(rw, 18)
If fg1.TextMatrix(rw, 23) = "s" Then S9 = S9 + fg1.TextMatrix(rw, 18)
Next



TableWord.Cell(6, 2).Range.Text = Str(O9)
TableWord.Cell(7, 2).Range.Text = Str(S9)
TableWord.Cell(8, 1).Range.Text = "������ ���.��:" + MainForm.Label8 + "�."
TableWord.Cell(8, 2).Range.Text = Me.Label13


Set DocWord = Nothing

'���������� ������ - Word
Set WordApp = Nothing


End Sub


Private Sub Command13_Click()
' ������ ������� arh_rep ��� ����� ������ �� �������� ������� �� �������� �������� �����
'Ca.Execute ("DELETE arh_rep.* FROM arh_rep")
Arc.Show
End Sub

Private Sub Command14_Click()
'������ arh_rep
Ca.Execute ("DELETE arh_rep.* FROM arh_rep")
'������ �������� FileListBox
MainForm.File1.Path = App.Path + "\data\Arhiv\"
MainForm.File1.Pattern = "*.amd"
'������� ���� ������ � �����
For j = 0 To MainForm.File1.ListCount - 1
'MsgBox (File1.List(i))
'������������ � ������ � �������� �� ������ ������ ������ �� Filter.Nm
�������� MainForm.File1.List(j), Filter.Nm, True, False
Next j

'��������� ������ �������� ������
Ca.Execute ("INSERT INTO arh_rep SELECT Adding.* FROM Adding IN '" + App.Path + "\Data\kvartplata.amd' WHERE (((Adding.KodKv)=" + Filter.Nm + "));")

Izv.Show
End Sub

Private Sub Command15_Click()
�������_Click
End Sub

Private Sub Command2_Click()
Dim Proc As Double
Dim rw As Integer

For rw = 1 To fg1.Rows - 1


If Me.fg1.TextMatrix(rw, 15) = "" Then
MsgBox "��� ���������� ������� ���������� ��������� �������." + vbNewLine + "���� ��� ������ ������� ������� ������������� ��, ���������� ���� �� ����"
Exit Sub
End If

MainForm.II = 0
MainForm.Pi = 0
MainForm.Ostatok = Me.fg1.TextMatrix(rw, 15)

'���� ������� �� ������� ����� ������� ��������� ��������
        
Ca.Execute ("UPDATE Adding SET Adding.ObPl = [Adding]![Shc_new]-[Adding]![Shc_old] WHERE (((Adding.Sch)='��') AND ((Adding.KodKv)=" + Filter.Nm + "))")

If Me.fg1.TextMatrix(rw, 30) = "��" Then

'���� ������� �� ������� ����� ������� ��������� ��������
Ca.Execute ("UPDATE Adding INNER JOIN TMP_LGOTA ON Adding.Key = TMP_LGOTA.UniKOd SET TMP_LGOTA.Plo = [Adding]![Shc_new]-[Adding]![Shc_old] WHERE (((Adding.Sch)='��') AND ((Adding.KodKv)=" + Filter.Nm + "))")
������ Me.fg1.TextMatrix(rw, 26)
End If

fg1.Row = rw

If Me.fg1.TextMatrix(rw, 6) = "" Then Me.fg1.TextMatrix(rw, 6) = MainForm.PrZ
Next rw

If MainForm.LGST = 1 Then Ca.Execute ("UPDATE Adding SET Adding.LgotaP = 1 where (adding.kodkv=" + Filter.Nm + ")")

SPR = ""
SposobR1.Show

'������������� ������� ������� �� ��������� ������ �����
'FG1.SetFocus


End Sub
Private Sub Command3_Click()
'Dim N, N1 As Integer
If MsgBox("�������� ����� ����������?", vbYesNo) = vbYes Then
Status = "��������"
��������



End If
End Sub
Private Sub Command4_Click()
K = 0
TMP.Open ("SELECT Adding.NameN, Adding.SummaI From Adding " + "Where(((Adding.KodKv) = " & Filter.Nm & ")" + " AND ((Adding.KodKat)=" + fg1.TextMatrix(fg1.Row, 22) + "))")
TMP.MoveFirst

Do While Not TMP.EOF
K = K + 1
TMP.MoveNext
Loop
TMP.Close
'MsgBox (Str(k))
If (Text1.Text = 0 Or K > 1) Then
If MsgBox("�� ������������� ������ ������� " + fg1.TextMatrix(fg1.Row, 3) + " �� " + fg1.TextMatrix(fg1.Row, 5) + "?", vbYesNo) = vbYes Then
Ca.Execute ("DELETE Adding.*, Adding.Key From Adding WHERE ((Adding.Key)=" + fg1.TextMatrix(fg1.Row, 26) + ")")
RS.Requery
fg1.Refresh
End If
Else
MsgBox ("������� ������ ! ���� �� ������� ������ �� ������ ������ = " + Text1.Text + " ���.")
End If
End Sub



Private Sub Command5_Click()
���������
'�����������
fg1.MergeCells = flexMergeRestrictAll
fg1.MergeCol(-1) = True
fg1.Refresh
fg1.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove
End Sub

Private Sub Command6_Click()





'

Set InfoRS = New ADODB.Recordset
Set InfoRS.ActiveConnection = Ca

InfoRS.CursorType = adOpenDynamic

InfoRS.LockType = adLockReadOnly






'Info.Show 1
Info.Caption = Me.Caption
Info.Label1 = Me.fg1.TextMatrix(Me.fg1.Row, 3)
Info.Label3 = Me.fg1.TextMatrix(Me.fg1.Row, 2)
'������ ��� ���
Info.Label2 = Me.fg1.TextMatrix(Me.fg1.Row, 21)
Info.Label6 = Me.fg1.TextMatrix(Me.fg1.Row, 20)
'��������� �������
Info.Label7 = Me.fg1.TextMatrix(Me.fg1.Row, 1)
'�����
Info.Label11 = Me.fg1.TextMatrix(Me.fg1.Row, 10)
'����������
Info.Label12 = Me.fg1.TextMatrix(Me.fg1.Row, 11)
'�������
Info.Label15 = Me.fg1.TextMatrix(Me.fg1.Row, 19)
'����� �������
Info.Label18 = Me.fg1.TextMatrix(Me.fg1.Row, 15)

'������ �������
Info.Label19 = Me.fg1.TextMatrix(Me.fg1.Row, 16)
'���������
Info.Label23 = Me.fg1.TextMatrix(Me.fg1.Row, 12)

'���������
Info.Label22 = Me.fg1.TextMatrix(Me.fg1.Row, 13)
'Key
Info.Label24 = Me.fg1.TextMatrix(Me.fg1.Row, 26)


MainForm.Pi = 0
'MainForm.Ostatok = me.fg1.TextMatrix(me.fg1.Row, 15)
MainForm.II = 0


If Me.fg1.TextMatrix(Me.fg1.Row, 30) = "��" Then ������ Me.fg1.TextMatrix(Me.fg1.Row, 26)


InfoRS.Open "SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.KodKls, tmp_lgota.NAME_KLS, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.Procent, tmp_lgota.Plo, tmp_lgota.Prop, tmp_lgota.Cocmin, tmp_lgota.OtheCode, tmp_lgota.parametr, tmp_lgota.itog, tmp_lgota.tarif, tmp_lgota.itog1,tmp_lgota.prim,tmp_lgota.plolg From tmp_lgota WHERE (((tmp_lgota.UniKOd)=" + Info.Label24 + ") AND ((tmp_lgota.KodKv)=" + Filter.Nm + " )) ORDER BY tmp_lgota.prim DESC"

Set Info.DG1.DataSource = InfoRS

For rw = 1 To Info.DG1.Rows - 1
If Info.DG1.TextMatrix(rw, 16) = 1 Then

Info.Label27 = Info.DG1.TextMatrix(rw, 15)
Info.DG1.Cell(flexcpBackColor, rw, 1, rw, Info.DG1.Cols - 1) = vbGreen
Else
Info.DG1.Cell(flexcpBackColor, rw, 1, rw, Info.DG1.Cols - 1) = vbBlue
End If
Next

Info.Show 1

End Sub

Private Sub Command7_Click()
������_Click
End Sub

Private Sub Command8_Click()
������_���_�����_Click
End Sub

Private Sub FG1_AfterDataRefresh()
����
���������
End Sub
Private Sub fg1_AfterMoveColumn(ByVal Col As Long, Position As Long)



' sort the data from first to last column
     fg1.Select 1, 0, 1, fg1.Cols - 1
     fg1.ColSort(1) = flexSortGenericAscending
     fg1.Select 1, 0
End Sub

Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
KODS_N = fg1.TextMatrix(fg1.Row, 2)
'////////////////////////////////
'��������
'////////////////////////////////

��������������
'��������_������
FGS = fg1.Row
�����������


'///////////////////////////////////
'Text1.Text = 0
'MsgBox (sal)
'FG1.TextMatrix(FG1.Row, 20) = Sal
'Text1.Text = Sal
'/////////////////////////

'���������_������

'/////////////////////////////////////
'Label1.Caption = FG1.TextMatrix(FG1.Row, 1)
'///////////////////////////////////
MainForm.��������
'//////////////////////////////////
����
Ca.Execute ("UPDATE Doc SET Doc.Summa = " + Str(fg1.TextMatrix(fg1.Row, 18)) + ", Doc.Stst = 1 WHERE (((Doc.Key)=" + fg1.TextMatrix(fg1.Row, 28) + "))")
End Sub

Private Sub fg1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)

'

On Error GoTo ������

'���������� ������� �����������
If fg1.TextMatrix(fg1.Row, 44) = 0 Then
Option1.Value = True
Option1.BackColor = &H80000004
Option1.Refresh

Option2.Value = False
Option2.BackColor = 0
Option2.Refresh
Else

Option1.Value = False
Option1.BackColor = 0
Option1.Refresh

Option2.Value = True
Option2.BackColor = &H80000004
Option2.Refresh

End If
'Option1.Value = FG1.TextMatrix(FG1.Row, 44)
'Option2.Value = FG1.TextMatrix(FG1.Row, 44)
'���������� �������
If fg1.TextMatrix(fg1.Row, 43) = "��" Then

Sch.Show
 Me.Command15.Visible = True
 
'Sch.Label20.Visible = True
'Sch.Label21.Visible = True
'Sch.Label22.Visible = True
'Sch.Label23.Visible = True
'Sch.Label24.Visible = True
'Sch.Label25.Visible = True

Sch.Label21.Caption = fg1.TextMatrix(fg1.Row, 42)
Sch.Label22.Caption = fg1.TextMatrix(fg1.Row, 41)
Sch.Label23.Caption = fg1.TextMatrix(fg1.Row, 42) - fg1.TextMatrix(fg1.Row, 41)
Sch.Label21.Refresh
Sch.Label22.Refresh
Sch.Label23.Refresh

Sch.Label10.Caption = Round((Me.fg1.TextMatrix(Me.fg1.Row, 42) - Me.fg1.TextMatrix(Me.fg1.Row, 41)) * Me.fg1.TextMatrix(Me.fg1.Row, 10), 2)
Sch.Label10.Refresh
Else
Me.Command15.Visible = False
Unload Sch
'Label20.Visible = False
'Label21.Visible = False
'Label22.Visible = False
'Label23.Visible = False
'Label24.Visible = False
'Label25.Visible = False

End If


Dim TMP_Lic As ADODB.Recordset
Set TMP_Lic = New ADODB.Recordset
If fg1.Row <> 0 Then
TMP_Lic.Open ("SELECT Adding.NameN, Adding.SummaI From Adding " + "Where(((Adding.KodKv) = " & Filter.Nm & ")" + " AND ((Adding.KodKat)=" + fg1.TextMatrix(fg1.Row, 22) + "))"), Ca, adOpenForwardOnly, adLockReadOnly
'WHERE (((Adding.KodKat)=1))"))
Set VS.DataSource = TMP_Lic
TMP_Lic.Close
Set TMP_Lic = Nothing
End If

'KODS = FG1.TextMatrix(FG1.Row, 22)
'KODS = 1
If Status <> "Text1" Then
Text1.Text = fg1.TextMatrix(fg1.Row, 20)
Text2.Text = fg1.TextMatrix(fg1.Row, 21)
Label1.Caption = fg1.TextMatrix(fg1.Row, 1)
'���������_��_����������
End If

��������_������
���������_��_����������

If fg1.TextMatrix(fg1.Row, 2) = "999" Or fg1.Col = 5 Then
fg1.Editable = 2
If fg1.TextMatrix(fg1.Row, 5) = "" Then fg1.TextMatrix(fg1.Row, 5) = MainForm.PeriodR


Else
fg1.Editable = 0
End If

Label9.Caption = Label1.Caption

������:
Select Case Err.Number
Case Is = 3021
MsgBox ("��� ����������. �� �������� ��������� ���������� ���������� ���������� (F<3>), ������� ������ �������������� ��� ������� ���������������� ��������� (�� ������ � �����)!")

Case Is = 0
Case Else
MsgBox (Err.Description)
End Select
If fg1.TextMatrix(fg1.Row, 38) <> "" Then
Label18.FontBold = True
Label18.Caption = fg1.TextMatrix(fg1.Row, 38)
Else
Label18.FontBold = False
Label18.Caption = "����������� �����������"
End If
Status = ""
End Sub



Private Sub FG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = Ca
Combo_RS.CursorType = adOpenForwardOnly
Combo_RS.LockType = adLockBatchOptimistic
'*******************************************************
'If (fg1.TextMatrix(0, fg1.Col)) = "���" Then
If fg1.Col = 2 Then
Combo_RS.Open "Nachisleniy"
Cl = ""
Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
'cl = cl + Combo_RS("Name_Kategor") + "|"
Cl = Cl + CStr(Combo_RS("Kod")) & vbTab & Combo_RS("Naim") + "|"
Combo_RS.MoveNext
Loop
fg1.ComboList = Cl
Combo_RS.Close
Else: fg1.ComboList = ""
End If
End Sub




Private Sub Fg1_DblClick()

If fg1.TextMatrix(fg1.Row, 23) = "+" Then
Command6_Click
End If

If fg1.TextMatrix(fg1.Row, 23) = "-" Then ArhoPL.Show 1



End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Command1_Click
End Sub



Private Sub FG1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Arhiv = True Then Cancel = True

If Col = 5 Or Col = 2 Or Col = 18 Or Col = 42 Then
fg1.Editable = 2
'fg1.TextMatrix(fg1.Row, 5) = MainForm.PeriodR
Else
Cancel = True
End If

End Sub

Private Sub Form_Load()

If Arhiv = True Then Command3.Enabled = False
If Arhiv = True Then Command4.Enabled = False
If Arhiv = True Then Command2.Enabled = False
If Arhiv = True Then Command11.Enabled = False
If Arhiv = True Then ����.Enabled = False
If Arhiv = True Then ��������������.Enabled = False

ops = 1
Me.Clik = 1
Dim Ras As ADODB.Recordset
'Set Ca = New ADODB.connection
 ' Ca.connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.amd;Persist Security Info=True"
 'Ca.Open "data/kvartplata.amd"
 
  Status = ""
  
  
  
Me.Caption = "����� " + Lic.Caption
  F = Filter.Nm
  
  F = MainForm.Fnum
  
  'If MainForm.arhivBooton = True Then F = 1
  
  Q = "SELECT Adding.KodKv AS �����, Adding.NameKat as ������, Adding.KodN AS ������, Adding.NameN AS ���_���, Adding.DataT AS �_������, Adding.DataR AS �_����, Adding.LgotaP AS �_�������, Adding.LgotaVid AS �_���, Adding.LgotaUSE AS �_USE, Adding.LgotaKod AS �_���, Adding.Tarif AS �����, Adding.Socmin AS ������, Adding.Propis AS ���������, Adding.Projiv AS ��������, Adding.ProLift AS �������, Adding.ObPl AS ��_���, Adding.PolPl AS ��_���, Adding.SummaB AS ����, Adding.SummaI AS ����, Adding.Formula, Adding.SaldoN, Adding.SaldoK, "
  Q = Q + "Adding.KodKat AS ������, Adding.Tip, Adding.TipKvKod, Adding.TipDomKod, Adding.Key, Adding.ispr,Adding.KodDoc, Adding.Parametr, Adding.Lig, Adding.OtheKol, Adding.OtheKol, Adding.TarifI, Adding.TarifD, Adding.SchetZ, Adding.FLOOR, Adding.kol, Adding.com, Adding.FormulaB, Adding.SummaBl, Adding.Shc_old, Adding.Shc_new, Adding.Sch, adding.KodConstanta FROM Adding "
  Q = Q + "WHERE (((Adding.KodKv)=" & F & "))ORDER BY Adding.KodKat"
  Qinf = "SELECT Max(Adding.SaldoN) AS [Max-SaldoN], Adding.NameKat, Adding.Tip, Adding.KodKat, Sum(Adding.SummaI) AS [Sum-SummaI], Max(Adding.SaldoK) AS [Max-SaldoK] From Adding Where (((Adding.KodKv) =" & F & ")) GROUP BY Adding.NameKat, Adding.Tip, Adding.KodKat"

  
  
Set Inf1 = New ADODB.Recordset
Set Inf1.ActiveConnection = Ca
  
'Set rsSaldo = New ADODB.Recordset

  
Set RS = New ADODB.Recordset
Set RS.ActiveConnection = Ca

RS.CursorType = adOpenForwardOnly
RS.LockType = adLockBatchOptimistic

Set TMP = New ADODB.Recordset
Set TMP.ActiveConnection = Ca
RS.CursorType = adOpenKeyset
RS.LockType = adLockPessimistic

Set MainOc = New ADODB.Recordset
Set MainOc.ActiveConnection = Ca

MainOc.CursorType = adOpenForwardOnly

Set Perebor = New ADODB.Recordset
Set Perebor.ActiveConnection = Ca

Set Perebor1 = New ADODB.Recordset
Set Perebor1.ActiveConnection = Ca

Set TMP_Lic = New ADODB.Recordset
Set TMP_Lic.ActiveConnection = Ca
TMP_Lic.CursorType = adOpenForwardOnly
TMP_Lic.LockType = adLockBatchOptimistic


RS.CursorType = adOpenForwardOnly
RS.LockType = adLockBatchOptimistic


Set Inf = New ADODB.Recordset
Set Inf.ActiveConnection = Ca

If F = "" Or F = "�����" Then Exit Sub


RS.Open (Q)



Perebor.Open (Q)
Inf.Open (Qinf)

Inf1.Open ("SELECT MainOccupant.Numer, TipDom.Name_Dom, TipKv.Name_Kv FROM ((KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.��� = MainOccupant.Dom) INNER JOIN TipDom ON KLS_PODR.Tip = TipDom.���) INNER JOIN TipKv ON MainOccupant.KV = TipKv.��� WHERE (((MainOccupant.Numer)=" + F + "))")




If Not Inf1.EOF Then Label19 = Inf1("Name_Dom") + vbNewLine + Inf1("Name_Kv")

Inf1.Close


Set Dat = New ADODB.Recordset
Set Dat.ActiveConnection = Ca
Dat.CursorType = adOpenForwardOnly
Dat.LockType = adLockBatchOptimistic
Dat.Open ("Settings")
Label8.Caption = Dat.Fields("TekData").Value

'Set Saldo0 = New ADODB.Recordset
'Set Saldo0.ActiveConnection = Ca
'Saldo0.CursorType = adOpenForwardOnly
'Saldo0.LockType = adLockBatchOptimistic

'************ ������ (������)***************************

' ��������� ������� Nach() � OPL(), ��� ������ �������� ����� ���� ��������� ������� "KodKat"
'//////////////////////////////////////
'��������
'////////////////////////////////////
Erase NACH, OPL




'If Inf.EOF = False Then Inf.MoveFirst
'Do While Not Inf.EOF
'If Inf.Fields("Tip").Value = "+" Then NACH(Inf.Fields("KodKat").Value) = Inf.Fields("Sum-SummaI").Value
'If Inf.Fields("Tip").Value = "-" Or Inf.Fields("Tip").Value = "s" Then OPL(Inf.Fields("KodKat").Value) = Inf.Fields("Sum-SummaI").Value
'Inf.MoveNext
'Loop


'*************������(�����)*****************************




'= flexSortGenericAscending

fg1.ColSort(2) = flexSortGenericAscending



'RS.MoveLast
'MsgBox (RS.Fields.Count)

Set fg1.DataSource = RS
'Set TS.DataSource = Rs

��������_������

����������

'�����������
fg1.MergeCells = flexMergeRestrictAll
fg1.MergeCol(-1) = True
fg1.Refresh
fg1.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove


fg1.AutoResize = False
 'Sal = FG1.TextMatrix(FG1.Row, 20)
Sal = 0
'Sal = FG1.TextMatrix(FG1.Row, 22)
Text1.Text = Str(Sal)
On Error GoTo S1
Text2.Text = Round(fg1.TextMatrix(fg1.Row, 21), 2)
S1:
Text2.Refresh
fg1.FocusRect = 3
fg1.AutoSearch = flexSearchFromCursor
Exit Sub


������:
Select Case Err.Number
Case Is = 3021
MsgBox ("��� ����������. �� �������� ��������� ���������� ���������� ���������� (F<3>), ������� ������ �������������� ��� ������� ���������������� ��������� (�� ������ � �����)!")
'��������
Case Is = 0
Case Else
MsgBox (Err.Description)
End Select

'�������



End Sub

Private Sub Form_Unload(Cancel As Integer)
Filter.Enabled = True
ops = 0
End Sub

Private Sub K����������_Click()
'FG1.TextMatrix(FG1.Row, 38) = InputBox("", "����������", FG1.TextMatrix(FG1.Row, 38))
Comentariy.Text1 = fg1.TextMatrix(fg1.Row, 38)
Comentariy.Show

End Sub

Private Sub Label18_Click()
'FG1.TextMatrix(FG1.Row, 38) = InputBox("", "����������", FG1.TextMatrix(FG1.Row, 38))
Comentariy.Text1 = fg1.TextMatrix(fg1.Row, 38)
Comentariy.Show
End Sub


Private Sub Option1_Click()
fg1.TextMatrix(fg1.Row, 44) = 0
Option1.BackColor = &H80000004
Option2.BackColor = 0
'Value = "��" Then Option1.Value = True
End Sub

Private Sub Option2_Click()
fg1.TextMatrix(fg1.Row, 44) = 1
Option2.BackColor = &H80000004
Option1.BackColor = 0
End Sub

Private Sub Text1_GotFocus()
Status = "Text1"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'MsgBox (Str(KeyAscii))
If KeyAscii = 13 Then Command10_Click
If KeyAscii = 27 Then Exit Sub
End Sub

Private Sub Text1_LostFocus()
 'FG1.TextMatrix(FG1.Row, 20) = Text1.Text
 

End Sub



'Private Sub TS_Click()
'MsgBox ("Col= " + Str(TS.Col) + "  Row=" + Str(TS.Row))
'End Sub

Private Sub V1_Click()
Set V1.DataSource = RS
MsgBox V1.Col
End Sub

Private Sub �����_Click()
Command1_Click
End Sub

Private Sub ������_���_�����_Click()

'Filter.Nm = FG.Cell(flexcpText, FG.Row, 0)
'Filter.Hide
'Filter.FG.Clear
If Filter.Nm = "" Then
MsgBox ("�� �� ������� ����������������")
Else
Kvart1.Show
Kvart1.Caption = Me.Caption

'Filter.Hide
'm_DS.m_RS.Close
'm_DS.m_Ca.Close
End If
End Sub

Private Sub ��������_����������_Click()
Command3_Click
End Sub

Private Sub ������������_Click()

���������������
Exit Sub



Dim TabN As Double
Dim NaKod As Integer
Dim ������� As Integer
Dim Potom(1000) As Double

Jdite.Show
Jdite.Label1.Refresh

For i = 1 To 1000
Potom(i) = 0
Next


������� = Filter.Fg.TextMatrix(Filter.Fg.Row, 10)
NaKod = fg1.TextMatrix(fg1.Row, 2)
i = 0




If MsgBox("�������� " + fg1.TextMatrix(fg1.Row, 3) + " ��  ��� ������� ����� �� ������ " + Filter.Fg.TextMatrix(Filter.Fg.Row, 5) + " ��� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 6), vbYesNo, "") = vbYes Then
MainOc.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(�������) + "))")
MainOc.MoveFirst
Do While Not MainOc.EOF
i = i + 1
TabN = MainOc.Fields("Numer").Value
Potom(i) = TabN
' �������� ������ ���� ������
'Ca.Execute ("DELETE Adding.KodKv, Adding.KodN From Adding WHERE (((Adding.KodKv)=" + Str(Tabn) + ") AND ((Adding.KodN)=" + Str(NaKod) + "))")
Ca.Execute ("")
MainOc.MoveNext
Loop

For i = 1 To 1000
If Potom(i) <> 0 Then MainForm.���������������� Potom(i)
Next

MainOc.Close
End If


Unload Jdite
End Sub

Private Sub ���������_Click()
Dolg = Round(Text2.Text, 2)

If Dolg <= 0 Then
If MsgBox("�� ������ ��������� ������� ��������� " + Str((Text2.Text)) + ". ������ ��������� �� ������� ������ �� ���������." + vbNewLine + "���������� ������ ��������� ?", vbYesNo) = vbNo Then
Exit Sub
Else
Dolg = 0
End If
End If


Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim stn As String


nameRP = "Izv1"

'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = True

'������ ����� �������� � Word-e
'Set DocWord = WordApp.Documents.Add

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate

'��������� ��������� ��������
On Error GoTo est
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)
est:
 
If Err.Number = 5356 Then
Err.Clear
nameRP = Trim(Trim(nameRP) + Trim(Str(Int(Rnd() * 1000))))

DocWord.SaveAs (App.Path + "\Temp\" + nameRP + ".doc")
End If
'���������, ���� �� ��������� ��������� ��������� ��������� Saved � ���� ��������� �� ���� ��������� - ��������� ��;
'If DocWord.Saved = False Then DocWord.Save

WordApp.Options.CheckSpellingAsYouType = False


Set TableWord = DocWord.Tables(1)
'.Add(DocWord.Range(), 10, 2)


'�������� ����� � ������ � �������
'(�����_������, �����_�������)



TableWord.Cell(10, 2).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 2) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 3) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 4)
If Dolg <> 0 Then
TableWord.Cell(13, 4).Range.Text = Dolg
TableWord.Cell(14, 6).Range.Text = Dolg
'Else
'TableWord.Cell(13, 4).Range.Text = ""
'TableWord.Cell(14, 6).Range.Text = ""
End If

'����������� ����� � ��������
stn = Filter.Fg.TextMatrix(Filter.Fg.Row, 11)
For i = 1 To Len(stn)
TableWord.Cell(12, i + 14).Range.Text = ""
TableWord.Cell(12, i + 14).Range.Text = Mid(stn, i, 1)
Next
'**********
' �����
TableWord.Cell(11, 2).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 5) + " �� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 9)
'����.�������
TableWord.Cell(8, 1).Range.Text = fg1.TextMatrix(fg1.Row, 1)
'����
TableWord.Cell(13, 2).Range.Text = MainForm.Label8 + " �."

End Sub

Private Sub ������_Click()

Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long




Dolg = Round(Label13, 2)
FormDolg.Text1 = Dolg

FormDolg.Show 1



If Dolg = -369.8985231 Then Exit Sub

If Dolg <= 0 Then
If MsgBox("�� ������ ��������� ������� ��������� " + Str((Dolg)) + ". ������ ��������� �� ������� ������ �� ���������." + vbNewLine + "���������� ������ ��������� ?", vbYesNo) = vbNo Then
Exit Sub
Else
'dolg = 0
Dolg = InputBox("������� �����" + vbNewLine + " � ������ �� " + vbNewLine + MainForm.Label8 + " �.", , Dolg)
End If
End If



nameRP = "IzvR"

'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = True

'������ ����� �������� � Word-e
'Set DocWord = WordApp.Documents.Add

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���



DocWord.Activate

'��������� ��������� ��������
On Error GoTo est
If Err.Number <> 5356 Then
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)

est:
 End If
If Err.Number = 5356 Then
Err.Clear
nameRP = Trim(Trim(nameRP) + Trim(Str(Int(Rnd() * 1000))))
DocWord.SaveAs (App.Path + "\Temp\" + nameRP + ".doc")
End If
WordApp.Options.CheckSpellingAsYouType = False
Set DocWord = WordApp.Documents.Open(App.Path + "\Temp\" + nameRP + ".doc")
DocWord.Activate
Set TableWord = DocWord.Tables(1)
TableWord.Cell(1, 1).Range.Text = MainForm.NamePr

TableWord.Cell(2, 1).Range.Text = MainForm.Bank

TableWord.Cell(3, 2).Range.Text = MainForm.BIK

TableWord.Cell(3, 4).Range.Text = MainForm.KS

TableWord.Cell(4, 4).Range.Text = MainForm.RS

TableWord.Cell(4, 2).Range.Text = MainForm.INN
'������� ����
TableWord.Cell(5, 2).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 11)

'���
TableWord.Cell(6, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 2) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 3) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 4)
If Dolg <> 0 Then
'TableWord.Cell(13, 4).Range.Text = dolg
'TableWord.Cell(14, 6).Range.Text = dolg
End If

'����������� ����� � ��������
'**********
' �����
TableWord.Cell(8, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 5) + " �� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 9)
'����.�������
'TableWord.Cell(13, 1).Range.Text = fg1.TextMatrix(fg1.Row, 1)
'�����

'����
TableWord.Cell(10, 2).Range.Text = MainForm.Label8 + " �."


'�������, ��������� � �.�.
TableWord.Cell(11, 1).Range.Text = "���.��.-" + fg1.TextMatrix(fg1.Row, 15) + "�*2 ���������-" + fg1.TextMatrix(fg1.Row, 12) + "�."


'DocWord.Tables(1).Rows.Add

 
'TableWord.Cell(15, 1).Range.Text = NumStr(Dolg, True)

'�������� �������
 '   Dim Tbl As Table
   ' Dim rng As Range
    
    
    With WordApp.ActiveDocument
 Set rng = .Paragraphs(.Paragraphs.Count).Range
 
 
 
'    Set rng = WordApp.ActiveDocument.Paragraphs(WordApp.ActiveDocument.Paragraphs.Count).Range
        
        
'��������� ������
'DocWord.Tables(1).Columns.Add 13
'DocWord.Tables(1).Rows.Add


K = 15

'������
DocWord.Tables(1).Rows.Add
If Val(Label10) >= 0 Then
TableWord.Cell(K + i, 1).Range.Text = "���� �� ������ " + MainForm.Label8 + " �."
TableWord.Cell(K + i, 2).Range.Text = Label10

End If

If Val(Label10) < 0 Then
TableWord.Cell(K + i, 1).Range.Text = "��������� �� ������ " + MainForm.Label8 + " �."
TableWord.Cell(K + i, 2).Range.Text = Label10
End If

K = 16
For i = 1 To fg1.Rows - 1
'If FG1.TextMatrix(I, 23) = "+" Then
DocWord.Tables(1).Rows.Add
'����.�������

If fg1.TextMatrix(i, 23) <> "+" Then TableWord.Cell(K + i, 1).Range.Text = fg1.TextMatrix(i, 3)
If fg1.TextMatrix(i, 23) = "+" Then TableWord.Cell(K + i, 1).Range.Text = fg1.TextMatrix(i, 3) + " (�� ��� = " + fg1.TextMatrix(i, 10) + "���.)"
'�����
TableWord.Cell(K + i, 2).Range.Text = fg1.TextMatrix(i, 18)
'������
TableWord.Cell(K + i, 3).Range.Text = fg1.TextMatrix(i, 23)
'End If
Next
        
        
        
        
        
        
        Set Tbl = .Tables(1)
    End With
    
    
'K = 14
If Val(Label10) <> 0 Then
DocWord.Tables(1).Rows.Add
'����.�������




'�����
'TableWord.Cell(14, 1).Range.Text = FG1.TextMatrix(FG1.Row, 1)
TableWord.Cell(K + 1, 1).Range.Text = "� ���� � ������:"
TableWord.Cell(K + 1, 2).Range.Text = Dolg
'K = 15
End If
    
    
    
    
       rng.ParagraphFormat.Alignment = wdAlignParagraphRight
       rng.InsertAfter NumStr(Dolg, True)
        
       
    
    Tbl.Range.Copy
    
    
    With rng
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
    
        .Collapse Direction:=wdCollapseEnd
        .Paste

 End With

End Sub

Private Sub ��������_Click()

Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long




Dolg = Round(Label13, 2)
FormDolg.Text1 = Dolg

FormDolg.Show 1



If Dolg = -369.8985231 Then Exit Sub

If Dolg <= 0 Then
If MsgBox("�� ������ ��������� ������� ��������� " + Str((Dolg)) + ". ������ ��������� �� ������� ������ �� ���������." + vbNewLine + "���������� ������ ��������� ?", vbYesNo) = vbNo Then
Exit Sub
Else
'dolg = 0
Dolg = InputBox("������� �����" + vbNewLine + " � ������ �� " + vbNewLine + MainForm.Label8 + " �.", , Dolg)
End If
End If



nameRP = "IzvR"

'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = True

'������ ����� �������� � Word-e
'Set DocWord = WordApp.Documents.Add

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���



DocWord.Activate

'��������� ��������� ��������
On Error GoTo est
If Err.Number <> 5356 Then
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)

est:
 End If
If Err.Number = 5356 Then
Err.Clear
nameRP = Trim(Trim(nameRP) + Trim(Str(Int(Rnd() * 1000))))
DocWord.SaveAs (App.Path + "\Temp\" + nameRP + ".doc")
End If
WordApp.Options.CheckSpellingAsYouType = False
Set DocWord = WordApp.Documents.Open(App.Path + "\Temp\" + nameRP + ".doc")
DocWord.Activate
Set TableWord = DocWord.Tables(1)
TableWord.Cell(1, 1).Range.Text = MainForm.NamePr

TableWord.Cell(2, 1).Range.Text = MainForm.Bank

TableWord.Cell(3, 2).Range.Text = MainForm.BIK

TableWord.Cell(3, 4).Range.Text = MainForm.KS

TableWord.Cell(4, 4).Range.Text = MainForm.RS

TableWord.Cell(4, 2).Range.Text = MainForm.INN
'������� ����
TableWord.Cell(5, 2).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 11)

'���
TableWord.Cell(6, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 2) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 3) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 4)
If Dolg <> 0 Then
'TableWord.Cell(13, 4).Range.Text = dolg
'TableWord.Cell(14, 6).Range.Text = dolg
End If

'����������� ����� � ��������
'**********
' �����
TableWord.Cell(8, 1).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 5) + " �� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 9)
'����.�������
'TableWord.Cell(13, 1).Range.Text = fg1.TextMatrix(fg1.Row, 1)
'�����

'����
TableWord.Cell(10, 2).Range.Text = MainForm.Label8 + " �."

'K = 14
If Val(Label10) <> 0 Then
DocWord.Tables(1).Rows.Add
'����.�������
'If Val(Label10) > 0 Then TableWord.Cell(14, 1).Range.Text = "���� �� ������ " + MainForm.Label8 + " �."
'If Val(Label10) < 0 Then TableWord.Cell(14, 1).Range.Text = "��������� �� ������ " + MainForm.Label8 + " �."
'�����
TableWord.Cell(14, 1).Range.Text = fg1.TextMatrix(fg1.Row, 1)
'TableWord.Cell(14, 1).Range.Text = "�������� ������"
TableWord.Cell(14, 2).Range.Text = Dolg
'K = 15
End If

'DocWord.Tables(1).Rows.Add

 
'TableWord.Cell(15, 1).Range.Text = NumStr(Dolg, True)

'�������� �������
 '   Dim Tbl As Table
   ' Dim rng As Range
    
    
    With WordApp.ActiveDocument
 Set rng = .Paragraphs(.Paragraphs.Count).Range
 
 
 
'    Set rng = WordApp.ActiveDocument.Paragraphs(WordApp.ActiveDocument.Paragraphs.Count).Range
        
        
        Set Tbl = .Tables(1)
    End With
    
       rng.ParagraphFormat.Alignment = wdAlignParagraphRight
       rng.InsertAfter NumStr(Dolg, True)
        
       
    
    Tbl.Range.Copy
    
    
    With rng
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
    
        .Collapse Direction:=wdCollapseEnd
        .Paste
    End With
    
      
End Sub

Private Sub ���������_���������_Click()
���������

End Sub



Private Sub ���������������_Click()
ArhoPL.Show 1
End Sub

Private Sub ����������_Click()
'FG1.TextMatrix(FG1.Row, 38) = InputBox("", "����������", FG1.TextMatrix(FG1.Row, 38))
Comentariy.Text1 = fg1.TextMatrix(fg1.Row, 38)
Comentariy.Show
End Sub

Private Sub �������_Click()
Command12_Click
End Sub

Private Sub ������_Click()
DropForm2.Show
    DropForm3.Show
    DropForm3.Move DropForm2.Width + 1, (DropForm2.Height - DropForm3.Height) / 2
   OtheOwner.Othe1 = 0
   
End Sub

Private Sub �������_Click()
V1.Visible = True

End Sub

Private Sub ����������_Click()
Dolg = Round(Text2.Text, 2)


'If MsgBox("�� ������ ��������� ������� ��������� " + Str((Text2.Text)) + ". ������ ��������� �� ������� ������ �� ���������." + vbNewLine + "���������� ������ ��������� ?", vbYesNo) = vbNo Then
Dolg = 0

Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long
Dim stn As String

nameRP = "Izv1"

'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = True

'������ ����� �������� � Word-e
'Set DocWord = WordApp.Documents.Add

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���
DocWord.Activate

'��������� ��������� ��������
On Error GoTo est
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)
est:
 
If Err.Number = 5356 Then
Err.Clear
nameRP = Trim(Trim(nameRP) + Trim(Str(Int(Rnd() * 1000))))

DocWord.SaveAs (App.Path + "\Temp\" + nameRP + ".doc")
End If
'���������, ���� �� ��������� ��������� ��������� ��������� Saved � ���� ��������� �� ���� ��������� - ��������� ��;
'If DocWord.Saved = False Then DocWord.Save

WordApp.Options.CheckSpellingAsYouType = False


Set TableWord = DocWord.Tables(1)
'.Add(DocWord.Range(), 10, 2)


'�������� ����� � ������ � �������
'(�����_������, �����_�������)



TableWord.Cell(10, 2).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 2) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 3) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 4)
If Dolg <> 0 Then
TableWord.Cell(13, 4).Range.Text = Dolg
TableWord.Cell(14, 6).Range.Text = Dolg
'Else
'TableWord.Cell(13, 4).Range.Text = ""
'TableWord.Cell(14, 6).Range.Text = ""
End If

'����������� ����� � ��������
stn = Filter.Fg.TextMatrix(Filter.Fg.Row, 11)
For i = 1 To Len(stn)
TableWord.Cell(12, i + 14).Range.Text = ""
TableWord.Cell(12, i + 14).Range.Text = Mid(stn, i, 1)
Next
'**********
' �����
TableWord.Cell(11, 2).Range.Text = Filter.Fg.TextMatrix(Filter.Fg.Row, 5) + " �� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 9)
'����.�������
'TableWord.Cell(8, 1).Range.Text = fg1.TextMatrix(fg1.Row, 1)
'����
'TableWord.Cell(13, 2).Range.Text = MainForm.Label8 + " �."

End Sub




Private Sub ���������_Click()
Command2_Click
End Sub

Private Sub ������_��_������_Click()
fg1.Enabled = False
Command10.Visible = True
'FGS = FG1.TextMatrix(FG1.Row, 22)
FGS = fg1.Row
TMPI = 1
Text1.Text = fg1.TextMatrix(fg1.Row, 20)
Text1.Enabled = True
Text1.SetFocus
End Sub

'Sub Saldo()

'Nac = 0
'nu = Nac
'End Sub

Private Sub ��������()
'Rs.UpdateBatch
'On Error Resume Next
If Not RS.EOF Then RS.MoveFirst

RS.AddNew
'MsgBox (Filter.Nm)
RS.Fields("�����") = Val(Filter.Nm)

RS("Formula") = "0"
RS("FormulaB") = "0"
RS("������") = 999

RS.UpdateBatch
RS.Requery
'Set FG1.DataSource = Rs

End Sub
Private Sub ��������()
Dim ComboQ As String

ComboQ = "Where(((Adding.KodKv) = " & Filter.Nm & "))"
'ComboQ = "Where(((Adding.KodKv) = " & Filter.Nm & ") and (Adding.KodN= " + FG1.TextMatrix(FG.Row, 3) + "))"

Ca.Execute ("UPDATE Adding INNER JOIN Nachisleniy ON Adding.KodN = Nachisleniy.Kod SET Adding.NameN = [Nachisleniy]![Naim], Adding.KodKat = [Nachisleniy]![���Kategor], Adding.Formula = [Nachisleniy]![Formula], Adding.FormulaB = [Nachisleniy]![FormulaB], Adding.Tip = [Nachisleniy]![Tip], Adding.NameKat = [Nachisleniy]![Kategor] " + ComboQ)
' ���� �������
Ca.Execute ("UPDATE Settings, Adding SET  Adding.DataR = [Settings]![TekData]" + ComboQ)
'������
Ca.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.Propis = [MainOccupant]![NLODGERF], Adding.Projiv = [MainOccupant]![NLODGER], Adding.ProLift = [MainOccupant]![NLODLIFT], Adding.ObPl = [MainOccupant]![COMSPACE], Adding.PolPl = [MainOccupant]![HABSPACE], Adding.TipKvKod = [MainOccupant]![KV], Adding.TipDomKod = [MainOccupant]![DomTip]" + ComboQ)
'����������
Ca.Execute ("UPDATE Adding SET Adding.Socmin =0 " + ComboQ)
Ca.Execute ("UPDATE Adding INNER JOIN Socmin ON (Adding.Propis = Socmin.koli) AND (Adding.KodKat = Socmin.KodKategor) SET Adding.Socmin = [Socmin]![Value]" + ComboQ)
'�����
Ca.Execute ("UPDATE Adding SET Adding.Tarif = 0 " + ComboQ)
Ca.Execute ("UPDATE Adding INNER JOIN Tarif ON (Tarif.KodDOM = Adding.TipDomKod) AND (Tarif.KodKV = Adding.TipKvKod) AND (Adding.KodKat = Tarif.KodKat) SET Adding.Tarif = [Tarif]![Value]" + ComboQ)
'������
'��������� ������ ���������� 0 ���� ������ ����������� �������
Ca.Execute ("UPDATE Adding SET Adding.ispr = 0 WHERE (((Adding.ispr)<>1))")
Ca.Execute ("UPDATE Adding LEFT JOIN Nachisleniy ON Adding.KodN = Nachisleniy.Kod SET Adding.LgotaVid = [Nachisleniy]![Vid]" + ComboQ)
'������������ �� /���
Ca.Execute ("UPDATE Adding INNER JOIN Nachisleniy ON Adding.KodN = Nachisleniy.Kod SET Adding.Lig = [Nachisleniy]![Lig]")
RS.Requery
End Sub
Private Sub ��������������()
Dim ertar As Label
Dim errObn As Label

Set TMP = New ADODB.Recordset
Set TMP.ActiveConnection = Ca
TMP.CursorType = adOpenForwardOnly
TMP.LockType = adLockBatchOptimistic





'On Error GoTo errObn
'������ �� Nachisleny
TMP.Open ("SELECT Nachisleniy.Kod, Nachisleniy.���Kategor, Nachisleniy.Kategor, Nachisleniy.Naim, Nachisleniy.Formula, Nachisleniy.FormulaB ,Nachisleniy.Vid, Nachisleniy.Tip, Nachisleniy.Lig, Nachisleniy.SchetZ, Nachisleniy.Sch From Nachisleniy WHERE (((Nachisleniy.Kod)=" + fg1.TextMatrix(fg1.Row, 2) + "))")
'MsgBox (fg1.TextMatrix(fg1.Row, 2))
fg1.TextMatrix(fg1.Row, 3) = TMP.Fields("Naim").Value
fg1.TextMatrix(fg1.Row, 22) = TMP.Fields("���Kategor").Value
fg1.TextMatrix(fg1.Row, 1) = TMP.Fields("Kategor").Value
fg1.TextMatrix(fg1.Row, 19) = TMP.Fields("Formula").Value
fg1.TextMatrix(fg1.Row, 7) = TMP.Fields("Vid").Value
fg1.TextMatrix(fg1.Row, 23) = TMP.Fields("Tip").Value
fg1.TextMatrix(fg1.Row, 30) = TMP.Fields("Lig").Value
fg1.TextMatrix(fg1.Row, 35) = TMP.Fields("SchetZ").Value
fg1.TextMatrix(fg1.Row, 39) = TMP.Fields("FormulaB").Value
fg1.TextMatrix(fg1.Row, 43) = TMP.Fields("Sch").Value
TMP.Close

'������ �� MainOccupant
TMP.Open ("SELECT MainOccupant.Numer, MainOccupant.NLODGER, MainOccupant.NLODGERF, MainOccupant.NLODLIFT, MainOccupant.COMSPACE, MainOccupant.dom, MainOccupant.FLOOR, MainOccupant.HABSPACE, MainOccupant.DomTip, MainOccupant.KV FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer GROUP BY MainOccupant.Numer, MainOccupant.NLODGER, MainOccupant.NLODGERF, MainOccupant.NLODLIFT, MainOccupant.COMSPACE, MainOccupant.dom, MainOccupant.FLOOR, MainOccupant.HABSPACE, MainOccupant.DomTip, MainOccupant.KV HAVING (((MainOccupant.Numer)=" + fg1.TextMatrix(fg1.Row, 0) + "))")
If TMP.Fields("NLODGER").Value <> "" Then fg1.TextMatrix(fg1.Row, 12) = TMP.Fields("NLODGER").Value Else fg1.TextMatrix(fg1.Row, 12) = 0 '���-�� �����������

fg1.TextMatrix(fg1.Row, 13) = TMP.Fields("NLODGER").Value '���-�� �����������
fg1.TextMatrix(fg1.Row, 12) = TMP.Fields("NLODGERF").Value '���-�� �����������
fg1.TextMatrix(fg1.Row, 14) = TMP.Fields("NLODLIFT").Value '���-�� ��� �����
If TMP.Fields("COMSPACE").Value <> "" Then fg1.TextMatrix(fg1.Row, 15) = TMP.Fields("COMSPACE").Value Else fg1.TextMatrix(fg1.Row, 15) = 0 '������� �����
fg1.TextMatrix(fg1.Row, 16) = TMP.Fields("HABSPACE").Value '������� ��������
fg1.TextMatrix(fg1.Row, 36) = TMP.Fields("FLOOR").Value '����
fg1.TextMatrix(fg1.Row, 25) = TMP.Fields("DomTip").Value '��� ����
fg1.TextMatrix(fg1.Row, 24) = TMP.Fields("KV").Value '��� ��������

'fg1.TextMatrix(fg1.Row, 25) = TMP.Fields("Dom").Value '��� ����

TMP.Close

RS.UpdateBatch

'������
TMP.Open ("SELECT Tarif.Value, Tarif.TarifI, Tarif.TarifD, Adding.KodKv, Adding.KodKv FROM Adding INNER JOIN Tarif ON (Adding.TipDomKod = Tarif.KodDOM) AND (Adding.TipKvKod = Tarif.KodKV) AND (Adding.KodKat = Tarif.KodKat) WHERE (((Adding.Key)=" + fg1.TextMatrix(fg1.Row, 26) + "))")
'TMP.Open ("SELECT Tarif.Value, Tarif.TarifI, Tarif.TarifD, Adding.KodKv, Adding.Key FROM (Adding INNER JOIN Tarif ON (Adding.KodKat = Tarif.KodKat) AND (Adding.TipKvKod = Tarif.KodKV)) INNER JOIN KLS_PODR ON (KLS_PODR.Tip = Tarif.KodDOM) AND (Adding.TipDomKod = KLS_PODR.���) WHERE (((Adding.Key)=" + fg1.TextMatrix(fg1.Row, 26) + "))")

'MsgBox (TMP.Fields("Value").Value)
On Error GoTo ertar
fg1.TextMatrix(fg1.Row, 10) = TMP.Fields("Value").Value '����� ��������
fg1.TextMatrix(fg1.Row, 33) = TMP.Fields("TarifI").Value '����� �������
fg1.TextMatrix(fg1.Row, 34) = TMP.Fields("TarifD").Value '����� ��������������
ertar:
TMP.Close
'����������

'TMP.Open ("SELECT Socmin.Value, Adding.KodKv FROM Adding INNER JOIN Socmin ON (Adding.Propis = Socmin.koli) AND (Adding.KodKat = Socmin.KodKategor) WHERE (((Adding.KodKv)=" + FG1.TextMatrix(FG1.Row, 0) + " AND ((Socmin.KodKategor)=3))")
TMP.Open ("SELECT Socmin.Value, Adding.KodKv, Socmin.KodKategor FROM Adding INNER JOIN Socmin ON (Adding.KodKat = Socmin.KodKategor) AND (Adding.Propis = Socmin.koli) WHERE (((Adding.KodKv)=" + fg1.TextMatrix(fg1.Row, 0) + ") AND ((Socmin.KodKategor)=" + fg1.TextMatrix(fg1.Row, 22) + "))")
'If TMP.EOF = False Then

             If TMP.EOF = False Then
fg1.TextMatrix(fg1.Row, 11) = TMP.Fields("Value").Value '����������
Else: fg1.TextMatrix(fg1.Row, 11) = 0
             End If
TMP.Close
End Sub

Private Sub ���������_������()
If Status = "Text1" Then Sal = Text1.Text
For rw = 1 To fg1.Rows - 1
If fg1.TextMatrix(rw, 22) = fg1.TextMatrix(FGS, 22) Then
fg1.TextMatrix(rw, 20) = Text1.Text
End If
Next rw


'SaldoQ = " Where(((Adding.KodKv) = " & Filter.Nm & ")" + " AND ((Adding.KodKat)=" + Str(KODS_Kat) + "))"
'Ca.Execute ("UPDATE Adding SET Adding.SaldoN = " + Str(Sal) + SaldoQ)
'��������_������
'Text1.Text = Sal



End Sub
Private Sub ���������_��_����������()
Dim SaldoNa As Double, SaldoKn As Double, ��������� As Double, �������� As Double, ��������� As Double

'Label1.Caption = FG1.TextMatrix(FG1.Row, 1)
'Text1.Text = FG1.TextMatrix(FG1.Row, 20)
'MsgBox ("���������_��_����������")
'On Error Resume Next

If fg1.Row <> 0 Then
��������� = fg1.TextMatrix(fg1.Row, 22)
SaldoNa = fg1.TextMatrix(fg1.Row, 20)
SaldoKn = fg1.TextMatrix(fg1.Row, 21)
��������� = Round(NACH(fg1.TextMatrix(fg1.Row, 22)), 2)
�������� = Round(OPL(fg1.TextMatrix(fg1.Row, 22)), 2)
End If

Label5.Caption = ���������
Label6.Caption = ��������


SaldoKn = Round(SaldoNa + ��������� - ��������, 2)
Text2.Text = SaldoKn
'Text2.Text = Sal + NACH(FG1.TextMatrix(FG1.Row, 22)) - OPL(FG1.TextMatrix(FG1.Row, 22))

' ���������� ������ ����������� ������ ������ ��������� ������ "FG1.TextMatrix(FG1.Row, 21) = SaldoKn"
' ���������� ��� ���� ����� �� ������ ��������� ������� ��� ��������� �������� � Fg1.TextMatrix(Fg1.Row, 22)
'���� �� ������� FG1
'����������� ������ �� ����� �� ������� "��������� ��������� �������" Fg1.TextMatrix(FG, 21) = SaldoKn
For Fg = 1 To fg1.Rows - 1

If fg1.TextMatrix(Fg, 22) = ��������� Then
fg1.TextMatrix(Fg, 21) = SaldoKn
'MsgBox Fg1.TextMatrix(FG, 3)
End If
Next

'MsgBox (Fg1.TextMatrix(Fg1.Row, 22))

End Sub
Private Sub ��������_������()
For i = 0 To 998
NACH(i) = 0
OPL(i) = 0
Next i

For Fg = 1 To fg1.Rows - 1
If fg1.TextMatrix(Fg, 18) <> "" Then
If fg1.TextMatrix(Fg, 23) = "+" Then NACH(fg1.TextMatrix(Fg, 22)) = NACH(fg1.TextMatrix(Fg, 22)) + fg1.TextMatrix(Fg, 18)
If fg1.TextMatrix(Fg, 23) = "-" Or fg1.TextMatrix(Fg, 23) = "s" Then OPL(fg1.TextMatrix(Fg, 22)) = OPL(fg1.TextMatrix(Fg, 22)) + fg1.TextMatrix(Fg, 18)
End If
Next



End Sub

Private Sub �������_Click()
Dim Tbl As Word.Table
Dim rng As Word.Range
Dim WordApp As Word.Application ' ��������� ����������
Dim DocWord As Word.Document ' ��������� ���������
'��������� ��������� ���������� � �������
' Generals �����
Dim TableWord As Word.Table
Dim O9 As Double
Dim S9 As Double
Dim rw As Long




Dolg = Round(Text2, 2)
FormDolg.Text1 = Dolg

'FormDolg.Show 1



'If Dolg = -369.8985231 Then Exit Sub

If Dolg > 0 Then
If MsgBox("�� ������� �������� ����� ���� " + Str((Dolg)) + ". ������ ������� ����������." + vbNewLine + "���������� ���������?", vbYesNo) = vbNo Then
Exit Sub
Else

' ���� ��� �������������

End If
End If
nameRP = "Dolg"

'111112222333335555577777779999555555666662222222
'������ ����� ��������� Word-a
Set WordApp = New Word.Application

'���������� ��������� Word-a �� True - �������,
'�� False - �� ������� (�������� ������ ����)
WordApp.Visible = True

'������ ����� �������� � Word-e
'Set DocWord = WordApp.Documents.Add

'// ���� ����� ������� ��������� ��������, �� ����� ����� ���
Set DocWord = WordApp.Documents.Open(App.Path + "\rep\" + nameRP + ".doc")
'���������� ���



DocWord.Activate

'��������� ��������� ��������
On Error GoTo est
If Err.Number <> 5356 Then
DocWord.SaveAs (App.Path + "\Temp\" + nameRP)

est:
 End If
If Err.Number = 5356 Then
Err.Clear
nameRP = Trim(Trim(nameRP) + Trim(Str(Int(Rnd() * 1000))))
DocWord.SaveAs (App.Path + "\Temp\" + nameRP + ".doc")
End If
WordApp.Options.CheckSpellingAsYouType = False
Set DocWord = WordApp.Documents.Open(App.Path + "\Temp\" + nameRP + ".doc")
DocWord.Activate


Set TableWord = DocWord.Tables(1)

TableWord.Cell(1, 1).Range.Text = MainForm.NamePr

TableWord.Cell(3, 1).Range.Text = MainForm.Adr


TableWord.Cell(2, 3).Range.Text = "�. ���������," + Filter.Fg.TextMatrix(Filter.Fg.Row, 5) + " �� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 9)


'������� ����


'���
TableWord.Cell(5, 2).Range.Text = Replace(Filter.Fg.TextMatrix(Filter.Fg.Row, 2) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 3) + " " + Filter.Fg.TextMatrix(Filter.Fg.Row, 4), "*", "")
TableWord.Cell(6, 2).Range.Text = MainForm.NamePr



TableWord.Cell(7, 2).Range.Text = fg1.TextMatrix(fg1.Row, 3)

TableWord.Cell(13, 1).Range.Text = MainForm.RukDol
TableWord.Cell(13, 3).Range.Text = MainForm.RukName
'+ " (�� ��� = " + FG1.TextMatrix(FG1.Row, 10) + "���.)"

'����������� ����� � ��������
'**********
' �����
'TableWord.Cell(8, 1).Range.Text = Filter.FG.TextMatrix(Filter.FG.Row, 5) + " �� �" + Filter.FG.TextMatrix(Filter.FG.Row, 9)
'����.�������
'TableWord.Cell(13, 1).Range.Text = fg1.TextMatrix(fg1.Row, 1)
'�����

'����
TableWord.Cell(8, 2).Range.Text = MainForm.Label8 + " �."


'�������, ��������� � �.�.
'TableWord.Cell(8, 4).Range.Text = "���.��.-" + fg1.TextMatrix(fg1.Row, 15) + "�*2 ���������-" + fg1.TextMatrix(fg1.Row, 12) + "�."


'DocWord.Tables(1).Rows.Add

 
'TableWord.Cell(15, 1).Range.Text = NumStr(Dolg, True)

'�������� �������
 '   Dim Tbl As Table
   ' Dim rng As Range
    
    
    With WordApp.ActiveDocument
 Set rng = .Paragraphs(.Paragraphs.Count).Range
 
 
 
'    Set rng = WordApp.ActiveDocument.Paragraphs(WordApp.ActiveDocument.Paragraphs.Count).Range
        
        
'��������� ������
'DocWord.Tables(1).Columns.Add 13
'DocWord.Tables(1).Rows.Add


K = 15

'������
'DocWord.Tables(1).Rows.Add
'If Val(Label10) >= 0 Then
'TableWord.Cell(8, 4).Range.Text = "���� �� ������ " + MainForm.Label8 + " �."

'MsgBox Dolg

TableWord.Cell(8, 4).Range.Text = Dolg
'Dolg = Label10
TableWord.Cell(9, 1).Range.Text = "(" + NumStr(Dolg, True) + ")"

       
        
        
        
        Set Tbl = .Tables(1)
    End With
    
    
    
'       rng.ParagraphFormat.Alignment = wdAlignParagraphRight
 '      rng.InsertAfter NumStr(Dolg, True)
        
       
    
    'Tbl.Range.Copy
    
    
   ' With rng
    '    .InsertParagraphAfter
     '   .InsertParagraphAfter
      '  .InsertParagraphAfter
       ' .InsertParagraphAfter
        '.InsertParagraphAfter
    
        '.Collapse Direction:=wdCollapseEnd
        '.Paste

 'End With



'111112222333335555577777779999555555666662222222
End Sub

Private Sub �������_Click()
If fg1.TextMatrix(fg1.Row, 43) <> "��" Then
MsgBox "��������� ���� ��� ���������� �� ������������ ������ �� ������ ��������."
Exit Sub
End If

SchVV.Show 1
End Sub

'Private Sub Return_KODS_KAT()
'TMP.Open ("Select Nachisleniy.���Kategor from Nachisleniy WHERE Nachisleniy.Kod=" + Str(KODS_N))
'KODS_Kat = TMP.Fields("���Kategor").Value
'TMP.Close
'End Sub

Private Sub �������_����������_Click()
Command4_Click
End Sub
Private Sub ���������()
            fg1.Select fg1.Row, 18
            fg1.EditCell
            fg1.TextMatrix(fg1.Row, 27) = 1
            
            
End Sub

Private Sub Lgota()
Dim Proc As Double

Select Case Err.Number
Case Is = 13
MsgBox ("��������� �������� �� �������� ������. ��������� ����")
'Text1_Validate = True
Case Else
MsgBox (Err.Description)
End Select
End Sub




Public Sub �������()
Dim rw As Integer '����� ������ �����

For rw = 1 To fg1.Rows - 1
               
'����������� me.fg1.TextMatrix(Rw, 26), True

Next rw
End Sub

Private Sub ����()
Dim rw As Integer

For rw = 1 To fg1.Rows - 1

If fg1.TextMatrix(rw, 28) <> 0 Then
'FG1.Cell(flexcpBackColor, Rw, 1, Rw, 28) = vbBlue
fg1.Cell(flexcpFontBold, rw, 1, rw, 28) = True


End If

If fg1.TextMatrix(rw, 27) = 1 Then

fg1.Cell(flexcpFontBold, rw, 18, rw, 18) = True
fg1.Cell(flexcpBackColor, rw, 18, rw, 18) = vbCyan
End If
If fg1.TextMatrix(rw, 23) = "+" Then fg1.Cell(flexcpForeColor, rw, 18, rw, 18) = vbBlue
If fg1.TextMatrix(rw, 23) = "-" Then fg1.Cell(flexcpForeColor, rw, 18, rw, 18) = vbRed
If fg1.TextMatrix(rw, 23) = "s" Then fg1.Cell(flexcpForeColor, rw, 18, rw, 18) = vbMagenta

If fg1.TextMatrix(rw, 43) = "��" Then
'fg1.Cell(flexcpForeColor, rw, 1, rw, 42) = RGB(50, 100, 50)
fg1.Cell(flexcpBackColor, rw, 1, rw, 42) = RGB(80, 250, 190)
End If
'RGB(80, 250, 150)
'RGB(300, 255, 200)
'RGB(200, 255, 200)
'vbGreen
Next rw



End Sub
Public Sub �����������1(ByVal Rw1, KEY As Double, Sposob As String)
Dim FormulaN As String
Dim FormulaB As String



If Sposob = "��������" Then
On Error GoTo ErrRas
FormulaN = Trim(fg1.TextMatrix(Rw1, 19))
FormulaB = Trim(fg1.TextMatrix(Rw1, 39))

'Ca.Execute ("UPDATE Adding SET Adding.SummaI = " + FormulaN + ", Adding.ispr = 0 WHERE (((Adding.Key)=" + Str(KEY) + "))")
'Ca.Execute ("UPDATE Adding SET Adding.SummaBl = " + FormulaB + ", Adding.ispr = 0 WHERE (((Adding.Key)=" + Str(KEY) + "))")

Ca.Execute ("UPDATE Adding SET Adding.SummaI = " + FormulaN + ", Adding.SummaBl = " + FormulaB + ", Adding.ispr = 0 WHERE (((Adding.Key)=" + Str(KEY) + "))")

End If

          If Sposob = "�������" Then
         ' On Error GoTo ErrRas

FormulaN = Trim(fg1.TextMatrix(Rw1, 19))
FormulaB = Trim(fg1.TextMatrix(Rw1, 39))

'Ca.Execute ("UPDATE Adding SET Adding.SummaI = " + FormulaN + " WHERE (((Adding.Key)=" + Str(KEY) + ") AND ((Adding.ispr)=0))")
'Ca.Execute ("UPDATE Adding SET Adding.SummaBl = " + FormulaB + " WHERE (((Adding.Key)=" + Str(KEY) + ") AND ((Adding.ispr)=0))")

'MsgBox Oplata(1)
Ca.Execute ("UPDATE Adding SET Adding.SummaI = " & FormulaN & ", Adding.SummaBl = " + FormulaB + " WHERE (((Adding.Key)=" + Str(KEY) + ") AND ((Adding.ispr)=0))")
             End If
             
             
ErrRas:
Select Case Err.Number
'Case Is = 1
'MsgBox ("��� ����������. �� �������� ��������� ���������� ���������� ���������� (F<3>), ������� ������ �������������� ��� ������� ���������������� ��������� (�� ������ � �����)!")
'��������
Case Is = 0
Case Else
MsgBox ("��� ������   " + Str(Err.Number) + "  " + Err.Description)
End Select
             
             
           End Sub


Private Sub ���������()
'Dim sum(101) As Double

Me.Label10.Caption = "���"
Dim Plus, Minus, SNM(101), sn, sk As Double
Dim rw As Integer
Erase SNM
For i = 0 To 100
SNM(i) = 0
Next i
sn = 0
sk = 0
Plus = 0
Minus = 0
For rw = 1 To fg1.Rows - 1
SNM(Val(fg1.TextMatrix(rw, 22))) = fg1.TextMatrix(rw, 20)
If fg1.TextMatrix(rw, 23) = "+" And fg1.TextMatrix(rw, 18) <> "" Then Plus = Plus + fg1.TextMatrix(rw, 18)
If fg1.TextMatrix(rw, 23) = "-" Or fg1.TextMatrix(rw, 23) = "s" Then Minus = Minus + fg1.TextMatrix(rw, 18)
Next
For i = 0 To 100
If SNM(i) <> 0 Then
sn = sn + SNM(i)
End If
Next i
sk = sn + Plus - Minus
Me.Label10.Caption = Round(sn, 2)
Me.Label11.Caption = Round(Plus, 2)
Me.Label12.Caption = Round(Minus, 2)
Me.Label13.Caption = Round(sk, 2)
End Sub
Private Sub ����������()
Dim NAdding(100), rw, Kat As Integer '����� ������ �����
For i = 1 To 100
NAdding(i) = 0
Next i

For rw = 1 To fg1.Rows - 1
Kat = fg1.TextMatrix(rw, 22)
'If FG1.TextMatrix(Rw, 22) =KAt Then
NAdding(Kat) = NAdding(Kat) + 1
Next rw

For rw = 1 To fg1.Rows - 1
Kat = fg1.TextMatrix(rw, 22)
If fg1.TextMatrix(rw, 22) = Kat Then
'MsgBox (Str(KAt) + "   " + Str(NAdding(KAt)))
fg1.TextMatrix(rw, 37) = NAdding(Kat)
End If
Next rw
End Sub


Private Sub ������������_Click()
Dim TabN As Double
Dim NaKod As Integer
Dim ������� As Integer
Dim Potom(1000) As Double



For i = 1 To 1000
Potom(i) = 0
Next


������� = Filter.Fg.TextMatrix(Filter.Fg.Row, 10)
NaKod = fg1.TextMatrix(fg1.Row, 2)
i = 0
If MsgBox("������� ���������� " + fg1.TextMatrix(fg1.Row, 3) + " � �� ���� ������� ������ �� ������ " + Filter.Fg.TextMatrix(Filter.Fg.Row, 5) + " ��� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 6), vbYesNo, "") = vbYes Then

Jdite.Show
Jdite.Label1.Refresh

MainOc.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(�������) + "))")
MainOc.MoveFirst
Do While Not MainOc.EOF
i = i + 1
TabN = MainOc.Fields("Numer").Value
Potom(i) = TabN
Ca.Execute ("DELETE Adding.KodKv, Adding.KodN From Adding WHERE (((Adding.KodKv)=" + Str(TabN) + ") AND ((Adding.KodN)=" + Str(NaKod) + "))")
MainOc.MoveNext
Jdite.Label1 = "���������� ���������"
Jdite.Label1 = Jdite.Label1 + "  >" + Str(i)
Jdite.Label1.Refresh
Loop
For i = 1 To 1000
If Potom(i) <> 0 Then MainForm.���������������� Potom(i)
Next
MainOc.Close
End If
Ca.Execute ("DELETE tmp_lgota.*, Adding.Key FROM tmp_lgota LEFT JOIN Adding ON tmp_lgota.UniKOd = Adding.Key WHERE (((Adding.Key) Is Null))")
Unload Jdite
End Sub
Private Sub �����������()

For rw = 1 To fg1.Rows - 1
If FGS <> rw And fg1.TextMatrix(rw, 22) = fg1.TextMatrix(FGS, 22) Then
fg1.TextMatrix(FGS, 20) = fg1.TextMatrix(rw, 20)
End If
Next
End Sub

Public Sub ViewArhiv(ByVal ������ As Integer)
Dim ArhivCn As ADODB.Connection
Dim ArhivRS As ADODB.Recordset
Dim NameArhiv As String
Dim bakName As String
Dim M As String
Dim G As String
Dim BazaName As String
Dim bazaN As String
Dim sn(100) As Double
Dim sk(100) As Double
Dim na As Double
Dim Ud As Double
Dim K As Integer
Dim SNa As Double
Dim SKn As Double
For i = 0 To 99
sn(i) = 0
sk(i) = 0
Next

'���������� ��� ������
bazaN = "kvartplata.amd"
BazaName = bazaN
If ������ <> 0 Then

M = MonthName(Month(MainForm.DR - 28 * ������), True)
G = Str(Year(MainForm.DR - 28 * ������))

Else
Exit Sub
End If
bakName = Left(BazaName, (Len(BazaName) - 14)) & "Data/Arhiv/" & G + M + ".amd"
NameArhiv = Replace(bakName, " ", "")
'����������
Set ArhivCn = New ADODB.Connection

Call BaseUnProtect(App.Path + "/" + NameArhiv, True)

ArhivCn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + NameArhiv + ";Persist Security Info=True"

 On Error GoTo Net
 ArhivCn.Open NameArhiv

'Call BaseProtect(App.Path + "/" + NameArhiv, True)


Set ArhivRS = New ADODB.Recordset

ArhivRS.Open ("SELECT Adding.SaldoN, Adding.NameKat, Adding.NameN, Adding.DataR, Adding.Tarif, Adding.ObPl, Adding.Propis, Adding.SummaI, Adding.SaldoK, Adding.KodKat, Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, KLS_PODR.NAIM_KLS, KLS_PODR.Num, MainOccupant.kv_num, Adding.Tip, Adding.com FROM (Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer) INNER JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((Adding.KodKv)=" + Filter.Nm + ")) order by Adding.KodKat"), ArhivCn
'Arhivme.Show
'Set Arhivme.VSA.DataSource = ArhivRS
Set ArhivRS = Nothing
Set ArhivCn = Nothing


'For rw = 1 To Arhivme.VSA.Rows - 1
'If Arhivme.VSA.TextMatrix(rw, 18) = "+" Then
'Arhivme.VSA.Cell(flexcpForeColor, rw, 1, rw, 18) = vbBlack
'na = na + Arhivme.VSA.TextMatrix(rw, 8)
'End If
'If Arhivme.VSA.TextMatrix(rw, 18) = "-" Then
'Arhivme.VSA.Cell(flexcpForeColor, rw, 1, rw, 18) = vbRed
'Arhivme.VSA.Cell(flexcpBackColor, Rw, 1, Rw, 18) = RGB(200, 255, 200)
'Ud = Ud + Arhivme.VSA.TextMatrix(rw, 8)
'End If
'If Arhivme.VSA.TextMatrix(rw, 18) = "s" Then
'Arhivme.VSA.Cell(flexcpForeColor, rw, 1, rw, 18) = vbMagenta
'Ud = Ud + Arhivme.VSA.TextMatrix(rw, 8)
'Arhivme.VSA.Cell(flexcpBackColor, Rw, 1, Rw, 18) = RGB(600, 255, 200)
'End If
'RGB(300, 255, 200)
'RGB(200, 255, 200)
'vbGreen

'sn(Arhivme.VSA.TextMatrix(rw, 10)) = Arhivme.VSA.TextMatrix(rw, 1)
'sk(Arhivme.VSA.TextMatrix(rw, 10)) = Arhivme.VSA.TextMatrix(rw, 9)
'Next rw

'Arhivme.Label6.Caption = Str(na)
'Arhivme.Label8.Caption = Str(Ud)

'������� ������ ��� ���� ���������
SNa = 0
SKn = 0
For i = 0 To 99
SNa = SNa + Round(sn(i), 2)
SKn = SKn + Round(sk(i), 2)
Next
'Arhivme.Label4.Caption = Str(SNa)
'Arhivme.Label10.Caption = Str(SKn)

'**************************
'If Arhivme.VSA.TextMatrix(1, 12) <> "" Then
'Arhivme.Label1 = Arhivme.VSA.TextMatrix(1, 12) + " " + Arhivme.VSA.TextMatrix(1, 13) + " " + Arhivme.VSA.TextMatrix(1, 14) + " " + Arhivme.VSA.TextMatrix(1, 15) + " ��� � " + Arhivme.VSA.TextMatrix(1, 16) + " �� �" + Arhivme.VSA.TextMatrix(1, 17)
'Else
'Arhivme.Label1 = "��� �������� ������"
'End If
'Arhivme.Label2 = "�������� ������ �� " + M + " " + Str(G) + " �."
'Arhivme.VSA.MergeCells = flexMergeRestrictAll
'Arhivme.VSA.MergeCol(-1) = True
'Arhivme.VSA.Refresh
'Arhivme.VSA.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove

'Arhivme.VSA.Cell(flexcpBackColor, Rw, 1, Rw, 18) = RGB(200, 255, 200)

Exit Sub
Net:
If Err.Number = 381 Then MsgBox ("��� �������� ������ " + M + " " + Str(G) + " �.")
If Err.Number <> 381 Then MsgBox Err.Description
Err.Clear


End Sub
Private Sub ���������������()
Dim TabN As Double
Dim NaKod As Integer
Dim ������� As Integer
Dim Potom(1000) As Double



For i = 1 To 1000
Potom(i) = 0
Next


������� = Filter.Fg.TextMatrix(Filter.Fg.Row, 10)
NaKod = fg1.TextMatrix(fg1.Row, 2)
i = 0
If MsgBox("�������� ���������� " + fg1.TextMatrix(fg1.Row, 3) + " ����  ������� ������ �� ������ " + Filter.Fg.TextMatrix(Filter.Fg.Row, 5) + " ��� �" + Filter.Fg.TextMatrix(Filter.Fg.Row, 6), vbYesNo, "") = vbYes Then

Jdite.Show
Jdite.Label1.Refresh

MainOc.Open ("SELECT MainOccupant.Numer, MainOccupant.Dom From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(�������) + "))")
MainOc.MoveFirst
Do While Not MainOc.EOF
i = i + 1
TabN = MainOc.Fields("Numer").Value
Potom(i) = TabN
'Ca.Execute ("DELETE Adding.KodKv, Adding.KodN From Adding WHERE (((Adding.KodKv)=" + Str(TabN) + ") AND ((Adding.KodN)=" + Str(NaKod) + "))")
Ca.Execute ("INSERT INTO Adding ( KodN, NameN, KodKat, NameKat, Formula, Tip, Lig, SchetZ, KodKv, LgotaVid ) SELECT nachisleniy.Kod, nachisleniy.Naim, nachisleniy.���Kategor, nachisleniy.Kategor, nachisleniy.Formula, nachisleniy.Tip, nachisleniy.Lig, nachisleniy.SchetZ, " + Str(TabN) + ", nachisleniy.Vid From Nachisleniy WHERE (((nachisleniy.Kod)=" + Str(NaKod) + "))")



MainOc.MoveNext
Jdite.Label1 = "���������� ���������"
Jdite.Label1 = Jdite.Label1 + "  >" + Str(i)
Jdite.Label1.Refresh
Loop
Jdite.Label1.FontSize = "10"
Jdite.Label1 = "���������� ��������� �������� ������ � ������, �������,�������"
For i = 1 To 1000
If Potom(i) <> 0 Then MainForm.���������������� Potom(i)
Next
MainOc.Close
End If
Ca.Execute ("DELETE tmp_lgota.*, Adding.Key FROM tmp_lgota LEFT JOIN Adding ON tmp_lgota.UniKOd = Adding.Key WHERE (((Adding.Key) Is Null))")
'������ ������
Ca.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv=MainOccupant.Numer SET Adding.Propis = MainOccupant!NLODGERF, Adding.Projiv = MainOccupant!NLODGER, Adding.ProLift = MainOccupant!NLODLIFT, Adding.ObPl = MainOccupant!COMSPACE, Adding.PolPl = MainOccupant!HABSPACE, Adding.TipKvKod = MainOccupant!KV, Adding.TipDomKod = MainOccupant!DomTip")
'������
Ca.Execute ("UPDATE Adding INNER JOIN Tarif ON (Adding.TipDomKod = Tarif.KodDOM) AND (Adding.TipKvKod = Tarif.KodKV) AND (Adding.KodKat = Tarif.KodKat) SET Adding.Tarif = [Tarif]![Value], Adding.TarifI = [Tarif]![TarifI], Adding.TarifD = [Tarif]![TarifD]")
Ca.Execute ("UPDATE Adding SET Adding.Tarif = 0 WHERE (((Adding.Tarif) Is Null))")
'����������
Ca.Execute ("UPDATE Adding INNER JOIN Socmin ON (Adding.Propis = Socmin.koli) AND (Adding.KodKat = Socmin.KodKategor) SET Adding.Socmin = [Socmin]![Value]")

MainForm.���������������� "All", True
MainForm.���������������� "All", False
Unload Jdite
End Sub


'��� ������� ���������� ����� ������ �� ��������� KodKoaegorii ��� �������� ���. �����
' X=Oplanta(1) �������� X �������� ����� ����� � �������� ��� ��������� ������� 1

Public Function Oplata(ByVal KodKoaegorii As Integer) As Double
Dim Fun As ADODB.Recordset
Set Fun = New ADODB.Recordset
   Fun.Open ("SELECT Sum(ADDING.SummaI) AS [Sum-SummaI] From ADDING GROUP BY ADDING.KodKat, ADDING.Tip, ADDING.KodKv HAVING (((ADDING.KodKat)=" + Str(KodKoaegorii) + ") AND ((ADDING.Tip)=" + Chr(34) + "-" + Chr(34) + "Or (ADDING.Tip)=" + Chr(34) + "s" + Chr(34) + ") AND ((ADDING.KodKv)=" + Filter.Nm + "))"), Ca
If Not Fun.EOF Then Oplata = Fun.Fields("Sum-SummaI").Value Else Oplata = 0
Fun.Close
End Function



