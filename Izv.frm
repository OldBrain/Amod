VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Izv 
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8892
   ClientLeft      =   12
   ClientTop       =   0
   ClientWidth     =   10920
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   741
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "� � � � �"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   960
      TabIndex        =   27
      Top             =   8280
      Width           =   1932
   End
   Begin VB.CommandButton Command6 
      Caption         =   "������"
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
      Left            =   5640
      TabIndex        =   25
      Top             =   8040
      Width           =   1692
   End
   Begin VB.CommandButton Command5 
      Caption         =   "0"
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
      Index           =   1
      Left            =   960
      TabIndex        =   23
      Top             =   8040
      Width           =   492
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
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
      Index           =   0
      Left            =   2400
      TabIndex        =   22
      Top             =   8040
      Width           =   492
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
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
      Index           =   1
      Left            =   1920
      TabIndex        =   21
      Top             =   8040
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
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
      Index           =   0
      Left            =   1440
      TabIndex        =   20
      Top             =   8040
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Height          =   252
      Left            =   3120
      Picture         =   "Izv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Width           =   2412
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5652
      Left            =   240
      TabIndex        =   19
      Top             =   1440
      Width           =   10500
      _cx             =   18521
      _cy             =   9970
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
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Izv.frx":011A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "������� ��� ������ ����� ��������.  ������� �� ������ <������>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   492
      Left            =   3360
      TabIndex        =   26
      Top             =   7440
      Width           =   5292
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "�������� �����������  ������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   960
      TabIndex        =   24
      ToolTipText     =   "��� ������ ������ ������ ����� ������� �������� ��������� �������� ����"
      Top             =   7440
      Width           =   2052
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "�� �"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   492
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   480
      TabIndex        =   17
      ToolTipText     =   """������"" � �/�� / ""�����"" 12-� ������� � �/��. ��� �������� � ������"
      Top             =   1200
      Width           =   732
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "�������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1320
      TabIndex        =   16
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2520
      TabIndex        =   15
      ToolTipText     =   """������"" � �/�� / ""�����"" 12-� ������� � �/��. ��� �������� � ������"
      Top             =   1200
      Width           =   492
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "���������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3480
      TabIndex        =   14
      Top             =   1200
      Width           =   1092
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "���������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   4920
      TabIndex        =   13
      ToolTipText     =   """������"" � �/�� / ""�����"" 12-� ������� � �/��. ��� �������� � ������"
      Top             =   1200
      Width           =   492
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "���������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5640
      TabIndex        =   12
      Top             =   1200
      Width           =   1092
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "���������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   6720
      TabIndex        =   11
      ToolTipText     =   """������"" � �/�� / ""�����"" 12-� ������� � �/��. ��� �������� � ������"
      Top             =   1200
      Width           =   492
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   372
      Left            =   7920
      TabIndex        =   9
      ToolTipText     =   """������"" � �/�� / ""�����"" 12-� ������� � �/��. ��� �������� � ������"
      Top             =   1200
      Width           =   492
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7320
      TabIndex        =   8
      Top             =   1200
      Width           =   492
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8
      X2              =   900
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   5760
      TabIndex        =   7
      Top             =   720
      Width           =   5052
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���������������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      TabIndex        =   6
      Top             =   720
      Width           =   2052
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1080
      TabIndex        =   5
      ToolTipText     =   """������"" � �.��  / ""�����"" 12-� ������� � �.��. ��� �������� � ������"
      Top             =   720
      Width           =   2892
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��� / �� �"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "����� "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   6372
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "����� �����������:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2412
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   10932
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
      Height          =   192
      Left            =   0
      Picture         =   "Izv.frx":01F8
      ToolTipText     =   "�������"
      Top             =   0
      Width           =   192
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   10890
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   720
      Picture         =   "Izv.frx":073A
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   960
      Picture         =   "Izv.frx":0E84
      Stretch         =   -1  'True
      ToolTipText     =   "������� ������ ���� ��������� ����� �� ���� ����� ��� ������ � �������� ���������"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   1200
      Picture         =   "Izv.frx":15CE
      Top             =   0
      Width           =   228
   End
End
Attribute VB_Name = "Izv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNastr As ADODB.Recordset
Dim rsDan As ADODB.Recordset
Dim rsLg As ADODB.Recordset
Dim oldKat, Filt As String
Dim ostPL As Double
Dim Itg As Double


Private Sub Command1_Click()
'������� ��������� ������
zRep = MainForm.Label3.Caption + " " + Label2.Caption + vbCrLf + "������ �/����� �" + Label6.Caption + " �� " + MainForm.Label8.Caption + " �." + vbCrLf
zRep = zRep + Me.Label7.Caption + "-" + Me.Label8.Caption + " " + Me.Label3.Caption + "-" + Me.Label4.Caption
zRep = zRep + vbCrLf + Me.Label11.Caption + "-" + Me.Label12.Caption + " " + Me.Label13.Caption + "-" + Me.Label16.Caption + " " + Me.Label20.Caption + "-" + Me.Label21.Caption + " " + Me.Label22.Caption + "-" + Me.Label23.Caption
zRep = zRep + " " + Me.Label14.Caption + "-" + Me.Label17.Caption


PrintW.Show
     With PrintW.VP
     
        PrintW.VP.StartDoc
        .FontSize = 12
        .Paragraph = zRep + vbNewLine + "_________________________________________________________________"
        .Paragraph = ""
        
        .FontSize = 8
        .RenderControl = fg1.hwnd
        .EndDoc
        
       End With


End Sub

Private Sub Command2_Click(Index As Integer)
fg1.Subtotal flexSTClear

'fg1.Sort = flexSortGenericAscending

fg1.OutlineBar = flexOutlineBarComplete
fg1.Subtotal flexSTSum, 1, 10, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
fg1.Subtotal flexSTSum, 1, 11, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
fg1.Subtotal flexSTSum, 1, 12, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
End Sub

Private Sub Command3_Click(Index As Integer)
fg1.Subtotal flexSTClear
fg1.OutlineBar = flexOutlineBarComplete
fg1.Subtotal flexSTSum, 1, 10, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
fg1.Subtotal flexSTSum, 1, 11, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
fg1.Subtotal flexSTSum, 1, 12, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)

fg1.Subtotal flexSTSum, 2, 10, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 2)
fg1.Subtotal flexSTSum, 2, 11, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 2)
fg1.Subtotal flexSTSum, 2, 12, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 2)

End Sub

Private Sub Command4_Click(Index As Integer)
fg1.Subtotal flexSTClear

fg1.OutlineBar = flexOutlineBarComplete
fg1.Subtotal flexSTSum, 1, 10, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
fg1.Subtotal flexSTSum, 1, 11, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
fg1.Subtotal flexSTSum, 1, 12, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)

fg1.Subtotal flexSTSum, 2, 10, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 2)
fg1.Subtotal flexSTSum, 2, 11, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 2)
fg1.Subtotal flexSTSum, 2, 12, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 3)

fg1.Subtotal flexSTSum, 3, 10, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 2)
fg1.Subtotal flexSTSum, 3, 11, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 2)
fg1.Subtotal flexSTSum, 3, 12, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 3)

End Sub

Private Sub Command5_Click(Index As Integer)
fg1.OutlineBar = 0
fg1.Subtotal flexSTClear
 
End Sub

Private Sub Command6_Click()
Msg.Show 1

End Sub

Private Sub Command7_Click()
fg1.Subtotal flexSTClear
fg1.OutlineBar = flexOutlineBarComplete
fg1.Subtotal flexSTSum, 0, 10, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
fg1.Subtotal flexSTSum, 0, 11, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
fg1.Subtotal flexSTSum, 0, 12, fg1.Cols, vbBlue, vbWhite, False, "� ���� " + fg1.TextMatrix(0, 1)
End Sub

Private Sub fg1_AfterMoveColumn(ByVal Col As Long, Position As Long)
fg1.Subtotal flexSTClear


End Sub

Private Sub Form_Load()


    
Filt = "Where (([arh_rep]![Tip] Like '*'))"
Filt = ""
'Where (((���_�����.������) Like "*"))

'********************************

Msg.Label1.Caption = "������ ����� ����� ��������!" + vbCrLf
Msg.Label1.Caption = vbCrLf + Msg.Label1.Caption + "�� ������ ��� ������ ���� �������� ������ � ������ ������� � ����� ������, � ��� �� ����������� ������� ������, � ����� ����� �������� �� ������ <0>,<1>,<2>,<3> ������ ������ ����������� � ������ �������� ����." + vbCrLf + vbCrLf + " ������: ���� �� ������ �������� ����� ������������ ����� ����� ���������� �� ������ ���� ����������:" + vbCrLf
Msg.Label1.Caption = Msg.Label1.Caption + "1. ��� ������ ���� ����������� ������� ��� �����" + vbCrLf + "2. ������ ������� <1>" + vbCrLf + "������ <0> ���������� ����� � �������� ���������" + vbCrLf
Msg.Label1.Caption = Msg.Label1.Caption + "�� ������ ����� ��������� �����, ������� �� ������ �� ������"


MakeWindow Me, True
'lblTitle.Caption = "������ �/�����"
lblTitle.Caption = MainForm.Label3.Caption + "/������ �/����� �� " + MainForm.Label8.Caption + " �."
Set rsNastr = New ADODB.Recordset
Set rsDan = New ADODB.Recordset
Set rsLg = New ADODB.Recordset

rsNastr.Open ("Settings"), Mconn
'��������� ��� ������ �/��
rsDan.Open ("SELECT MainOccupant.Numer, KLS_PODR.NAIM_KLS, MainOccupant.OLDNUM, MainOccupant.BanKN, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.NLODGERF, MainOccupant.NLODGER, MainOccupant.COMSPACE, MainOccupant.Priv, MainOccupant.Kv_num,  MainOccupant.Floor FROM MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.��� WHERE (((MainOccupant.Numer)=" + Filter.Nm + "))"), Mconn

'��������� ��� �����
'rsLg.Open ("SELECT Adding.NameKat AS [��������� �������], TMP_LGOTA.NAME_KLS AS [������], TMP_LGOTA.Use AS ����������, TMP_LGOTA.Procent AS �������, Adding.ObPl AS [����� �������], Adding.Tarif AS �����, TMP_LGOTA.Koll AS [���������� �������], TMP_LGOTA.PloLG AS [������� �������] FROM Adding INNER JOIN TMP_LGOTA ON Adding.Key = TMP_LGOTA.UniKOd WHERE (((Adding.KodKv)=" + Filter.Nm + ") AND ((TMP_LGOTA.Prim)>0))"), Mconn
'rsLg.Open ("SELECT arh_rep.DataT, arh_rep.NameKat, arh_rep.Tip, arh_rep.KodN, arh_rep.NameN, arh_rep.ObPl, arh_rep.Propis, arh_rep.Tarif, arh_rep.TarifI, arh_rep.Shc_new, arh_rep.SummaI FROM arh_rep"), Mconn, adOpenKeyset, adLockPessimistic
'rsLg.Open ("SELECT arh_rep.DataR AS ����, arh_rep.NameKat AS ���������, arh_rep.KodN AS ���, arh_rep.NameN AS ����������, arh_rep.ObPl AS �������, arh_rep.Propis AS ���������, arh_rep.Tarif AS [�����(���)], arh_rep.TarifI AS [�����(���)], arh_rep.Shc_new AS �������, IIf([arh_rep]![Tip]='+',[arh_rep]![SummaI],0) AS ���������, IIf([arh_rep]![Tip]='-',[arh_rep]![SummaI],0) AS ��������, IIf([arh_rep]![Tip]='s',[arh_rep]![SummaI],0) AS �������� FROM arh_rep ORDER BY arh_rep.DataR"), Mconn

'���� �����
If Lic.TipArh = "all" Then rsLg.Open ("SELECT Year([arh_rep]![DataR]) as ���, Month([arh_rep]![DataR]) as �����, arh_rep.NameKat AS ���������, arh_rep.KodN AS ���, arh_rep.NameN AS ����������, arh_rep.ObPl AS �������, arh_rep.Propis AS ���������, arh_rep.Tarif AS [�����(���)], arh_rep.TarifI AS [�����(���)], arh_rep.Shc_new AS �������, IIf([arh_rep]![Tip]='+',[arh_rep]![SummaI],0) AS ���������, IIf([arh_rep]![Tip]='-',[arh_rep]![SummaI],0) AS ��������, IIf([arh_rep]![Tip]='s',[arh_rep]![SummaI],0) AS �������� FROM arh_rep " + Filt + " ORDER BY arh_rep.DataR"), Mconn
'������ ������
If Lic.TipArh = "opl" Then rsLg.Open ("SELECT [arh_rep]![DataR] AS ������, [arh_rep]![NameN] AS ������������, [arh_rep]![SummaI] AS ��������, [arh_rep]![Com] AS ����������� From arh_rep WHERE (((arh_rep.Tip)=" + "'-'" + ")) ORDER BY [arh_rep]![DataR] DESC"), Mconn

'������ ����������
If Lic.TipArh = "nach" Then rsLg.Open ("SELECT [arh_rep]![DataR] AS ������, [arh_rep]![NameN] AS ������������, [arh_rep]![SummaI] AS ���������, [arh_rep]![Com] AS ����������� From arh_rep WHERE (((arh_rep.Tip)=" + "'+'" + ")) ORDER BY [arh_rep]![DataR] DESC"), Mconn


'MsgBox (rsLg.RecordCount)
'Label10.Caption = rsLg.RecordCount


'Str (Year([arh_rep]![DataR]))
Set fg1.DataSource = rsLg



fg1.Refresh

'****************************
 
fg1.AllowUserResizing = flexResizeBoth
fg1.Sort = flexSortGenericAscending
'fg1.Cols = G
fg1.ExplorerBar = flexExMove
fg1.MergeCells = flexMergeRestrictAll
fg1.MergeCol(-1) = True
fg1.MergeCol(fg1.Cols - 1) = False
fg1.MergeCol(-1) = True
'�����������
fg1.MergeCells = flexMergeRestrictAll
fg1.MergeCol(-1) = True
fg1.Refresh
fg1.Sort = flexSortGenericAscending
fg1.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove
fg1.RowHeight(0) = 500
fg1.WordWrap = True
fg1.Cell(flexcpAlignment, 0, 0, 0, fg1.Cols - 1) = flexAlignCenterCenter


  ' ���������� ������� ������ (��� �������)
'FG.MergeCells = flexMergeRestrictAll



 
 '       FG1.FixedCols = 0
        'FG1.GridLinesFixed = flexGridExplorer
       'FG1.AllowUserResizing = flexResizeBoth
        
        'FG1.Editable = 2
        
  '      FG1.ExplorerBar = flexExMoveRows Or flexExSortShowAndMove


      


 
 
 
'************************
 
 
 
 
 
rsNastr.MoveFirst
'Label1.Caption = rsNastr("NamePred")
Label2.Caption = rsNastr("Bank") + " ���:" + rsNastr("Bik") + " �/��" + rsNastr("Ks") + " �/��" + rsNastr("Rs")
Label4.Caption = rsDan("NAIM_KLS")
Label6.Caption = rsDan("OLDNUM") + " / " + rsDan("BankN")
'Label6.ToolTipText = "<������> � �/��" + vbNewLine + "<�����> 12-� ������� � �/��. ��� �������� � ������"
Label8.Caption = rsDan("FAM") + " " + rsDan("Im") + " " + rsDan("Ot")


Label12.Caption = rsDan("kv_num")

Label16.Caption = rsDan("Comspace")
Label17.Caption = rsDan("Floor")
Label21.Caption = rsDan("NlodgerF")
Label23.Caption = rsDan("Nlodger")
'Label19.Caption = MainForm.Label8.Caption + " �."
'Me.Caption = "������ �/����� �� " + MainForm.Label8.Caption + " �."

'Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsNastr.Close
rsDan.Close
Set rsNastr = Nothing
Set rsDan = Nothing
Msg.Label1.Caption = ""
End Sub

Private Sub imgTitleHelp_Click()
Unload Me
End Sub


Private Sub Label1_Click()
Msg.Show 1
End Sub

Private Sub VS_Click()
'VS.MergeCells = flexMergeRestrictAll
        
       ' sort the data from first to last column
 '       VS.Select 1, 0, 1, VS.Cols - 1
  '      VS.Sort = flexSortGenericAscending
   '     VS.Select 1, 0
        ' calculate subtotals
 '    VS.Subtotal flexSTClear

'VS.Subtotal flexSTSum, 1, 11, VS.Cols, vbBlue, vbWhite, False, "� ���� �� ������"
'VS.Subtotal flexSTSum, 2, 11, VS.Cols, vbBlue, vbWhite, False, "� ���� �� ���������"
'VS.Subtotal flexSTSum, 3, 11, VS.Cols, vbBlue, vbWhite, False, "� ���� "
Me.Show
VS.Refresh
End Sub

Private Sub Label9_Click()
Msg.Show 1
End Sub
