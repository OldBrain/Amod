VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form SCH_ET 
   ClientHeight    =   8172
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12252
   ControlBox      =   0   'False
   Icon            =   "SCH_ET.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   681
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1021
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Tarifs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   5520
      TabIndex        =   8
      Text            =   "0"
      Top             =   720
      Width           =   732
   End
   Begin KvPay.xpcmdbutton xpcmdbutton2 
      Height          =   492
      Left            =   11280
      TabIndex        =   2
      Top             =   480
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   868
      Caption         =   "�����"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   6732
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   12132
      _cx             =   21399
      _cy             =   11874
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483624
      ForeColorSel    =   -2147483646
      BackColorBkg    =   -2147483637
      BackColorAlternate=   16777215
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
      SelectionMode   =   3
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
      FormatString    =   $"SCH_ET.frx":0CCA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   5
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
      DataMode        =   0
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
      Begin VB.CommandButton Command2 
         Caption         =   "��������� ������ ���������"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Width           =   7332
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   4680
         TabIndex        =   14
         Top             =   -240
         Width           =   120
      End
   End
   Begin KvPay.xpcmdbutton xpcmdbutton1 
      Height          =   492
      Left            =   7560
      TabIndex        =   11
      Top             =   480
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   868
      Caption         =   "������"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton3 
      Height          =   492
      Left            =   6480
      TabIndex        =   12
      Top             =   480
      Width           =   1092
      _ExtentX        =   1926
      _ExtentY        =   868
      Caption         =   "������� � XL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KvPay.xpcmdbutton xpcmdbutton4 
      Height          =   492
      Left            =   8520
      TabIndex        =   13
      Top             =   480
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   868
      Caption         =   "��������� "
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
   Begin KvPay.xpcmdbutton xpcmdbutton5 
      Height          =   492
      Left            =   9720
      TabIndex        =   15
      Top             =   480
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   868
      Caption         =   "��������� � �/�� "
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
   Begin VB.Line Line6 
      X1              =   460
      X2              =   460
      Y1              =   40
      Y2              =   60
   End
   Begin VB.Line Line5 
      X1              =   20
      X2              =   460
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Line Line4 
      X1              =   530
      X2              =   530
      Y1              =   40
      Y2              =   90
   End
   Begin VB.Line Line3 
      X1              =   20
      X2              =   20
      Y1              =   40
      Y2              =   90
   End
   Begin VB.Line Line2 
      X1              =   20
      X2              =   530
      Y1              =   90
      Y2              =   90
   End
   Begin VB.Line Line1 
      X1              =   20
      X2              =   530
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label Q 
      Caption         =   "Select"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   7920
      Visible         =   0   'False
      Width           =   11052
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   5520
      TabIndex        =   7
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Adr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   6
      Top             =   720
      Width           =   4332
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��������� �������"
      Height          =   252
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label KodKat 
      BackStyle       =   0  'Transparent
      Caption         =   "KodKat"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   372
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������ ���������"
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
      TabIndex        =   1
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   12210
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
      Height          =   228
      Left            =   0
      Picture         =   "SCH_ET.frx":0DB7
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   456
      Left            =   960
      Picture         =   "SCH_ET.frx":13D9
      Stretch         =   -1  'True
      ToolTipText     =   "������� ������ ���� ��������� ����� �� ���� ����� ��� ������ � �������� ���������"
      Top             =   0
      Width           =   288
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   240
      Picture         =   "SCH_ET.frx":1B23
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   1560
      Picture         =   "SCH_ET.frx":226D
      Top             =   0
      Width           =   228
   End
End
Attribute VB_Name = "SCH_ET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCH As ADODB.Recordset
Dim OPL As ADODB.Recordset
Dim ColI As ADODB.Recordset






Private Sub Command2_Click()
Command2.Visible = False
rsCH.Open (Me.Q.Caption), Mconn
Set FG.DataSource = rsCH
FG.Sort = flexSortCustom
'FG.Cols = 9

'���� �� ������� ����� �������� ������
For R = 1 To FG.Rows - 1
'MsgBox (FG.TextMatrix(R, 1))
' ��������� ������ �� Adding
Kodo = FG.TextMatrix(R, 1)
OPL.Open ("SELECT Adding.Tip, Adding.KodKv, Sum(Adding.SummaI) AS [Sum-SummaI] From Adding GROUP BY Adding.KodKat, Adding.Tip, Adding.KodKv HAVING (((Adding.KodKat)=" + Me.KodKat.Caption + ") AND ((Adding.Tip)='-') AND ((Adding.KodKv)=" + Kodo + "))"), Mconn, adOpenStatic

'OPL
If OPL.BOF = False Or OPL.EOF = False Then

'��������� ��� ����������� ���������� ��������� �� ��������� ��������� � ���.�����
ColI.Open ("SELECT Adding.Tip, Adding.KodKv, Adding.KodKat, Count(Adding.KodKv) AS [Count-KodKv] From Adding Where (((Adding.Sch) = '��'))GROUP BY Adding.Tip, Adding.KodKv, Adding.KodKat HAVING (((Adding.Tip)='+') AND ((Adding.KodKv)=" + Kodo + ") AND ((Adding.KodKat)=" + Me.KodKat.Caption + "))"), Mconn

 ' ��� ������� ������ �� ��������� ��� ������� ����� ������ �������� �� ������ ���������
 ' ����� ������ �� ���������� ��������� ������� ������� �� ���������� ColI
 
 FG.TextMatrix(R, 9) = Round(OPL("Sum-SummaI") / ColI("Count-KodKv"), 2)
'FG.TextMatrix(R, 9) = OPL("Sum-SummaI")
 ColI.Close
End If
OPL.Close

Next
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If FG.TextMatrix(Row, 7) > FG.TextMatrix(Row, 8) Then
If MsgBox("������� ������ �������� ������ ����������, �� �������?", vbYesNo, "") = vbNo Then FG.TextMatrix(Row, 10) = FG.TextMatrix(Row, 9)

End If
End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> 7 And Col <> 8 Then
'MsgBox "� ���� ���� ����� ������� ������ ������ ��������!"
Msg.Label1.Caption = "� ���� ���� ����� ������� ������ ������ ��������!"
Msg.Show 1

Cancel = True
End If


End Sub

Private Sub Form_Load()


'Me.KodKat.Caption = s�h_kat


MakeWindow Me, True
Set rsCH = New ADODB.Recordset
Set OPL = New ADODB.Recordset
Set ColI = New ADODB.Recordset



End Sub

Private Sub imgTitleHelp_Click()
Unload Me
End Sub



Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub xpcmdbutton1_Click()
MainMenu.Enabled = False
ScReport.Show
Unload Me
End Sub

Private Sub xpcmdbutton2_Click()

If MsgBox("���������� ������ ��������� � ������� �����?", vbYesNo, "") = vbYes Then
'FG.TextMatrix(R, 8) = "������� �������"
For R = 1 To FG.Rows - 1
'MsgBox (FG.TextMatrix(R, 8))
Nch = Replace(FG.TextMatrix(R, 8), ",", ".")

Mconn.Execute ("UPDATE Adding SET Adding.Shc_new = " + Nch + " WHERE (((Adding.Tip)='+') AND ((Adding.KodKv)=" + FG.TextMatrix(R, 1) + ") AND ((Adding.KodKat)=" + Me.KodKat.Caption + ") AND ((Adding.Sch)='��'))")
'UPDATE Adding SET Adding.Shc_new = "+FG.TextMatrix(R,8)+" WHERE (((Adding.Tip)='+') AND ((Adding.KodKv)="+Kodo+") AND ((Adding.KodKat)="+[Me].[KodKat].[Caption]+") AND ((Adding.Sch)='��'))

Next
Msg.Label1.Caption = "�� �������� ������ ���������, �� �������� ����������� ������� �����."
Msg.Show 1
MainMenu.Enabled = True
Unload Me

Else
MainMenu.Enabled = True
Unload Me
End If

End Sub

Private Sub xpcmdbutton3_Click()
Pod.Show
Pod.Label1 = "��������� ���� ������� ������ � XL"

For i = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = i
 For j = 1 To 1000
    Next j
   Next

FG.Subtotal flexSTClear
For i = 250 To 500
    Pod.ProgressBar1.Value = i
 For j = 1 To 1000
    Next j
   Next

FG.DataRefresh
For i = 500 To 750
    Pod.ProgressBar1.Value = i
    
 For j = 1 To 1000
    Next j
   Next

�Exl
For i = 750 To 1000
    Pod.ProgressBar1.Value = i
    
 For j = 1 To 1000
    Next j
   Next

Unload Pod




'MainMenu.Enabled = True
End Sub

Private Sub xpcmdbutton4_Click()
If MsgBox("��������� ������ �������� �� ������ ������ �� ������>" + Me.Tarifs.Text + " ���.  ����? (������ �� ������� ����� ����� ��������� �� ����� � ���������� ��������� ��������� � ������ �������� �� ������� �����.) !!��������!! ���� � ������� ����� ������� ����� ������ �������� �� ��������� ���� ��������� ������� ������ ����� ��������� ��������������� ���������� ���������. �.�. ���� ������ ��������� 1000 ���., � � ������� ����� ��� �������� �� �� ������ ������� �������� ����� ��������� �� 500 ���.)", vbYesNo, "") = vbYes Then
'Mconn.Execute ("UPDATE Adding SET Adding.Shc_new = [Adding]![Shc_old] WHERE (((Adding.Shc_new)=0))")

'���� �� ������� ����� �������� ������
' = "������� ����������"
'FG.TextMatrix(R, 8) = "������� �������"
'FG.TextMatrix(R, 9)= "��������"
'FG.TextMatrix(R, 6)= "���������"
'me.Tarifs.Text  - �����

For R = 1 To FG.Rows - 1
' ����������� ������� ������� �� ������ ������
FG.TextMatrix(R, 8) = Round(FG.TextMatrix(R, 7) + FG.TextMatrix(R, 9) / Val(Me.Tarifs.Text), 2)

Next
End If
End Sub

Private Sub xpcmdbutton5_Click()

If MsgBox("���������� ������ ��������� � ������� �����?", vbYesNo, "") = vbYes Then
'FG.TextMatrix(R, 8) = "������� �������"
For R = 1 To FG.Rows - 1
'MsgBox (FG.TextMatrix(R, 8))
Nch = Replace(FG.TextMatrix(R, 8), ",", ".")

Mconn.Execute ("UPDATE Adding SET Adding.Shc_new = " + Nch + " WHERE (((Adding.Tip)='+') AND ((Adding.KodKv)=" + FG.TextMatrix(R, 1) + ") AND ((Adding.KodKat)=" + Me.KodKat.Caption + ") AND ((Adding.Sch)='��'))")
'UPDATE Adding SET Adding.Shc_new = "+FG.TextMatrix(R,8)+" WHERE (((Adding.Tip)='+') AND ((Adding.KodKv)="+Kodo+") AND ((Adding.KodKat)="+[Me].[KodKat].[Caption]+") AND ((Adding.Sch)='��'))

Next
Msg.Label1.Caption = "�� �������� ������ ���������, �� �������� ����������� ������� �����."
Msg.Show 1
MainMenu.Enabled = True
'Unload Me

Else
'MainMenu.Enabled = True
'Unload Me
End If
End Sub
Sub �Exl()
   Const ��������� = 1
   Dim RS As New ADODB.Recordset
   Dim ex1 As Object ' Excel.Application
   Dim wb As Object ' Excel.Workbook
   Dim ws As Object ' Excel.Worksheet
   Dim i As Long, j As Long, K As Long, r������ As String
   Dim v As Variant
   
   
'rs.CursorType = adOpenStatic
'rs.LockType = adLockReadOnly



   Set ex1 = CreateObject("Excel.Application")  'New Excel.Application
   Set wb = ex1.Workbooks.Add
   Set ws = wb.Sheets(1)
'   Set rs = Rs_kat.Clone

 'Set rs = Rs_kat
  ' rs.Filter = Rs_kat.Filter
   'rs.Sort = Rs_kat.Sort
 ' k = FG1.Rows - 1
  ' Rs_kat.MoveLast
'   r������ = "A" & (��������� + 1) & ":" & XCol_(k) & Rs_kat.RecordCount + ���������
   
   r������ = "A" & (��������� + 1) & ":" & XCol_(FG.Cols - 1) & FG.Rows + ���������
   ReDim v(FG.Rows, FG.Cols) '����� �������
'   If rs.RecordCount > 0 Then

   'If Rs_kat.RecordCount > 0 Then
   If FG.Rows > 0 Then
    '  Rs_kat.MoveFirst
      'i = 0
      'Do Until Rs_kat.EOF
         For co = 1 To FG.Cols - 1
         For rw = 0 To FG.Rows - 1
             'v(i, j) = Rs_kat.Fields(j).Value
             v(rw, co) = FG.TextMatrix(rw, co)
             
         Next rw
         Next co
         'Rs_kat.MoveNext
      'Loop
      ex1.Visible = True   '��� �����
      
      ws.Range(r������) = v
      
 End If
End Sub

Function XCol_(ByVal Column_ As Long) As String
    If (Column_ < 0) Then Column_ = 0
    If (Column_ < 26) Then
        XCol_ = Chr(Column_ + Asc("A"))
    ElseIf (Column_ < 676) Then
        XCol_ = Chr((Column_ \ 26) + Asc("A") - 1) & Chr((Column_ Mod 26) + Asc("A"))
    Else
        XCol_ = "ZZ"
    End If
End Function

