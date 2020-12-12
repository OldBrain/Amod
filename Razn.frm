VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Razn 
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form7"
   ScaleHeight     =   7920
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid LC 
      Height          =   1935
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   6375
      _cx             =   11245
      _cy             =   3413
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
      FormatString    =   $"Razn.frx":0000
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
   Begin VSFlex8Ctl.VSFlexGrid RZ 
      Height          =   3615
      Left            =   840
      TabIndex        =   2
      Top             =   3240
      Width           =   5295
      _cx             =   9340
      _cy             =   6376
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
      FormatString    =   $"Razn.frx":00FD
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
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Razn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rw, Cl As Integer
Dim Kat As ADODB.Recordset
Dim Lic As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim Summ, s, S1, SRz As Double

'Dim lbl(50) As Label




Private Sub Command1_Click()
Unload Me
Doc.Show
Doc.Enabled = True
End Sub

Private Sub Command2_Click()
Unload Me
Doc.Show
Doc.Enabled = True

End Sub

Private Sub Form_Load()
'Set mconn = New ADODB.Connection
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'mconn.Open "data/Kvartplata.mdb"

Set Kat = New ADODB.Recordset
Set Kat.ActiveConnection = mconn
Kat.CursorType = adOpenStatic
Kat.LockType = adLockBatchOptimistic
Kat.Open "SELECT Kategor.Name_Kategor, Kategor.Nac, 0 AS Сумма From Kategor WHERE (((Kategor.Nac)<>0))"
 


Set Lic = New ADODB.Recordset
Set Lic.ActiveConnection = mconn
Lic.CursorType = adOpenStatic

Lic.LockType = adLockBatchOptimistic
Lic.Open ("SELECT Adding.KodKv, Adding.SaldoN, Adding.NameKat, Sum(IIf([Adding]![Tip]=" + Chr(34) + "+" + Chr(34) + ",[SummaI],0)) AS Начислено, Sum(IIf([Adding]![Tip]=" + Chr(34) + "-" + Chr(34) + ",[SummaI],0)) AS Оплата, Sum(IIf([Adding]![Tip]=" + Chr(34) + "s" + Chr(34) + ",[SummaI],0)) AS Субсидии, Adding.SaldoK From Adding GROUP BY Adding.KodKv, Adding.SaldoN, Adding.NameKat, Adding.SaldoK Having (((Adding.kodkv) =" + Doc.FG.TextMatrix(Doc.FG.Row, 5) + ")) ORDER BY Adding.NameKat")

'+ Doc.FG.TextMatrix(Doc.FG.Row, 5) +


Set RZ.DataSource = Kat
Set LC.DataSource = Lic

LC.MergeCells = flexMergeRestrictAll
LC.MergeCol(-1) = True
LC.MergeCol(LC.Cols - 1) = False




RZ.Editable = flexEDKbdMouse
RZ.Cols = 4

i = 0
For Rw = 1 To RZ.Rows - 1

If RZ.TextMatrix(Rw, 2) = Doc.FG.TextMatrix(Doc.FG.Row, 3) Then
RZ.TextMatrix(Rw, 3) = Doc.FG.TextMatrix(Doc.FG.Row, 7)
i = i + 1
End If

Next

If i = 0 Then RZ.AddItem vbTab & Doc.FG.TextMatrix(Doc.FG.Row, 4) & vbTab & Doc.FG.TextMatrix(Doc.FG.Row, 3) & vbTab & Doc.FG.TextMatrix(Doc.FG.Row, 7)
цвет

Label1 = Doc.FG.TextMatrix(Doc.FG.Row, 6)
Итоги
'RZ.ColSort(3) = flexSortNumericAscending
End Sub

Private Sub RZ_AfterEdit(ByVal Row As Long, ByVal Col As Long)

If RZ.TextMatrix(RZ.Row, 3) = "" Then RZ.TextMatrix(RZ.Row, 3) = 0
If InStr(1, RZ.TextMatrix(RZ.Row, 3), ",") Then
'MsgBox (StringCleaner(RZ.TextMatrix(RZ.Row, 3), ",", "."))
RZ.TextMatrix(RZ.Row, 3) = StringCleaner(RZ.TextMatrix(RZ.Row, 3), ",", ".")

End If
Пересчет
RZ.TextMatrix(Cl, 3) = Summ - S1
Итоги
Label7.Refresh
End Sub

Private Sub RZ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
s = Val(RZ.TextMatrix(Cl, 3))
End Sub

Private Sub RZ_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If RZ.Col <> 3 Then Cancel = True
End Sub
Private Sub цвет()

For Rw = 1 To RZ.Rows - 1

If RZ.TextMatrix(Rw, 3) <> "0" Then
RZ.Cell(flexcpBackColor, Rw, 3) = RGB(200, 255, 200)
RZ.Cell(flexcpFontBold, Rw, 3, Rw, 3) = True
Cl = Rw
Summ = Val(RZ.TextMatrix(Rw, 3))
End If
Next Rw

For Rw = 1 To LC.Rows - 1
LC.Cell(flexcpForeColor, Rw, 4) = vbBlue
LC.Cell(flexcpForeColor, Rw, 5) = vbRed
LC.Cell(flexcpForeColor, Rw, 6) = vbGreen
'RGB(200, 100, 200)

Next


End Sub
Private Sub Итоги()
Dim sn, sk, na, op, su As Double
sn = 0
sk = 0
na = 0
op = 0
su = 0

For Rw = 1 To LC.Rows - 1
sn = sn + Round(LC.TextMatrix(Rw, 2), 2)
na = na + Round(LC.TextMatrix(Rw, 4), 2)
op = op + Round(LC.TextMatrix(Rw, 5), 2)
su = su + Round(LC.TextMatrix(Rw, 6), 2)
sk = sk + Round(LC.TextMatrix(Rw, 7), 2)
Next

Label2 = sn
Label3 = na
Label4 = op
Label5 = su
Label6 = sk

SRz = 0
For Rw = 1 To RZ.Rows - 1
SRz = SRz + Val(RZ.TextMatrix(Rw, 3))
Next
Label7 = SRz
End Sub


Sub Пересчет()
S1 = 0
For Rw = 1 To RZ.Rows - 1

If Rw <> Cl Then S1 = S1 + Val(RZ.TextMatrix(Rw, 3))
Next
MsgBox (Str(Cl) + "  " + Str(S1))
End Sub
Function StringCleaner(s As String, _
        Search As String, zam As String) As String
                Dim i As Integer, res As String
                res = s
                Do While InStr(res, Search)
                        i = InStr(res, Search)
                        res = Left(res, i - 1) & _
                                zam & Mid(res, i + 1)
                Loop
                StringCleaner = res
End Function


