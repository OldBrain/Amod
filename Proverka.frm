VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Proverka 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7608
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   12804
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   634
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1067
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid Pr2 
      Height          =   2052
      Left            =   360
      TabIndex        =   14
      Top             =   4920
      Width           =   12132
      _cx             =   21399
      _cy             =   3619
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000016&
      Height          =   375
      Left            =   1320
      Picture         =   "Proverka.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid pr1 
      Height          =   1212
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   12612
      _cx             =   22246
      _cy             =   2138
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
      FormatString    =   $"Proverka.frx":0344
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
   Begin VSFlex8Ctl.VSFlexGrid PR 
      Height          =   1212
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   12612
      _cx             =   22246
      _cy             =   2138
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
      ForeColor       =   255
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   0
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
      FormatString    =   $"Proverka.frx":0462
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "Закрыть"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "На эти записи нет документе в реестре"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   16
      Top             =   4680
      Width           =   12132
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
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
      Left            =   8520
      TabIndex        =   15
      Top             =   7200
      Width           =   1092
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
      Picture         =   "Proverka.frx":058C
      Top             =   0
      Width           =   192
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "В результате проверки выявлены следующие ошибки:"
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
      Left            =   240
      TabIndex        =   13
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   12330
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   4800
      Picture         =   "Proverka.frx":06B3
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   5160
      Picture         =   "Proverka.frx":0DFD
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   4560
      Picture         =   "Proverka.frx":1547
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   3240
      TabIndex        =   12
      Top             =   360
      Width           =   195
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   8
      X2              =   848
      Y1              =   190
      Y2              =   190
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   9000
      TabIndex        =   10
      Top             =   2400
      Width           =   2292
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Не разнесено всего на сумму:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5760
      TabIndex        =   9
      Top             =   2400
      Width           =   3132
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   7920
      TabIndex        =   8
      Top             =   4320
      Width           =   2292
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Общая сумма расхождений"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3960
      TabIndex        =   7
      Top             =   4320
      Width           =   2772
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "(введены непосредственно из л/сч)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   4440
      TabIndex        =   6
      Top             =   2760
      Width           =   3732
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "(Не разнесены)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8040
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Есть в л/сч но нет в документах"
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
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4092
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Есть в документах, но отсутствуют в л/сч следующие начисления"
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
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7695
   End
End
Attribute VB_Name = "Proverka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Prov As ADODB.Recordset
Dim Prov1 As ADODB.Recordset
Dim Prov2 As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim s As Double


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Pod.Show
Pod.Label1 = "Подождите идет экспорт данных в XL"

For I = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = I
 For J = 1 To 1000
    Next J
   Next

pr1.Subtotal flexSTClear
For I = 250 To 500
    Pod.ProgressBar1.Value = I
 For J = 1 To 1000
    Next J
   Next

pr1.DataRefresh
For I = 500 To 750
    Pod.ProgressBar1.Value = I
    
 For J = 1 To 1000
    Next J
   Next

ВыводВExel
For I = 750 To 1000
    Pod.ProgressBar1.Value = I
    
 For J = 1 To 1000
    Next J
   Next

Unload Pod
End Sub

Private Sub Form_Load()
MakeWindow Me, True


'Set mconn = New ADODB.Connection
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'mconn.Open "data/Kvartplata.mdb"

Set Prov = New ADODB.Recordset
Set Prov.ActiveConnection = Mconn
Prov.CursorType = adOpenStatic
Prov.LockType = adLockBatchOptimistic

Set Prov1 = New ADODB.Recordset
Set Prov1.ActiveConnection = Mconn
Prov1.CursorType = adOpenStatic
Prov1.LockType = adLockBatchOptimistic


Set Prov2 = New ADODB.Recordset
Set Prov2.ActiveConnection = Mconn
Prov2.CursorType = adOpenStatic
Prov2.LockType = adLockBatchOptimistic



'Prov1.Open ("SELECT Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.NameN, Adding.SummaI, Adding.Tip, Adding.Com FROM MainOccupant RIGHT JOIN (Adding LEFT JOIN Doc ON Adding.KodDoc = Doc.Key) ON MainOccupant.Numer = Adding.KodKv WHERE (((Adding.Tip)='-') AND ((Doc.Cod) Is Null))")
'("Proverka")

'Изменяем условия проверки Теперь нулевые суммы на выводятся
Prov1.Open ("SELECT Adding.KodKv, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, Adding.NameN, Adding.SummaI, Adding.Tip, Adding.Com FROM MainOccupant RIGHT JOIN (Adding LEFT JOIN Doc ON Adding.KodDoc = Doc.Key) ON MainOccupant.Numer = Adding.KodKv WHERE (((Adding.SummaI)<>0) AND ((Adding.Tip)='-') AND ((Doc.Cod) Is Null))")


Prov.Open ("SELECT Doc.Cod, Doc.DataR, Doc.KodN, Doc.NameN, Doc.KodKv, Doc.NameKv, Doc.Summa, Doc.Key, Doc.Com, Doc.Stst FROM Doc LEFT JOIN Adding ON Doc.Key = Adding.KodDoc WHERE (((Adding.KodDoc) Is Null))")

'Проверка на записи в DOC без записей в реестре

Prov2.Open ("SELECT doc.Cod, doc.DataR, doc.KodN, doc.NameN, doc.KodKv, doc.NameKv, doc.Summa, doc.Com FROM doc LEFT JOIN ReestrDoc ON doc.Cod = ReestrDoc.Cod WHERE (((ReestrDoc.Cod) Is Null))")



Set PR.DataSource = Prov
Set pr1.DataSource = Prov1
Set Pr2.DataSource = Prov2

цвет
Prov.Close
End Sub

Private Sub цвет()
Dim rw As Integer
s = 0
For rw = 1 To pr1.Rows - 1
s = s + pr1.TextMatrix(rw, 6)
'MsgBox (FG.TextMatrix(Rw, 6))
If pr1.TextMatrix(rw, 2) <> pr1.TextMatrix(rw, 4) Then
pr1.Cell(flexcpFontBold, rw, 1, rw, 5) = True
End If
Next rw
Label7 = s

s = 0
For rw = 1 To PR.Rows - 1
s = s + PR.TextMatrix(rw, 7)


Next rw
Label9 = s

'*************
s = 0
For rw = 1 To Pr2.Rows - 1
'MsgBox (Pr2.TextMatrix(rw, 6))
s = s + Pr2.TextMatrix(rw, 7)
Next rw
Label1 = s




End Sub

Sub ВыводВExel()
   Const НачСтрока = 1
   Dim RS As New ADODB.Recordset
   Dim ex1 As Object ' Excel.Application
   Dim wb As Object ' Excel.Workbook
   Dim ws As Object ' Excel.Worksheet
   Dim I As Long, J As Long, K As Long, rДанные As String
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
'   rДанные = "A" & (НачСтрока + 1) & ":" & XCol_(k) & Rs_kat.RecordCount + НачСтрока
   
   rДанные = "A" & (НачСтрока + 1) & ":" & XCol_(pr1.Cols - 1) & pr1.Rows + НачСтрока
   ReDim v(pr1.Rows, pr1.Cols) 'Забыл указать
'   If rs.RecordCount > 0 Then

   'If Rs_kat.RecordCount > 0 Then
   If pr1.Rows > 0 Then
    '  Rs_kat.MoveFirst
      'i = 0
      'Do Until Rs_kat.EOF
         For co = 1 To pr1.Cols - 1
         For rw = 0 To pr1.Rows - 1
             'v(i, j) = Rs_kat.Fields(j).Value
             v(rw, co) = pr1.TextMatrix(rw, co)
             
         Next rw
         Next co
         'Rs_kat.MoveNext
      'Loop
      ex1.Visible = True   'Еще забыл
      
      ws.Range(rДанные) = v
      
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

Private Sub imgTitleHelp_Click()
Command1_Click
End Sub

Private Sub imgTitleHelp_DblClick()
Command1_Click
End Sub

