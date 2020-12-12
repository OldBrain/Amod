VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form VibTablDoc 
   Caption         =   "Окно выбора"
   ClientHeight    =   7692
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11040
   LinkTopic       =   "Form8"
   ScaleHeight     =   7692
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnEnh1 
      Caption         =   "Отметить всех"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Отмена"
      Height          =   255
      Left            =   7080
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Включить отмеченных в документ"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Снять отметку со всех"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VSFlex8Ctl.VSFlexGrid VSv 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10575
      _cx             =   18653
      _cy             =   11245
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
      FormatString    =   $"VibTablDoc.frx":0000
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
Attribute VB_Name = "VibTablDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsV As ADODB.Recordset
Dim Doc As ADODB.Recordset
Dim rsDan As ADODB.Recordset
'Dim mconn As ADODB.Connection



Private Sub BtnEnh1_Click()
For R = 1 To VSv.Rows - 1
VSv.Cell(flexcpChecked, R, 0) = flexChecked
Next
End Sub



Private Sub BtnEnh11_Click()

End Sub

Private Sub Command2_Click()
For R = 1 To VSv.Rows - 1
VSv.Cell(flexcpChecked, R, 0) = flexUnchecked
Next
End Sub

Private Sub Command3_Click()
Set rsDan = New ADODB.Recordset


Jdite.Show
For R = 1 To VSv.Rows - 1
If VSv.Cell(flexcpChecked, R, 0) = flexChecked Then

'MsgBox "INSERT INTO TablDoc ( Cod, TabNum, Fam, Im, Ot, KvNum, Kodn, FItog ) SELECT " + Str(ReestrTablDoc.N + 1) + " AS Выражение1 ," + VSv.TextMatrix(r, 1) + " AS Выражение2 ," + VSv.TextMatrix(r, 2) + " AS Выражение3 ," + VSv.TextMatrix(r, 3) + " AS Выражение4 ," + VSv.TextMatrix(r, 4) + " AS Выражение5 ," + VSv.TextMatrix(r, 5) + " AS Выражение6 ," + Str(ShapkaTBL.KNac) + " AS Выражение7 ," + ShapkaTBL.Text2 + " AS Выражение8"
Doc.AddNew

Doc.Fields("Cod") = ReestrTablDoc.n + 1
Doc.Fields("TabNum") = VSv.TextMatrix(R, 1)
Doc.Fields("Fam") = VSv.TextMatrix(R, 2)
Doc.Fields("Im") = VSv.TextMatrix(R, 3)
Doc.Fields("Ot") = VSv.TextMatrix(R, 4)
Doc.Fields("KvNum") = VSv.TextMatrix(R, 5)
Doc.Fields("Kodn") = ShapkaTBL.KNac
Doc.Fields("Формула") = ShapkaTBL.Text2



'************************************
'Если расчитываем простои то...
If ShapkaTBL.Prostoy = True Then
' Открываем данные из Аддинг
kg = Val(ShapkaTBL.Combo1.Text)
rsDan.Open ("SELECT Adding.*, Adding.KodKv, Adding.KodN From Adding WHERE (((Adding.KodKv)=" + Str(VSv.TextMatrix(R, 1)) + ") AND ((Adding.Kodkat)=" + Str(kg) + "))"), Mconn

' Ячейка S1 - тариф
If rsDan.Fields("tarif") <> "" Then Doc.Fields("S1") = rsDan.Fields("tarif")
' Ячейка S2 - площадь
If rsDan.Fields("obpl") <> "" Then Doc.Fields("S2") = rsDan.Fields("obpl")


' Ячейка S3 - количество дней в месяц
Doc.Fields("S3") = DateDiff("d", MainForm.DR, DateAdd("m", 1, MainForm.DR))

' Если это документ на простой то заполняем F5 = "Prostoy"
Doc.Fields("F5") = "Prostoy"

rsDan.Close
End If
'************************************

'************************************
'Если расчитываем ODN...
If ShapkaTBL.Odn = True Then
' Открываем данные из Аддинг
kg = Val(ShapkaTBL.Combo1.Text)

'rsDan.Open ("SELECT Adding.*, Adding.KodKv, Adding.KodN From Adding WHERE (((Adding.KodKv)=" + Str(VSv.TextMatrix(R, 1)) + ") AND ((Adding.Kodkat)=" + Str(kg) + "))"), Mconn


rsDan.Open ("SELECT Adding.*, Adding.KodN, MainOccupant.COMSPACE FROM Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer WHERE (((Adding.KodKv)=" + Str(VSv.TextMatrix(R, 1)) + ") AND ((Adding.KodKat)=" + Str(kg) + "))"), Mconn

'rsDan.MoveLast


      If rsDan.EOF = False Or rsDan.BOF = False Then 'Проверка что рекорсет не пустой


' Ячейка S1 - тариф
If rsDan.Fields("tarif") <> "" Then Doc.Fields("S1") = rsDan.Fields("tarif")

' Ячейка S2 - площадь
'If rsDan.Fields("obpl") <> "" Then Doc.Fields("S2") = rsDan.Fields("obpl")
If rsDan.Fields("obpl") <> "" Then Doc.Fields("S2") = rsDan.Fields("COMSPACE")



' Ячейка S3 - ОБЩАЯ площадь
Doc.Fields("S3") = ShapkaTBL.Text3.Text

' Ячейка S4 - РАЗНИЦА ДАННЫЕ ОБЩЕГО СЧЕТЧИКА-СУМАРНЫЕ СЧЕТЧИКИ
Doc.Fields("S4") = CVar(ShapkaTBL.Text5.Text) - ShapkaTBL.ODN_Sc
'VOL (ShapkaTBL.Text5.Text) - VOL(ShapkaTBL.Text4.Text)


' Если это документ на простой то заполняем F5 = "Prostoy"
Doc.Fields("f5") = "odn"


End If
rsDan.Close

                            End If

'rsDan.Close
'***************************************


Doc.UpdateBatch

'mconn.Execute ("INSERT INTO TablDoc ( Cod, TabNum, Fam, Im, Ot, KvNum, Kodn, FItog ) SELECT " + Str(ReestrTablDoc.N + 1) + " AS Выражение1 ," + VSv.TextMatrix(r, 1) + " AS Выражение2 ," + VSv.TextMatrix(r, 2) + " AS Выражение3 ," + VSv.TextMatrix(r, 3) + " AS Выражение4 ," + VSv.TextMatrix(r, 4) + " AS Выражение5 ," + VSv.TextMatrix(r, 5) + " AS Выражение6 ," + Str(ShapkaTBL.KNac) + " AS Выражение7 ," + ShapkaTBL.Text2 + " AS Выражение8")
End If
Next

Unload Jdite
Unload ShapkaTBL
Unload VibTablDoc
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
PodD = ShapkaTBL.Combo4.Text
'Set mconn = New ADODB.Connection
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'mconn.Open "data/Kvartplata.mdb"
Set RsV = New ADODB.Recordset
Set RsV.ActiveConnection = Mconn

Set Doc = New ADODB.Recordset
Set Doc.ActiveConnection = Mconn

Doc.Open ("Tabldoc"), Mconn, adOpenKeyset, adLockPessimistic

If (PodD = "" Or PodD = "*") Then
RsV.Open ("SELECT MainOccupant.Numer, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, MainOccupant.FLOOR, MainOccupant.Dom, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(ShapkaTBL.KAdres) + "))")
Else
RsV.Open ("SELECT MainOccupant.Numer, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, MainOccupant.FLOOR, MainOccupant.Dom, MainOccupant.podyezd From MainOccupant WHERE (((MainOccupant.Dom)=" + Str(ShapkaTBL.KAdres) + ") AND ((MainOccupant.podyezd)=" + PodD + "))")
End If


VSv.Cols = 7
Set VSv.DataSource = RsV

For R = 1 To VSv.Rows - 1
VSv.Cell(flexcpChecked, R, 0) = flexChecked
Next

End Sub

