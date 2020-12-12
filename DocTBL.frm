VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form DocTBL 
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10776
   LinkTopic       =   "Form8"
   ScaleHeight     =   8040
   ScaleWidth      =   10776
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Закрыть"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Разнести"
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Печать"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Расчитать"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Проставить всем"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   10455
      _cx             =   18441
      _cy             =   9763
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"DocTBL.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      DataMode        =   3
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   1
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
   Begin VB.Label Label10 
      Caption         =   "Формула тек.ячейки:"
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
      Left            =   360
      TabIndex        =   15
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   1440
      Width           =   8175
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
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
      Left            =   3480
      TabIndex        =   13
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Кол-во строк документа"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   11
      Top             =   7440
      Width           =   945
   End
   Begin VB.Label Label5 
      Caption         =   "И того по документу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Документ №"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "DocTBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cmb As ADODB.Recordset
Dim Inf As ADODB.Recordset
Dim Nac As ADODB.Recordset
'Public mconn As ADODB.Connection
Dim Dan As String
Public L As String
Public Naz As String
Dim Sum As Double
Dim Kol As Integer
Dim Fr As String
Dim Kt As String





Private Sub Command1_Click()

If Fg.Col <= 7 Then
MsgBox ("Эти данные изменять нельзя")
Exit Sub
End If
Dan = Fg.TextMatrix(Fg.Row, Fg.Col)
Naz = Fg.TextMatrix(0, Fg.Col)
Dan = Replace(Dan, ",", ".")
If MsgBox("Проставить всем значение " + Naz + " = " + Dan, vbYesNo) = vbYes Then

If Val(Dan) = 0 Then Dan = Chr(34) + Dan + Chr(34)
Q = "UPDATE TablDoc SET TablDoc." + Naz + " = " + Dan + " WHERE (((TablDoc.Cod)=" + L + "))"
'MsgBox Q

Mconn.Execute (Q)
Set Fg.DataSource = Cmb
End If

End Sub

Private Sub Command2_Click()

If Fg.Rows <= 1 Then Exit Sub

If MsgBox("Выполнить расчет? ", vbYesNo) = vbYes Then
Inf.MoveFirst
Do While Not Inf.EOF
'If Inf("f1") <> "" Or Inf("f1") <> "Нет формулы" Then

For rw = 1 To Fg.Rows - 1

If Fg.TextMatrix(rw, 15) = Inf("TablKod") Then

'********S1
If Inf("f1") = "Оплата по текущей категории расчета" Then
Fg.TextMatrix(rw, 8) = Oplata(Inf("TabNum"), Kt, "-")
End If
If Inf("f1") = "Начислено по текущей категории расчета" Then
Fg.TextMatrix(rw, 8) = Oplata(Inf("TabNum"), Kt, "+")
End If
If Inf("f1") = "Субсидия по текущей категории расчета" Then
Fg.TextMatrix(rw, 8) = Oplata(Inf("TabNum"), Kt, "s")
End If

If Inf("f1") = "Оплачено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 8) = Oplata(Inf("TabNum"), "All", "-")
End If

If Inf("f1") = "Начислено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 8) = Oplata(Inf("TabNum"), "All", "+")
End If

If Inf("f1") = "Субсидии всего. по лиц/сч" Then
Fg.TextMatrix(rw, 8) = Oplata(Inf("TabNum"), "All", "s")
End If

'******* S2
'If FG.TextMatrix(Rw, 15) = Inf("TablKod") Then

If Inf("f2") = "Оплата по текущей категории расчета" Then
Fg.TextMatrix(rw, 9) = Oplata(Inf("TabNum"), Kt, "-")
End If

If Inf("f2") = "Начислено по текущей категории расчета" Then
Fg.TextMatrix(rw, 9) = Oplata(Inf("TabNum"), Kt, "+")
End If

If Inf("f2") = "Субсидия по текущей категории расчета" Then
Fg.TextMatrix(rw, 9) = Oplata(Inf("TabNum"), Kt, "s")
End If

If Inf("f2") = "Оплачено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 9) = Oplata(Inf("TabNum"), "All", "-")
End If

If Inf("f2") = "Начислено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 9) = Oplata(Inf("TabNum"), "All", "+")
End If

If Inf("f2") = "Субсидии всего. по лиц/сч" Then
Fg.TextMatrix(rw, 9) = Oplata(Inf("TabNum"), "All", "s")
End If

'************ S3

If Inf("f3") = "Оплата по текущей категории расчета" Then
Fg.TextMatrix(rw, 10) = Oplata(Inf("TabNum"), Kt, "-")
End If

If Inf("f3") = "Начислено по текущей категории расчета" Then
Fg.TextMatrix(rw, 10) = Oplata(Inf("TabNum"), Kt, "+")
End If

If Inf("f3") = "Субсидия по текущей категории расчета" Then
Fg.TextMatrix(rw, 10) = Oplata(Inf("TabNum"), Kt, "s")
End If

If Inf("f3") = "Оплачено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 10) = Oplata(Inf("TabNum"), "All", "-")
End If

If Inf("f3") = "Начислено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 10) = Oplata(Inf("TabNum"), "All", "+")
End If

If Inf("f3") = "Субсидии всего. по лиц/сч" Then
Fg.TextMatrix(rw, 10) = Oplata(Inf("TabNum"), "All", "s")
End If

'************ S4

If Inf("f4") = "Оплата по текущей категории расчета" Then
Fg.TextMatrix(rw, 11) = Oplata(Inf("TabNum"), Kt, "-")
End If

If Inf("f4") = "Начислено по текущей категории расчета" Then
Fg.TextMatrix(rw, 11) = Oplata(Inf("TabNum"), Kt, "+")
End If

If Inf("f4") = "Субсидия по текущей категории расчета" Then
Fg.TextMatrix(rw, 11) = Oplata(Inf("TabNum"), Kt, "s")
End If

If Inf("f4") = "Оплачено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 11) = Oplata(Inf("TabNum"), "All", "-")
End If

If Inf("f4") = "Начислено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 11) = Oplata(Inf("TabNum"), "All", "+")
End If

If Inf("f4") = "Субсидии всего. по лиц/сч" Then
Fg.TextMatrix(rw, 11) = Oplata(Inf("TabNum"), "All", "s")
End If


'************ S5

If Inf("f5") = "Оплата по текущей категории расчета" Then
Fg.TextMatrix(rw, 12) = Oplata(Inf("TabNum"), Kt, "-")
End If

If Inf("f5") = "Начислено по текущей категории расчета" Then
Fg.TextMatrix(rw, 12) = Oplata(Inf("TabNum"), Kt, "+")
End If

If Inf("f5") = "Субсидия по текущей категории расчета" Then
Fg.TextMatrix(rw, 12) = Oplata(Inf("TabNum"), Kt, "s")
End If

If Inf("f5") = "Оплачено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 12) = Oplata(Inf("TabNum"), "All", "-")
End If

If Inf("f5") = "Начислено всего. по лиц/сч" Then
Fg.TextMatrix(rw, 12) = Oplata(Inf("TabNum"), "All", "+")
End If

If Inf("f5") = "Субсидии всего. по лиц/сч" Then
Fg.TextMatrix(rw, 12) = Oplata(Inf("TabNum"), "All", "s")
End If


'***************************
End If
Next
'End If
Inf.MoveNext
Loop

'************************************
For rw = 1 To Fg.Rows - 1
Form = Fg.TextMatrix(rw, 13)
KodTbl = Fg.TextMatrix(rw, 15)
On Error GoTo ErRas
Mconn.Execute ("UPDATE TablDoc SET TablDoc.Итог = " + Form + " WHERE (((TablDoc.TablKod)=" + KodTbl + "))")
ErRas:
If Err.Number <> 0 Then
MsgBox "Ошибка в формуле " + Form + " для " + Fg.TextMatrix(rw, 2) + "  " + Fg.TextMatrix(rw, 3) + "  " + Fg.TextMatrix(rw, 4) + " " + Fg.TextMatrix(rw, 5)
e = Fg.TextMatrix(rw, 15)
Err.Clear
End If
Next rw
Set Fg.DataSource = Cmb
For rw = 1 To Fg.Rows - 1
If Fg.TextMatrix(rw, 15) = e Then Fg.Cell(flexcpForeColor, rw, 1, rw, 15) = vbRed
'Else FG.Cell(flexcpForeColor, 1, 1, FG.Rows - 1, 15) = vbGreen
Next
Расчет
End If
End Sub

Private Sub Command3_Click()
PrintW.Show
        
     With PrintW.VP
        PrintW.VP.StartDoc
        .FontSize = 12
 .Paragraph = Label1 + "  " + Label2 + " " + ReestrTablDoc.VStbd.TextMatrix(ReestrTablDoc.VStbd.Row, 5) + vbNewLine + "Адрес: " + Label4 + vbNewLine + "Начисление: " + ReestrTablDoc.VStbd.TextMatrix(ReestrTablDoc.VStbd.Row, 4) + vbNewLine + Label5 + "  " + Label6 + " руб. " + Label7 + "  " + Label8
        .Paragraph = ""
        .FontSize = 8
        .RenderControl = Fg.hwnd
        .EndDoc
        
       End With
End Sub

Private Sub Command4_Click()

If Fg.Rows <= 1 Then Exit Sub
If MsgBox("Разнести данные в лиц.счета", vbYesNo) = vbYes Then
Jdite.Show
Jdite.Label1 = "Подождите, разношу начисления в лицевые счета"


Jdite.Label1.Refresh
Jdite.Label1 = Jdite.Label1 + vbNewLine + "*"
Jdite.Label1.Refresh

If Err.Number = 0 Then Else MsgBox Err.Description

Mconn.Execute ("INSERT INTO Adding ( TablDoc, KodKv, KodN, SummaB ) SELECT TablDoc.TablKod, TablDoc.TabNum, TablDoc.Kodn, TablDoc.Итог FROM TablDoc LEFT JOIN Adding ON TablDoc.TablKod = Adding.TablDoc WHERE (((TablDoc.Cod)=" + Fg.TextMatrix(1, 1) + ") AND ((Adding.TablDoc) Is Null))")


'Mconn.Execute ("INSERT INTO Adding ( prostoy ) SELECT TablDoc.S4 FROM TablDoc LEFT JOIN Adding ON TablDoc.TablKod = Adding.TablDoc WHERE (((TablDoc.Cod)=" + FG.TextMatrix(1, 1) + ") AND ((Adding.TablDoc) Is Null) AND ((TablDoc.F5)='Prostoy'))")

Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh


If Err.Number = 0 Then Else MsgBox Err.Description

Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh

'mconn.Execute ("UPDATE nachisleniy INNER JOIN (Adding INNER JOIN TablDoc ON Adding.TablDoc = TablDoc.TablKod) ON nachisleniy.Kod = Adding.KodN SET Adding.NameN = [nachisleniy]![Naim], Adding.KodKat = [nachisleniy]![КодKategor], Adding.NameKat = [nachisleniy]![Kategor], Adding.Formula = [nachisleniy]![Formula], Adding.Tip = [nachisleniy]![Tip], Adding.LgotaVid = [Nachisleniy]![Vid] WHERE (((TablDoc.Cod)=" + FG.TextMatrix(1, 1) + "))")
'mconn.Execute ("UPDATE (nachisleniy INNER JOIN (Adding INNER JOIN TablDoc ON Adding.TablDoc = TablDoc.TablKod) ON nachisleniy.Kod = Adding.KodN) INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.NameN = [nachisleniy]![Naim], Adding.KodKat = [nachisleniy]![КодKategor], Adding.NameKat = [nachisleniy]![Kategor], Adding.Formula = [nachisleniy]![Formula], Adding.FormulaB = [nachisleniy]![FormulaB], Adding.Tip = [nachisleniy]![Tip], Adding.LgotaVid = [Nachisleniy]![Vid], Adding.Propis = [MainOccupant]![NLODGERF], Adding.Projiv = [MainOccupant]![NLODGER], Adding.ProLift = [MainOccupant]![NLODLIFT], Adding.PolPl = [MainOccupant]![HABSPACE], Adding.ObPl = [MainOccupant]![COMSPACE], Adding.FLOOR = [MainOccupant]![FLOOR] WHERE (((TablDoc.Cod)=" + FG.TextMatrix(1, 1) + "))")
Mconn.Execute ("UPDATE (nachisleniy INNER JOIN (Adding INNER JOIN TablDoc ON Adding.TablDoc = TablDoc.TablKod) ON nachisleniy.Kod = Adding.KodN) INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.NameN = [nachisleniy]![Naim], Adding.KodKat = [nachisleniy]![КодKategor], Adding.NameKat = [nachisleniy]![Kategor], Adding.Formula = [nachisleniy]![Formula], Adding.FormulaB = [nachisleniy]![FormulaB], Adding.Tip = [nachisleniy]![Tip], Adding.LgotaVid = [Nachisleniy]![Vid], Adding.Propis = [MainOccupant]![NLODGERF], Adding.Projiv = [MainOccupant]![NLODGER], Adding.ProLift = [MainOccupant]![NLODLIFT], Adding.PolPl = [MainOccupant]![HABSPACE], Adding.ObPl = [MainOccupant]![COMSPACE], Adding.FLOOR = [MainOccupant]![FLOOR], Adding.SummaB = [TablDoc]![Итог] WHERE (((TablDoc.Cod)=" + Fg.TextMatrix(1, 1) + "))")



If Len(ReestrTablDoc.VStbd.TextMatrix(ReestrTablDoc.VStbd.Row, 5)) > 49 Then MsgBox "Слишком длинный коментарий"
Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh
Da = Replace(ReestrTablDoc.VStbd.TextMatrix(ReestrTablDoc.VStbd.Row, 2), ".", "/")
Mconn.Execute ("UPDATE Adding SET Adding.KodDoc = -1, Adding.Com = " + Chr(34) + ReestrTablDoc.VStbd.TextMatrix(ReestrTablDoc.VStbd.Row, 5) + Chr(34) + ", Adding.DataR = " + Da + " WHERE (((Adding.TablDoc)=" + Fg.TextMatrix(1, 1) + "))")
Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh
'Тариф
Mconn.Execute ("UPDATE Adding INNER JOIN Tarif ON (Adding.TipDomKod = Tarif.KodDOM) AND (Adding.TipKvKod = Tarif.KodKV) AND (Adding.KodKat = Tarif.KodKat) SET Adding.Tarif = [Tarif]![Value], Adding.TarifI = [Tarif]![TarifI], Adding.TarifD = [Tarif]![TarifD]")
Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh
'Соцминимум
Mconn.Execute ("UPDATE Adding INNER JOIN Socmin ON (Adding.Propis = Socmin.koli) AND (Adding.KodKat = Socmin.KodKategor) SET Adding.Socmin = [Socmin]![Value]")
Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh
' Теперь заполняем нулями пустые соцминимумы для Adding
Mconn.Execute ("UPDATE Adding SET Adding.Socmin = 0 WHERE (((Adding.Socmin) Is Null))")
Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh
Mconn.Execute ("UPDATE Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd SET tmp_lgota.Cocmin = [Adding]![Socmin]")
Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh
Mconn.Execute ("UPDATE Adding SET Adding.KodDoc = -1 WHERE (((Adding.KodDoc)=0) AND ((Adding.TablDoc)<>0))")


Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh

'Проставим сальдо на начало
If Err.Number = 0 Then Else MsgBox Err.Description
Mconn.Execute ("UPDATE (Saldo_Arh INNER JOIN Adding ON (Saldo_Arh.KodKat = Adding.KodKat) AND (Saldo_Arh.KodKV = Adding.KodKv)) INNER JOIN TablDoc ON Adding.TablDoc = TablDoc.TablKod SET Adding.SaldoN = [Saldo_Arh]![SK] WHERE (((TablDoc.Cod)=" + Fg.TextMatrix(1, 1) + "))")

Jdite.Label1 = Jdite.Label1 + "**"
Jdite.Label1.Refresh
If Err.Number = 0 Then Else MsgBox Err.Description

For rw = 1 To Fg.Rows - 1
Jdite.Label1.Caption = "-" + Str(rw) + "-" + vbNewLine + "Расчитываю сальдо л/счета >" + Fg.TextMatrix(rw, 2)
Jdite.Label1.Refresh
MainForm.RSaldoN Fg.TextMatrix(rw, 2)
'MainForm.RSaldoN FG.TextMatrix(Rw, 2)
MainForm.RSaldoK Str(Fg.TextMatrix(rw, 2))
'Pod.Label1.FontItalic = True
MainForm.КоличествоСальдо Str(Fg.TextMatrix(rw, 2))
'MainForm.RSaldoK Str(FG.TextMatrix(Rw, 5))
Next


Unload Jdite
MsgBox "Данные разнесены успешно"
End If

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Расчет
End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Fg.TextMatrix(Fg.Row, 13) = InputBox("", "Формула расчета", Fg.TextMatrix(Fg.Row, 13))
'FG.TextMatrix(FG.Row, 13) =

End Sub

Private Sub FG_Click()
'MsgBox Oplata(FG.TextMatrix(FG.Row, 2), Kt)
Dan = Fg.TextMatrix(Fg.Row, Fg.Col)
Naz = Fg.TextMatrix(0, Fg.Col)
Fg.ColComboList(13) = "..."
Fr = ""
Label9.Caption = "Нет формулы"
If Naz = "S1" Then Fr = "F1"
If Naz = "S2" Then Fr = "F2"
If Naz = "S3" Then Fr = "F3"
If Naz = "S4" Then Fr = "F4"
If Naz = "S5" Then Fr = "F5"

If Fr <> "" Then
Inf.MoveFirst
Do While Not Inf.EOF
If (Fg.TextMatrix(Fg.Row, 1) = Inf("Cod") And Inf("TabNum") = Fg.TextMatrix(Fg.Row, 2)) Then
If Inf(Fr) <> "" Then Label9.Caption = Inf(Fr) Else Label9.Caption = "Нет формулы"
End If
Inf.MoveNext
Loop
End If
End Sub

Private Sub FG_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Fg.Col <= 7 Then
MsgBox ("Эти данные изменять нельзя")
Exit Sub
End If
If Fg.Col < 13 Then
Me.Enabled = False
TablFormula.Show
End If
End Sub

Private Sub Form_Load()


'Set mconn = New ADODB.Connection
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'mconn.Open "data/Kvartplata.mdb"
'Set rs_Tit = New ADODB.Recordset

'Recordset для фильтров
Set Cmb = New ADODB.Recordset
Set Cmb.ActiveConnection = Mconn
Cmb.CursorType = adOpenDynamic
Cmb.LockType = adLockPessimistic

Set Inf = New ADODB.Recordset
Set Inf.ActiveConnection = Mconn
Inf.CursorType = adOpenDynamic
Inf.LockType = adLockBatchOptimistic




Label3.Caption = ReestrTablDoc.VStbd.TextMatrix(ReestrTablDoc.VStbd.Row, 4) + " " + ReestrTablDoc.VStbd.TextMatrix(ReestrTablDoc.VStbd.Row, 5)
L = ReestrTablDoc.VStbd.TextMatrix(ReestrTablDoc.VStbd.Row, 1)
Label2.Caption = L
Label4 = ReestrTablDoc.VStbd.TextMatrix(ReestrTablDoc.VStbd.Row, 10)

Inf.Open ("SELECT TablDoc.TablKod, TablDoc.Cod, TablDoc.TabNum,TablDoc.S1, TablDoc.F1,TablDoc.S2, TablDoc.F2, TablDoc.S3,TablDoc.F3,TablDoc.S4, TablDoc.F4, TablDoc.S5,TablDoc.F5 From TablDoc WHERE (((TablDoc.Cod)=" + L + "))")

Cmb.Open "SELECT TablDoc.Cod, TablDoc.TabNum, TablDoc.Fam, TablDoc.Im, TablDoc.Ot, TablDoc.KvNum, TablDoc.Kodn, TablDoc.S1, TablDoc.S2, TablDoc.S3, TablDoc.S4, TablDoc.S5, TablDoc.Формула, TablDoc.Итог, TablDoc.TablKod From TablDoc WHERE (((TablDoc.Cod)=" + L + "))"
Set Fg.DataSource = Cmb

Set Nac = New ADODB.Recordset
Set Nac.ActiveConnection = Mconn
Nac.Open ("SELECT nachisleniy.Kod, nachisleniy.КодKategor FROM nachisleniy")

If Fg.Rows > 1 Then
Nac.MoveFirst
Do While Not Nac.EOF
If Nac("Kod") = Fg.TextMatrix(1, 7) Then Kt = Nac("КодKategor")
Nac.MoveNext
Loop
End If
Nac.Close
Set Nac = Nothing
Расчет
End Sub
Private Sub Расчет()
Kol = 0
Sum = 0
For rw = 1 To Fg.Rows - 1
Kol = Kol + 1
Sum = Sum + Fg.TextMatrix(rw, 14)
Next
Label6 = Sum
Label8 = Kol
End Sub

