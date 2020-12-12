VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ZPlan 
   Caption         =   "Экономическое обоснование тарифов"
   ClientHeight    =   7488
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10632
   LinkTopic       =   "Form4"
   ScaleHeight     =   7488
   ScaleWidth      =   10632
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Тариф"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   0
      Width           =   3612
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      Left            =   8160
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   360
      Width           =   1692
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   6600
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   360
      Width           =   972
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10632
      _ExtentX        =   18754
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
   Begin VSFlex8Ctl.VSFlexGrid Fg 
      Height          =   6492
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10452
      _cx             =   18436
      _cy             =   11451
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Plan.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
      OwnerDraw       =   2
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   3
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4440
      Top             =   2232
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Plan.frx":0195
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Plan.frx":02A7
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Plan.frx":03B9
            Key             =   "Save"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   252
      Left            =   7560
      TabIndex        =   8
      Top             =   7200
      Width           =   1572
   End
   Begin VB.Label Label3 
      Caption         =   "Месяц"
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
      Left            =   7560
      TabIndex        =   6
      Top             =   360
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "Год"
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
      Left            =   6240
      TabIndex        =   5
      Top             =   360
      Width           =   372
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
      ForeColor       =   &H80000001&
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   4092
   End
End
Attribute VB_Name = "ZPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Dom As Integer ' код дома для фильтра выбранного дома
Public God As String
Public Mes As String
Public t As String

Dim RsPlan As ADODB.Recordset ' Рекордсет для Z_Doma_Plan
Dim RsDomTar As ADODB.Recordset 'Рекордсет для получения данных о тарифе для дома
Dim RsKat As ADODB.Recordset 'Рекордсет для получения данных о категории расчета для анализа

Dim Kat As String
Dim Tip As String








Private Sub Combo1_Click()
'Открываем выбранный год

Me.God = Me.Combo1.Text
Set Me.Fg.DataSource = Nothing
Me.Fg.DataRefresh
RsPlan.Close
RsPlan.Open ("SELECT Z_Doma_Plan.DomKod, Z_Doma_Plan.DomTip, Z_Doma_Plan.Tarif, Z_Doma_Plan.vid1, Z_Doma_Plan.vid2, Z_Doma_Plan.NameR, Z_Doma_Plan.Summa, Z_Doma_Plan.percent, Z_Doma_Plan.Код, Z_Doma_Plan.god, Z_Doma_Plan.Mes From Z_Doma_Plan WHERE (((Z_Doma_Plan.DomKod)=" + Str(Dom) + ") AND ((Z_Doma_Plan.god)='" + Trim(God) + "') AND ((Z_Doma_Plan.Mes)='" + Mes + "'))"), Mconn, adOpenStatic, adLockBatchOptimistic
Set Me.Fg.DataSource = RsPlan
Me.Строимгрид
End Sub

Private Sub Combo2_Click()
'Открываем выбранный месяц

Me.Mes = Me.Combo2.Text
Set Me.Fg.DataSource = Nothing
Me.Fg.DataRefresh
RsPlan.Close
RsPlan.Open ("SELECT Z_Doma_Plan.DomKod, Z_Doma_Plan.DomTip, Z_Doma_Plan.Tarif, Z_Doma_Plan.vid1, Z_Doma_Plan.vid2, Z_Doma_Plan.NameR, Z_Doma_Plan.Summa, Z_Doma_Plan.percent, Z_Doma_Plan.Код, Z_Doma_Plan.god, Z_Doma_Plan.Mes From Z_Doma_Plan WHERE (((Z_Doma_Plan.DomKod)=" + Str(Dom) + ") AND ((Z_Doma_Plan.god)='" + Trim(God) + "') AND ((Z_Doma_Plan.Mes)='" + Mes + "'))"), Mconn, adOpenStatic, adLockBatchOptimistic
Set Me.Fg.DataSource = RsPlan
Me.Строимгрид
End Sub

Private Sub Command1_Click()
F4.Text1.Text = t
F4.Show 1, Me

End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)

If t <> 0 Then Me.Fg.TextMatrix(Fg.Row, 8) = (Me.Fg.TextMatrix(Fg.Row, 7) / t) * 100
'Me.Строимгрид
End Sub

Private Sub Fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Me.Строимгрид
End Sub




Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Fg.Col = 7 Or Fg.Col = 8 Then Fg.Editable = flexEDKbdMouse Else Fg.Editable = flexEDNone

End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

'Yes
If MsgBox("Изменить тариф для всего периода " + Me.Combo2.Text + " " + Me.Combo1.Text + " ? <Да>-Заменить для всех, <Нет>-Только для текущей строки, <Отмена>-Не менять", vbYesNoCancel) = vbYes Then
Me.t = InputBox("Внимание после нажатия <Ок> тариф будет изменен по всем записям за период " + Me.Combo2.Text + " " + Me.Combo1.Text, "Введите значение тарифа", Me.Fg.TextMatrix(Fg.Row, 3))
Me.t = Replace(Me.t, ".", ",")

For i = 1 To Fg.Rows - 1
Me.Fg.TextMatrix(i, 3) = Me.t
Next i
Else
Me.Fg.TextMatrix(Fg.Row, 3) = InputBox("Внимание после нажатия <Ок> тариф будет изменен только для текущей строки", "Введите значение тарифа", Me.Fg.TextMatrix(Fg.Row, 3))
End If

'No
'If MsgBox("Изменить тариф для всего периода " + Me.Combo2.Text + " " + Me.Combo1.Text + " ? <Да>-Заменить для всех, <Нет>-Только для текущей строки, <Отмена>-Не менять", vbYesNoCancel) = vbNo Then
'Me.Fg.TextMatrix(Fg.Row, 3) = InputBox("Внимание после нажатия <Ок> тариф будет изменен только для текущей строки", "Введите значение тарифа", Me.Fg.TextMatrix(Fg.Row, 3))
'End If

'Cancel
'If MsgBox("Изменить тариф для всего периода " + Me.Combo2.Text + " " + Me.Combo1.Text + " ? <Да>-Заменить для всех, <Нет>-Только для текущей строки, <Отмена>-Не менять", vbYesNoCancel) = vbCancel Then
'Exit Sub
'End If

End Sub

Private Sub FG_Click()
If Fg.Col = 7 Or Fg.Col = 8 Or Fg.Col = 3 Then
'Or Fg.Col = 12
Fg.Editable = flexEDKbdMouse
Else
Fg.Editable = flexEDNone
End If
End Sub

Private Sub FG_DblClick()
If Fg.Col = 7 Or Fg.Col = 8 Then Fg.Editable = flexEDKbdMouse Else Fg.Editable = flexEDNone

'MsgBox (Fg.ColComboList(Fg.Col))

End Sub

Private Sub FG_EnterCell()
If Fg.Col = 7 Or Fg.Col = 8 Then Fg.Editable = flexEDKbdMouse Else Fg.Editable = flexEDNone

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Me.Зпись
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.KEY
        Case "New"
            'ToDo: Add 'New' button code.
           ' MsgBox "Add 'New' button code."
            Z_New_Plan.Show 1, Me
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            'MsgBox "Add 'Delete' button code."
                    If Me.Fg.Row <> 0 Then
            If MsgBox("Удалить строку <<" + Me.Fg.TextMatrix(Me.Fg.Row, 6) + ">>?", vbYesNo) = vbYes Then
           Mconn.Execute ("DELETE Z_Doma_Plan.Код From Z_Doma_Plan WHERE (((Z_Doma_Plan.Код)=" + Me.Fg.TextMatrix(Me.Fg.Row, 9) + "))")
           Me.Fg.Cell(flexcpBackColor, Me.Fg.Row, 1, Me.Fg.Row, Fg.Cols - 1) = vbRed
           
           
            Else
            
            End If
           Else
            MsgBox ("Нет записей")
            End If
        Case "Save"
            'ToDo: Add 'Save' button code.
            'MsgBox "Add 'Save' button code."
         Unload Me
    End Select
End Sub


Private Sub Form_Load()

Me.Mes = MonthName(Month(MainForm.PeriodR), False)
Me.God = Str(Year(MainForm.PeriodR))
Me.Combo1.Text = Me.God
Me.Combo2.Text = Me.Mes

Me.Combo2.AddItem "Январь"
Me.Combo2.AddItem "Февраль"
Me.Combo2.AddItem "Март"
Me.Combo2.AddItem "Апрель"
Me.Combo2.AddItem "Май"
Me.Combo2.AddItem "Июнь"
Me.Combo2.AddItem "Июль"
Me.Combo2.AddItem "Август"
Me.Combo2.AddItem "Сентябрь"
Me.Combo2.AddItem "Октябрь"
Me.Combo2.AddItem "Ноябрь"
Me.Combo2.AddItem "Декабрь"

Me.Combo1.AddItem Me.God + 1

For i = 0 To 10
Me.Combo1.AddItem Me.God - i
Next i

' Получаем данные о категории расчета для анализа затрат
Set RsKat = New ADODB.Recordset
RsKat.Open ("SELECT Settings.zatrKat FROM Settings"), Mconn
RsKat.MoveFirst
Kat = RsKat("zatrKat")
RsKat.Close

'Получаем тариф
Set RsDomTar = New ADODB.Recordset
RsDomTar.Open ("SELECT Tarif.KodKat, KLS_PODR.КОД, Max(Tarif.Value) AS [Max-Value], KLS_PODR.NAIM_KLS, Tarif.KodDOM FROM Tarif INNER JOIN KLS_PODR ON Tarif.KodDOM = KLS_PODR.Tip GROUP BY Tarif.KodKat, KLS_PODR.КОД, KLS_PODR.NAIM_KLS, Tarif.KodDOM HAVING (((Tarif.KodKat)=" + Kat + ") AND ((KLS_PODR.КОД)=" + Str(Dom) + "));"), Mconn
t = RsDomTar("Max-Value")
Tip = RsDomTar("KodDOM")
Me.Label1.Caption = "Адрес " + RsDomTar("NAIM_KLS") + " Тариф " + Str(t)
RsDomTar.Close


Me.Command1.Caption = "Тариф=" + t + "р. Изменить?"

Set RsPlan = New ADODB.Recordset
'RsPlan.Open ("SELECT Z_Doma_Plan.DomKod, Z_Doma_Plan.DomTip, Z_Doma_Plan.Tarif, Z_Doma_Plan.vid1, Z_Doma_Plan.vid2, Z_Doma_Plan.NameR, Z_Doma_Plan.Summa, Z_Doma_Plan.percent, Z_Doma_Plan.КОД, Z_Doma_Plan.god, Z_Doma_Plan.mes From Z_Doma_Plan WHERE (((Z_Doma_Plan.DomKod)=" + Str(Dom) + "))"), Mconn, adOpenStatic, adLockBatchOptimistic
RsPlan.Open ("SELECT Z_Doma_Plan.DomKod, Z_Doma_Plan.DomTip, Z_Doma_Plan.Tarif, Z_Doma_Plan.vid1, Z_Doma_Plan.vid2, Z_Doma_Plan.NameR, Z_Doma_Plan.Summa, Z_Doma_Plan.percent, Z_Doma_Plan.Код, Z_Doma_Plan.god, Z_Doma_Plan.Mes From Z_Doma_Plan WHERE (((Z_Doma_Plan.DomKod)=" + Str(Dom) + ") AND ((Z_Doma_Plan.god)='" + Trim(God) + "') AND ((Z_Doma_Plan.Mes)='" + Mes + "'))"), Mconn, adOpenStatic, adLockBatchOptimistic





Set Fg.DataSource = RsPlan

'Fg.Cols = Fg.Cols + 1
'Fg.TextMatrix(0, 12) = "...."
'Fg.ColComboList(12) = "..."
'Fg.AddItem "+"
' FG  строчка, колонка
'Fg.TextMatrix(1, 1) = Str(Dom)



End Sub
Public Sub Строимгрид()

'Изменяет ширину столбцов или строк, чтобы соответствовать высоте содержимого ячейки.
Me.Fg.AutoSize 0, Me.Fg.Cols - 1, False, 0
Me.Fg.WordWrap = True
Fg.Editable = flexEDKbdMouse

'RsPlan.Open ("SELECT Z_Doma_Plan.DomKod, Z_Doma_Plan.DomTip, Z_Doma_Plan.Tarif, Z_Doma_Plan.vid1, Z_Doma_Plan.vid2, Z_Doma_Plan.NameR, Z_Doma_Plan.Summa, Z_Doma_Plan.percent, Z_Doma_Plan.Код, Z_Doma_Plan.god, Z_Doma_Plan.Mes From Z_Doma_Plan WHERE (((Z_Doma_Plan.DomKod)=" + Str(Dom) + ") AND ((Z_Doma_Plan.god)='" + Trim(God) + "') AND ((Z_Doma_Plan.Mes)='" + Mes + "'))"), Mconn, adOpenStatic, adLockBatchOptimistic


End Sub

Public Sub Зпись()

If MsgBox("Сохранить данные ?", vbYesNo) = vbYes Then

'SELECT Z_Doma_Plan.DomKod, Z_Doma_Plan.DomTip, Z_Doma_Plan.Tarif, Z_Doma_Plan.vid1, Z_Doma_Plan.vid2, Z_Doma_Plan.NameR, Z_Doma_Plan.Summa, Z_Doma_Plan.percent From Z_Doma_Plan WHERE (((Z_Doma_Plan.DomKod)=" + Str(Dom) + "))"

For i = 1 To Me.Fg.Rows - 1

If Me.Fg.TextMatrix(i, 0) = "+" Then
RsPlan.AddNew
RsPlan("DomKod") = Dom
RsPlan("DomTip") = Tip
RsPlan("Tarif") = t
RsPlan("vid1") = Me.Fg.TextMatrix(i, 4)
RsPlan("vid2") = Me.Fg.TextMatrix(i, 5)
RsPlan("NameR") = Me.Fg.TextMatrix(i, 6)

RsPlan("Summa") = Me.Fg.TextMatrix(i, 7)
RsPlan("percent") = Me.Fg.TextMatrix(i, 8)


End If

RsPlan.UpdateBatch
Next

'RsPlan.UpdateBatch
RsPlan.Close

End If ' End if msgbox
End Sub
