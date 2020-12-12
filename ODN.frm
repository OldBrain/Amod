VERSION 5.00
Begin VB.Form ODN 
   Caption         =   "Ввод исходных данных расчета"
   ClientHeight    =   5268
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6936
   Icon            =   "ODN.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   5268
   ScaleWidth      =   6936
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   252
      Left            =   4080
      TabIndex        =   22
      Top             =   3240
      Width           =   252
   End
   Begin VB.TextBox Text7 
      Height          =   288
      Left            =   1440
      TabIndex        =   21
      Text            =   "0"
      Top             =   1800
      Width           =   852
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Расчет от прописанных"
      Height          =   372
      Left            =   2760
      TabIndex        =   19
      Top             =   3960
      Width           =   2292
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Расчет от площади"
      Height          =   252
      Left            =   240
      TabIndex        =   18
      Top             =   3960
      Value           =   -1  'True
      Width           =   2052
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1920
      TabIndex        =   17
      Text            =   "S1"
      Top             =   3240
      Width           =   1932
   End
   Begin VB.TextBox Text5 
      Height          =   372
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   6732
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   5280
      TabIndex        =   13
      Text            =   "0"
      Top             =   600
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   6732
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      Left            =   4320
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   1800
      Width           =   2532
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   5160
      TabIndex        =   8
      Text            =   "0"
      Top             =   1200
      Width           =   1452
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1800
      TabIndex        =   6
      Text            =   "0"
      Top             =   1200
      Width           =   852
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2400
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Продолжить"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   6732
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   1560
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   4692
   End
   Begin VB.Label Label10 
      Caption         =   "Тариф на площадь"
      Height          =   252
      Left            =   4560
      TabIndex        =   23
      Top             =   3240
      Width           =   2172
   End
   Begin VB.Label Label9 
      Caption         =   "Норматив"
      Height          =   252
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label8 
      Caption         =   "Формула расчета"
      Height          =   372
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   1572
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Коментарий к документу"
      Height          =   252
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   6612
   End
   Begin VB.Label Label6 
      Caption         =   "Прописано"
      Height          =   372
      Left            =   4080
      TabIndex        =   12
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "Выбор начисления"
      Height          =   372
      Left            =   2400
      TabIndex        =   10
      Top             =   1800
      Width           =   1932
   End
   Begin VB.Label Label4 
      Caption         =   "Введите п лощадь мест общего пользования"
      Height          =   612
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   1812
   End
   Begin VB.Label Label3 
      Caption         =   "Тариф"
      Height          =   372
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Общая площадь дома"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "Выбор дома"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1212
   End
End
Attribute VB_Name = "ODN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ComboStreet As ADODB.Recordset 'Рекордсет для адреса
Dim ComboN As ADODB.Recordset 'Рекордсет для начисления
Dim SumPl As ADODB.Recordset 'Рекордсет для подчсета площади
Dim rs_kat As ADODB.Recordset 'Реестр табличных документов
Dim Tbl_str As ADODB.Recordset 'Строки табличных документов

Dim CodDom As Integer ' Код выбранного адреса
Dim CodN As Integer ' Код выбранного начисления

Dim ObUslug As Currency
Dim Tar As Double ' Тариф ОДН
Dim KodN As Integer ' Код начисления
Dim Cod As Integer ' Порядковай номер документа в реестре
Dim Formula As String ' Формула расчета
Dim n As Integer ' Вспомогательный счетчик
Dim CodD As Integer 'Код дома
'S1
Dim Obpl As Double ' Общая площадь дома S1
Dim OProp As Integer ' Общее количество прописанных S1

'S2
Dim OU As Double ' Объем услуг Площадь мест общего пользования S2

Dim OSum As Double ' Общая сумма по услугам OU*Tar
'S4
Dim Normativ As Double ' Норматив S4
'S5
Dim TarifODN As Double ' Тариф ОДН S5



Private Sub Check1_Click()
If Check1.Value = 1 Then
Text3.Enabled = False
Text7.Enabled = False
Text3.Text = 0
Text7.Text = 0
Me.Text6.Text = "Round((S5*S3),2)"
Formula = Trim(Me.Text6.Text)
Else
Text3.Enabled = True
Text7.Enabled = True
Me.Text6.Text = "Round((S4*S2*S3/S1),2)*S5"
Formula = Trim(Me.Text6.Text)
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 0 Then KeyAscii = 0
End Sub

Private Sub Combo1_LostFocus()

'If Trim(Combo1.Text) = "Выбери адрес" Or Trim(Combo1.Text) = "" Then Combo1.SetFocus
'CodDom = Val(Combo1.Text)
'MsgBox (CodDom)
End Sub



Private Sub Combo1_Validate(Cancel As Boolean)



CodDom = Val(Combo1.Text)
CodDom = Val(Left(Combo1.Text, InStr(1, Combo1.Text, " ")))

'Подсчет общей площади дома
If CodDom <> 0 Then
Set SumPl = New ADODB.Recordset

SumPl.Open ("SELECT MainOccupant.Dom, Sum(MainOccupant.COMSPACE) AS [Sum-COMSPACE], Sum(MainOccupant.HABSPACE) AS [Sum-HABSPACE], Sum(MainOccupant.NLODGER) AS [Sum-NLODGER], Sum(MainOccupant.NLODGERF) AS [Sum-NLODGERF] From MainOccupant GROUP BY MainOccupant.Dom HAVING (((MainOccupant.Dom)=" + Str(CodDom) + "))"), Mconn
Text1.Text = SumPl("Sum-COMSPACE")

Text4.Text = SumPl("Sum-NLODGER")

Me.Text1.Text = Replace(Me.Text1.Text, ",", ".")
Obpl = Val(Text1.Text)

OProp = Val(Text4.Text)


End If




End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 0 Then KeyAscii = 0
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)

CodN = Val(Combo2.Text)

End Sub

Private Sub Command1_Click()

'MsgBox (Formula)

Set rs_kat = New ADODB.Recordset

If Trim(Combo2.Text) = "Выбор начисления" Or Trim(Combo2.Text) = "" Then
MsgBox ("Вы не выбрали начисление!")
Combo2.SetFocus

If Trim(Combo1.Text) = "Выбери адрес" Or Trim(Combo1.Text) = "" Then
MsgBox ("Вы не выбрали адрес!")
Combo1.SetFocus

 

End If
End If


If Trim(Combo2.Text) <> "Выбор начисления" And Trim(Combo1.Text) <> "Выбери адрес" Then
If Trim(Combo2.Text) <> "" And Trim(Combo1.Text) <> "" Then

rs_kat.CursorType = adOpenKeyset
rs_kat.LockType = adLockOptimistic
'Определяем номер локумента следующий за максимальным
rs_kat.Open ("SELECT ReestrTablDoc.Cod, ReestrTablDoc.Data, ReestrTablDoc.NachCod, ReestrTablDoc.Nach, ReestrTablDoc.Coment, ReestrTablDoc.Summa, ReestrTablDoc.Status, ReestrTablDoc.Tip, ReestrTablDoc.KodDom, ReestrTablDoc.Adres FROM ReestrTablDoc"), Mconn




' Добавляем запись в реестр документов

rs_kat.AddNew
'Rs_kat("Cod") = n + 1
If Trim(Me.Text5.Text) = "" Then Me.Text5.Text = "Коментарий отсутствует"
rs_kat("Coment") = Me.Text5.Text
rs_kat("Data") = MainForm.DR
rs_kat("NachCod") = Val(Me.Combo2.Text)

KodN = Val(Me.Combo2.Text) ' Код начисления

rs_kat("Nach") = Me.Combo2.Text
rs_kat("Summa") = 0
rs_kat("Status") = 0
rs_kat("Tip") = "ODN"
rs_kat("KodDom") = Val(Me.Combo1.Text)
rs_kat("Adres") = Me.Combo1.Text

'MsgBox (Left(Me.Combo1.Text, 3))




CodD = Val(Left(Me.Combo1.Text, 3)) ' Код дома
rs_kat.UpdateBatch



'Определяем номер локумента
If Not rs_kat.EOF Then rs_kat.MoveFirst
Do While Not rs_kat.EOF
If rs_kat("Cod").Value = "" Then
rs_kat.Delete
rs_kat.MoveFirst
End If
N1 = rs_kat("Cod").Value
If N1 > n Then n = N1
rs_kat.MoveNext
Loop


rs_kat.Close

Cod = n ' Порядковай номер документа в реестре




' Добавляем строки табличных документов

Set Tbl_str = New ADODB.Recordset
 
 

If Me.Option1 Then 'Если расчет от площади
 
'Tbl_str.Open ("INSERT INTO TablDoc ( TabNum, Fam, Im, Ot, KvNum, Kodn, Cod, Формула, S1, S2, S3, S4, S5 ) SELECT MainOccupant.Numer, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, " + Str(KodN) + " AS Выражение2, " + Str(Cod) + " AS Выражение1, '" + Formula + "' AS Выражение3, " + Str(Obpl) + " AS Выражение4, " + Str(OProp) + " AS Выражение5, MainOccupant.COMSPACE, MainOccupant.NLODGER, " + Str(OU * Tar) + " AS Выражение8 FROM KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom WHERE (((MainOccupant.Dom)=" + Str(CodD) + "))"), Mconn

Tbl_str.Open ("INSERT INTO TablDoc ( TabNum, Fam, Im, Ot, KvNum, Kodn, Cod, Формула, S1, S2, S3, S4, S5 ) SELECT MainOccupant.Numer, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, " + Str(KodN) + " AS Выражение2, " + Str(Cod) + " AS Выражение1, '" + Formula + "' AS Выражение3, " + Str(Obpl) + " AS Выражение4, " + Str(ObUslug) + " AS Выражение5, MainOccupant.COMSPACE, " + Str(Normativ) + ", " + Str(TarifODN) + " AS Выражение8 FROM KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom WHERE (((MainOccupant.Dom)=" + Str(CodD) + "))"), Mconn

End If


If Me.Option2 Then 'Если расчет от прописанных

Tbl_str.Open ("INSERT INTO TablDoc ( TabNum, Fam, Im, Ot, KvNum, Kodn, Cod, Формула, S1, S2, S3, S4, S5 ) SELECT MainOccupant.Numer, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.kv_num, " + Str(KodN) + " AS Выражение2, " + Str(Cod) + " AS Выражение1, '" + Formula + "' AS Выражение3, " + Str(OProp) + " AS Выражение4, " + Str(ObUslug) + " AS Выражение5, MainOccupant.NLODGER, " + Str(Normativ) + ", " + Str(TarifODN) + " AS Выражение8 FROM KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom WHERE (((MainOccupant.Dom)=" + Str(CodD) + "))"), Mconn

End If


ReestrTablDoc.Show

Unload Me

End If
End If


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()


If Me.Option1 Then
Me.Text6.Text = "Round((S4*S2*S3/S1),2)*S5"
Formula = Trim(Me.Text6.Text)
End If

If Me.Option2 Then
Me.Text6.Text = "Round((S4*S2*S3/S1),2)*S5"
Formula = Trim(Me.Text6.Text)
End If






'Назначаем комбобокс для выбора начисления

Combo2.Text = "Выбор начисления"

Set ComboN = New ADODB.Recordset
ComboN.Open ("SELECT nachisleniy.Kod, nachisleniy.Naim, nachisleniy.Tip From Nachisleniy WHERE (((nachisleniy.Tip)='+'))"), Mconn


Cl = ""
ComboN.MoveFirst

Do While Not ComboN.EOF
If Trim(Cl) <> "" Then Combo2.AddItem Cl
Cl = CStr(ComboN("Kod")) & "  " & ComboN("Naim")
ComboN.MoveNext
Loop




'Назначаем комбобокс для выбора адреса
Combo1.Text = "Выбери адрес"


Set ComboStreet = New ADODB.Recordset
ComboStreet.Open ("SELECT KLS_PODR.КОД, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip FROM KLS_PODR"), Mconn


Cl = ""
ComboStreet.MoveFirst

Do While Not ComboStreet.EOF
If Trim(Cl) <> "" Then Combo1.AddItem Cl
Cl = CStr(ComboStreet("Код")) & "  " & ComboStreet("Naim_kls") & " дом № " & ComboStreet("Num")
ComboStreet.MoveNext
Loop

'Event combo1()



End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Visible = True
ReestrTablDoc.Show
End Sub


Private Sub Label8_Click()
Me.Text6.Enabled = True
End Sub

Private Sub Option1_Validate(Cancel As Boolean)
If Me.Option1 Then
Me.Text6.Text = "Round((S4*S2*S3/S1),2)*S5"
Formula = Trim(Me.Text6.Text)
End If

If Me.Option2 Then
Me.Text6.Text = "Round(S4*S2*S3/S1,2)*S5"
Formula = Trim(Me.Text6.Text)
End If
Me.Text6.Refresh
End Sub

Private Sub Option2_Validate(Cancel As Boolean)
If Me.Option1 Then
Me.Text6.Text = "Round((S4*S2*S3/S1),2)*S5"
Formula = Trim(Me.Text6.Text)
End If

If Me.Option2 Then
Me.Text6.Text = "Round((S4*S2*S3/S1),2)*S5"
Formula = Trim(Me.Text6.Text)
End If
Me.Text6.Refresh
End Sub

Private Sub Text1_Validate(Cancel As Boolean)

Me.Text1.Text = Replace(Me.Text1.Text, ",", ".")

Obpl = Val(Text1.Text)

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Text2.Text = Replace(Trim(Text2.Text), ",", ".")
TarifODN = Val(Text2.Text)
Tar = Val(Text2.Text)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
Text3.Text = Replace(Trim(Text3.Text), ",", ".")
ObUslug = Val(Text3.Text)

OU = Val(Text3.Text)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
OProp = Val(Text4.Text)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
Formula = Trim(Me.Text6.Text)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
Text7.Text = Replace(Trim(Text7.Text), ",", ".")
Normativ = Val(Text7.Text)
Tar = Val(Text7.Text)
End Sub
