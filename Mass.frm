VERSION 5.00
Begin VB.Form Mass 
   BackColor       =   &H8000000A&
   Caption         =   "Обработка данных"
   ClientHeight    =   2268
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4980
   DrawMode        =   1  'Blackness
   FillColor       =   &H00808000&
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2265
   ScaleMode       =   0  'User
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
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
      Left            =   960
      TabIndex        =   3
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Далее"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Mass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim mconn, DBFConn As ADODB.Connection
Public RsPer, rsClon As ADODB.Recordset


Private Sub Command1_Click()

Set RsPer = New ADODB.Recordset
Set RsPer.ActiveConnection = Mconn
'Set rsClon = New ADODB.Recordset
'Set rsClon.ActiveConnection = mconn

RsPer.CursorType = adOpenForwardOnly
RsPer.LockType = adLockBatchOptimistic
'rsClon.CursorType = adOpenForwardOnly
'rsClon.LockType = adLockBatchOptimistic

Me.Enabled = False

'If MsgBox("Да - С удалением прежних данных об ОТВЕТСТВЕННЫХ КВАРТИРОСЪЕМЩИКАХ, Нет-добавить к уже имеющемся", vbYesNo) = vbYes Then
Mconn.Execute ("DELETE MainOccupant.* FROM MainOccupant")
Mconn.Execute ("delete lgota.* FROM lgota")
'End If
'If MsgBox("Да - С удалением справочника ДОМОВ, Нет-добавить к уже имеющемуся", vbYesNo) = vbYes Then
Mconn.Execute ("DELETE KLS_PODR.* FROM KLS_PODR")
'If MsgBox("Да - С удалением справочника ЛЬГОТ, Нет-добавить к уже имеющемуся", vbYesNo) = vbYes Then
Mconn.Execute ("DELETE KLS_PRIV.* FROM KLS_PRIV")
'If MsgBox("Да - С удалением справочника НАЧИСЛЕНИЙ, Нет-добавить к уже имеющемуся", vbYesNo) = vbYes Then
Mconn.Execute ("DELETE Nachisleniy.* FROM Nachisleniy")



Mconn.Execute ("INSERT INTO KLS_PODR ( КОД, NAIM_KLS, Tip, Num ) SELECT KLS_podr1.N_KLS, KLS_podr1.NAIM_KLS, KLS_podr1.KOD_KLS, 0 AS Выражение1 FROM KLS_podr1")
Mconn.Execute ("INSERT INTO KLS_PRIV ( N_KLS, NAME_KLS, LPKV, LPOTOPL, LPCOMM, LPTEH, LPMUSOR, USEKV, USEOTOPL, USECOMM, USETEH, USEMUSOR ) SELECT KLS_priv1.N_KLS, KLS_priv1.NAIM_KLS, KLS_priv1.LPROCSPACE, KLS_priv1.LPROCOTP, KLS_priv1.LPROCCOM, KLS_priv1.LPROCTECH, KLS_priv1.LPROCMUSOR, KLS_priv1.USESPACE, KLS_priv1.USEOTP, KLS_priv1.USECOM, KLS_priv1.USETECH, KLS_priv1.USEMUSOR FROM KLS_priv1")
'тип дома
Mconn.Execute ("DELETE tipdom.* FROM tipdom")
Mconn.Execute ("INSERT INTO TipDom ( Код, Name_Dom ) SELECT kls_home.N_KLS, kls_home.NAIM_KLS FROM kls_home")
'тип квартиры
Mconn.Execute ("DELETE tipkv.* FROM tipkv")
Mconn.Execute ("INSERT INTO TipKv ( Код, Name_Kv ) SELECT KLS_HAB.N_KLS, KLS_HAB.NAIM_KLS FROM KLS_HAB")

'счета затрат
Mconn.Execute ("DELETE Schet.* FROM Schet")
Mconn.Execute ("INSERT INTO Schet ( Schet, Schet_Name ) SELECT kls_zak.N_KLS, kls_zak.NAIM_KLS FROM kls_zak")

Mconn.Execute ("UPDATE All_k SET All_k.NAPARTMENT = '0' WHERE (((All_k.NAPARTMENT) Is Null))")
'mconn.Execute ("INSERT INTO MainOccupant ( Dom, OldNum, FAM, IM, OT, NLODGER, NROOM, COMSPACE, HABSPACE, PRIVILEGE, KV, BIRTHDAY, NORDER, NLODGERF, kv_num ) SELECT Val([DOM]) AS Выражение1, All_k.TABN, All_k.FAM, All_k.IM, All_k.OT, All_k.NLODGER, All_k.NROOM, All_k.COMSPACE, All_k.HABSPACE, All_k.PRIVILEGE, All_k.HABITATE, All_k.BIRTHDAY, All_k.NORDER, All_k.NLODGERF, All_k.NAPARTMENT From All_k WHERE (((All_k.NUMHABIT)=1))")
 Mconn.Execute ("INSERT INTO MainOccupant ( Dom, OldNum, FAM, IM, OT, NLODGER, NROOM, COMSPACE, HABSPACE, PRIVILEGE, KV, BIRTHDAY, NORDER, NLODGERF, kv_num ) SELECT Val([DOM]) AS Выражение1, All_k.TABN, All_k.FAM, All_k.IM, All_k.OT, All_k.NLODGER, All_k.NROOM, All_k.COMSPACE, All_k.HABSPACE, All_k.PRIVILEGE, All_k.HABITATE, All_k.BIRTHDAY, All_k.NORDER, All_k.NLODGERF, val(All_k.NAPARTMENT) From All_k WHERE (((All_k.NUMHABIT)=1) and All_k.dom is not null)")
 'WHERE (((All_k.NUMHABIT) = 1))")
'mconn.Execute ("Доб_MainOccupant")
'RsPer.Open (Доб_MainOccupant)
Mconn.Execute ("INSERT INTO Nachisleniy ( Kod, Naim, Formula, Tip, Lig, Kategor, КодKategor ) SELECT KLS_VO.VID, KLS_VO.NAIMVID, KLS_VO.FORM, " + Chr(34) + "/" + Chr(34) + " AS Выражение1, No AS Выражение2, " + Chr(34) + "Не определена" + Chr(34) + " AS Выражение3, 0 AS Выражение4 FROM KLS_VO")
'Доб.тип дома
Mconn.Execute ("UPDATE KLS_PODR INNER JOIN MainOccupant ON KLS_PODR.КОД = MainOccupant.Dom SET MainOccupant.DomTip = [KLS_PODR]![Tip]")


Label2 = "Ж Д И Т Е"
Label2.Refresh

Mconn.Execute ("DELETE tmp_Z.* FROM TMP_Z")
'mconn.Execute ("INSERT INTO Tmp_Z ( Dom, TabN, Vid, SummaI ) SELECT Val([DOM]) AS Выражение1, All_z.TABN, All_z.VID, All_z.MONEY FROM All_z")
Mconn.Execute ("INSERT INTO Tmp_Z ( Dom, TabN, Vid, SummaI ) SELECT Val([DOM]) AS Выражение1, All_z.TABN, All_z.VID, All_z.MONEY FROM All_z Where All_z.dom is not null")

'If MsgBox("Да - С удалением НАЧИСЛЕНИЙ ЛИЦЕВЫХ СЧЕТОВ ТЕКУЩЕГО МЕСЯЦА, Нет-добавить к уже имеющемуся", vbYesNo) = vbYes Then
Mconn.Execute ("DELETE Adding.* FROM Adding")

Label2 = "< Ж Д И Т Е >"
Label2.Refresh

Mconn.Execute ("INSERT INTO Adding ( KodKv, KodN, SummaI ) SELECT MainOccupant.Numer, Tmp_Z.Vid, Tmp_Z.SummaI FROM Tmp_Z INNER JOIN MainOccupant ON (Tmp_Z.Dom = MainOccupant.Dom) AND (Tmp_Z.TabN = MainOccupant.OLDNUM)")
'Прочие
Mconn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.Propis = [MainOccupant]![NLODGERF], Adding.Projiv = [MainOccupant]![NLODGER], Adding.ProLift = [MainOccupant]![NLODLIFT], Adding.ObPl = [MainOccupant]![COMSPACE], Adding.PolPl = [MainOccupant]![HABSPACE], Adding.FLOOR = [MainOccupant]![FLOOR], Adding.TipKvKod = [MainOccupant]![KV], Adding.TipDomKod = [MainOccupant]![DomTip]")

Mconn.Execute ("INSERT INTO Lgota ( Numer, NomNum, OhteCode, LPKV, LPTEH, LPOTOPL, LPCOMM, LPMUSOR, USEKV, USETEH, USEOTOPL, USECOMM, USEMUSOR, NAME_KLS ) SELECT KLS_PRIV.N_KLS, MainOccupant.Numer, 0 AS Выражение1, KLS_PRIV.LPKV, KLS_PRIV.LPTEH, KLS_PRIV.LPOTOPL, KLS_PRIV.LPCOMM, KLS_PRIV.LPMUSOR, KLS_PRIV.USEKV, KLS_PRIV.USETEH, KLS_PRIV.USEOTOPL, KLS_PRIV.USECOMM, KLS_PRIV.USEMUSOR, KLS_PRIV.NAME_KLS FROM KLS_PRIV INNER JOIN MainOccupant ON KLS_PRIV.N_KLS = MainOccupant.PRIVILEGE WHERE (((MainOccupant.PRIVILEGE)<>0 And (MainOccupant.PRIVILEGE) Is Not Null))")

'Совместно проживающие
'Удаляем
Mconn.Execute ("DELETE OtheOwner.* FROM OtheOwner")
'Добавляем
'mconn.Execute ("INSERT INTO OtheOwner ( Dom, OldNum, FAM, IM, OT, PRIVILEGE, BIRTHDAY, KV ) SELECT Val([DOM]) AS Выражение1, All_k.TABN, All_k.FAM, All_k.IM, All_k.OT, All_k.PRIVILEGE, All_k.BIRTHDAY, All_k.NAPARTMENT From All_k WHERE (((All_k.NUMHABIT)<>1) and all_k.dom is not null)")
 Mconn.Execute ("INSERT INTO OtheOwner ( OldNum, FAM, IM, OT, PRIVILEGE, BIRTHDAY, KV ) SELECT All_k.TABN, All_k.FAM, All_k.IM, All_k.OT, All_k.PRIVILEGE, All_k.BIRTHDAY, All_k.NAPARTMENT From All_k WHERE (((All_k.NUMHABIT)<>1))")
'Проставляем номера
'mconn.Execute ("UPDATE MainOccupant RIGHT JOIN OtheOwner ON MainOccupant.OLDNUM = OtheOwner.OLDNUM SET OtheOwner.Numer = [MainOccupant]![Numer] WHERE (((OtheOwner.Numer) Is Not Null))")
Mconn.Execute ("UPDATE MainOccupant RIGHT JOIN OtheOwner ON MainOccupant.OLDNUM = OtheOwner.OLDNUM SET OtheOwner.Numer = [MainOccupant]![Numer]")
'Постоянные начисления

Mconn.Execute ("DELETE Constanta.* FROM Constanta")
Mconn.Execute ("INSERT INTO Constanta ( Numer, KodNach ) SELECT MainOccupant.Numer, All_l.VID FROM MainOccupant INNER JOIN All_l ON MainOccupant.OLDNUM = All_l.TABN")

'Льготы СМ-См
Mconn.Execute ("UPDATE KLS_PRIV SET KLS_PRIV.USEKV = " + Chr(34) + "См на всех" + Chr(34) + " WHERE (((KLS_PRIV.USEKV)=" + Chr(34) + "СМ на всех" + Chr(34) + "))")
Mconn.Execute ("UPDATE KLS_PRIV SET KLS_PRIV.USEteh = " + Chr(34) + "См на всех" + Chr(34) + " WHERE (((KLS_PRIV.USEteh)=" + Chr(34) + "СМ на всех" + Chr(34) + "))")
Mconn.Execute ("UPDATE KLS_PRIV SET KLS_PRIV.USEotopl = " + Chr(34) + "См на всех" + Chr(34) + " WHERE (((KLS_PRIV.USEotopl)=" + Chr(34) + "СМ на всех" + Chr(34) + "))")
Mconn.Execute ("UPDATE KLS_PRIV SET KLS_PRIV.USEcomm = " + Chr(34) + "См на всех" + Chr(34) + " WHERE (((KLS_PRIV.USEcomm)=" + Chr(34) + "СМ на всех" + Chr(34) + "))")
Mconn.Execute ("UPDATE KLS_PRIV SET KLS_PRIV.USEmusor = " + Chr(34) + "См на всех" + Chr(34) + " WHERE (((KLS_PRIV.USEmusor)=" + Chr(34) + "СМ на всех" + Chr(34) + "))")


'mconn.Execute ("")

ЗапЛьгот
MsgBox ("Далее необходимо настроить справочники и выполнить расчет всех лицевых счетов")
MsgBox ("ВНИМАНЕЕ!! проверте справочник льгот. В результате переноса возможно изменение буквы <C> параметров <СМ на 1-го > и <СМ на всех> на маленькую букву <с>. Для устранения ошибки просто пройдите по справочнику льгот и переназначте эти параметры при помощи падающего меню")
Mass.Show
Mass.продолжить


'Me.Hide
'Nachisleniy.Show
'Nachisleniy.Visible = False
'Nachisleniy.Command4.Visible = True
'Nachisleniy.Command4.Enabled = True


'Nachisleniy.ControlBox



'ЗапЛьгот

'Me.Enabled = True
'Me.Command1.Visible = False
'Me.Command2.Visible = False
'Me.Command3.Visible = True
'Label2 = "Перенос окончен"
'Label2.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
ЗапЛьгот

Filter.Nm = 0
RsPer.Open ("Adding")

RsPer.MoveFirst
Do While Not RsPer.EOF
i = 1
Filter.Nm = Str(RsPer.Fields("KodKv").Value)
'Lic.ЛучшаяЛгота RsPer.Fields("Key").Value, True

RsPer.MoveNext

Label2 = Str(i)
Label2.Refresh
i = i + 1
Loop
RsPer.Close

Label2 = "Перенос ШАГ №3 ЗАВЕРШЕН"



'Lic.ЛучшаяЛгота Lic.FG1.TextMatrix(Lic.FG1.Row, 26), True
End Sub

Private Sub Form_Load()
Set Mconn = New ADODB.Connection

Mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
Mconn.Open "data/Kvartplata.mdb"
Me.Command3.Visible = False

Set RsPer = New ADODB.Recordset
Set RsPer.ActiveConnection = Mconn
RsPer.CursorType = adOpenForwardOnly
RsPer.LockType = adLockBatchOptimistic


End Sub

' Процедура заполнения файла TMP_Lgot для последующего выбора     *
' наилучшего процкнта льготы для каждого начисления               *
'******************************************************************

Private Sub ЗапЛьгот()



'Удаляие старые льготы для  [Filter].[nm]
Mconn.Execute ("DELETE tmp_lgota.* FROM tmp_lgota")
'mconn.Execute ("DELETE tmp_lgota.KodKv From tmp_lgota WHERE (((tmp_lgota.KodKv)=" + [Filter].[Nm] + "))")

Mconn.Execute ("UPDATE Adding INNER JOIN Kategor ON Adding.KodKat = Kategor.Код SET Adding.Parametr = Kategor.Parametr")


'Добавляем  льготы для "квартплата" [Filter].[nm]
Mconn.Execute ("INSERT INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEKV, Lgota.LPKV, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Квартплата" + Chr(34) + "))")

'Добавляем  льготы для "Отопление" [Filter].[nm]
Mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEotopl, Lgota.LPotopl, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Отопление" + Chr(34) + "))")

'Добавляем  льготы для "Техобслуживание" [Filter].[nm]
Mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEteh, Lgota.LPteh, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Техобслуживание" + Chr(34) + "))")

'Добавляем  льготы для "Мусор" [Filter].[nm]
Mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEmusor, Lgota.LPmusor, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Мусор" + Chr(34) + "))")

'Добавляем  льготы для "Коммунальные услуги" [Filter].[nm]
Mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEcomm, Lgota.LPcomm, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Коммунальные услуги" + Chr(34) + "))")

End Sub
Public Sub продолжить()

ЗапЛьгот
Me.Enabled = True
Me.Command1.Visible = False
Me.Command2.Visible = False
Me.Command3.Visible = True
Label2 = "Перенос окончен"
Label2.Refresh
End Sub


