VERSION 5.00
Begin VB.Form Расчет1 
   Caption         =   "Расчет"
   ClientHeight    =   2112
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   2112
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Пожалуйста подождите. Идет расчет."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Расчет1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public KEY, n, Proc As Double
Dim R As Integer
Dim Ni(100), KodK, SaldoN(100), SaldoK(100), Nac(100), Ud(100) As Double

Public Isprav As Integer 'Isprav = 1 С учетом исправлений 2 Без учета испр
'Dim mconn As ADODB.Connection
Dim Ras As ADODB.Recordset
Dim TMP As ADODB.Recordset
Dim Formula As String



Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()




Filter.Enabled = False
Command1.Enabled = False
SposobR.Command1.Visible = False
SposobR.Command2.Visible = False
SposobR.Command3.Visible = False

SposobR.Label1 = "Начало расчета"
SposobR.Label1.Refresh



'Set conn = New ADODB.Connection
'conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' conn.Open "data/Kvartplata.mdb"
Set Ras = New ADODB.Recordset
Set Ras.ActiveConnection = Mconn
Ras.CursorType = adOpenForwardOnly
Ras.LockType = adLockBatchOptimistic

'Ras.Open ("Adding")
n = 0
I = 0
MainForm.EOk = 0
For R = 2 To Filter.Fg.Rows - 1
SposobR.ProgressBar1.Value = R

'Если дни не заполнены то ставим норму

Mconn.Execute ("UPDATE Adding SET Adding.DnP = [Adding]![DnF] WHERE (((Adding.DnP)=0))")

'N- номер счета ключевое поле НЕ СТАРЫЙ


        If Filter.Fg.Cell(flexcpChecked, R, 7) = 1 Then ' ТОЛЬКО ОТМЕЧЕННЫЕ
        
n = Val(Filter.Fg.TextMatrix(R, 0))
'MsgBox (Str(N))
        
        
        I = I + 1
'1= flexChecked)
Ras.Open ("SELECT Adding.Obpl ,Adding.KodKv, Adding.Formula, Adding.FormulaB, Adding.Key, Adding.LgotaP, Adding.tip, Adding.saldon, Adding.saldok, Adding.summai, Adding.kol, Adding.kodkat, Adding.Lig From Adding WHERE (((Adding.KodKv)=" + Str(n) + "))"), Mconn, adOpenKeyset, adLockPessimistic
    'On Error Resume Next
If Ras.RecordCount > 0 Then Ras.MoveFirst
Formula = "0"
KEY = 0


'ЗапЛьготОтм




                                       Do While Not Ras.EOF
                                       
                                       
DoEvents
  
MainForm.Pi = 0
'MainForm.Ostatok = Lic.FG1.TextMatrix(Lic.FG1.Row, 15)
MainForm.II = 0
                                       
SposobR.Refresh
MainForm.R1 = R

SposobR.Label1 = "Счет № " + Filter.Fg.TextMatrix(R, 1) + "   " + Filter.Fg.TextMatrix(R, 2) + " " + Left(Filter.Fg.TextMatrix(R, 3), 1) + ". " + Left(Filter.Fg.TextMatrix(R, 4), 1) + "."
SposobR.Label1.Refresh


'If Int(i / 100) - i / 100 = 0 Then
'MsgBox(Str(i))
'SposobR.Hide
'SposobR.Show
'End If
                                       
                                       
                                       
                                       
If Ras.Fields("Formula").Value <> "" Then Formula = Ras.Fields("Formula").Value Else Formula = "SummaI"

If Ras.Fields("FormulaB").Value <> "" Then FormulaB = Ras.Fields("FormulaB").Value Else Formula = "SummaI"

KEY = Ras.Fields("Key").Value


'ЛучшаяЛгота1 KEY, True

         If SposobR.Dolgo = True Then

If SposobR.RasRr = True Then
SposobR.RasRr = False
Unload SposobR
Filter.Enabled = True

Exit Sub
Unload Me
End If


'Если счетчик не заполнен то ставим норму

'Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Prois] WHERE (((Adding.Sch)='Да') AND (([Adding]![Shc_old]-[Adding]![Shc_new])=0) AND ((Adding.Key)=" + Str(KEY) + "))")
Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Propis] WHERE (((Adding.Sch)='Да') AND ((Adding.Shc_new)=0) AND ((Adding.Key)=" + Str(KEY) + "))")

'Если счетчик то площадь равна разнице показаний счетчика
Mconn.Execute ("UPDATE Adding SET Adding.ObPl = [Adding]![Shc_new]-[Adding]![Shc_old] WHERE (((Adding.Sch)='Да') AND ((Adding.Key)=" + Str(KEY) + "))")


MainForm.Ostatok = Ras.Fields("Obpl").Value
MainForm.II = 0
MainForm.Pi = 0
'MainForm.РЛ Val(KEY), True
If Ras("Lig") = "Да" Then

'Если счетчик не заполнен то ставим норму

'Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Prois] WHERE (((Adding.Sch)='Да') AND (([Adding]![Shc_old]-[Adding]![Shc_new])=0) AND ((Adding.Key)=" + Str(KEY) + "))")
Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Prois] WHERE (((Adding.Sch)='Да') AND ((Adding.Shc_new)=0) AND ((Adding.Key)=" + Str(KEY) + "))")

'Если счетчик то площадь равна разнице показаний счетчика
Mconn.Execute ("UPDATE Adding INNER JOIN TMP_LGOTA ON Adding.Key = TMP_LGOTA.UniKOd SET TMP_LGOTA.Plo = [Adding]![Shc_new]-[Adding]![Shc_old] WHERE (((Adding.Sch)='Да') AND ((Adding.Key)=" + Str(KEY) + "))")

Расчет Str(KEY)
End If
'If MainForm.Двойник = True Then
'MainForm.Ostatok = Ras.Fields("Obpl").Value
'MainForm.Pi = 0
'MainForm.II = 0
'MainForm.РЛ Val(KEY), False

'Расчет KEY
'End If

'MainForm.ViborLLg Val(KEY)


'Ras.Fields("LgotaP").Value = MainForm.PrZ

Ras.UpdateBatch


     End If
     
     
'MainForm.Количество N
'Filter.Nm , N

'Заполняем ЛучшаяЛгота
'Lic.ЛучшаяЛгота Str(KEY), True

'Isprav =  2 Без учета испр
If Isprav = 2 Then
'MsgBox (Str(Isprav) + "  " + Str(KEY))
'MainForm.РЛ N, Val(KEY), Proc

'Если счетчик не заполнен то ставим норму
'Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Prois] WHERE (((Adding.Sch)='Да') AND (([Adding]![Shc_old]-[Adding]![Shc_new])=0) AND ((Adding.Key)=" + Str(KEY) + "))")
Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Prois] WHERE (((Adding.Sch)='Да') AND ((Adding.Shc_new)=0) AND ((Adding.Key)=" + Str(KEY) + "))")
'Если счетчик то площадь равна разнице показаний счетчика
Mconn.Execute ("UPDATE Adding SET Adding.ObPl = [Adding]![Shc_new]-[Adding]![Shc_old] WHERE (((Adding.Sch)='Да') AND ((Adding.Key)=" + Str(KEY) + "))")

Mconn.Execute ("UPDATE Adding SET Adding.SummaI = " + Formula + ", Adding.SummaBl = " + FormulaB + ", Adding.Ispr = 0  WHERE (((Adding.Key)=" + Str(KEY) + "))")
'mconn.Execute ("UPDATE Adding SET Adding.SummaI = " + Formula + ", Adding.ispr = 0 WHERE (((Adding.Key)=" + Str(KEY) + "))")
End If
'Isprav =  1 C учетом испр

If Isprav = 1 Then

'Если счетчик не заполнен то ставим норму
'Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Propis] WHERE (((Adding.Sch)='Да') AND (([Adding]![Shc_old]-[Adding]![Shc_new])=0) AND ((Adding.Key)=" + Str(KEY) + "))")
Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Propis] WHERE (((Adding.Sch)='Да') AND ((Adding.Shc_new)=0) AND ((Adding.Key)=" + Str(KEY) + "))")

'Если дни не заполнены то ставим норму



'Если счетчик то площадь равна разнице показаний счетчика
Mconn.Execute ("UPDATE Adding SET Adding.ObPl = [Adding]![Shc_new]-[Adding]![Shc_old] WHERE (((Adding.Sch)='Да') AND ((Adding.Key)=" + Str(KEY) + "))")




'MainForm.РЛ N, Val(KEY), Proc
Mconn.Execute ("UPDATE Adding SET Adding.SummaI = " + Formula + ", Adding.SummaBl = " + FormulaB + " WHERE (((Adding.Key)=" + Str(KEY) + ") and (Adding.Ispr=0))")
End If





Ras.MoveNext
                                              Loop



        End If
        
        
        
        
        
'////////////////////////////

        If Filter.Fg.Cell(flexcpChecked, R, 7) = 1 Then ' ТОЛЬКО ОТМЕЧЕННЫЕ
If Ras.RecordCount > 0 Then Ras.MoveFirst
Do While Not Ras.EOF

KodK = Ras.Fields("KodKat").Value
SaldoN(KodK) = Ras.Fields("SaldoN").Value
If Ras.Fields("KodKat").Value = KodK Then Ni(KodK) = Ni(KodK) + 1
If Ras.Fields("Tip").Value = "+" Then Nac(KodK) = Nac(KodK) + Ras.Fields("SummaI").Value
If Ras.Fields("Tip").Value = "-" Or Ras.Fields("Tip").Value = "s" Then Ud(KodK) = Ud(KodK) + Ras.Fields("SummaI").Value
SaldoK(KodK) = SaldoN(KodK) + Nac(KodK) - Ud(KodK)
Ras.MoveNext
Loop




If Ras.RecordCount > 0 Then Ras.MoveFirst
Do While Not Ras.EOF
KodK = Ras.Fields("KodKat").Value
If Ras.Fields("KodKat").Value = KodK Then Ras.Fields("kol").Value = Ni(KodK)
Ras.Fields("saldoK").Value = SaldoK(KodK)

Ras.UpdateBatch
Ras.MoveNext
Loop

For J = 0 To 100
Ni(J) = 0
SaldoK(J) = 0
SaldoN(J) = 0
Nac(J) = 0
Ud(J) = 0
Next J


Ras.Close


    End If
'/////////////////////////////
        
        
        
        
        
        
SposobR.Caption = Str(I) + " Ошибок - " + Str(MainForm.EOk)
'If I / 50 - Int(I / 50) = 0 Then
'SposobR.Show

Filter.Fg.Cell(flexcpChecked, R, 8) = True
Filter.Fg.Cell(flexcpChecked, R, 7) = flexUnchecked
'End If



'Filter.FG.TextMatrix(r, 7) = False


Next
       
SposobR.Show
SposobR.Label1 = "Расчет окончен. Количество расчитанных лицевых счетов = " + Str(I)
SposobR.Label1.Refresh
       
SposobR.Enabled = True



SposobR.Command4.Visible = True

       
       
'Label1 = "Расчет окончен. Количество расчитанных лицевых счетов = " + Str(i)
'Command1.Enabled = True
Расчет1.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Filter.Enabled = True
Filter.Show

End Sub

'  О Т М Е Ч Е Н Н Ы Е
'******************************************************************
' Процедура заполнения файла TMP_Lgot для последующего выбора     *
' наилучшего процкнта льготы для каждого начисления               *
'******************************************************************
'Расчет для отмеченных
Private Sub ЗапЛьготОтм()

'mconn.Execute ("UPDATE Adding INNER JOIN Kategor ON Adding.KodKat = Kategor.Код SET Adding.Parametr = Kategor.Parametr")
'Удаляие старые льготы для  [Filter].[nm]

Mconn.Execute ("DELETE tmp_lgota.KodKv From tmp_lgota WHERE (((tmp_lgota.KodKv)=" + Str(n) + "))")
'Добавляем  льготы для "квартплата" [Filter].[nm]
Mconn.Execute ("INSERT INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif, Parametr ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEKV, Lgota.LPKV, Adding.Tarif, Adding.Parametr FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Квартплата" + Chr(34) + ") and (Lgota.NomNum)=" + Str(n) + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
'Добавляем  льготы для "Отопление" [Filter].[nm]
Mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif, Parametr ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEotopl, Lgota.LPotopl, Adding.Tarif, Adding.Parametr FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Отопление" + Chr(34) + ") and (Lgota.NomNum)=" + Str(n) + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
'Добавляем  льготы для "Техобслуживание" [Filter].[nm]
Mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif, Parametr ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEteh, Lgota.LPteh, Adding.Tarif, Adding.Parametr FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Техобслуживание" + Chr(34) + ") and (Lgota.NomNum)=" + Str(n) + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
'Добавляем  льготы для "Мусор" [Filter].[nm]
Mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif, Parametr ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEmusor, Lgota.LPmusor, Adding.Tarif, Adding.Parametr FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Мусор" + Chr(34) + ") and (Lgota.NomNum)=" + Str(n) + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
'Добавляем  льготы для "Коммунальные услуги" [Filter].[nm]
Mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif, Parametr ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEcomm, Lgota.LPcomm, Adding.Tarif, Adding.Parametr FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Коммунальные услуги" + Chr(34) + ") and (Lgota.NomNum)=" + Str(n) + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
End Sub


Public Sub УДАЛИТЬ_ЛучшаяЛгота1_1(ByVal UniK As Double, Zapis As Boolean)

If UniK = 0 Then
MsgBox ("Не выбрано начисление")
Exit Sub
End If

Dim klsKod, othe, KEY, GoodKLS, OtheKol, KolL As Integer
Dim Plo, prop, Procent, Socmin, Socmin1, Socmin2, Tarif, Itog, tmpItog, ItogOdin As Double
Dim Use, Vid, Parametr As String





Set GoodL = New ADODB.Recordset
Set GoodL.ActiveConnection = Mconn
GoodL.CursorType = adOpenForwardOnly
GoodL.LockType = adLockBatchOptimistic

Set Odin = New ADODB.Recordset
Set Odin.ActiveConnection = Mconn
Odin.CursorType = adOpenForwardOnly
Odin.LockType = adLockBatchOptimistic





Mconn.Execute ("UPDATE Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd SET tmp_lgota.parametr = [Adding]![Parametr] WHERE (((Adding.KodKv)=" + Filter.Nm + "))")
GoodL.Open ("SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.KodKls, tmp_lgota.NAME_KLS, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.Procent, tmp_lgota.Plo, tmp_lgota.Prop, tmp_lgota.Cocmin, tmp_lgota.OtheCode, tmp_lgota.Parametr, tmp_lgota.tarif From tmp_lgota WHERE (((tmp_lgota.KodKv)=" + Filter.Nm + "))")

Lic.Itog1 = 1
Itog = 1
tmpItog = 1



If GoodL.EOF = False Then GoodL.MoveFirst

                             
                             
                           Do While Not GoodL.EOF ' Перебор по данному лиц.счету
                          
                           
'KEY- Код строки начесления в лц.счете уникальный.  В данном запросе равен
'количеству льгот ВОЗМОЖНЫХ к применеию для данного
'начисления + Othe ЛЬГОТЫ СОВМ.ПРОЖИВАЮЩИХ

        KEY = GoodL.Fields("UniKod").Value
        
                                              If KEY = UniK Or UniK = 0 Then
                             
Plo = GoodL.Fields("Plo").Value
prop = GoodL.Fields("Prop").Value
Procent = GoodL.Fields("Procent").Value
Use = GoodL.Fields("Use").Value
Vid = GoodL.Fields("LgotaVid").Value
othe = GoodL.Fields("OtheCode").Value
'Parametr = GoodL.Fields("Parametr").Value
Socmin = GoodL.Fields("Cocmin").Value
klsKod = GoodL.Fields("KodKls").Value
Tarif = GoodL.Fields("tarif").Value
'OtheKol = 0
'KolL = 0
'MsgBox (Str(Socmin) + "  " + Str(Tarif))


'*******************   1. Расчет производится от Общей площади


'If Parametr = "Общ.площадь" Then

'Вся площадь
If Use = "Вся площадь" Then
Itog = (100 - Procent) / 100
OtheKol = 0
End If



  'См на всех
            
            If Use = "См на всех" Then
                    If Plo > Socmin Then
                      Itog = (((Socmin * (100 - Procent)) / 100) + (Plo - Socmin)) / Plo
                      KolL = 0
                    Else
                      Itog = (100 - Procent) / 100
                      KolL = 0
                    End If
                    
            
            End If
            
 'СМ на 1-го
 
  
 
                    If Trim(Use) = Trim("СМ на 1-го") Then
                     
                       TMP.Open ("SELECT Socmin.Value, Socmin.koli FROM Adding INNER JOIN Socmin ON Adding.KodKat = Socmin.KodKategor WHERE (((Socmin.koli)=1) and ((Adding.key)=" + Str(KEY) + ")) ")
                       Socmin1 = TMP.Fields("Value").Value
                       TMP.Close
                       
                         
                              If Plo > Socmin1 Then
                                  Itog = (((Socmin1 * (100 - Procent)) / 100) + (Plo - Socmin1)) / Plo
                                 KolL = 0
                         
                
                              Else
                                  Itog = (100 - Procent) / 100
                                  KolL = 0
                
                              End If
                              
                    End If
'См/кол.жил

                    If Trim(Use) = Trim("См/кол.жил") Then
                    
                    'MsgBox (Str(Socmin2) + " = " + Str(Socmin) + " / " + Str(Prop))
                    Socmin2 = Socmin / prop

                               If Plo > Socmin2 Then
                     '          MsgBox (Str(Plo) + ">" + Str(Socmin2))
                                  Itog = (((Socmin2 * (100 - Procent)) / 100) + (Plo - Socmin2)) / Plo
                      '            MsgBox (Str(Itog))
                                  KolL = 0
                              Else
                                  Itog = (100 - Procent) / 100
                                  KolL = 0
                              End If
                    
                    End If
                    
               
                    
           
            
'End If



'*******************   1. Расчет производится от кол-ва прописанных

'If Parametr = "Кол-во пропис." Then

                      If Use = "На всех" Then
                      Itog = (100 - Procent) / 100
                      KolL = 0
                      End If

' "На 1-го НЕ ЖИЛПЛОЩАДЬ





                      If Use = "На 1-го" Then
                      

                      'TMP.Open ("SELECT Socmin.Value, Socmin.koli FROM Adding INNER JOIN Socmin ON Adding.KodKat = Socmin.KodKategor WHERE (((Socmin.koli)=1) and ((Adding.key)=" + Str(KEY) + ")) ")
                      'Socmin1 = TMP.Fields("Value").Value
                       'TMP.Close
                       
      '//////////////// Подсчет процента льгот НА ОДНОГО для НЕСКОЛЬКИХ //////////////////////
      ItogOdin = 0
      'odin.Open ("SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.KodKls, tmp_lgota.NAME_KLS, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.Procent, tmp_lgota.OtheCode From tmp_lgota WHERE (((tmp_lgota.UNIKOD)=" + Str(KEY) + ") AND ((tmp_lgota.Use)=" + Chr(34) + "На 1-го" + Chr(34) + ") and ((tmp_lgota.KodKls)=" + Str(klsKod) + ")) ORDER BY tmp_lgota.OtheCode")
      Q = "SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.LgotaVid, tmp_lgota.Use, Max(tmp_lgota.Procent) AS [Procent], tmp_lgota.OtheCode From tmp_lgota GROUP BY tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.OtheCode Having (((tmp_lgota.UniKOd) = " + Str(KEY) + ") And ((tmp_lgota.Use) = " + Chr(34) + "На 1-го" + Chr(34) + ")) ORDER BY tmp_lgota.OtheCode"
      Odin.Open (Q)
      
      
                         Odin.MoveFirst
                         
                         
                         
                OldOthe = Odin.Fields("OtheCode").Value
      P = 0
      'KolL = 0
                
                      Do While Not Odin.EOF
      Procent1 = (100 - Odin.Fields("Procent").Value) / 100
      
      If Odin.Fields("OtheCode").Value <> OldOthe Then
      P = P + Procent1
      'Itog = 1 - (ItogOdin + (1 * procent1 * (Prop - 1)) / Prop)
      KolL = KolL + 1
      'MsgBox ("<>  " + Str(Procent1) + "%   Лготчиков=" + Str(KolL) + "  " + Str(P) + "%  Льгота=" + Str(klsKod))
      Else
      'ItogOdin = 1 - ((1 * procent1 * (Prop - 1)) / Prop)
      Itog = ItogOdin
      P = Procent1
      KolL = 1
     'MsgBox ("=  " + Str(Procent1) + "%   Лготчиков=" + Str(KolL) + "  " + Str(P) + "%  Льгота=" + Str(klsKod))
      End If
                        Odin.MoveNext
                            Loop
      Odin.Close
      Itog = (P + (prop - KolL)) / prop
      'MsgBox ("ВОТ И ВСЕ ИТОГ " + Str(Itog) + "Льготчиков " + Str(KolL))
      
      '                        If Kol * Tarif > Socmin1 Then
       '                           Itog = (((Socmin1 * Tarif * (100 - Procent)) / 100) + (Kol * Tarif - Socmin1)) / (Kol * Tarif)
        '                          KolL = 1
         '                     Else
          '                        Itog = (100 - Procent) / 100
           '                       KolL = 1
            '                  End If
                              
                    End If

'******************************************************************
'На одного


                      If Use = "На одного" Then
                      

                      'TMP.Open ("SELECT Socmin.Value, Socmin.koli FROM Adding INNER JOIN Socmin ON Adding.KodKat = Socmin.KodKategor WHERE (((Socmin.koli)=1) and ((Adding.key)=" + Str(KEY) + ")) ")
                      'Socmin1 = TMP.Fields("Value").Value
                       'TMP.Close
                       
      '//////////////// Подсчет процента льгот НА ОДНОГО для НЕСКОЛЬКИХ //////////////////////
      ItogOdin = 0
      'Odin.Open ("SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.KodKls, tmp_lgota.NAME_KLS, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.Procent, tmp_lgota.OtheCode From tmp_lgota WHERE (((tmp_lgota.UNIKOD)=" + Str(KEY) + ") AND ((tmp_lgota.Use)=" + Chr(34) + "На одного" + Chr(34) + ") and ((tmp_lgota.KodKls)=" + Str(klsKod) + ")) ORDER BY tmp_lgota.OtheCode")
      Q = "SELECT tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.LgotaVid, tmp_lgota.Use, Max(tmp_lgota.Procent) AS [Procent], tmp_lgota.OtheCode From tmp_lgota GROUP BY tmp_lgota.UniKOd, tmp_lgota.KodKv, tmp_lgota.LgotaVid, tmp_lgota.Use, tmp_lgota.OtheCode Having (((tmp_lgota.UniKOd) = " + Str(KEY) + ") And ((tmp_lgota.Use) = " + Chr(34) + "На одного" + Chr(34) + ")) ORDER BY tmp_lgota.OtheCode"
      Odin.Open (Q)
      'Set t1.V1.DataSource = odin
      't1.Show
      
      

      '**********************************
      'Формула для расчета % льготы лиц/сч. при прим НА ОДНОГО
      '1. Общий процент льгот выч. как сумма процентов льгот каждого льготчика
      '   данной квартиры (при условии что льгота НА ОДНОГО проставлена разным
      '   жильцам, определяется проверкой на несовпадение поля OtheCod, при переборе
      '   данных RecordSet Odin  для данного лиц.счета) :
      '
      '    P=P1+P2+....Pn
      '
      '
      '2. Универсальная и проверенная мнойц формула выглядит так:
      '
      '      Itog = (P + (Prop - KolL)) / Prop, где
      '
      '    Itog - Коэффициент на который надо умножит начисление без льгот
      '           для получения суммы С УЧЕТОМ ЛЬГОТ по лиц счету. Иными словами
      '           это процент льготы всего лиц.счета [Площадь*Тариф*Itog] даст
      '           сумму уже с учетом льгот. ВАЖНО! Пложадь ВСЯ ДАЖЕ СВЫШЕ СМ
      '
      '    P - процент льгот  P=P1+P2+....Pn
      '
      '    KolL - Количество льготчиков РЕАЛЬНО ПРИМЕНЯЕМЫХ В РАСЧЕТЕ
      '
      '    Prop - Кол-во прописанных
      '
      ' 3.при открытии запроса Odin производим группировку с выбором наибольшего процента
      '
      '
      '**********************************
                         Odin.MoveFirst
                         
                         
                         
                OldOthe = Odin.Fields("OtheCode").Value
      P = 0
      'KolL = 0
                
                      Do While Not Odin.EOF
      Procent1 = (100 - Odin.Fields("Procent").Value) / 100
      
      If Odin.Fields("OtheCode").Value <> OldOthe Then
      P = P + Procent1
      'Itog = 1 - (ItogOdin + (1 * procent1 * (Prop - 1)) / Prop)
      KolL = KolL + 1
      
    '  MsgBox ("<>  " + Str(Procent1) + "%   Лготчиков=" + Str(KolL) + "  " + Str(P) + "%  Льгота=" + Str(klsKod))
      Else
      'ItogOdin = 1 - ((1 * procent1 * (Prop - 1)) / Prop)
      Itog = ItogOdin
      P = Procent1
      KolL = 1
     'MsgBox ("=  " + Str(Procent1) + "%   Лготчиков=" + Str(KolL) + "  " + Str(P) + "%  Льгота=" + Str(klsKod))
      End If
                        Odin.MoveNext
                            Loop
      Odin.Close
      
      Itog = (P + (prop - KolL)) / prop
      'MsgBox ("=  " + Str(Procent1) + "%   Лготчиков=" + Str(KolL) + "  " + Str(P) + "%  Льгота=" + Str(klsKod))
      
      'MsgBox ("ВОТ И ВСЕ ИТОГ " + Str(Itog) + "Льготчиков " + Str(KolL))
      
      '                        If Kol * Tarif > Socmin1 Then
       '                           Itog = (((Socmin1 * Tarif * (100 - Procent)) / 100) + (Kol * Tarif - Socmin1)) / (Kol * Tarif)
        '                          KolL = 1
         '                     Else
          '                        Itog = (100 - Procent) / 100
           '                       KolL = 1
            '                  End If
                              
                    End If



'********************************************************************


'Выбор наименьшего процента
If tmpItog > Itog Then
tmpItog = Itog
GoodKLS = klsKod
'MsgBox (Str(GoodKLS))
End If

'MsgBox (Use + " " + Parametr + "" + Str(Procent) + " % " + Str(tmpItog))

If Zapis = True Then
'MsgBox ("Запишем  " + Str(Itog))
Mconn.Execute ("UPDATE tmp_lgota SET tmp_lgota.itog =" + Str(Itog) + " WHERE (((tmp_lgota.UniKOd)=" + Str(UniK)) + " AND ((tmp_lgota.KodKls)=" + Str(klsKod) + ")))"
Mconn.Execute ("UPDATE Adding SET Adding.othekol = " + Str(KolL) + " WHERE (((Adding.Key)=" + Str(UniK) + "))")
End If
'MsgBox (Use + " " + Parametr + "" + Str(Procent) + " % " + Str(Itog))
                                                          End If
GoodL.MoveNext


'tmpItog = Itog



                                

                                Loop

Lic.Itog1 = tmpItog
'MsgBox (Str(Lic.Itog1))
If Zapis = True Then

'mconn.Execute ("UPDATE Adding SET Adding.LgotaP = " + Str(tmpItog) + " WHERE (((Adding.Key)=" + Str(UniK) + "))")
Mconn.Execute ("UPDATE Adding SET Adding.OtheKol =1 ,Adding.LgotaP = " + Str(tmpItog) + " WHERE (((Adding.Key)=" + Str(UniK) + "))")
End If
'
GoodL.Close


End Sub

