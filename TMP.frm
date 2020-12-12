VERSION 5.00
Begin VB.Form TMP 
   Caption         =   "TMP"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "TMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lgota As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim Ok



Private Sub ЛГСМНАВСЕХ(ByVal Key11 As Double)
Dim P  As Double
Dim Rsl1 As ADODB.Recordset


Set Rsl1 = New ADODB.Recordset
Set Rsl1.ActiveConnection = mconn


Rsl1.CursorType = adOpenDynamic
Rsl1.LockType = adLockOptimistic



Rsl1.Open ("SELECT tmp_lgota.UniKOd, tmp_lgota.Itog1 From tmp_lgota WHERE (((tmp_lgota.UniKOd)=" + Str(Key11) + "))")


PL = 0.65526

Rsl1.MoveFirst
Do While Not Rsl1.EOF
Rsl1.Fields("Itog1").Value = PL
Rsl1.MoveNext
Loop
Rsl1.Close
End Sub

Private Sub ЗапЛьгот()
Ok = 0
Filter.Nm = Str(276025)
mconn.Execute ("DELETE tmp_lgota.KodKv From tmp_lgota WHERE (((tmp_lgota.KodKv)=" + [Filter].[Nm] + "))")
'Добавляем  льготы для "квартплата" [Filter].[nm]
mconn.Execute ("INSERT INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEKV, Lgota.LPKV, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Квартплата" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
'Добавляем  льготы для "Отопление" [Filter].[nm]
mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEotopl, Lgota.LPotopl, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Отопление" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
'Добавляем  льготы для "Техобслуживание" [Filter].[nm]
mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEteh, Lgota.LPteh, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Техобслуживание" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
'Добавляем  льготы для "Мусор" [Filter].[nm]
mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEmusor, Lgota.LPmusor, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Мусор" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
'Добавляем  льготы для "Коммунальные услуги" [Filter].[nm]
mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEcomm, Lgota.LPcomm, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Коммунальные услуги" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )AND ((Adding.Lig)=" + Chr(34) + "Да" + Chr(34) + ")")
Ok = 1
End Sub

