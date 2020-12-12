VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form DropForm3 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5535
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   120
      Picture         =   "DropF3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   5040
      Width           =   855
   End
   Begin VSFlex8Ctl.VSFlexGrid DgT 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _cx             =   12515
      _cy             =   7435
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
      ForeColor       =   -2147483647
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
      FormatString    =   $"DropF3.frx":0442
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   2
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
      ShowComboButton =   2
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
      BackColorFrozen =   16711680
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   600
      TabIndex        =   3
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   6810
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "DropF3.frx":0588
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   600
      Picture         =   "DropF3.frx":0CD2
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   240
      Picture         =   "DropF3.frx":141C
      Top             =   0
      Width           =   285
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
      Height          =   195
      Left            =   0
      Picture         =   "DropF3.frx":1B66
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "DropForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsL As ADODB.Recordset
Dim RsSet As ADODB.Recordset
Dim Comb As String


'Dim mconn As ADODB.Connection




Private Sub Command1_Click()
'RSL.Update

Kvart.IzmLgot = True

Mconn.Execute ("UPDATE Lgota INNER JOIN KLS_PRIV ON Lgota.Numer = KLS_PRIV.N_KLS SET Lgota.LPKV = [KLS_PRIV]![LPKV], Lgota.LPTEH = [KLS_PRIV]![LPTEH], Lgota.LPOTOPL = [KLS_PRIV]![LPOTOPL], Lgota.LPCOMM = [KLS_PRIV]![LPCOMM], Lgota.LPMUSOR = [KLS_PRIV]![LPMUSOR], Lgota.USEKV = [KLS_PRIV]![USEKV], Lgota.USETEH = [KLS_PRIV]![USETEH], Lgota.USEOTOPL = [KLS_PRIV]![USEOTOPL], Lgota.USECOMM = [KLS_PRIV]![USECOMM], Lgota.USEMUSOR = [KLS_PRIV]![USEMUSOR]")
Mconn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv=MainOccupant.Numer SET Adding.Propis = MainOccupant!NLODGERF, Adding.Projiv = MainOccupant!NLODGER, Adding.ProLift = MainOccupant!NLODLIFT, Adding.ObPl = MainOccupant!COMSPACE, Adding.PolPl = MainOccupant!HABSPACE, Adding.TipKvKod = MainOccupant!KV, Adding.TipDomKod = MainOccupant!DomTip where Adding.KodKv= " + Filter.Nm)
MainForm.ЗапЛьгот

Unload DropForm2
Unload DropForm3

End Sub


Private Sub Command2_Click()
'DgT_DblClick
DropForm3.DgT.RemoveItem (DropForm3.DgT.Row)
DgT.Refresh

End Sub

Private Sub DgT_AfterDataRefresh()
DgT.ColComboList(5) = Comb
End Sub

Private Sub DgT_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)



' if this is a date column, edit it with the date picker control
If Arhiv = True Then Cancel = True
If Col <> 5 Then Cancel = True
   
'If Col = 5 Then
        
'End If
End Sub


Private Sub Form_Load()

MakeWindow Me, True
lblTitle.Caption = "Льготы лиц/сч.№" + Filter.Nm

    
    DgT.Editable = True
    DgT.DataMode = flexDMBoundImmediate
        

If Arhiv = True Then Command2.Enabled = False

Set RsL = New ADODB.Recordset
Set RsL.ActiveConnection = Mconn

Set RsSet = New ADODB.Recordset
Set RsSet.ActiveConnection = Mconn

RsL.CursorType = adOpenForwardOnly
RsL.LockType = adLockBatchOptimistic

RsSet.CursorType = adOpenForwardOnly
RsSet.LockType = adLockBatchOptimistic

'MsgBox (filter.nm)
RsL.Open ("SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Lgota.DaatN, Lgota.DaatK, Lgota.OhteCode FROM Lgota WHERE (Lgota.NomNum= " + Filter.Nm + ")")

RsSet.Open ("SELECT OtheOwner.Numer, OtheOwner.FAM, OtheOwner.IM, OtheOwner.OT, OtheOwner.OhteCode From OtheOwner WHERE (((OtheOwner.Numer)=" + Filter.Nm + "))")



DgT.FixedCols = 0


Set DgT.DataSource = RsL
For Rw = 1 To DgT.Rows - 1
If DgT.TextMatrix(Rw, DgT.Cols - 1) <> "" Then
If DgT.TextMatrix(Rw, DgT.Cols - 1) <> 0 Then DgT.Cell(flexcpForeColor, Rw, 1, Rw, DgT.Cols - 1) = vbBlue
End If
Next



Comb = "#" + "0" + ";" + "Ответственный квартиросъемщик" + "|"
On Error GoTo NoFirst
RsSet.MoveFirst
NoFirst:

Do While Not RsSet.EOF
'If RsSet("fam") = "" Then
'RsSet("fam") = "_"
'RsSet.Update
'End If

'If RsSet("Im") = "" Then
'RsSet("Im") = "_"
'RsSet.Update
'End If

'If RsSet("Ot") = "" Then
'RsSet("Ot") = "_"
'RsSet.Update
'End If



If RsSet("fam") <> "" Then F = RsSet("fam")
If RsSet("Im") = "" Then Im = RsSet("Im")
If RsSet("Ot") = "" Then O = RsSet("Ot")

If F <> "" Then
Comb = Comb + "#" + Str(RsSet("OhteCode")) + ";" + F + " " + Im + " " + O + "|"
Else
Comb = Comb + "#" + Str(RsSet("OhteCode")) + ";" + "Фамилия не указана" + "|"
End If
RsSet.MoveNext
Loop
 DgT.ColComboList(5) = Comb
End Sub

Private Sub Form_Unload(Cancel As Integer)
RsL.Close
End Sub
Private Sub before_add()
'RSL.MoveLast
'RSL.AddNew
'RSL.MoveLast
End Sub
Private Sub dtPick_Change()
    
    ' update grid value whenever the data changes
    DgT.Text = dtPick.Value
    
End Sub

Private Sub dtPick_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' close date picker when user hits escape or return
    Select Case KeyCode
        Case vbKeyEscape
            FG = dtPick.Tag
            dtPick.Visible = False
        Case vbKeyReturn
            dtPick.Visible = False
    End Select
    
End Sub

Private Sub dtPick_LostFocus()

    ' hide date picker when user is done with it
    dtPick.Visible = False
    
End Sub
'***********************************


'******************************************************************
' Процедура заполнения файла TMP_Lgot для последующего выбора     *
' наилучшего процкнта льготы для каждого начисления               *
'******************************************************************

'Private Sub ЗапЛьгот()

'Удаляие старые льготы для  [Filter].[nm]
'mconn.Execute ("DELETE tmp_lgota.KodKv From tmp_lgota WHERE (((tmp_lgota.KodKv)=" + [Filter].[Nm] + "))")

'Добавляем  льготы для "квартплата" [Filter].[nm]
'mconn.Execute ("INSERT INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEKV, Lgota.LPKV, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Квартплата" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )")

'Добавляем  льготы для "Отопление" [Filter].[nm]
'mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEotopl, Lgota.LPotopl, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Отопление" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )")

'Добавляем  льготы для "Техобслуживание" [Filter].[nm]
'mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEteh, Lgota.LPteh, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Техобслуживание" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )")

'Добавляем  льготы для "Мусор" [Filter].[nm]
'mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEmusor, Lgota.LPmusor, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Мусор" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )")

'Добавляем  льготы для "Коммунальные услуги" [Filter].[nm]
'mconn.Execute ("INSERT  INTO tmp_lgota ( KodKv, KodKls, NAME_KLS, LgotaVid, UniKOd, Plo, Prop, Cocmin, OtheCode, Use, Procent, tarif ) SELECT Lgota.NomNum, Lgota.Numer, Lgota.NAME_KLS, Adding.LgotaVid, Adding.Key, Adding.ObPl, Adding.Propis, Adding.Socmin, Lgota.OhteCode, Lgota.USEcomm, Lgota.LPcomm, Adding.Tarif FROM Adding INNER JOIN Lgota ON Adding.KodKv = Lgota.NomNum WHERE (((Adding.LgotaVid)=" + Chr(34) + "Коммунальные услуги" + Chr(34) + ") and (Lgota.NomNum)=" + [Filter].[Nm] + " )")
'End Sub
Private Sub lblTitle_Click()

End Sub
