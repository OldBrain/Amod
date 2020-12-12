VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form DobLgot 
   Caption         =   "Form8"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12600
   LinkTopic       =   "Form8"
   ScaleHeight     =   7440
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid NS1 
      Height          =   1815
      Left            =   6240
      TabIndex        =   9
      Top             =   4680
      Width           =   6255
      _cx             =   11033
      _cy             =   3201
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
   Begin VSFlex8Ctl.VSFlexGrid VS2 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   5655
      _cx             =   9975
      _cy             =   1931
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
   Begin VB.CommandButton Command3 
      Caption         =   "Удалить лишние"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   6720
      Width           =   3975
   End
   Begin VSFlex8Ctl.VSFlexGrid VS1 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   13335
      _cx             =   23521
      _cy             =   2566
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
      Caption         =   "Отмена"
      Height          =   495
      Left            =   8880
      TabIndex        =   3
      Top             =   6720
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Разнести"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6720
      Width           =   4215
   End
   Begin VSFlex8Ctl.VSFlexGrid VS 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   13335
      _cx             =   23521
      _cy             =   2355
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Неверное сальдо на конец"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   4080
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Лишние льготы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   13095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Отсутствующие льготы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   13215
   End
End
Attribute VB_Name = "DobLgot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim m_DS As FlexADO
Dim RS_Ad, Rs_Dell As ADODB.Recordset
Dim NS As ADODB.Recordset

'Dim mconn As ADODB.Connection




Private Sub Command1_Click()
Jdite.Label1 = "Пожалуйста подождите"
Jdite.Show
Jdite.Label1.Refresh
MainForm.ДОБЛьготыВАддинг "All", True
Jdite.Hide
If MsgBox("Недостающие льготы разнесены успешно, перезаписать все ставки?", vbOKCancel) = vbOK Then
Jdite.Show
Jdite.Label1.Refresh
MainForm.ДОБЛьготыВАддинг "All", False
Unload Jdite
Command1.Visible = False
Command2.Caption = "Выйти в предыдущее меню"
Else
Command1.Visible = False
Command3.Visible = False
Command2.Caption = "Выйти в предыдущее меню"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Jdite.Label1 = "Пожалуйста подождите"
Jdite.Show
Jdite.Label1.Refresh
Mconn.Execute ("DELETE tmp_lgota.*, Adding.Key FROM tmp_lgota LEFT JOIN Adding ON tmp_lgota.UniKOd = Adding.Key WHERE (((Adding.Key) Is Null))")
Command1.Visible = False
Command3.Visible = False
Command2.Caption = "Выйти в предыдущее меню"
Unload Jdite
End Sub

Private Sub Form_Load()

'Set Mconn = New ADODB.Connection
Set RS_Ad = New ADODB.Recordset
Set Rs_Dell = New ADODB.Recordset
Set Rs_Kol = New ADODB.Recordset
Set NS = New ADODB.Recordset

'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True;Persist Security Info=True"
'mconn.Open "data/Kvartplata.mdb"

RS_Ad.Open ("SELECT Lig_Adding.NomNum, Lig_Adding.Numer, Lig_Adding.Lig, Lig_Adding.Key, Lig_Adding.FAM, Lig_Adding.IM, Lig_Adding.OT, Lig_Adding.NAME_KLS, Lig_Adding.DaatN, Lig_Adding.DaatK, Lig_Adding.OhteCode, Lig_Adding.LPKV, Lig_Adding.LPTEH, Lig_Adding.LPOTOPL, Lig_Adding.LPCOMM, Lig_Adding.LPMUSOR, Lig_Adding.USEKV, Lig_Adding.USETEH, Lig_Adding.USEOTOPL, Lig_Adding.USECOMM, Lig_Adding.USEMUSOR FROM Lig_Adding LEFT JOIN tmp_lgota ON (Lig_Adding.LgotaVid = tmp_lgota.LgotaVid) AND (Lig_Adding.Key = tmp_lgota.UniKOd) WHERE (((tmp_lgota.UniKOd) Is Null) AND ((tmp_lgota.LgotaVid) Is Null))"), Mconn
Rs_Dell.Open ("SELECT tmp_lgota.* FROM tmp_lgota LEFT JOIN Adding ON tmp_lgota.UniKOd = Adding.Key WHERE (((Adding.Key) Is Null))"), Mconn

Rs_Kol.Open ("SELECT КоличествоПоКатегориям.[Count-Key], Adding.Kol, [Adding]![Kol]-[КоличествоПоКатегориям]![Count-Key] AS Расхождение FROM КоличествоПоКатегориям INNER JOIN Adding ON (КоличествоПоКатегориям.KodKv = Adding.KodKv) AND (КоличествоПоКатегориям.KodKat = Adding.KodKat) Where ((([Adding]![Kol] - [КоличествоПоКатегориям]![Count-Key]) <> 0)) ORDER BY [Adding]![Kol]-[КоличествоПоКатегориям]![Count-Key]"), Mconn
NS.Open ("Проверка_сальдо"), Mconn

Set VS.DataSource = RS_Ad
Set VS1.DataSource = Rs_Dell
Set VS1.DataSource = Rs_Kol
Set NS1.DataSource = NS

End Sub

