VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form SubsAddr 
   Caption         =   "œÓ‚ÂÍ‡ ‡‰ÂÒÓ‚"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12480
   LinkTopic       =   "Form8"
   ScaleHeight     =   6420
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   12135
      _cx             =   21405
      _cy             =   9975
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
Attribute VB_Name = "SubsAddr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim mconn As ADODB.Connection
Dim Addr As ADODB.Recordset
Dim AddrC As ADODB.Recordset
Dim Cl As String
Dim a As Integer


Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
FG.TextMatrix(FG.Row, 4) = Val(FG.TextMatrix(FG.Row, 4))
FG.TextMatrix(FG.Row, 5) = Val(FG.TextMatrix(FG.Row, 4))

AddrC.MoveFirst
Do While Not AddrC.EOF
If FG.TextMatrix(FG.Row, FG.Col) = AddrC(" Œƒ") Then
FG.TextMatrix(FG.Row, 6) = AddrC("NAIM_KLS")
FG.TextMatrix(FG.Row, 7) = AddrC("Num")
FG.TextMatrix(Rw, 5) = AddrC(" Œƒ")
End If
AddrC.MoveNext
Loop



'œ–Œ—“¿¬»“‹ Œƒ€
End Sub

Private Sub FG_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

If FG.Col <> 4 Then
FG.ComboList = ""
Exit Sub
a = 0
End If
FG.ComboSearch = flexCmbSearchAll
If FG.Col = 4 Then
FG.ComboList = Cl
a = 1
Else
FG.ComboList = ""
End If

End Sub

Private Sub Form_Load()
'Set mconn = New ADODB.Connection
'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
'mconn.Open "data/Kvartplata.mdb"

Set Addr = New ADODB.Recordset
Set Addr.ActiveConnection = mconn




Set AddrC = New ADODB.Recordset
Set AddrC.ActiveConnection = mconn

AddrC.Open ("SELECT KLS_PODR. Œƒ, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM KLS_PODR"), mconn, adOpenDynamic, adLockPessimistic


Cl = ""
AddrC.MoveFirst
Do While Not AddrC.EOF
Cl = Cl + Str(AddrC(" Œƒ")) + "  ÛÎ." + AddrC("NAIM_KLS") + "ƒÓÏ π" + AddrC("NUM") + "|"
AddrC.MoveNext
Loop
Addr.Open ("SELECT JAK. Ó‰, JAK.NYLIC, JAK.NDOM, JAK.Kod_ylic, KLS_PODR. Œƒ, KLS_PODR.NAIM_KLS, KLS_PODR.Num FROM JAK LEFT JOIN KLS_PODR ON (JAK.NDOM = KLS_PODR.Num) AND (JAK.NYLIC = KLS_PODR.NAIM_KLS) ORDER BY KLS_PODR.NAIM_KLS")
Set FG.DataSource = Addr

œ–Œ—“¿¬»“‹ Œƒ€
End Sub
Private Sub œ–Œ—“¿¬»“‹ Œƒ€()
For Rw = 1 To FG.Rows - 1
If FG.TextMatrix(Rw, 2) = FG.TextMatrix(Rw, 6) And FG.TextMatrix(Rw, 3) = FG.TextMatrix(Rw, 7) Then FG.TextMatrix(Rw, 4) = FG.TextMatrix(Rw, 5)
'MsgBox FG.TextMatrix(Rw, 2) + "  " + FG.TextMatrix(Rw, 6)
Next
End Sub

