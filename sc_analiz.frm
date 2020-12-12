VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form sc_analiz 
   Caption         =   "Form3"
   ClientHeight    =   7716
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11760
   LinkTopic       =   "Form3"
   ScaleHeight     =   7716
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmAdr 
      Height          =   288
      Left            =   3480
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   3372
   End
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   6132
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   11292
      _cx             =   19918
      _cy             =   10816
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
   Begin VB.Label L1 
      Caption         =   "Àäðåñ"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   732
   End
End
Attribute VB_Name = "sc_analiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comboRs As ADODB.Recordset
Private Sub Form_Load()
CmAdr.Text = "*"
Set comboRs = New ADODB.Recordset
comboRs.Open ("SELECT KLS_PODR.ÊÎÄ, KLS_PODR.NAIM_KLS From KLS_PODR ORDER BY KLS_PODR.ÊÎÄ"), Mconn

comboRs.MoveFirst
Do While Not comboRs.EOF
'CmAdr.AddItem = CStr(comboRs("Êîä")) & "  " & comboRs("NAIM_KLS")
comboRs.MoveNext
Loop




End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True

End Sub
