VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Schet1 
   Caption         =   "Справочник статей затрат(план)"
   ClientHeight    =   5952
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11772
   LinkTopic       =   "Form7"
   ScaleHeight     =   5952
   ScaleWidth      =   11772
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Статьи затрат(ФАКТ)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5052
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   252
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11772
      _ExtentX        =   20765
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
            Key             =   "OOFL1"
            ImageKey        =   "OOFL1"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   5172
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   11172
      _cx             =   19706
      _cy             =   9123
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
      FormatString    =   $"Schet.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   1
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
      DataMode        =   1
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
      Left            =   3930
      Top             =   2730
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Schet.frx":0128
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Schet.frx":023A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Schet.frx":034C
            Key             =   "OOFL1"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Schet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_kat, Prov As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim PR As String


Private Sub Command1_Click()
Kzatrat.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)

    'On Error Resume Next
    Select Case Button.KEY
        Case "New"
        
        Dim n, N1 As Integer
If MsgBox("Добавить новую запись?", vbYesNo) = vbYes Then




rs_kat.AddNew
'rs_kat("Schet") = n + 1
rs_kat("vid2") = "Введите вид работ"
rs_kat.UpdateBatch
FG.DataRefresh
rs_kat.MoveLast
End If

     
       Case "Delete"
            DelItem = FG.TextMatrix(FG.Row, 1)
            
            
' Удаляем из zatr_tarif


If FG.TextMatrix(FG.Row, 5) = 1 Then
MsgBox ("Данный вид работ добавлен на основании перечней обязательных и дополнительных работ и услуг по содержанию и ремонту общего имущества собственников помещений в многоквартирном доме, предусмотренных Правилами проведения органом местного самоуправления открытого конкурса по отбору управляющей организации для управления многоквартирным домом, утвержденными постановлением Правительства РФ №75 от 06.02.2006 г. УДАЛЯТЬ НЕЛЬЗЯ! ")
Exit Sub
End If
 
 
With rs_kat
'DelItem = FG.TextMatrix(FG.Row, 1)

If MsgBox("Вы хотите удалить счет " + FG.TextMatrix(FG.Row, 1) + "  " + FG.TextMatrix(FG.Row, 2) + "?", vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''
.MoveFirst
Do While Not .EOF
If rs_kat("Код") = DelItem Then .Delete
If .EOF = False Then .MoveNext Else .MoveLast
Loop

Zconn.Execute ("DELETE Perecen.Код, Perecen.* From Perecen WHERE (((trim(Perecen.код))= " + Trim(FG.TextMatrix(FG.Row, 1)) + "))")

.UpdateBatch
FG.DataRefresh
If .EOF Then .MoveLast

End If
End With
        Case "OOFL1"
            Unload Me
    End Select
    
End Sub


Private Sub Form_Load()
КоннектЗ
'Set mconn = New ADODB.Connection

 
  'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
     
  Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Zconn
rs_kat.CursorType = adOpenForwardOnly
rs_kat.LockType = adLockBatchOptimistic
rs_kat.Open "SELECT Perecen.Код, Perecen.vid1, Perecen.vid2, Perecen.NameR, Perecen.sys FROM Perecen"

FG.Editable = flexEDKbdMouse
FG.DataMode = flexDMBoundImmediate



FG.Sort = flexSortUseColSort

'FG.HighLight = flexHighlightWithFocus

Set FG.DataSource = rs_kat



End Sub

Private Sub Form_Unload(Cancel As Integer)
rs_kat.UpdateBatch


MainForm.zatr_tarif
Sprav.Enabled = True

End Sub


