VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form zatr_tarif 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Справочник счетов затрат"
   ClientHeight    =   5952
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   6240
   LinkTopic       =   "Form7"
   ScaleHeight     =   5952
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
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
      Height          =   372
      Left            =   5160
      MaskColor       =   &H0080FF80&
      TabIndex        =   9
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Категории   затрат"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4332
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   252
   End
   Begin VSFlex8Ctl.VSFlexGrid FG 
      Height          =   4452
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   5772
      _cx             =   10181
      _cy             =   7853
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
      FormatString    =   $"zatr_tarif.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
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
      Editable        =   0
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
            Picture         =   "zatr_tarif.frx":0100
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "zatr_tarif.frx":0212
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "zatr_tarif.frx":0324
            Key             =   "OOFL1"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4920
      TabIndex        =   8
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4080
      TabIndex        =   7
      Top             =   720
      Width           =   732
   End
   Begin VB.Label Label5 
      Caption         =   "руб. из"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3120
      TabIndex        =   6
      Top             =   720
      Width           =   852
   End
   Begin VB.Label Label4 
      Caption         =   "Распределено"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1812
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   972
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "123 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   348
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9132
   End
End
Attribute VB_Name = "zatr_tarif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Double
Dim Sp As Double
Public zatr_tarif As ADODB.Recordset

'Dim mconn As ADODB.Connection
Dim PR As String


Private Sub Command1_Click()
Schet1.Show (1)
End Sub




Private Sub Command2_Click()
Unload Me
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Fg.TextMatrix(Row, Col) = "" Then Fg.TextMatrix(Row, Col) = 0
s = 0
Sp = 0
For i = 1 To Fg.Rows - 1
'MsgBox ((FG.TextMatrix(i, 7)))
s = s + (Fg.TextMatrix(i, 7))


Fg.TextMatrix(i, 8) = Round(((Fg.TextMatrix(i, 7)) / Tarif.Tar) * 100, 4)
Sp = Sp + Fg.TextMatrix(i, 8)
Me.Label7.Caption = Str(Round(Sp, 2)) + " %"
Next i
Label3.Caption = s
'MsgBox (CSng(Label3.Caption))
If CSng(Label3.Caption) > CSng(Me.Label6.Caption) Then
MsgBox ("!!ВНИМАНИЕ!! Ошибка. Привышен тариф " + (Label3.Caption) + " > " + (Me.Label6.Caption))
Me.Label3.BackColor = &HFF&
Else
Me.Label3.BackColor = &H8000000F
End If
If Round(Sp, 2) = 100 Then
Me.BackColor = &H8000&
Me.Label1.BackColor = &H8000&
Me.Label2.BackColor = &H8000&
Me.Label3.BackColor = &H8000&
Me.Label4.BackColor = &H8000&
Me.Label5.BackColor = &H8000&
Me.Label6.BackColor = &H8000&
Me.Label7.BackColor = &H8000&
Me.Command2.Visible = True
End If
If Round(Sp, 2) < 100 Then
Me.BackColor = &H8000000F
Me.Label1.BackColor = &H8000000F
Me.Label2.BackColor = &H8000000F
Me.Label3.BackColor = &H8000000F
Me.Label4.BackColor = &H8000000F
Me.Label5.BackColor = &H8000000F
Me.Label6.BackColor = &H8000000F
Me.Label7.BackColor = &H8000000F
Me.Command2.Visible = False
End If
If Round(Sp, 2) > 100 Then
Me.BackColor = &HFF&
Me.Label1.BackColor = &HFF&
Me.Label2.BackColor = &HFF&
Me.Label3.BackColor = &HFF&
Me.Label4.BackColor = &HFF&
Me.Label5.BackColor = &HFF&
Me.Label6.BackColor = &HFF&
Me.Label7.BackColor = &HFF&
Me.Command2.Visible = False
End If


End Sub

Private Sub Form_Load()
Me.Label1.Caption = "Пожалуйста проставьте долю тарифа каждой категории затрат в рублях."
'Set mconn = New ADODB.Connection
Me.Command2.Visible = False
 
  'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
     
  Set zatr_tarif = New ADODB.Recordset
Set zatr_tarif.ActiveConnection = Mconn
zatr_tarif.CursorType = adOpenForwardOnly
zatr_tarif.LockType = adLockBatchOptimistic

 
Mconn.Execute ("UPDATE Tarif INNER JOIN zatr_tarif ON Tarif.Код = zatr_tarif.kod_tar SET zatr_tarif.tarif = [Tarif]![Value]")
Mconn.Execute ("UPDATE zatr_tarif INNER JOIN Schet ON zatr_tarif.Schet = Schet.Schet SET zatr_tarif.Schet_Name = [Schet]![Schet_Name]")

zatr_tarif.Open "SELECT zatr_tarif.key, zatr_tarif.kod_tar, zatr_tarif.num_tar, zatr_tarif.tarif, zatr_tarif.Schet, zatr_tarif.Schet_Name, zatr_tarif.Summa, zatr_tarif.Procent From zatr_tarif Where (((zatr_tarif.kod_tar) = " + Tarif.kod_tar + ")) ORDER BY zatr_tarif.Schet"

Fg.Editable = flexEDKbdMouse
Fg.DataMode = flexDMBoundImmediate



Fg.Sort = flexSortUseColSort
Fg.AutoResize = False

Set Fg.DataSource = zatr_tarif
Me.Label6.Caption = Tarif.Tar
s = 0
For i = 1 To Fg.Rows - 1
s = s + Fg.TextMatrix(i, 7)
Next i
Label3.Caption = s

s = 0
Sp = 0
For i = 1 To Fg.Rows - 1
'MsgBox ((FG.TextMatrix(i, 7)))
s = s + (Fg.TextMatrix(i, 7))


Fg.TextMatrix(i, 8) = Round(((Fg.TextMatrix(i, 7)) / Tarif.Tar) * 100, 4)
Sp = Sp + Fg.TextMatrix(i, 8)
Me.Label7.Caption = Str(Round(Sp, 2)) + " %"
Next i
Label3.Caption = s
'MsgBox (CSng(Label3.Caption))
If CSng(Label3.Caption) > CSng(Me.Label6.Caption) Then
MsgBox ("!!ВНИМАНИЕ!! Ошибка. Привышен тариф " + (Label3.Caption) + " > " + (Me.Label6.Caption))
Me.Label3.BackColor = &HFF&
Else
Me.Label3.BackColor = &H8000000F
End If
 
If Round(Sp, 2) = 100 Then
Me.BackColor = &H8000&
Me.Label1.BackColor = &H8000&
Me.Label2.BackColor = &H8000&
Me.Label3.BackColor = &H8000&
Me.Label4.BackColor = &H8000&
Me.Label5.BackColor = &H8000&
Me.Label6.BackColor = &H8000&
Me.Label7.BackColor = &H8000&
Me.Command2.Visible = True
End If
If Round(Sp, 2) < 100 Then
Me.BackColor = &H8000000F
Me.Label1.BackColor = &H8000000F
Me.Label2.BackColor = &H8000000F
Me.Label3.BackColor = &H8000000F
Me.Label4.BackColor = &H8000000F
Me.Label5.BackColor = &H8000000F
Me.Label6.BackColor = &H8000000F
Me.Label7.BackColor = &H8000000F
Me.Command2.Visible = False
End If
If Round(Sp, 2) > 100 Then
Me.BackColor = &HFF&
Me.Label1.BackColor = &HFF&
Me.Label2.BackColor = &HFF&
Me.Label3.BackColor = &HFF&
Me.Label4.BackColor = &HFF&
Me.Label5.BackColor = &HFF&
Me.Label6.BackColor = &HFF&
Me.Label7.BackColor = &HFF&
Me.Command2.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
zatr_tarif.UpdateBatch
Tarif.Enabled = True


End Sub


