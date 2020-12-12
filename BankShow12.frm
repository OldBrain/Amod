VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form BankShow12 
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9096
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   11868
   ControlBox      =   0   'False
   Icon            =   "BankShow12.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   758
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   989
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Все не разнесенные"
            Key             =   "Key1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "В том числе, номера л/сч другого подразделения"
            Key             =   "Key2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "В том числе, не указаны номера л/сч"
            Key             =   "Key3"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00BDC6BB&
      Caption         =   "XL"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton Image1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Ок"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   5775
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   11535
      _cx             =   20346
      _cy             =   10186
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"BankShow12.frx":030A
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
      ExplorerBar     =   1
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
      BackColorFrozen =   255
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   11055
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
      Height          =   156
      Left            =   0
      Picture         =   "BankShow12.frx":03F3
      Top             =   0
      Width           =   156
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resizable Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   10890
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   0
      Picture         =   "BankShow12.frx":063D
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   360
      Picture         =   "BankShow12.frx":0D87
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   0
      Picture         =   "BankShow12.frx":14D1
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   360
      Width           =   285
   End
End
Attribute VB_Name = "BankShow12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim S1 As Double
Dim S2 As Double
Dim S3 As Double

Private Sub BtnEnh11_Click()

End Sub

Private Sub Command3_Click()
ВыводВExel
End Sub



Private Sub Form_Load()

Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set rs3 = New ADODB.Recordset


rs1.Open ("SELECT Mid([Bank]![NewNum],9,2) AS Район, Mid([Bank]![NewNum],7,2) AS [ЖЭК N], Bank.NewNum, Bank.ADR, Bank.FIO, Bank.SUMMA, Bank.PLDATE, Bank.LSCHET From Bank Where (((Mid([Bank]![NewNum], 7, 2)) <> " + MainForm.Jak + ") And ((Bank.NewNum) <> '0')) ORDER BY Bank.ADR"), Mconn, adOpenKeyset, adLockPessimistic

If rs1.RecordCount > 0 Then
rs1.MoveFirst
Do While Not rs1.EOF
S1 = S1 + rs1("Summa")
rs1.MoveNext
Loop
End If

rs2.Open ("SELECT Bank.NewNum, Bank.LSCHET, Bank.ADR, Bank.FIO, Bank.SUMMA, Bank.PLDATE From Bank Where (((Bank.NewNum) = '0')) ORDER BY Bank.ADR"), Mconn, adOpenKeyset, adLockPessimistic

If rs2.RecordCount > 0 Then
rs2.MoveFirst
Do While Not rs2.EOF
S2 = S2 + rs2("Summa")
rs2.MoveNext
Loop
End If

rs3.Open ("SELECT Bank.LSCHET, Bank.NewNum, Bank.ADR, Bank.FIO, Bank.SUMMA, Bank.PLDATE FROM Bank LEFT JOIN MainOccupant ON Bank.NewNum = MainOccupant.BanKN Where (((MainOccupant.BanKN) Is Null)) ORDER BY Bank.ADR"), Mconn, adOpenKeyset, adLockPessimistic

If rs3.RecordCount > 0 Then
rs3.MoveFirst
Do While Not rs3.EOF
S3 = S3 + rs3("Summa")
rs3.MoveNext
Loop
End If

MakeWindow Me, True
fg1.Width = Me.Width / 15.40107
fg1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top

'**************************************************************
'****************************************************************

lblTitle = "Импорт оплаты из банка. Файл > " + BankImport.File1.FileName
'Label1.Caption = "Просмотр файла >" + BankImport.File1.FileName + ". Для продолжения нажмите <<Далее>>"

Label1.Caption = TabStrip1.SelectedItem + " на сумму >" + Str(S3)

lblTitle.Caption = lblTitle.Caption + "Общая суммареестра  > " + Str(BankShow.SummI)
'TabStrip1.Index = 1


Set fg1.DataSource = rs3
End Sub


Private Sub Image1_Click()

'Добавляем Realdata



'Msg.Show vbModal
'Msg.Label1.Caption = "Данные добавлены в реестр документов оплаты №" + Str(BankShow.Cod) + vbNewLine + "Сумма реестра банка=" + Str(BankShow.SummI) + vbNewLine + "Разнесенная сумма=" + Str(BankShow.SummI - S3) + vbNewLine + "Отклонение=" + Str(Round(S3, 2))
'Msg.Label1.Refresh

 'Unload ReestrDoc

'ReestrDoc.Fg.Refresh
Kill BankShow.For_Dell

Unload Me

Msg.Show

Msg.Label1.Caption = "Данные добавлены в реестр документов оплаты №" + Str(BankShow.Cod) + vbNewLine + "Сумма реестра банка=" + Str(BankShow.SummI) + vbNewLine + "Разнесенная сумма=" + Str(Round(BankShow.SummI - S3, 2)) + vbNewLine + "Отклонение=" + Str(Round(S3, 2))
Msg.Label1.Refresh

Unload BankShow
Msg.Label1.Refresh


ReestrDoc.Enabled = True
End Sub
Sub ВыводВExel()
   Const НачСтрока = 1
   Dim RS As New ADODB.Recordset
   Dim ex1 As Object ' Excel.Application
   Dim wb As Object ' Excel.Workbook
   Dim ws As Object ' Excel.Worksheet
   Dim i As Long, j As Long, K As Long, rДанные As String
   Dim v As Variant
   
   Set ex1 = CreateObject("Excel.Application")  'New Excel.Application
   Set wb = ex1.Workbooks.Add
   Set ws = wb.Sheets(1)
   
   rДанные = "A" & (НачСтрока + 1) & ":" & XCol_(fg1.Cols - 1) & fg1.Rows + НачСтрока
   ReDim v(fg1.Rows, fg1.Cols) 'Забыл указать
   
   If fg1.Rows > 0 Then
            For co = 1 To fg1.Cols - 1
         For rw = 0 To fg1.Rows - 1

             v(rw, co) = fg1.TextMatrix(rw, co)
             
         Next rw
         Next co
      ex1.Visible = True   'Еще забыл
      
      ws.Range(rДанные) = v
      
 End If
End Sub

Function XCol_(ByVal Column_ As Long) As String
    If (Column_ < 0) Then Column_ = 0
    If (Column_ < 26) Then
        XCol_ = Chr(Column_ + Asc("A"))
    ElseIf (Column_ < 676) Then
        XCol_ = Chr((Column_ \ 26) + Asc("A") - 1) & Chr((Column_ Mod 26) + Asc("A"))
    Else
        XCol_ = "ZZ"
    End If
End Function


Private Sub imgTitleHelp_Click()
Form2.Label1 = "   Импорт оплаты коммунальных услуг, из файла данных предоставленных банком. Это окно предназначено для предварительного просмотра импортируемых данных." + vbNewLine + "   Кроме того, Вы можете перенести данные в XL для дальнейшей печати"
Form2.Show
End Sub

Private Sub imgTitleMain_Click()
ChangeState Me
End Sub

Private Sub lblTitle_Click()
ChangeState Me
End Sub
Private Sub Form_Resize()
fg1.Width = Me.Width / 15.40107
   fg1.Height = Me.Height / 20
Image1.Top = Me.Height / 16.16477
Image1.Left = 3
Command3.Top = Image1.Top
'BtnEnh1.Top = Image1.Top

Command3.Left = Image1.Left + Image1.Width
End Sub


Private Sub TabStrip1_Click()



If TabStrip1.SelectedItem = "В том числе, номера л/сч другого подразделения" Then
Set fg1.DataSource = rs1
Label1.Caption = TabStrip1.SelectedItem + " на сумму >" + Str(S1)
End If

If TabStrip1.SelectedItem = "В том числе, не указаны номера л/сч" Then
Set fg1.DataSource = rs2
Label1.Caption = TabStrip1.SelectedItem + " на сумму >" + Str(S2)
End If

If TabStrip1.SelectedItem = "Все не разнесенные" Then
Set fg1.DataSource = rs3
Label1.Caption = TabStrip1.SelectedItem + " на сумму >" + Str(S3)
End If

'
'SelectedItem

End Sub

