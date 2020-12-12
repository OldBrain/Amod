VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Cal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Выбор периода"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _Version        =   524288
      _ExtentX        =   7858
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2004
      Month           =   6
      Day             =   10
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Cyr"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Cyr"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Cyr"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Cal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Calendar1_AfterUpdate()
Calendar1.Day = "01"
MainForm.TMP1 = Calendar1.Value
End Sub

Private Sub Calendar1_KeyPress(KeyAscii As Integer)
Calendar1.Day = "01"
MainForm.TMP1 = Calendar1.Value
End Sub

Private Sub Calendar1_NewMonth()
Calendar1.Day = "01"
'Calendar1.SetFocus
Calendar1.ShowDateSelectors = True
MainForm.TMP1 = Calendar1.Value
End Sub

Private Sub Calendar1_NewYear()
Calendar1.Day = "01"
MainForm.TMP1 = Calendar1.Value
End Sub

Private Sub CancelButton_Click()
MainForm.TMP1 = Calendar1.Value
Cal.Hide
End Sub

Private Sub Form_Load()
Calendar1.Day = "01"
End Sub

Private Sub OKButton_Click()
Calendar1.Day = "01"
Socmin.FG1.Cell(flexcpText) = Calendar1.Value
'MainForm.TMP1 = Calendar1.Value
Cal.Hide
End Sub

