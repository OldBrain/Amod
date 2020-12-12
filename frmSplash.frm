VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H80000011&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4452
   ClientLeft      =   276
   ClientTop       =   1428
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   6960
      Picture         =   "frmSplash.frx":2CE14
      ScaleHeight     =   204
      ScaleWidth      =   204
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   7200
      Picture         =   "frmSplash.frx":2D1D5
      ScaleHeight     =   204
      ScaleWidth      =   204
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   7215
      _ExtentX        =   12721
      _ExtentY        =   445
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   5000
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   7095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright, 2005, Астрахань, Бугоров Андрей Владимирович. Все права сохранены"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Квартплата +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ПОЖАЛУЙСТА ПОДОЖДИТЕ, ИДЕТ ЗАГРУЗКА ПРОГРАММЫ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   3120
      Width           =   5535
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7335
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim x As Double

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

Me.Label1.Caption = "Расчет и анализ коммунальных платежей" + vbNewLine + "Консультации по тел: +79881733600" + vbNewLine + "E-Mail: bestonline@list.ru"


    lblVersion.Caption = "Версия " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    
 'Me.Picture3.TextHeight "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
Main
    
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub



 Public Sub Main()
 
  
    frmSplash.Show
    frmSplash.Refresh
    
    
    
 ' ВВОД ПАРОЛЯ
    
 ' Проверяем надо ли нам это
    
    
    
    
    Load MainForm
     
     
  
   
    Unload frmSplash
    MainForm.Show
End Sub

