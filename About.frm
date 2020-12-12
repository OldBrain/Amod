VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4980
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   4680
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "About.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ок"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
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
      Left            =   0
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   5850
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
      Left            =   1200
      Picture         =   "About.frx":030A
      Top             =   120
      Visible         =   0   'False
      Width           =   156
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   120
      Picture         =   "About.frx":0554
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   600
      Picture         =   "About.frx":0C9E
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   360
      Picture         =   "About.frx":13E8
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Click()

'Dim AboutBox As New AboutBox
'With AboutBox
 '   .Title = "Some Application"
  '  .Version = "Version 1.2.3.4"
   ' .Company = "Some Company (R)"
    '.Copyright = "(C) Some Company 1900-2004"
 '   .Description = "The quick brown fox jumps over a lazy dog again and again"
  '  .License = "This sample is hosted at <A HREF=""http://vbrussian.com/Examples.asp?ID=100"">vbrussian.com</A>. Visit it for more info!"
   ' .hWndOwner = Me.hWnd
   ' Set .Icon = Me.Icon
   ' .AboutBox
'End With
End Sub

Private Sub Form_Load()
MakeWindow Me, False
lblTitle = "О программе"

'Label1.Caption = ""
Label1.Caption = "   <<КВАРТПЛАТА +>>" + vbNewLine + " (C) Copyright, 2005, Астрахань, Бугоров Андрей Владимирович. Консультации +79881733600." + vbNewLine + "Все права сохранены"
 

End Sub

Private Sub Label1_Click()
'Dim AboutBox As New AboutBox
'With AboutBox
 '   .Title = "Some Application"
  '  .Version = "Version 1.2.3.4"
   ' .Company = "Some Company (R)"
    '.Copyright = "(C) Some Company 1900-2004"
    '.Description = "The quick brown fox jumps over a lazy dog again and again"
'    .License = "This sample is hosted at <A HREF=""http://vbrussian.com/Examples.asp?ID=100"">vbrussian.com</A>. Visit it for more info!"
 '   .hWndOwner = Me.hWnd
  '  Set .Icon = Me.Icon
   ' .AboutBox
'End With

End Sub
