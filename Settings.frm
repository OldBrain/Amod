VERSION 5.00
Begin VB.Form Settings 
   BackColor       =   &H00BDC6BB&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8655
   ClientLeft      =   2730
   ClientTop       =   3330
   ClientWidth     =   12225
   ControlBox      =   0   'False
   ForeColor       =   &H00404000&
   HasDC           =   0   'False
   Icon            =   "Settings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Реквизиты"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   1320
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   2160
      Width           =   10815
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "При импорте оплаты из банка работать только по 12-значным  номерам л/сч"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   32
      Top             =   600
      Width           =   3015
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Показывать сведения о договорах "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   31
      Top             =   8400
      Width           =   3855
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d MMMM yyyy ""г."""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Код Неопознанных сумм:"
      Top             =   8160
      Width           =   2775
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3000
      TabIndex        =   29
      Text            =   "Text11"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   10080
      TabIndex        =   28
      Text            =   "Combo2"
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   2640
      TabIndex        =   27
      Text            =   "Combo1"
      Top             =   720
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Показывать уникальные номера л/сч"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   24
      Top             =   8040
      Width           =   3855
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8040
      TabIndex        =   22
      Text            =   "Text10"
      Top             =   7560
      Width           =   3975
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d MMMM yyyy ""г."""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      HideSelection   =   0   'False
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text8"
      Top             =   7920
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddddd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      HideSelection   =   0   'False
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   6720
      Width           =   8535
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   6240
      Width           =   10095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   5160
      Width           =   8655
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4560
      Width           =   10095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3480
      Width           =   8655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2880
      Width           =   10095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Width           =   12015
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00BDC6BB&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   1815
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
      Picture         =   "Settings.frx":030A
      Top             =   0
      Visible         =   0   'False
      Width           =   195
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
      Left            =   240
      TabIndex        =   36
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   11850
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   1680
      Picture         =   "Settings.frx":0554
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   960
      Picture         =   "Settings.frx":0C9E
      Stretch         =   -1  'True
      ToolTipText     =   "Двойной щелчек мышы развернет форму во весь экран или вернет в исходное состояние"
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   120
      Picture         =   "Settings.frx":13E8
      Top             =   240
      Width           =   285
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Адрес:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ЖЭК №:"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dddddd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8040
      TabIndex        =   26
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Код района:"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dddddd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Путь к архивным копиям"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   23
      Top             =   7560
      Width           =   3375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Расчет ведется начиная с"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   7560
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Период расчета:"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dddddd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   808
      X2              =   808
      Y1              =   368
      Y2              =   456
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   8
      X2              =   808
      Y1              =   488
      Y2              =   488
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   8
      X2              =   8
      Y1              =   400
      Y2              =   488
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Фамилия Имя Отчество"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Должность"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   504
      X2              =   808
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   320
      X2              =   8
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ответственное лицо"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   808
      X2              =   808
      Y1              =   256
      Y2              =   352
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   576
      X2              =   808
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   8
      X2              =   8
      Y1              =   288
      Y2              =   384
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   272
      X2              =   8
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   8
      X2              =   808
      Y1              =   384
      Y2              =   384
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Руководитель финансовой службы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   4200
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Фамилия Имя Отчество"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Должность"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   808
      X2              =   808
      Y1              =   144
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   504
      X2              =   808
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   8
      X2              =   808
      Y1              =   272
      Y2              =   272
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   8
      X2              =   8
      Y1              =   176
      Y2              =   272
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   256
      X2              =   8
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Руководитель предприятия"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   11055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Фамилия Имя Отчество"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Должность"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Наименование предприятия"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   7335
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KodJak As String
Option Explicit
Dim Rs_kat As ADODB.Recordset
Dim RsRay As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim I As Integer


Private Sub CancelButton_Click()
Settings.Hide
End Sub


Private Sub Combo1_Click()
Rs_kat("Ray") = Right(Combo1.Text, 2)
Rs_kat.UpdateBatch
End Sub

Private Sub Combo2_Click()
Rs_kat("Jak") = Combo2.Text
Rs_kat.UpdateBatch
End Sub


Private Sub Command1_Click()
Rekvizit.Show 1
End Sub

Private Sub Form_Load()
MenuNastr.Hide
MakeWindow Me, False
lblTitle = "Реквизиты предприятия"
Check1.BackColor = RGB(207, 207, 207)
Check2.BackColor = RGB(207, 207, 207)
Check3.BackColor = RGB(207, 207, 207)
Text9.BackColor = RGB(207, 207, 207)
Text12.BackColor = RGB(207, 207, 207)
Text8.BackColor = RGB(207, 207, 207)

For I = 1 To 99
KodJak = Trim(Str(I))
If Len(KodJak) = 1 Then KodJak = "0" + KodJak
Combo2.AddItem KodJak
Next I


' open connection
' Set mconn = New ADODB.Connection
 ' mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
 ' mconn.Open "data/Kvartplata.mdb"
    
Set Rs_kat = New ADODB.Recordset
Set RsRay = New ADODB.Recordset
Set Rs_kat.ActiveConnection = Mconn
 
Rs_kat.CursorType = adOpenForwardOnly
Rs_kat.LockType = adLockBatchOptimistic
Rs_kat.Open ("Settings")
RsRay.Open ("AstrRay"), Mconn

RsRay.MoveFirst
Combo2.Text = Rs_kat("Jak")
Combo1.Text = Rs_kat("Ray")
Do While Not RsRay.EOF
Combo1.AddItem RsRay("Name") + " " + RsRay("KodAstrRay")
RsRay.MoveNext
Loop



Rs_kat.MoveFirst
Set Text1.DataSource = Rs_kat
Text1.DataField = "NamePred"

Set Text2.DataSource = Rs_kat
Text2.DataField = "DolgnRuk"

Set Text3.DataSource = Rs_kat
Text3.DataField = "FIORuk"

Set Text4.DataSource = Rs_kat
Text4.DataField = "DolgnFin"

Set Text5.DataSource = Rs_kat
Text5.DataField = "FIOFin"

Set Text6.DataSource = Rs_kat
Text6.DataField = "DolgnOtv"

Set Text7.DataSource = Rs_kat
Text7.DataField = "FIOOtv"

Set Text8.DataSource = Rs_kat
Text8.DataField = "TekData"

Set Text9.DataSource = Rs_kat
Text9.DataField = "BeginData"

Set Text10.DataSource = Rs_kat
Text10.DataField = "Arhiv"
'If Rs_kat.Fields("pokaz").Value = 1 Then Check1.DataChanged = True Else Check1.DataChanged = False
'Text1.DataField = "NamePred"

Set Text11.DataSource = Rs_kat
Text11.DataField = "Neo"

Set Text13.DataSource = Rs_kat
Text13.DataField = "Adres"


Check1.Value = Rs_kat.Fields("pokaz").Value
Check2.Value = Rs_kat.Fields("Dogovor").Value
Check3.Value = Rs_kat.Fields("Bank12").Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rs_kat.Close
MenuNastr.Show

'Mconn.Close
'MainMenu.Enabled = True
End Sub

Private Sub OKButton_Click()
'MsgBox (Text1.Text)
Rs_kat.Fields("pokaz").Value = Check1.Value
Rs_kat.Fields("Dogovor").Value = Check2.Value
Rs_kat.Fields("Bank12").Value = Check3.Value

MainForm.Bank12 = Check3.Value

MainForm.Dog = Check2.Value
Rs_kat.UpdateBatch
Mconn.Execute ("UPDATE MainOccupant, Settings SET MainOccupant.Jak = [Settings]![Jak], MainOccupant.Ray = [Settings]![Ray]")


'Rs_kat.Close
'mconn.Close
Settings.Hide
Unload Me

End Sub

