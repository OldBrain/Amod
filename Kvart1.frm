VERSION 5.00
Begin VB.Form Kvart1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "Данные о квартире"
   ClientHeight    =   7536
   ClientLeft      =   60
   ClientTop       =   816
   ClientWidth     =   10488
   ControlBox      =   0   'False
   FillColor       =   &H00400000&
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Kvart1.frx":0000
   ScaleHeight     =   7536
   ScaleWidth      =   10488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "..........."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   39
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4800
      TabIndex        =   37
      Text            =   "Док.на квартиру "
      Top             =   6120
      Width           =   5655
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   35
      Text            =   "Дата рожд."
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   33
      Text            =   "Паспортные данные"
      Top             =   5640
      Width           =   6375
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   32
      Text            =   "Номер телефона"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      TabIndex        =   30
      Text            =   "Номер в соцзащите"
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2040
      TabIndex        =   28
      Text            =   "Документ на льготу"
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
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
      Left            =   9000
      TabIndex        =   25
      Text            =   "Кол-во"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
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
      Left            =   4320
      TabIndex        =   23
      Text            =   "Кол-во"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   21
      Text            =   "Кол-во"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
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
      Left            =   9240
      TabIndex        =   19
      Text            =   "Этаж"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
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
      Left            =   6600
      TabIndex        =   17
      Text            =   "Жил.плдщ."
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
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
      Left            =   3000
      TabIndex        =   14
      Text            =   "Общ.пл."
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7800
      TabIndex        =   11
      Text            =   "Combo3"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7080
      TabIndex        =   9
      Text            =   "№кв."
      Top             =   1440
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Kvart1.frx":4DF7D
      Left            =   720
      List            =   "Kvart1.frx":4DF7F
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   1440
      Width           =   5655
   End
   Begin VB.TextBox Text4 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6960
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   3
      Text            =   "О"
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3480
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   2
      Text            =   "И"
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   1
      Text            =   "Ф"
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BDC6BB&
      Caption         =   "Закрыть <F12>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Line Line16 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   10560
      X2              =   10560
      Y1              =   5400
      Y2              =   6600
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   5400
      Y2              =   6600
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Док.на квартиру"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   3120
      TabIndex        =   38
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Дата рождения"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Паспорт"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   3000
      TabIndex        =   34
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Телефон"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "№ соцзащиты"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   5520
      TabIndex        =   29
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Документ на льготу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Прочие данные"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   5040
      Width           =   10455
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "в т.ч.для начисления за услуги пользования лифтом"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   5280
      TabIndex        =   24
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Этаж"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   8160
      TabIndex        =   22
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "прописано"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "проживает"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Жилая площадь"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Общая площадь"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   10560
      X2              =   10560
      Y1              =   3600
      Y2              =   5040
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3600
      Y2              =   5040
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Данные о квартире"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   3240
      Width           =   10575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Льготы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   1920
      Width           =   10455
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   10560
      X2              =   10560
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   10560
      X2              =   10560
      Y1              =   360
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   360
      Y2              =   960
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ул."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Кв.№"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Адрес"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   10455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ответственный квартиросъемщик"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   10455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Отмена 
         Caption         =   "Отмена"
         Shortcut        =   {F11}
      End
      Begin VB.Menu Выход 
         Caption         =   "Выход"
         Index           =   13
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Kvart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ComboTipKv As ADODB.Recordset
Dim rs_kat As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim Addrconn As ADODB.Recordset
Dim Combo_RS As ADODB.Recordset

Dim f1, sq, sq1, SQDOM As String
Public F As String
Public Old As Double




Private Sub Combo1_LostFocus()

Combo_RS.MoveFirst
Do While Not Combo_RS.EOF
If Combo1.Text = Combo_RS("NAME_KLS") Then rs_kat("PRIVILEGE") = Combo_RS("N_KLS")
Combo_RS.MoveNext
Loop

End Sub

Private Sub Combo1_Validate(Cancel As Boolean)


'Rs_kat.Fields("PRIVILEGE") = Combo1.ItemData
'.DataField
End Sub

Private Sub Combo2_LostFocus()

Addrconn.MoveFirst
Do While Not Addrconn.EOF
If Combo2.Text = Addrconn("NAIM_KLS") + " дом № " + Addrconn("Num") Then
rs_kat("DOM") = Addrconn("Код")
rs_kat("DOMTip") = Addrconn("Tip")
End If
Addrconn.MoveNext
Loop


End Sub



Private Sub Combo3_LostFocus()
ComboTipKv.MoveFirst
Do While Not ComboTipKv.EOF
If Combo3.Text = ComboTipKv("NAME_KV") Then rs_kat("KV") = ComboTipKv("КОД")
ComboTipKv.MoveNext
Loop
End Sub

Private Sub Command1_Click()
'Set Filter = Nothing
rs_kat.Fields("oldnum").Value = Text4.Text
rs_kat.UpdateBatch

'Прочие
If Lic.ops <> 1 Then
Mconn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.Propis = [MainOccupant]![NLODGERF], Adding.Projiv = [MainOccupant]![NLODGER], Adding.ProLift = [MainOccupant]![NLODLIFT], Adding.ObPl = [MainOccupant]![COMSPACE], Adding.PolPl = [MainOccupant]![HABSPACE], Adding.TipKvKod = [MainOccupant]![KV], Adding.TipDomKod = [MainOccupant]![DomTip]" + f1)
Mconn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv=MainOccupant.Numer SET Adding.Propis = MainOccupant!NLODGERF, Adding.Projiv = MainOccupant!NLODGER, Adding.ProLift = MainOccupant!NLODLIFT, Adding.ObPl = MainOccupant!COMSPACE, Adding.PolPl = MainOccupant!HABSPACE, Adding.TipKvKod = MainOccupant!KV, Adding.TipDomKod = MainOccupant!DomTip " + f1)
MainForm.ЗапЛьгот
End If



Unload Me
If Lic.ops = 1 Then
Lic.Show
'Lic.FG1.Refresh
Else
Filter.Show
End If
'Kvart.Hide
End Sub



Private Sub Command2_Click()
Unload Me
Filter.Show
Kvart.Hide
End Sub

Private Sub Command3_Click()
    DropForm2.Show
    DropForm3.Show
    DropForm3.Move DropForm2.Width + 1, (DropForm2.Height - DropForm3.Height) / 2
   'OtheOwner.othe = 0
   
    'DropForm3.Move DropForm2.Left + DropForm1.Width + 500, DropForm1.Top + DropForm2.Height
'Viblgot.Show
'VBLGOT.Show
End Sub

Private Sub Form_Activate()
Form_Load
End Sub

Private Sub Form_Initialize()
If F <> "" Then Kvart.Hide
End Sub

Private Sub Form_Load()
' open connection
 ' Set mconn = New ADODB.Connection
  'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'mconn.Open "data/Kvartplata.mdb"
  ' Рекордсет для сетки
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
'f = Filter.nm
If Filter.ad <> 1 Then F = Filter.Nm
'MsgBox (F)
If F <> "" Then f1 = "WHERE (((MainOccupant.Numer)=" & F & "))"
f2 = "WHERE (((OtheOwner.Numer)=" & F & "))"
f3 = "SELECT MainOccupant.Numer, MainOccupant.Dom, MainOccupant.KV, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.NLODLIFT"
f3 = f3 + ", MainOccupant.NLODGER, MainOccupant.NLODGERF, MainOccupant.NROOM, MainOccupant.COMSPACE, MainOccupant.HABSPACE, MainOccupant.PRIVILEGE, MainOccupant.HABITATE, MainOccupant.BIRTHDAY, MainOccupant.NORDER, MainOccupant.KITCHSPACE, MainOccupant.BATHSPACE, MainOccupant.CORRSPACE, MainOccupant.TOILSPACE, MainOccupant.BALCSPACE, MainOccupant.NFAMILY, MainOccupant.DATARECEIV, MainOccupant.PASSPORT, MainOccupant.TELEPHONE, MainOccupant.LDOK, MainOccupant.LDATEBEG, MainOccupant.LDATEEND, MainOccupant.NAPARTMENT, MainOccupant.FLOOR, MainOccupant.SocNum, MainOccupant.COMM, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim, TipKv.Name_Kv, MainOccupant.KV_num, MainOccupant.DomTip,MainOccupant.OLDNUM "
sq = f3 & "FROM (MainOccupant LEFT JOIN KLS_PODR ON MainOccupant.Dom = KLS_PODR.КОД) LEFT JOIN TipKv ON MainOccupant.KV = TipKv.Код " & f1
' Рекордсет для подписи (адреса)
Set Addrconn = New ADODB.Recordset
Set Addrconn.ActiveConnection = Mconn






rs_kat.CursorType = adOpenDynamic
rs_kat.LockType = adLockBatchOptimistic

' Рекордсет для падающего списка лгот
Set Combo_RS = New ADODB.Recordset
Set Combo_RS.ActiveConnection = Mconn

Set ComboTipKv = New ADODB.Recordset
Set ComboTipKv.ActiveConnection = Mconn
'Combo_RS.CursorType = adOpenForwardOnly
'Combo_RS.LockType = adLockBatchOptimistic

Combo_RS.Open ("SELECT KLS_PRIV.N_KLS, KLS_PRIV.NAME_KLS FROM KLS_PRIV")
ComboTipKv.Open ("TipKV")

rs_kat.Open (sq)


'If Filter.ad = 1 Then Command2.Enabled = False



'MsgBox (Rs_kat.Fields("FAM").Value)

Addrconn.Open ("KLS_PODR")
Показать

'Combo1.ItemData (2)




End Sub
Sub Показать()
On Error Resume Next
If rs_kat.Fields("FAM").Value <> "" Then Text1.Text = rs_kat.Fields("FAM").Value
If rs_kat.Fields("IM").Value <> "" Then Text2.Text = rs_kat.Fields("IM").Value
If rs_kat.Fields("OT").Value <> "" Then Text3.Text = rs_kat.Fields("OT").Value
If rs_kat.Fields("OLDNUM").Value <> "" Then Text4.Text = rs_kat.Fields("OLDNUM").Value
If rs_kat.Fields("kv_num").Value <> "" Then Text5.Text = rs_kat.Fields("kv_num").Value

If rs_kat.Fields("ComSpace").Value <> "" Then Text6.Text = rs_kat.Fields("ComSpace").Value
If rs_kat.Fields("HabSpace").Value <> "" Then Text7.Text = rs_kat.Fields("HabSpace").Value
If rs_kat.Fields("Floor").Value <> "" Then Text8.Text = rs_kat.Fields("Floor").Value
If rs_kat.Fields("NLODGER").Value <> "" Then Text9.Text = rs_kat.Fields("NLODGER").Value
If rs_kat.Fields("NLODGERF").Value <> "" Then Text10.Text = rs_kat.Fields("NLODGERF").Value
If rs_kat.Fields("NLODLIFT").Value <> "" Then Text11.Text = rs_kat.Fields("NLODLIFT").Value
If rs_kat.Fields("LDOK").Value <> "" Then Text12.Text = rs_kat.Fields("LDOK").Value
If rs_kat.Fields("SocNum").Value <> "" Then Text13.Text = rs_kat.Fields("SocNum").Value

'If Rs_kat.Fields("LDateBeg").Value <> "" Then TextD1.Text = Rs_kat.Fields("LDateBeg").Value
'If Rs_kat.Fields("LDateEnd").Value <> "" Then Textd2.Text = Rs_kat.Fields("LDateEnd").Value
If rs_kat.Fields("Telephone").Value <> "" Then Text14.Text = rs_kat.Fields("Telephone").Value
If rs_kat.Fields("Passport").Value <> "" Then Text15.Text = rs_kat.Fields("Passport").Value
If rs_kat.Fields("BIRTHDAY").Value <> "" Then Text16.Text = rs_kat.Fields("BIRTHDAY").Value
If rs_kat.Fields("norder").Value <> "" Then Text17.Text = rs_kat.Fields("norder").Value

'Combo_RS.MoveFirst
'Do While Not Combo_RS.EOF
'If Combo_RS.Fields("N_KLS").Value = Rs_kat.Fields("PRIVILEGE").Value Then Combo1.Text = Combo_RS.Fields("NAME_KLS").Value
'Combo_RS.MoveNext
'Loop

'Combo_RS.MoveFirst
'Do While Not Combo_RS.EOF
'Combo1.AddItem Combo_RS("NAME_KLS")
'Combo_RS.MoveNext
'Loop



'Combo_RS.MoveFirst
'Do While Not Combo_RS.EOF
'If Combo_RS.Fields("N_KLS").Value = Rs_kat.Fields("PRIVILEGE").Value Then Combo1.Text = Combo_RS.Fields("NAME_KLS").Value
'Combo_RS.MoveNext
'Loop

If Not rs_kat("Num") Then Combo2.Text = rs_kat.Fields("NAIM_KLS") + " дом № " + rs_kat("Num")
Addrconn.MoveFirst
Do While Not Addrconn.EOF
Combo2.AddItem Addrconn("NAIM_KLS") + " дом № " + Addrconn("Num")
Addrconn.MoveNext
Loop


Combo3.Text = rs_kat.Fields("NAME_KV")
ComboTipKv.MoveFirst
Do While Not ComboTipKv.EOF
Combo3.AddItem ComboTipKv("NAME_KV")
ComboTipKv.MoveNext
Loop

End Sub

Private Sub Text1_Change()
'Text1.Text = Rs_kat.Fields("Numer").Value
End Sub

Private Sub Text1_LostFocus()
rs_kat.Fields("FAM").Value = Text1.Text
End Sub

Private Sub Text2_LostFocus()
rs_kat.Fields("IM").Value = Text2.Text
End Sub
Private Sub Text3_LostFocus()
rs_kat.Fields("OT").Value = Text3.Text
End Sub

Private Sub Text4_GotFocus()


Old = Val(Text4)
End Sub

Private Sub Text4_LostFocus()
Set Rs_Add = New ADODB.Recordset
Set Rs_Add.ActiveConnection = Mconn
 
Rs_Add.CursorType = adOpenForwardOnly
Rs_Add.LockType = adLockBatchOptimistic
Rs_Add.Open "MainOccupant"

n = 0
Rs_Add.MoveFirst
Do While Not Rs_Add.EOF
N1 = Rs_Add.Fields("oldnum").Value
If Rs_Add.Fields("oldnum").Value = Text4.Text And Rs_Add.Fields("oldnum").Value <> Old Then
MsgBox ("Такой номер уже имеется! Введите другой")
Kvart.Text4 = Old
Kvart.Text4.Enabled = True
Kvart.Text4.SetFocus
Exit Sub
End If
Rs_Add.MoveNext
Loop



rs_kat.Fields("OLDNUM").Value = Text4.Text
End Sub

Private Sub Выход_Click(Index As Integer)
Command1_Click
End Sub
Private Sub Text5_LostFocus()
rs_kat.Fields("kv_num").Value = Text5.Text
End Sub
Private Sub Text6_LostFocus()
rs_kat.Fields("ComSpace").Value = Text6.Text
If Lic.ops = 1 Then
Lic.fg1.TextMatrix(Lic.fg1.Row, 15) = Text6.Text
Lic.fg1.Refresh
End If
End Sub
Private Sub Text7_LostFocus()
rs_kat.Fields("HabSpace").Value = Text7.Text
End Sub
Private Sub Text8_LostFocus()
rs_kat.Fields("Floor").Value = Text8.Text
End Sub
Private Sub Text9_LostFocus()
rs_kat.Fields("NLODGER").Value = Text9.Text
End Sub
Private Sub Text10_LostFocus()
rs_kat.Fields("NLODGERF").Value = Text10.Text
End Sub
Private Sub Text11_LostFocus()
rs_kat.Fields("NLODLIFT").Value = Text11.Text
End Sub
Private Sub Text12_LostFocus()
rs_kat.Fields("LDOK").Value = Text12.Text
End Sub
Private Sub Text13_LostFocus()
rs_kat.Fields("SocNum").Value = Text13.Text
End Sub
Private Sub TextD1_LostFocus()
On Error Resume Next
rs_kat.Fields("LDateBeg").Value = TextD1.Text
End Sub
Private Sub TextD2_LostFocus()
On Error Resume Next
rs_kat.Fields("LDateEnd").Value = Textd2.Text
End Sub
Private Sub Text14_LostFocus()
rs_kat.Fields("Telephone").Value = Text14.Text
End Sub
Private Sub Text15_LostFocus()
rs_kat.Fields("Passport").Value = Text15.Text
End Sub
Private Sub Text16_LostFocus()
On Error Resume Next
rs_kat.Fields("BIRTHDAY").Value = Text16.Text
End Sub
Private Sub Text17_LostFocus()
rs_kat.Fields("norder").Value = Text17.Text
End Sub
