VERSION 5.00
Begin VB.Form Kvart 
   BorderStyle     =   0  'None
   ClientHeight    =   6744
   ClientLeft      =   1572
   ClientTop       =   2232
   ClientWidth     =   10656
   ControlBox      =   0   'False
   FillColor       =   &H00400000&
   ForeColor       =   &H80000017&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   888
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9720
      TabIndex        =   46
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   285
      Left            =   75
      Picture         =   "Kvart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Печать карточки квартиросъемщика (F5)"
      Top             =   90
      Width           =   555
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7320
      TabIndex        =   43
      Text            =   "Доп"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Льготы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Да"
      CausesValidation=   0   'False
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   855
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2595
      TabIndex        =   10
      Text            =   "Документ на льготу"
      Top             =   2640
      Width           =   5055
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7920
      TabIndex        =   9
      Text            =   "Номер в соцзащите"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1800
      TabIndex        =   11
      Text            =   "0,00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton BtnEnh1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Закрыть"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Text            =   "Жил.плдщ."
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9360
      TabIndex        =   13
      Text            =   "Этаж"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Нет"
      CausesValidation=   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   735
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   2160
      TabIndex        =   21
      Text            =   "Док.на квартиру "
      Top             =   5760
      Width           =   3975
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   2520
      TabIndex        =   19
      Text            =   "0"
      Top             =   5280
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   4080
      TabIndex        =   20
      Text            =   "Паспортные данные"
      Top             =   5280
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   360
      TabIndex        =   18
      Text            =   "000-00-00"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Text            =   "0"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Text            =   "Кол-во"
      Top             =   4080
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   7800
      TabIndex        =   5
      Text            =   "Combo3"
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox Text5 
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
      Height          =   345
      Left            =   7080
      TabIndex        =   4
      Text            =   "№кв."
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "Kvart.frx":011A
      Left            =   720
      List            =   "Kvart.frx":011C
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   1680
      Width           =   5655
   End
   Begin VB.TextBox Text4 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   345
      Left            =   120
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   17
      Text            =   "Text4"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   6960
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   345
      Left            =   3600
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   120
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Подъезд"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   6.6
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8880
      TabIndex        =   47
      Top             =   4200
      Width           =   852
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Доп. льготн.  площадь"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5880
      TabIndex        =   44
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Квартира приватизирована ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   1440
      TabIndex        =   42
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Данные о квартире"
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
      Left            =   660
      TabIndex        =   41
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Документ на льготу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "№ соцзащиты"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6240
      TabIndex        =   39
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Общая площадь"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Жилая площадь"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3000
      TabIndex        =   37
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Этаж"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8280
      TabIndex        =   36
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "проживает"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image imgTitleHelp 
      Height          =   156
      Left            =   10320
      Picture         =   "Kvart.frx":011E
      Top             =   120
      Width           =   156
   End
   Begin VB.Image imgTitleLeft 
      Height          =   360
      Left            =   5400
      Picture         =   "Kvart.frx":0368
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleRight 
      Height          =   360
      Left            =   6000
      Picture         =   "Kvart.frx":0AB2
      Top             =   0
      Width           =   228
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   7200
      Picture         =   "Kvart.frx":11FC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   285
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      X1              =   8
      X2              =   704
      Y1              =   440
      Y2              =   440
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Док.на квартиру"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   34
      Top             =   5760
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   33
      Top             =   5040
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6600
      TabIndex        =   32
      Top             =   5040
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      X1              =   8
      X2              =   704
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Прочие данные"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   4560
      Width           =   10455
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      X1              =   184
      X2              =   880
      Y1              =   440
      Y2              =   440
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "в т.ч.для начисления за услуги пользования лифтом"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   4440
      TabIndex        =   29
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "прописано"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2160
      TabIndex        =   28
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      X1              =   8
      X2              =   704
      Y1              =   280
      Y2              =   280
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Данные о квартире"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   10575
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      X1              =   8
      X2              =   704
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      X1              =   8
      X2              =   704
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ул."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Кв.№"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   1680
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   10455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ответственный квартиросъемщик:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   10335
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
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
Attribute VB_Name = "Kvart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ComboTipKv As ADODB.Recordset
Dim rs_kat As ADODB.Recordset
Dim RsL As ADODB.Recordset
'Dim mconn As ADODB.Connection
Dim Addrconn As ADODB.Recordset
Dim Combo_RS As ADODB.Recordset
Dim Addi As ADODB.Recordset
Public IzmLgot As Boolean
Dim f1, sq, sq1, SQDOM As String
Public F, Q, t As String
Public Old, OldPlo, OldProp As Double
Dim IzmfAM As Boolean

Dim Temp
Dim flgResize As Boolean
Dim OldCursorPos As PointAPI
Dim NewCursorPos As PointAPI




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

Private Sub BtnEnh1_Click()
Dim Сообщение As String
Dim i As Integer


If Q <> "" Then
If MsgBox("Был установлен тариф на электроэнергию =" + t + "  для всех проживающих по адресу " + Combo2.Text + " ОБНОВИТЬ ДЛЯ ВСЕХ", vbYesNo) = vbYes Then
Mconn.Execute (Q)
End If
End If
Сообщение = ""
i = 0

If Trim(Text1.Text) = "" Then
MsgBox ("Не бывает квартиросъемщика, без фамилии!")
Text1.SetFocus
Exit Sub
End If

If Trim(Text2.Text) = "" Then
MsgBox ("Безымянный квартиросъемщик!")
Text2.SetFocus
Exit Sub
End If

If Trim(Text3.Text) = "" Then
MsgBox ("Введите отчество!")
Text3.SetFocus
Exit Sub
End If


If Val(Trim(Text6.Text)) = 0 Then
'MsgBox ("Без общей площади квартиры, начисление квартплаты будет невозможно!")
i = i + 1
Сообщение = Сообщение + Str(i) + ".  " + " Без общей площади квартиры, начисление квартплаты будет невозможно!" + vbNewLine

'Text6.SetFocus
'Exit Sub
End If

If Val(Trim(Text8.Text)) = 0 Then
'MsgBox ("Если не указан этаж, то Вы не сможете использовать этот параметр, при начислении услуг пользования лифтом!")
i = i + 1
Сообщение = Сообщение + Str(i) + ".  " + " Если не указан этаж, то Вы не сможете использовать этот параметр, при начислении услуг пользования лифтом!" + vbNewLine
'Text6.SetFocus
'Exit Sub
End If

If Val(Trim(Text10.Text)) = 0 Then
i = i + 1
Сообщение = Сообщение + Str(i) + ".  " + " Нет прописанных жильцов!" + vbNewLine
End If


If Val(Trim(Text9.Text)) = 0 Then
i = i + 1
Сообщение = Сообщение + Str(i) + ".  " + " Нет проживающих!" + vbNewLine
End If



If Сообщение <> "" Then MsgBox (Сообщение)

'Set Filter = Nothing

rs_kat.Fields("oldnum").Value = Trim(Text4.Text)
rs_kat.UpdateBatch

'Прочие
If Lic.ops <> 1 Then

Mconn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv = MainOccupant.Numer SET Adding.Propis = [MainOccupant]![NLODGERF], Adding.Projiv = [MainOccupant]![NLODGER], Adding.ProLift = [MainOccupant]![NLODLIFT], Adding.ObPl = [MainOccupant]![COMSPACE], Adding.PolPl = [MainOccupant]![HABSPACE], Adding.TipKvKod = [MainOccupant]![KV], Adding.TipDomKod = [MainOccupant]![DomTip], Adding.Dop = [MainOccupant]![Dop]" + f1)
Mconn.Execute ("UPDATE Adding INNER JOIN MainOccupant ON Adding.KodKv=MainOccupant.Numer SET Adding.Propis = MainOccupant!NLODGERF, Adding.Projiv = MainOccupant!NLODGER, Adding.ProLift = MainOccupant!NLODLIFT, Adding.ObPl = MainOccupant!COMSPACE, Adding.PolPl = MainOccupant!HABSPACE, Adding.TipKvKod = MainOccupant!KV, Adding.TipDomKod = MainOccupant!DomTip " + f1)
Mconn.Execute ("UPDATE Adding INNER JOIN TMP_Lgota ON Adding.KodKv = TMP_Lgota.KodKv SET TMP_Lgota.Plo = [Adding]![ObPl], TMP_Lgota.Prop = [Adding]![Propis] WHERE (((Adding.KodKv)=" + F + "))")

If IzmLgot = True Then

'Jdite.Show
'Jdite.Label1.FontSize = 8
Pod.ProgressBar1.min = 1
Pod.ProgressBar1.Max = 1000
Pod.ProgressBar1.Visible = True



Pod.Show

Pod.Label1.Font = 8
Pod.Label1.Caption = " Пожалуйста подождите. Идет пересчет данных о льготах, т.к. данные лиц.счета были изменены "
Pod.Refresh
Pod.Label1.Refresh

For i = Pod.ProgressBar1.min To 250
    Pod.ProgressBar1.Value = i
   Next
'Pod.Refresh


'Jdite.Label1.Caption = "  Пожалуйста подождите. Идет пересчет данных о льготах, т.к. данные лиц.счета были изменены "

'Jdite.Label1.Refresh
'Обновляем соцминимум
Mconn.Execute ("UPDATE Adding INNER JOIN Socmin ON (Adding.KodKat = Socmin.KodKategor) AND (Adding.Propis = Socmin.koli) SET Adding.Socmin = [Socmin]![Value]+Adding.DOP WHERE (((Adding.KodKv)=" + Filter.Nm + "))")

Pod.ProgressBar1.Value = 400

Mconn.Execute ("UPDATE Tarif INNER JOIN Adding ON (Tarif.KodDOM = Adding.TipDomKod) AND (Tarif.KodKV = Adding.TipKvKod) AND (Tarif.KodKat = Adding.KodKat) SET Adding.Tarif = [Tarif]![Value], Adding.TarifI = [Tarif]![TarifI], Adding.TarifD = [Tarif]![TarifD] WHERE (((Adding.KodKv)=" + Filter.Nm + "))")

Pod.ProgressBar1.Value = 500
Mconn.Execute ("UPDATE Adding INNER JOIN tmp_lgota ON Adding.Key = tmp_lgota.UniKOd SET tmp_lgota.Cocmin = [Adding]![Socmin] WHERE (((Adding.KodKv)=" + Filter.Nm + "))")
Pod.ProgressBar1.Value = 600


' Если данных нет то добавляем норматив

'Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Propis] WHERE (((Adding.Sch)='Да') AND (([Adding]![Shc_old]-[Adding]![Shc_new])=0) AND ((Adding.KodKv)=" + Filter.Nm + "))")
'Если проставлен признак норматива то пересчитываем
Mconn.Execute ("UPDATE Adding SET Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Propis] WHERE (((Adding.nr)=True) AND ((Adding.Sch)='Да') AND ((Adding.KodKv)=" + Filter.Nm + "))")

'Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Propis] WHERE (((Adding.Sch)='Да') AND ((Adding.Shc_new)=0) AND ((Adding.KodKv)=" + Filter.Nm + "))")

'Mconn.Execute ("UPDATE Adding SET Adding.nr = true ,Adding.Shc_new = [Adding]![Shc_old]+[Adding]![norm]*[Adding]![Propis] WHERE (((Adding.Sch)='Да') AND ((Adding.Shc_new)=0) AND ((Adding.KodKv)=" + Filter.Nm + "))")



'Если счетчик то площадь равна разнице показаний счетчика
 Mconn.Execute ("UPDATE Adding SET Adding.ObPl = [Adding]![Shc_new]-[Adding]![Shc_old] WHERE (((Adding.Sch)='Да') AND ((Adding.KodKv)=" + Filter.Nm + "))")



' Если были изменения то пересчет

Set Addi = New ADODB.Recordset
Set Addi.ActiveConnection = Mconn
Addi.CursorType = adOpenKeyset
Addi.LockType = adLockPessimistic

On Error GoTo ПустойАддинг


Addi.Open ("SELECT Adding.KodKv, Adding.LgotaP, Adding.Key, Adding.ObPl From Adding WHERE (((Adding.KodKv)=" + [Filter].[Nm] + "))")
Pod.ProgressBar1.Value = 650

Addi.MoveFirst
Do While Not Addi.EOF

MainForm.II = 0
MainForm.Pi = 0
MainForm.Ostatok = rs_kat.Fields("ComSpace").Value
MainForm.РЛ Addi.Fields("key").Value, True

If MainForm.Двойник = True Then
MainForm.Pi = 0
MainForm.II = 0
MainForm.Ostatok = rs_kat.Fields("ComSpace").Value
MainForm.РЛ Addi.Fields("key").Value, False
End If

Pod.ProgressBar1.Value = 700
MainForm.ViborLLg Addi.Fields("key").Value

Addi.Fields("LgotaP").Value = MainForm.PrZ
Addi.UpdateBatch
Addi.MoveNext
Loop

Pod.ProgressBar1.Value = 800
End If
End If

If IzmfAM = True Then
Pod.ProgressBar1.Value = 900
Pod.Show
Pod.Refresh
'Unload Filter
End If

Unload Me

Pod.ProgressBar1.Value = 1000

ПустойАддинг:



'Pod.ProgressBar1.Value = 1199
Unload Pod
Unload Me
If Lic.ops = 1 Then
Lic.Show
'Lic.FG1.Refresh
Else

'Возврат фильтра
Filter.Fg.FlexDataSource = Filter.m_DS

Filter.Show
Unload Pod


' Возвращаем курсор на место
Filter.Fg.Row = Filter.CL5
Filter.Fg.SetFocus
Filter.Fg.Select Filter.CL5, 2, Filter.CL5, 3
SendKeys "{left}"


End If
'Kvart.Hide


End Sub

Private Sub BtnEnh11_Click()

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



Private Sub Combo2_Validate(Cancel As Boolean)
IzmLgot = True
End Sub

Private Sub Combo3_LostFocus()
ComboTipKv.MoveFirst
Do While Not ComboTipKv.EOF
If Combo3.Text = ComboTipKv("NAME_KV") Then rs_kat("KV") = ComboTipKv("КОД")
ComboTipKv.MoveNext
Loop
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
IzmLgot = True
End Sub

Private Sub Command2_Click()
Electro.Show

'Unload Me
'Filter.Show
'Kvart.Hide
End Sub

Private Sub Command1_Click()
    ostrovodrepinit Me, rs_kat
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
'Form_Load
End Sub

Private Sub Form_Initialize()
'If F <> "" Then Kvart.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = 116 Then Command1_Click
    End If
End Sub

Private Sub Form_Load()
MakeWindow Me, True
lblTitle.Left = Command1.Left + Command1.Width + 4
'imgTitleMaxRestore.Picture = imgTitleMaximize.Picture



Q = ""
IzmfAM = False

IzmLgot = False
'IzmfAM = False
' open connection
 ' Set mconn = New ADODB.Connection
  'mconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/kvartplata.mdb;Persist Security Info=True"
  'mconn.Open "data/Kvartplata.mdb"
  ' Рекордсет для сетки
Set rs_kat = New ADODB.Recordset
Set rs_kat.ActiveConnection = Mconn
'f = Filter.nm
If Filter.ad <> 1 Then F = Filter.Nm Else F = Filter.nNum
'MsgBox (F)
If F <> "" Then f1 = "WHERE (((MainOccupant.Numer)=" & F & "))"
f2 = "WHERE (((OtheOwner.Numer)=" & F & "))"
f3 = "SELECT MainOccupant.Numer, MainOccupant.Dom,MainOccupant.Priv, MainOccupant.KV, MainOccupant.FAM, MainOccupant.IM, MainOccupant.OT, MainOccupant.NLODLIFT"
f3 = f3 + ", MainOccupant.NLODGER, MainOccupant.NLODGERF, MainOccupant.NROOM, MainOccupant.COMSPACE, MainOccupant.HABSPACE, MainOccupant.PRIVILEGE, MainOccupant.HABITATE, MainOccupant.BIRTHDAY, MainOccupant.NORDER, MainOccupant.KITCHSPACE, MainOccupant.BATHSPACE, MainOccupant.CORRSPACE, MainOccupant.TOILSPACE, MainOccupant.BALCSPACE, MainOccupant.NFAMILY, MainOccupant.DATARECEIV, MainOccupant.PASSPORT, MainOccupant.TELEPHONE, MainOccupant.LDOK, MainOccupant.LDATEBEG, MainOccupant.LDATEEND, MainOccupant.NAPARTMENT, MainOccupant.FLOOR, MainOccupant.SocNum, MainOccupant.COMM, KLS_PODR.NAIM_KLS, KLS_PODR.Num, KLS_PODR.Tip, KLS_PODR.Tip_Naim, TipKv.Name_Kv, MainOccupant.KV_num, MainOccupant.DomTip,MainOccupant.OLDNUM, MainOccupant.dop, MainOccupant.podyezd "
'Normann's-addition
f3 = f3 & ",MainOccupant.BanKN "
'Normann's-addition-end
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

rs_kat.Open (sq), Mconn, adOpenKeyset



'If Filter.ad = 1 Then Command2.Enabled = False



'MsgBox (Rs_kat.Fields("FAM").Value)

Addrconn.Open ("KLS_PODR")
Показать

If Filter.ad <> 1 Then OldPlo = rs_kat.Fields("comSpace").Value Else OldPlo = 0
If Filter.ad <> 1 And rs_kat.Fields("NLODGERF").Value <> "" Then OldProp = rs_kat.Fields("NLODGERF").Value Else OldProp = 0

'If Filter.ad <> 1 Then Oldf = Rs_kat.Fields("FAM").Value
'If Filter.ad <> 1 Then OldiM = Rs_kat.Fields("im").Value
'If Filter.ad <> 1 Then Oldiot = Rs_kat.Fields("ot").Value
'MsgBox (Str(OldPlo) + " " + Str(OldProp))
'Combo1.ItemData (2)




End Sub
Sub Показать()
On Error Resume Next
Dim nnn As Integer

'For nnn = 0 To (rs_kat.Fields.Count - 1)
'    Debug.Print nnn & vbTab & rs_kat.Fields(nnn).name
'Next nnn

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

If rs_kat.Fields("priv").Value <> "" Then
If rs_kat.Fields("priv").Value = "Да" Then Option1.Value = True
If rs_kat.Fields("priv").Value = "Нет" Then Option2.Value = True

If rs_kat.Fields("dop").Value <> "" Then Me.Text18 = rs_kat.Fields("dop").Value Else Me.Text18 = 0

If rs_kat.Fields("podyezd").Value <> "" Then Me.Text19 = rs_kat.Fields("podyezd").Value Else Me.Text19 = 0

End If



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
' Normann-------------------------------------------Normann
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

Private Sub Form_Unload(Cancel As Integer)
Filter.Enabled = True


End Sub

Private Sub Option1_Click()
'MsgBox (Option1.CausesValidation)
'MsgBox (Option2.CausesValidation)
rs_kat.Fields("Priv").Value = "Да"
End Sub

Private Sub Option2_Click()
rs_kat.Fields("Priv").Value = "Нет"
End Sub

Private Sub Text1_Change()
'Text1.Text = Rs_kat.Fields("Numer").Value
End Sub

Private Sub Text1_LostFocus()
If Len(Text1.Text) < 20 Then
rs_kat.Fields("FAM").Value = Text1.Text
Else
MsgBox ("Слишком длинная фамилия")
rs_kat.Fields("FAM").Value = Left(Text1.Text, 20)
Text1 = Left(Text1.Text, 20)
Text1.Refresh

End If

End Sub

Private Sub Text1_Validate(Cancel As Boolean)
IzmfAM = True
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
IzmLgot = True
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
IzmLgot = True
End Sub

Private Sub Text18_LostFocus()
If Trim(Text18.Text) = "Доп" Then Text18.Text = 0
If Trim(Text18.Text) <> "" Then rs_kat.Fields("Dop").Value = Replace(Text18.Text, ".", ",")
rs_kat.UpdateBatch
IzmLgot = True
End Sub

Private Sub Text19_Change()
If Me.Text19 = "" Then Me.Text19 = 0
End Sub

Private Sub Text19_LostFocus()
If Me.Text19 = "" Then Me.Text19 = 0
rs_kat.Fields("podyezd").Value = Int(Text19)
rs_kat.UpdateBatch
Mconn.Execute ("UPDATE Adding SET Adding.podyezd = " + Text19 + " WHERE (((Adding.KodKv)=" + F + "))")
End Sub

Private Sub Text2_LostFocus()
If Len(Text2.Text) < 20 Then
rs_kat.Fields("IM").Value = Text2.Text
Else
MsgBox ("Слишком длинное имя")
rs_kat.Fields("IM").Value = Left(Text2.Text, 20)
Text2 = Left(Text2.Text, 20)
Text2.Refresh

End If

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
IzmfAM = True
End Sub

Private Sub Text3_LostFocus()
If Len(Text3.Text) < 20 Then
rs_kat.Fields("OT").Value = Text3.Text
Else
MsgBox ("Слишком длинное отчество")
rs_kat.Fields("OT").Value = Left(Text3.Text, 20)
Text3 = Left(Text3.Text, 20)
Text3.Refresh

End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
IzmfAM = True
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

Private Sub Text6_Change()
'Text6.FillColor = RGB(207, 207, 207)
End Sub
Private Sub Text6_Validate(Cancel As Boolean)
IzmLgot = True

End Sub

Private Sub Text8_Validate(Cancel As Boolean)
IzmLgot = True
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
IzmLgot = True
End Sub

Private Sub Выход_Click(Index As Integer)
'Command1_Click
End Sub
Private Sub Text5_LostFocus()
rs_kat.Fields("kv_num").Value = Text5.Text
End Sub
Private Sub Text6_LostFocus()
If InStr(1, Text6.Text, ".", vbTextCompare) <> 0 Then
MsgBox ("Для разделения разрядов используется запятая")
Text6.SetFocus
Exit Sub
End If
rs_kat.Fields("ComSpace").Value = Text6.Text
If Lic.ops = 1 Then
Lic.fg1.TextMatrix(Lic.fg1.Row, 15) = Text6.Text
Lic.fg1.Refresh
End If
End Sub
Private Sub Text7_LostFocus()
If InStr(1, Text7.Text, ".", vbTextCompare) <> 0 Then
MsgBox ("Для разделения разрядов используется запятая")
Text7.SetFocus
Exit Sub
End If
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
If Len(Text12.Text) >= 250 Then Text12.Text = Left(Text12.Text, 250)
If Len(Text12.Text) = 0 Then Text12.Text = " "

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
