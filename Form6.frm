VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MkSQ"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form6"
   ScaleHeight     =   8400
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  1.Наименование льготы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CheckBox Check17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "17. Прописка"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   3720
      TabIndex        =   34
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CheckBox Check16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "16. Дата прописки"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CheckBox Check15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "15. Этаж"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   3720
      TabIndex        =   32
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CheckBox Check14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "14.Площадь балкона"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CheckBox Check13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "13.Площадь туалета"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CheckBox Check12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "12.Площадь коридоров"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "11.Площадь ванной"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "10.Площадь кухни"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ок"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   7
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  9.Счет затрат"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   6000
      Width           =   3015
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  8.Начисление"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  7.Кол-во прописанных"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0C0C0&
      Caption         =   " 6. Кол-во проживающих"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  5.Общая площадь"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  4.Жилая площадь"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  2.Адрес"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   3255
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  3.Фамилия имя отчество"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   0
      Left            =   3600
      TabIndex        =   25
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   1
      Left            =   3600
      TabIndex        =   26
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   2
      Left            =   3600
      TabIndex        =   27
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   3
      Left            =   3600
      TabIndex        =   28
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   4
      Left            =   3600
      TabIndex        =   29
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3480
      TabIndex        =   30
      Top             =   960
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   5
      Left            =   3600
      TabIndex        =   35
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   6
      Left            =   3600
      TabIndex        =   36
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   7
      Left            =   3600
      TabIndex        =   37
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   3480
      TabIndex        =   38
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Index           =   9
      Left            =   120
      TabIndex        =   43
      Top             =   6600
      Width           =   3255
      Begin VB.CheckBox Check18 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Сумма начисления"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Выбор данных для отчета"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Прочие данные"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   3480
      TabIndex        =   39
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Основные сведения"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Площадь подсобных помещений"
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
      Height          =   615
      Left            =   3480
      TabIndex        =   31
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sq1 As String

Private Sub Command1_Click()

Call MakeSQ(sq1)
Form6.Hide

Reports.sq = sq1
Unload Me
Analizlgot.Show


'Form6.Hide
'Form111.Show
End Sub

Private Sub Command2_Click()
Form6.Hide
Form1.Show
End Sub

Sub MakeSQ(sq1)
Dim Tbl As String

Tbl = "J_ALL"
sq1 = "SELECT "

' ************ 1
If Form6.Check1.Value Then
            sq1 = sq1 + Tbl + ".[Наименование_льготы], "
        Else
            sq1 = sq1
        End If
' ************ 2
If Form6.Check2.Value Then
            sq1 = sq1 + Tbl + ".NAIM_KLS, "
        Else
            sq1 = sq1
        End If
        
' ************ 3
If Form6.Check3.Value Then
            sq1 = sq1 + Tbl + ".[ФИО], "
        Else
            sq1 = sq1
        End If

' ************ 4
If Form6.Check4.Value Then
            sq1 = sq1 + Tbl + ".[ЖилаяПЛ], "
        Else
            sq1 = sq1
        End If

' ************ 5
If Form6.Check5.Value Then
            sq1 = sq1 + Tbl + ".[Площадь], "
        Else
            sq1 = sq1
       End If
       
  ' ************ 6
If Form6.Check6.Value Then
            sq1 = sq1 + Tbl + ".[Проживает], "
        Else
            sq1 = sq1
       End If
       
' ************ 7
If Form6.Check7.Value Then
            sq1 = sq1 + Tbl + ".[Прописано], "
        Else
            sq1 = sq1
       End If
       
 ' ************ 8
If Form6.Check8.Value Then
            sq1 = sq1 + Tbl + ".[Начисление], "
        Else
            sq1 = sq1
       End If
       
 ' ************ 9
If Form6.Check9.Value Then
            sq1 = sq1 + Tbl + ".[Счет_затрат], "
        Else
            sq1 = sq1
       End If
       
  ' ************ 18
If Form6.Check18.Value Then
            sq1 = sq1 + Tbl + ".[Сумма], "
        Else
            sq1 = sq1
       End If
       
       ' ************ 10
If Form6.Check10.Value Then
            sq1 = sq1 + Tbl + ".[Кухня], "
        Else
            sq1 = sq1
       End If
       
 ' ************ 11
If Form6.Check11.Value Then
            sq1 = sq1 + Tbl + ".[Ванная], "
        Else
            sq1 = sq1
       End If
       
       
 ' ************ 12
If Form6.Check12.Value Then
            sq1 = sq1 + Tbl + ".[Коридор], "
        Else
            sq1 = sq1
       End If
       
 ' ************ 13
If Form6.Check13.Value Then
            sq1 = sq1 + Tbl + ".[Туалет], "
        Else
            sq1 = sq1
       End If
       
       
' ************ 14
If Form6.Check14.Value Then
            sq1 = sq1 + Tbl + ".[Балкон], "
        Else
            sq1 = sq1
       End If
       
       
 ' ************ 15
If Form6.Check15.Value Then
            sq1 = sq1 + Tbl + ".[Этаж], "
        Else
            sq1 = sq1
       End If
       
' ************ 16
If Form6.Check16.Value Then
            sq1 = sq1 + Tbl + ".[ДатаПрописки], "
        Else
            sq1 = sq1
       End If
       
' ************ 17
If Form6.Check17.Value Then
            sq1 = sq1 + Tbl + ".[Прописка], "
        Else
            sq1 = sq1
       End If
   
'*********************************************
        sq1 = Left(sq1, Len(sq1) - 2) + " "
sq1 = sq1 + "FROM " + Tbl

End Sub


