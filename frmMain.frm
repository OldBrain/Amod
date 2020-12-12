VERSION 5.00
Begin VB.Form Closc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2508
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5616
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2508
   ScaleWidth      =   5616
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Квартплата V1.0 "
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox PctFon 
      Height          =   735
      Left            =   840
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   684
      ScaleWidth      =   1164
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   600
      Top             =   360
   End
   Begin VB.Image ImgSM 
      Height          =   240
      Index           =   0
      Left            =   3480
      Picture         =   "frmMain.frx":EFF8
      Top             =   1560
      Width           =   240
   End
   Begin VB.Image ImgSS 
      Height          =   180
      Index           =   0
      Left            =   4080
      Picture         =   "frmMain.frx":F4B4
      Top             =   1440
      Width           =   180
   End
   Begin VB.Image ImgSH 
      Height          =   360
      Index           =   0
      Left            =   2880
      Picture         =   "frmMain.frx":F8C0
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image ImgCenter 
      Height          =   480
      Left            =   4320
      Picture         =   "frmMain.frx":FF12
      Top             =   240
      Width           =   480
   End
   Begin VB.Image ImgB 
      Height          =   480
      Index           =   100
      Left            =   3480
      Picture         =   "frmMain.frx":106E4
      Top             =   360
      Width           =   480
   End
   Begin VB.Image ImgS 
      Height          =   192
      Index           =   100
      Left            =   2880
      Picture         =   "frmMain.frx":10FEF
      Top             =   360
      Width           =   192
   End
End
Attribute VB_Name = "Closc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Вот, собственно, программа - скринсейвер.
'Можно использовать любые картинки любых размеров.
'единственное ограничение - они должны быть квадратными
'т.е. другими словами можно сказать, что она поддерживает скины.
'после компиляции меняем разрешение на .SCR и копируем в
'"?:\WINDOWS\System32\" - завтавка готова

'Alexey Spirin (Alexey@VBRussian.com)

Option Explicit

Dim NullX As Integer, NullY As Integer, R As Integer, cnt As Single, SHC As Single, SMC As Single, SSC As Single, StartX As Integer, StartY As Integer


Sub MakeCircle() ' функция рисует циферблат

'рисуем минутные точки

For cnt = 0 To 59
    'сначали был ImgS с индексом 100
    'здесь загружаем новые ImgS с индексами с 0 до 59
    Load ImgS(cnt)
    'устанавливаем позиции.
    'фраза  "- ImgS(cnt).Width / 2" нужна для сдвига
    'картинки на половину её размеров, т.к. Left & Top
    'коррдинаты верхнего левого угла, а не центра картинки
    ImgS(cnt).Left = NullX + Cos(Rad(cnt) * 6) * R - ImgS(cnt).Width / 2
    ImgS(cnt).Top = NullY + Sin(Rad(cnt) * 6) * R - ImgS(cnt).Width / 2
    'делаем её видимой
    ImgS(cnt).Visible = True
Next

'рисуем часовые точки
'здесь все тоже самое

For cnt = 0 To 11
    Load ImgB(cnt)
    ImgB(cnt).Left = NullX + Cos(Rad(cnt) * 30) * R - ImgB(cnt).Width / 2
    ImgB(cnt).Top = NullY + Sin(Rad(cnt) * 30) * R - ImgB(cnt).Width / 2
    ImgB(cnt).Visible = True
    'Используем ZOrder для указания того, что эти картинки должны быть
    'поверх картинок минут - чем меньше ZOrder - тем "выше" объект
    ImgB(cnt).ZOrder 0
Next

End Sub


Private Sub Command1_Click()
Dim AboutBox As New AboutBox
With AboutBox
    .Title = " Расчет и анализ коммунальных платежей населения"
    .Version = "Версия: " + Str(App.Major) + "." + Str(App.Minor) + "." + Str(App.Revision)
    .Company = "Квартплата +  (C) Copyright, 2005, Астрахань"
    .Copyright = " Бугоров Андрей Владимирович"
    .Description = "Комплексная автоматизация бухучета"
    .License = "Связь с автором E-Mail:bestonline@list.ru телефоны:+79881733-600"
    .hWndOwner = Me.hwnd
    'Set .Icon = Me.Icon
    .AboutBox
End With
End Sub

Private Sub Form_Click()
'Завершаем работу при клике на форму
'Unload Me
Command1_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Завершаем работу при нажатии любой кнопки на форме
Unload Me
'MainForm.Show

End Sub

Private Sub Form_Load()

'StartX, StartY - координаты мыши при загрузке программы.
'Сначала выставляем их в -1, чтобы потом было ясно, что при
'первом срабатывании MouseMove, оно действительно первое
StartX = -1
StartY = -1

'подгоняем размеры формы под размеры экрана, двигаем в
'крайний левый угол и делаем поверх всех
Width = Screen.Width
Height = Screen.Height
Move 0, 0
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H1 Or &H2

'загружаем на форму картинку фона с нужными размерами
PaintPicture PctFon.Picture, 0, 0, Width, Height

'определяем координаты центра формы
NullX = (ScaleWidth) / 2
NullY = (ScaleHeight) / 2

'выравниваем центральную картинку
ImgCenter.Left = NullX - ImgCenter.Width / 2
ImgCenter.Top = NullY - ImgCenter.Height / 2

'убираем начальные картинки циферблата
ImgB(100).Visible = False
ImgS(100).Visible = False

'Определяем радиус циферблата.
R = (NullY - ImgB(100).Width)

'Определяем количество картинок для часов, минут и секунд
'Длина часовой стрелки должна быть не больше 2/3 радиуса
'Длина минутной - чуть меньше радиуса
'а длина секундной - не больше 3/4 радиуса
'Зачем все это нужно и почему бы не прописать сразу?
'а потому что во-первых программа должна работать с разными разрешениями,
'во-вторых - можно использовать другие картинки - так что таким образом
'программа обретает универсальность
SHC = Int(2 / 3 * R / ImgSH(0).Width)
SMC = Int((R - ImgB(100).Width * 1.5) / ImgSM(0).Width)
SSC = Int(3 / 4 * R / ImgSS(0).Width)

'описанным выше методом загружаем картинки часов, минут и секунд
For cnt = 1 To SHC - 1
    Load ImgSH(cnt)
    ImgSH(cnt).Visible = True
    ImgSH(cnt).ZOrder 0
Next

For cnt = 1 To SMC - 1
    Load ImgSM(cnt)
    ImgSM(cnt).Visible = True
    ImgSM(cnt).ZOrder 0
Next

For cnt = 1 To SSC - 1
    Load ImgSS(cnt)
    ImgSS(cnt).Visible = True
    ImgSS(cnt).ZOrder 0
Next

'делаем центральную картинку поверх других
ImgCenter.ZOrder 0

'запускаем прорисовку циферблата
MakeCircle

'запускаем прорисовку стрелок (т.к. таймер сработает только через 1 сек).
Rotate

Command1.Caption = "Версия: " + Str(App.Major) + "." + Str(App.Minor) + "." + Str(App.Revision)

End Sub

Function Rad(Grad) As Double 'функция для перевода градусов в радианы
Rad = Grad / 180 * 3.141592654
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If StartX = -1 Or StartY = -1 Then  'если начальные координаты мыши неизвестны,
                                    'то устанавливаем текущие в качестве начальных
 '   StartX = X
 '  StartY = Y
'Else 'в противном случае следим- если мыша сдвинулась больше чем на 5 пикселей - закрываем программу
 '   If Abs(StartX - X) > 150 Then Unload Me
  '  If Abs(StartY - Y) > 150 Then Unload Me
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'после выгрузки формы завершаем программу
'End
MainMenu.Show

End Sub

Private Sub Timer_Timer() 'каждую секнду запускает прорисовку стрелок
Rotate
End Sub


Sub Rotate() 'прорисовка стрелок.

'здесь в общем то ничего сложного - работа с тригонометрией и координатами.
'каждая картинка передвигается на положенное ей место.
'единственное что может вызывать вопрос - так это почему из секунд и минут
'вычитается 15, а из часов - 3 ?
'это связано с тригонометрией. Нужно вычесть четверть периода-тогда нормально
'все отобржаться будет. Для интереса - попробуйте убрать эти вычитания.

Dim s As Single, M   As Single, H  As Single

s = Second(Now) - 15
M = Minute(Now) + Second(Now) / 60 - 15
H = Hour(Now) + Minute(Now) / 60 - 3

For cnt = 0 To SHC - 1
    ImgSH(cnt).Left = NullX + Cos(Rad(H) * 30) * cnt * ImgSH(cnt).Width - ImgSH(cnt).Width / 2
    ImgSH(cnt).Top = NullY + Sin(Rad(H) * 30) * cnt * ImgSH(cnt).Width - ImgSH(cnt).Width / 2
Next

For cnt = 0 To SMC - 1
    ImgSM(cnt).Left = NullX + Cos(Rad(M) * 6) * cnt * ImgSM(cnt).Width - ImgS(cnt).Width / 2
    ImgSM(cnt).Top = NullY + Sin(Rad(M) * 6) * cnt * ImgSM(cnt).Width - ImgS(cnt).Width / 2
Next

For cnt = 0 To SSC - 1
    ImgSS(cnt).Left = NullX + Cos(Rad(s) * 6) * cnt * ImgSS(cnt).Width - ImgSS(cnt).Width / 2
    ImgSS(cnt).Top = NullY + Sin(Rad(s) * 6) * cnt * ImgSS(cnt).Width - ImgSS(cnt).Width / 2
Next

End Sub
