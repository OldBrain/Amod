VERSION 5.00
Begin VB.Form UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
    'тут типа SQL
    PauseTime = 5
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents
    Loop
    'ну и хватит пожалуй юзера мучать
    Unload Me
End Sub

