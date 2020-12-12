Attribute VB_Name = "RusLat"
' Декларация функций и констант АПИ
   Declare Function ActivateKeyboardLayout Lib "user32" _
           (ByVal HKL As Long, ByVal flags As Long) As Long
   Public Const kb_lay_ru As Long = 68748313
   Public Const kb_lay_en As Long = 67699721
       
   ' Переключить на русский язык
   'X = ActivateKeyboardLayout&(kb_lay_ru, 0)

   ' Переключить на английский язык
   'X = ActivateKeyboardLayout&(kb_lay_en, 0)

