VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMessageSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©1999 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' **************************************************************
'  Implemented Methods
' **************************************************************
Public Function WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
'   Static Result As Long
'
'   Select Case msg
'      Case Else
'         ' Insure all other messages are handled properly.
'         Result = InvokeWindowProc(hWnd, msg, wp, lp)
'   End Select
'
'   ' Return appropriate value.
'   IMessageSink_WindowProc = Result
End Function

