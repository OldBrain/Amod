Attribute VB_Name = "modAboutBox"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyLong Lib "MSVBVM60.DLL" Alias "GetMem4" (ByVal Source As Long, Dest As Long)
'Private Declare Sub CopyMemory Lib "kernel32" (Dest As Any, Src As Any, ByVal Length As Long)
Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
'Private Type LITEM
'    mask As Long
'    iLink As Long
'    state As Long
'    stateMask As Long
'    szID As String * 96
'    szUrl As String * 4168
'End Type
'Private Type NMHDR
'    hWndFrom As Long
'    idFrom As Long
'    code As Long
'End Type
'Private Type NMLINK
'    hdr As NMHDR
'    item As LITEM
'End Type
Private Const WM_INITDIALOG As Long = &H110
Private Const WM_NOTIFY As Long = &H4E
Private Const NM_CLICK As Long = -2
Private m_Company As String, m_Copyright As String, m_Version As String, m_License As String
Private m_DlgProc As Long

Public Function DialogProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case Msg
Case WM_INITDIALOG:
    SetDlgItemText hWnd, &H3500&, m_Company & " %s"
    If Len(m_License) > 0 Then _
        SetDlgItemText hWnd, &H3512&, m_License
    SetWindowText FindWindowEx(hWnd, GetDlgItem(hWnd, &H350B&), "Static", vbNullString), m_Copyright
    DialogProc = CallWindowProc(m_DlgProc, hWnd, Msg, wParam, lParam)
    SetDlgItemText hWnd, &H350B&, m_Version
Case WM_NOTIFY:
    If (Len(m_License) > 0) And (wParam = &H3512&) Then
        Dim code As Long
        CopyLong lParam + 8, code
        If code = NM_CLICK Then
            ShellExecute hWnd, 0, lParam + 124, 0, 0, 1
        Else
            DialogProc = CallWindowProc(m_DlgProc, hWnd, Msg, wParam, lParam)
        End If
    Else
        DialogProc = CallWindowProc(m_DlgProc, hWnd, Msg, wParam, lParam)
    End If
Case Else:
    DialogProc = CallWindowProc(m_DlgProc, hWnd, Msg, wParam, lParam)
End Select
End Function

Public Sub InitVariables(Company As String, Copyright As String, Version As String, License As String, ByVal DlgProc As Long)
m_Company = Company: m_Copyright = Copyright: m_Version = Version: m_License = License
m_DlgProc = DlgProc
End Sub
