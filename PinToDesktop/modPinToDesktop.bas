Attribute VB_Name = "modPinToDesktop"
'---------------------------------------------------------------------------------------
'
' Module    : modPinToDesktop
' DateTime  : 28/06/2005 09:14 PM
' Update    : 29/06/2005 02:11 PM (Added unpin sub)
' Author    : Carlos Alberto S.
' Purpose   : Module to easily pin or stick an application to Windows' desktop.
'             It will work with "Active Desktop" on or off and with "Show desktop icons"
'             on or off too. The WinKey+D or WinKey+M won't minimize the pinned application.
'             Tested only in Windows XP.
' Credits   : Original code by Jesse Seidel (Dr. Fire). The original code works
'              just when Active Desktop is on. Thanks "herd" for the explanation.
'              http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=57755&lngWId=1
' Article   : http://msdn.microsoft.com/msdnmag/issues/0600/w2kui2/
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Sub PinToDesktop(FormToPin As Form)

    '"Progman" would be the parent window of the SHELLDLL_DefView class that
    '   itself would be either hosting the SysListView32 in classic view
    '   or the Internet Explorer_Server in Active Desktop view.
    
    Dim progman As Long
    progman = FindWindow("progman", vbNullString)
    SetParent FormToPin.hWnd, progman

End Sub

Public Sub UnPinFromDesktop(FormToUnPin As Form)

    Dim lngExplorer As Long
    
    lngExplorer = FindWindow("ExploreWClass", vbNullString)
    SetParent FormToUnPin.hWnd, lngExplorer

End Sub

'Another way to pin or stick to desktop (limitation: if the application is running and the user
'   set the desktop to "Show icons" the application will be hidden).

Public Sub StickToDesktop(FormToStick As Form)

    Dim progman As Long
    Dim shelldlldefview As Long
    Dim tDesktop As Long
    
    progman = FindWindow("progman", vbNullString)
    shelldlldefview = FindWindowEx(progman, 0&, "shelldll_defview", vbNullString)
    SetParent FormToStick.hWnd, shelldlldefview

End Sub

