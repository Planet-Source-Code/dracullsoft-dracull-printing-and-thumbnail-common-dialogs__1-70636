Attribute VB_Name = "modCommonDlg"
' ======================================================================================
' File  :     cCommonDialog
' Author:     Thomas Nick Andersen (DiceSix)
' Date  :     05-05-2008
' --------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------
' Copyright © 2008 DracullSoft aka DiceSix
' Author or copyright holders can not be held responsible for anything.
' Based on work by Steve McMahon and Bruce McKinney and VBnet, Randy Birch
' --------------------------------------------------------------------------------------
' Purpose: Update to Steve's all in one class but with support of 2000 / XP / Vista
'          and with ability to set initial view for Open and Save to show thumbnails
'
' NB If you do specify the OFN_EXPLORER flag when
' ======================================================================================
Option Explicit

'var for exposed property
Private m_lvInitialView As Long
Private m_bvCenterView  As Boolean

'this is the version 5+ definition of
'the OPENFILENAME structure containing
'three additional members providing
'additional options on Windows 2000
'or later. The SetOSVersion routine
'will assign either OSV_LENGTH (76)
'or OSVEX_LENGTH (88) to the OSV_VERSION_LENGTH
'variable declared above. This variable, rather
'than Len(OFN) is used to assign the required
'value to the OPENFILENAME structure's nStructSize
'member which tells the OS if extended features
'- primarily the Places Bar - are supported.
Public Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hwndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
    pvReserved        As Long
    dwReserved        As Long
    flagsEx           As Long    ' needed for 2000 xp and vista bar

End Type


Private Type RECT
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type

Private Type NMHDR
  hwndFrom  As Long
  idfrom    As Long
  code      As Long
End Type

Private Type OFNOTIFY
  hdr       As NMHDR
  lpOFN     As OPENFILENAME
  pszFile   As String
End Type

'windows messages & notifications etc
Private Const WM_COMMAND = &H111
Private Const WM_NOTIFY As Long = &H4E&
Private Const WM_INITDIALOG As Long = &H110
Private Const CDN_FIRST As Long = -601
Private Const CDN_INITDONE As Long = (CDN_FIRST - &H0&)
'// Notifications from Open or Save dialog

'Private Declare Function SetWindowText Lib "user32" _
'   Alias "SetWindowTextA" _
'  (ByVal hwnd As Long, _
'   ByVal lpString As String) As Long
'

Private Declare Function FindWindowEx Lib "user32" _
   Alias "FindWindowExA" _
  (ByVal hwndParent As Long, _
   ByVal hWndChildAfter As Long, _
   ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long

Private Declare Function GetParent Lib "user32" _
  (ByVal hwnd As Long) As Long

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
    
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
        ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal bRepaint As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32" _
   Alias "ShellExecuteA" _
  (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long


'______ *** Hot URL *** __________
Public Sub DoShellExecute(sTopic As String, sFile As Variant, _
                            sParams As Variant, sDirectory As Variant, _
                            nShowCmd As Long)

  'execute the passed operation, passing
  'the desktop as the window to receive
  'any error messages
   Call ShellExecute(GetDesktopWindow(), _
                     sTopic, _
                     sFile, _
                     sParams, _
                     sDirectory, _
                     nShowCmd)

End Sub


Public Property Let OFN_SetInitialView(ByVal initview As Long)

   m_lvInitialView = initview
   
End Property


Public Function FARPROC(pfn As Long) As Long
  
  'A dummy procedure that receives and returns
  'the return value of the AddressOf operator.
 
  'Obtain and set the address of the callback
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)

  FARPROC = pfn

End Function


Public Function OFNHookProc(ByVal hwnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByRef lParam As OFNOTIFY) As Long

   Dim hwndParent As Long
   Dim hwndLv As Long
   Static bLvSetupDone As Boolean
   'Debug.Print "uMsg:" & uMsg
   Select Case uMsg
      Case WM_INITDIALOG
        'Initdialog is set when the dialog has been created and is ready to
        'be displayed, so set our flag to prevent re-executing the code
        'in the wm_notify message. This is required as the dialog receives a
        'number of WM_NOTIFY messages throughout the life of the dialog. If this is not
        'done, and the user chooses a different view, on the next WM_NOTIFY message
        'the listview would be reset to the initial view, probably ticking off
        'the user. The variable is declared static to preserve values between
        'calls; it will be automatically reset on subsequent showing of the dialog.
         bLvSetupDone = False
         
        'other WM_INITDIALOG code here, such as caption or button changing, or
        '__centering the dialog.
'            If m_bvCenterView = False Then
'              pvDoCenterHook hwnd, uMsg, wParam, lParam
'            End If
'
      Case WM_NOTIFY
               
            If bLvSetupDone = False Then
               
              'hwnd is the handle to the dialog
              'hwndParent is the handle to the common control
              'hwndLv is the handle to the listview itself
               hwndParent = GetParent(hwnd)
               hwndLv = FindWindowEx(hwndParent, 0, "SHELLDLL_DefView", vbNullChar)
               
               If hwndLv > 0 Then
                  Call SendMessage(hwndLv, WM_COMMAND, ByVal m_lvInitialView, ByVal 0&)
                 
                 'since we found the lv hwnd, assume the
                 'command was received and set the flag
                 'to prevent recalling this routine
                 
                 ' Each time the folder is changed we reset the thumbnail view.
                 ' If you dont want this behaviour but will allow desktop.ini file to take over
                 ' comment out the next line.
                '__ bLvSetupDone = True
               End If  'hwndLv

            End If  'bLvSetupDone
            
        Case Else
         
   End Select

End Function

'Private Function pvDoCenterHook(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByRef lParam As OFNOTIFY) As Long
'
'   Dim hwndParent As Long
'   Dim rc As RECT
'   Dim newLeft As Long
'   Dim newTop As Long
'   Dim dlgWidth As Long
'   Dim dlgHeight As Long
'   Dim scrWidth As Long
'   Dim scrHeight As Long
'
'
'
'         '__obtain the handle to the parent dialog
'         hwndParent = GetParent(hwnd)
'
'         If hwndParent <> 0 Then
'           ' If lParam.hdr.code = CDN_INITDONE Then
'
'            'Just to prove the handle was obtained,
'            'change the dialog's caption.
'             Call SetWindowText(hwndParent, "I'm Hooked on Hooked Dialogs!")
'
'            'Position the dialog in the centre of
'            'the screen. First get the current dialog size.
'             Call GetWindowRect(hwndParent, rc)
'
'            '(To show the calculations involved, I've
'            'used variables instead of creating a
'            'one-line MoveWindow call)
'             dlgWidth = rc.Right - rc.Left
'             dlgHeight = rc.Bottom - rc.Top
'
'             scrWidth = Screen.Width \ Screen.TwipsPerPixelX
'             scrHeight = Screen.Height \ Screen.TwipsPerPixelY
'
'             newLeft = (scrWidth - dlgWidth) \ 2
'             newTop = (scrHeight - dlgHeight) \ 2
'
'            '..and set the new dialog position.
'             Call MoveWindow(hwndParent, newLeft, newTop, dlgWidth, dlgHeight, True)
'          'End If
'        End If
'End Function
'


