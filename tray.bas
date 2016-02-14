Attribute VB_Name = "tray"
Option Explicit
Private Const APP_SYSTRAY_ID = 927 ' # Identifier

Private Const NOTIFYICON_VERSION = &H3

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

' Icon flags
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10
   
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

' Icon
Public Const BALOON_ICO_NONE = 0
Public Const BALOON_ICO_INFO = 1
Public Const BALOON_ICO_WARNING = 2
Public Const BALOON_ICO_ERROR = 3

' NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504
Private NOTIFYICONDATA_SIZE As Long

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const WM_LBUTTONUP = &H202   '514
Public Const WM_RBUTTONUP = &H205   '517
Public Const WM_MOUSEMOVE = &H200   '512
Public Const WM_LBUTTONDBLCLK = &H203


Public Sub ShellTrayAdd(strTip As String)
   
   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = principal.IconoBandeja.hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
      .dwState = NIS_SHAREDICON
      .hIcon = principal.IconoBandeja.Picture
      .uCallbackMessage = WM_MOUSEMOVE
      .szTip = strTip & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
      
   End With
   
   Call Shell_NotifyIcon(NIM_ADD, nid)
   Call Shell_NotifyIcon(NIM_SETVERSION, nid)
       
End Sub


Public Sub ShellTrayRemove()

   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
      
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = principal.IconoBandeja.hwnd
      .uID = APP_SYSTRAY_ID
   End With
   
   Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub


Public Sub ShellTrayModifyTip(nIconIndex As Long, strTitle As String, strMess As String)

   Dim nid As NOTIFYICONDATA

   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = principal.IconoBandeja.hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = nIconIndex
      .szInfoTitle = strTitle & vbNullChar
      .szInfo = strMess & vbNullChar
   End With

   Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub


Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select

End Sub


Private Function IsShellVersion(ByVal version As Long) As Boolean

   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
      ReDim bBuffer(nBufferSize - 1) As Byte
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
         IsShellVersion = nVerMajor >= version
      End If
   End If
  
End Function
