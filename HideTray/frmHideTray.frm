VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmHideTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   720
      Top             =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   55
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hide all the tray icons and the system clock.


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private TrayNot As Long, Shell_Tray As Long
Private Down As Boolean

Const RSP_SIMPLE_SERVICE = 1
Const RSP_UNREGISTER_SERVICE = 0
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Private Sub HideApp(Hide As Boolean)
    'to hide the App in the CTRL-ALT-SUPR Window (95/98)
    On Error Resume Next
    Dim ProcessID As Long, retval As Long
    ProcessID = GetCurrentProcessId()

    If Hide Then
        retval = RegisterServiceProcess(ProcessID, RSP_SIMPLE_SERVICE)
    Else
        retval = RegisterServiceProcess(ProcessID, RSP_UNREGISTER_SERVICE)
    End If
End Sub

Private Sub ToTray()
 
    Shell_Tray = FindWindow("Shell_TrayWnd", vbNullString)
    TrayNot = FindWindowEx(Shell_Tray, 0, "TrayNotifyWnd", vbNullString)

    SetParent Me.hwnd, TrayNot
    ' push my form into the tray notify window
End Sub


Private Sub Form_Load()
    HideApp True
    ToTray
    Timer1_Timer
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'to Exit: CTRL+ Left Click -> CTRL+ Right Click

If Button = vbLeftButton Then
    If Shift And vbCtrlMask Then Down = True
Else
    If Down Then
        If Shift And vbCtrlMask Then
            Unload Me
        End If
    End If
End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetParent Me.hwnd, &H0
'    SetParent Clock, TrayNot
End Sub


Private Sub Timer1_Timer()
    'new clock
    Label1.Caption = Format$(Time$, "HH:MM")
    Label1.Left = ((Me.ScaleWidth + Label1.Width) / Screen.TwipsPerPixelX) + 100
End Sub
