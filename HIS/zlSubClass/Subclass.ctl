VERSION 5.00
Begin VB.UserControl Subclass 
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   4500
   ToolboxBitmap   =   "Subclass.ctx":0000
   Begin VB.Image imgIcon 
      Height          =   420
      Left            =   0
      Picture         =   "Subclass.ctx":00FA
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "Subclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "SoftCircuits Subclass Control"
Option Explicit

Private m_hWnd As Long
Private m_Messages() As Long
Private m_NumMessages As Integer

Event WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)

'Hooks or unhooks the specified message
Public Property Let Messages(nMessage As Long, bSubclass As Boolean)
    Dim i As Integer, j As Integer
    'Look up existing entry for this message
    For i = 1 To m_NumMessages
        If m_Messages(i) = nMessage Then
            If bSubclass Then
                'Message already subclassed
                Exit Property
            Else
                'Remove this message
                m_NumMessages = m_NumMessages - 1
                For j = i To m_NumMessages
                    m_Messages(j) = m_Messages(j + 1)
                Next j
                ReDim Preserve m_Messages(m_NumMessages)
                Exit Property
            End If
        End If
    Next i
    'Add message if not found
    If bSubclass Then
        'Add new hook for this window
        m_NumMessages = m_NumMessages + 1
        ReDim Preserve m_Messages(m_NumMessages)
        m_Messages(m_NumMessages) = nMessage
    End If
End Property

'Returns True if the specified message is currently hooked
Public Property Get Messages(nMessage As Long) As Boolean
Attribute Messages.VB_Description = "Specifies which messages are passed to the WndProc event"
Attribute Messages.VB_MemberFlags = "400"
    Dim i As Integer
    'Look up entry for this message
    For i = 1 To m_NumMessages
        If m_Messages(i) = nMessage Then
            Messages = True
            Exit Property
        End If
    Next i
    'No entry for this message
    Messages = False
End Property

'Hook specified window
Public Property Let hWnd(hWndNew As Long)
    'Only if hWnd has changed
    If hWndNew <> m_hWnd Then
        'Clear existing hook (if any)
        If m_hWnd <> 0 Then
            UnhookWindow m_hWnd
        End If
        m_hWnd = hWndNew
        'Hook new window (if any)
        If m_hWnd <> 0 Then
            HookWindow m_hWnd, Me
        End If
        'Note: No need to call PropertyChanged
        'because this property is not saved
    End If
End Property

'Return currently-hooked window
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Specifies the handle of the window to be subclassed"
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = m_hWnd
End Property

'Call default window procedure
Public Function CallWndProc(Msg As Long, wParam As Long, lParam As Long) As Long
Attribute CallWndProc.VB_Description = "Invokes the original window procedure for the subclassed window"
    If m_hWnd <> 0 Then
        CallWndProc = WinProc.CallWndProc(m_hWnd, Msg, wParam, lParam)
    End If
End Function

'Invoke WndProc event (called from BAS-module WndProc)
Friend Function RaiseWndProc(Msg As Long, wParam As Long, lParam As Long) As Long
    Dim Result As Long
    RaiseEvent WndProc(Msg, wParam, lParam, Result)
    RaiseWndProc = Result
End Function

'Force design-time control to size of icon
Private Sub UserControl_Resize()
    Size imgIcon.Width, imgIcon.Height
End Sub

'Unhook window if still hooked
Private Sub UserControl_Terminate()
    If m_hWnd <> 0 Then
        UnhookWindow m_hWnd
    End If
End Sub
