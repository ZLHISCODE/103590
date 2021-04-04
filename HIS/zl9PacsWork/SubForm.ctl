VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.UserControl SubForm 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   6240
   ToolboxBitmap   =   "SubForm.ctx":0000
   Begin VB.Label labClose 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "¡Á"
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label labEvent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin XtremeSuiteControls.ShortcutCaption scTitle 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      _Version        =   589884
      _ExtentX        =   9975
      _ExtentY        =   529
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌו"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColorLight=   16761024
      GradientColorDark=   16744576
   End
End
Attribute VB_Name = "SubForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Event OnResize()
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event OnKeyDown(KeyCode As Integer, Shift As Integer)
Public Event OnKeyUp(KeyCode As Integer, Shift As Integer)
Public Event OnKeyPress(KeyAscii As Integer)

Public Event OnClick()
Public Event OnDblClick()

Public Event OnEnterFocus()
Public Event OnExitFocus()
Public Event OnClose()


Private mBorderColor As OLE_COLOR
Private mblnDrawBorder As Boolean



Property Get Title() As String
    Title = scTitle.Caption
End Property

Property Let Title(value As String)
    scTitle.Caption = value
End Property



Property Get BorderColor() As OLE_COLOR
    BorderColor = mBorderColor
End Property

Property Let BorderColor(value As OLE_COLOR)
    mBorderColor = value
    
    Call UserControl_Paint
    Call AdjustFace
End Property



Property Get DrawBorder() As Boolean
    DrawBorder = mblnDrawBorder
End Property

Property Let DrawBorder(value As Boolean)
    mblnDrawBorder = value
    
    Call UserControl.Refresh
    Call UserControl_Paint
    Call AdjustFace
End Property



Private Sub AdjustFace()
    On Error Resume Next
    scTitle.Left = 0
    scTitle.Top = 0
    scTitle.Width = UserControl.Width
    
    
    labEvent.Left = 0
    labEvent.Top = 0
    labEvent.Width = scTitle.Width - labClose.Width
    labEvent.Height = scTitle.Height
    
    labClose.Left = UserControl.Width - labClose.Width
    labClose.Top = 30
    
End Sub


Private Sub DrawBoder()
    UserControl.Line (5, 0)-(UserControl.Width - 10, UserControl.Height - 10), mBorderColor, B
End Sub




Private Sub labClose_Click()
    On Error Resume Next
    
    RaiseEvent OnClose
End Sub

Private Sub labClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labClose.ForeColor = vbRed
End Sub

Private Sub labClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     labClose.ForeColor = vbWhite
End Sub

Private Sub labEvent_Click()
    Call UserControl_Click
End Sub


Private Sub labEvent_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub labEvent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub labEvent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub labEvent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labEvent.ForeColor = vbWhite
    
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
    On Error Resume Next
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnClick
End Sub

Private Sub UserControl_DblClick()
    On Error Resume Next
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnDblClick
End Sub

Private Sub UserControl_EnterFocus()
    RaiseEvent OnEnterFocus
End Sub

Private Sub UserControl_ExitFocus()
    On Error Resume Next
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnExitFocus
End Sub

Private Sub UserControl_Initialize()
    Call AdjustFace
End Sub


Private Sub UserControl_InitProperties()
    scTitle.Caption = "SubForm"
    mBorderColor = &H80000006
    mblnDrawBorder = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnKeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnKeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnKeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    On Error Resume Next
    
    If mblnDrawBorder Then Call DrawBoder
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    scTitle.Caption = PropBag.ReadProperty("Title", "SubForm")
    mBorderColor = PropBag.ReadProperty("BorderColor", &H80000006)
    mblnDrawBorder = PropBag.ReadProperty("DrawBorder", True)
End Sub

Private Sub UserControl_Resize()
    Call AdjustFace
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    Call PropBag.WriteProperty("Title", scTitle.Caption, "SubForm")
    Call PropBag.WriteProperty("BorderColor", mBorderColor, &H80000006)
    Call PropBag.WriteProperty("DrawBorder", mblnDrawBorder, True)
End Sub
