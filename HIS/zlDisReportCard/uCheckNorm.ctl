VERSION 5.00
Begin VB.UserControl uCheckNorm 
   BackColor       =   &H80000005&
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   FillColor       =   &H80000005&
   ScaleHeight     =   1830
   ScaleWidth      =   2175
   Begin VB.Shape shpFocus 
      BorderStyle     =   3  'Dot
      Height          =   240
      Left            =   1125
      Top             =   1080
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FFC0C0&
      Caption         =   "散居"
      Height          =   195
      Left            =   825
      TabIndex        =   0
      Top             =   420
      Width           =   465
   End
   Begin VB.Shape shpClick 
      BorderColor     =   &H80000001&
      Height          =   210
      Left            =   540
      Top             =   1095
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape shpBoder 
      Height          =   240
      Left            =   465
      Top             =   645
      Width           =   225
   End
   Begin VB.Label lblChecked 
      BackColor       =   &H00FFFFFF&
      Caption         =   "√"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   870
      TabIndex        =   1
      Top             =   735
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "uCheckNorm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mblnChecked As Boolean
Private mblnCheckedVisible As Boolean
Private mintChedkType As Integer
Public Event Change()
Public Enum eCheckType
    eSingle = 0
    eMulti = 1
End Enum

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get CheckType() As eCheckType
    CheckType = mintChedkType
End Property
Public Property Let CheckType(ByVal val As eCheckType)
    mintChedkType = val
    PropertyChanged "CheckType"
End Property

Public Property Get CheckedVisible() As Boolean
    CheckedVisible = mblnCheckedVisible
End Property
Public Property Let CheckedVisible(ByVal val As Boolean)
    mblnCheckedVisible = val
    If mblnCheckedVisible Then
        UserControl.Width = shpBoder.Width + lblCaption.Width + 45
    Else
        BoxVisible = False
        UserControl.Width = lblCaption.Width + 30
    End If
    PropertyChanged "CheckedVisible"
End Property

Public Property Get BoxVisible() As Boolean
    BoxVisible = shpBoder.Visible
End Property
Public Property Let BoxVisible(ByVal val As Boolean)
    shpBoder.Visible = val
    If val Then
        CheckedVisible = True
    End If
    PropertyChanged "BoxVisible"
End Property

Public Property Get Checked() As Boolean
    Checked = mblnChecked
End Property
Public Property Let Checked(ByVal val As Boolean)
    On Error Resume Next
    mblnChecked = val
    If mblnCheckedVisible Then
        lblChecked.Visible = mblnChecked
    End If
    
    If val Then
        If CheckType = eSingle Then
            If UserControl.Parent.Visible Then '当控件为单选时，将其它同类控件置为未选中
                Dim ot As Object, oi As Integer
                For Each ot In UserControl.ParentControls
                    If TypeName(ot) = UserControl.Name And ot.Name = UserControl.Extender.Name Then
                        If ot.hWnd <> UserControl.Extender.hWnd Then
                            ot.Checked = Not val
                        End If
                    End If
                Next
            End If
        End If
    End If
    Err.Clear
    PropertyChanged "Checked"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal val As OLE_COLOR)
    On Error Resume Next
    UserControl.BackColor = val
    lblChecked.BackColor = val
    lblCaption.BackColor = val
    If val = shpClick.BackColor Then
        shpClick.BackColor = &HFF&
    End If
    Err.Clear
    PropertyChanged "BackColor"
End Property

Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property
Public Property Let Caption(ByVal val As String)
    lblCaption.Caption = val
    PropertyChanged "Caption"
End Property

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub

Private Sub lblChecked_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub

Private Sub lblChecked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub

Private Sub UserControl_GotFocus()
    shpFocus.Visible = True
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    UserControl.Width = shpBoder.Width + lblCaption.Width + 45
    UserControl.Height = shpBoder.Height + 30
    shpBoder.Move 15, 15
    lblCaption.Move shpBoder.Left + shpBoder.Width + 15, (UserControl.Height - lblCaption.Height) / 2, UserControl.Width - shpBoder.Width - 45
    '底色
    shpBoder.BackColor = UserControl.BackColor
    shpClick.BackColor = UserControl.BackColor
    lblChecked.BackColor = UserControl.BackColor
    lblCaption.BackColor = UserControl.BackColor
    Err.Clear
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeySpace Or KeyCode = vbKeyReturn) And BoxVisible Then
        shpClick.Visible = True
    End If
End Sub

Public Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        If shpClick.Visible Then shpClick.Visible = False
        Checked = Not mblnChecked
        RaiseEvent Change
    End If
End Sub

Private Sub UserControl_LostFocus()
    shpFocus.Visible = False
End Sub

Public Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If BoxVisible Then shpClick.Visible = True
End Sub

Public Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If shpClick.Visible Then shpClick.Visible = False
    If x > UserControl.Width Or y > UserControl.Height Then Exit Sub
    If BoxVisible Then
        Checked = Not mblnChecked
        RaiseEvent Change
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    shpClick.Move shpBoder.Left + 15, shpBoder.Top + 15
    UserControl.Height = shpBoder.Height + 30
    lblChecked.Move shpBoder.Left + 30, shpBoder.Top + 30
    If CheckedVisible Then
        lblCaption.Move shpBoder.Left + shpBoder.Width + 15, (UserControl.Height - lblCaption.Height) / 2, UserControl.Width - shpBoder.Width - 45
    Else
        lblCaption.Move shpBoder.Left, (UserControl.Height - lblCaption.Height) / 2, UserControl.Width - 30
    End If
    shpFocus.Move lblCaption.Left - 15, lblCaption.Top - 15, lblCaption.Width + 15, lblCaption.Height + 30
End Sub
Private Sub UserControl_InitProperties()
    Checked = False
    BackColor = vbWindowBackground
    Caption = UserControl.Name
    CheckType = eMulti
    BoxVisible = True
    CheckedVisible = True
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    Caption = PropBag.ReadProperty("Caption", UserControl.Name)
    CheckType = PropBag.ReadProperty("CheckType", eSingle)
    BoxVisible = PropBag.ReadProperty("BoxVisible", True)
    CheckedVisible = PropBag.ReadProperty("CheckedVisible", True)
    If CheckedVisible Then
        UserControl.Width = shpBoder.Width + lblCaption.Width + 45
    Else
        UserControl.Width = lblCaption.Width + 30
    End If
    Checked = PropBag.ReadProperty("Checked", False)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    PropBag.WriteProperty "Checked", Checked, False
    PropBag.WriteProperty "BackColor", BackColor, &HFFFFFF
    PropBag.WriteProperty "Caption", Caption, UserControl.Name
    PropBag.WriteProperty "CheckType", CheckType, eSingle
    PropBag.WriteProperty "BoxVisible", BoxVisible, True
    PropBag.WriteProperty "CheckedVisible", CheckedVisible, True
    
    PropertyChanged "Checked"
    PropertyChanged "BackColor"
    PropertyChanged "Caption"
    PropertyChanged "CheckType"
    PropertyChanged "BoxVisible"
    PropertyChanged "CheckedVisible"
End Sub

