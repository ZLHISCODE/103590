VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.UserControl ucCommandBar 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12630
   ScaleHeight     =   5595
   ScaleWidth      =   12630
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   2100
      Left            =   375
      ScaleHeight     =   2100
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   375
      Width           =   3840
      Begin VB.PictureBox picCommandbar 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   390
         ScaleHeight     =   1065
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   405
         Width           =   2370
         Begin XtremeCommandBars.CommandBars cbsMain 
            Left            =   375
            Top             =   120
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            VisualTheme     =   2
         End
      End
   End
End
Attribute VB_Name = "ucCommandBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Public Event Resize()
Public Event Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
Public Event Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Public Event GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
Public Event ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Public Event InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
Public Event Customization(ByVal Options As XtremeCommandBars.ICustomizeOptions)

Private mbytBorderStyle As Byte

'######################################################################################################################
Public Property Get ActiveMenuBar() As CommandBar
    Set ActiveMenuBar = cbsMain.ActiveMenuBar
    
End Property

Public Property Get BorderStyle() As Byte
    BorderStyle = mbytBorderStyle
End Property

Public Property Let LargeIcon(ByVal blnData As Boolean)
    
    cbsMain.Options.LargeIcons = blnData
    Call UserControl_Resize
    
End Property

Public Property Let BorderStyle(ByVal bytData As Byte)
    mbytBorderStyle = bytData
    
    If mbytBorderStyle = 1 Then
        UserControl.BackColor = &H8000000D
    Else
        UserControl.BackColor = picBack.BackColor
    End If
End Property

Public Function FindControl(ByVal lngKey As Long) As CommandBarControl
    Set FindControl = cbsMain.FindControl(, lngKey)
End Function

Public Sub RefreshCtl()
    cbsMain.RecalcLayout
End Sub


'######################################################################################################################
Private Sub cbsMain_Customization(ByVal Options As XtremeCommandBars.ICustomizeOptions)
    RaiseEvent Customization(Options)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    RaiseEvent Execute(Control)
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    RaiseEvent GetClientBordersWidth(Left, Top, Right, Bottom)
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    RaiseEvent InitCommandsPopup(CommandBar)
End Sub

Private Sub cbsMain_Resize()
    RaiseEvent Resize
End Sub

Private Sub cbsMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    RaiseEvent ResizeClient(Left, Top, Right, Bottom)
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    RaiseEvent Update(Control)
End Sub

Private Sub picBack_Resize()
    On Error Resume Next
    picCommandbar.Move -120, 0, picBack.Width + 120, picBack.Height
End Sub

Private Sub UserControl_Initialize()
    mbytBorderStyle = 1
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    If cbsMain.Options.LargeIcons = False Then
        UserControl.Height = 390 + 30
    Else
        UserControl.Height = 525 + 30
    End If
    
    picBack.Move 15, 15, UserControl.Width - 30, UserControl.Height - 30
    
End Sub

Public Property Get ObjCommandBar() As Object
    Set ObjCommandBar = cbsMain
End Property


