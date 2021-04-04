VERSION 5.00
Begin VB.UserControl Progress 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   345
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Progress.ctx":0000
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   170
      Left            =   1800
      TabIndex        =   0
      Top             =   90
      Width           =   1005
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'######################################################################################
'##ģ �� ����Progress.ctl
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��5��20��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ����һ���Զ���ķ����Ľ������ؼ���
'##��    ����
'######################################################################################

Option Explicit
Private mvarValue As Single

Public Property Get Value() As Single
    Value = mvarValue
End Property

Public Property Let Value(vData As Single)
    mvarValue = vData
    lblProgress.Caption = Format(vData, "0%")
    DrawProgress mvarValue, UserControl.hdc, 0, 0, ScaleWidth / Screen.TwipsPerPixelX, ScaleHeight / Screen.TwipsPerPixelY
    Refresh
    PropertyChanged "Value"
End Property

Public Sub Cls()
    UserControl.Cls
End Sub

Private Sub UserControl_Initialize()
    mvarValue = 0#
End Sub

Private Sub UserControl_Resize()
    lblProgress.Move 0, (ScaleHeight - lblProgress.Height) / 2, ScaleWidth
End Sub
