VERSION 5.00
Begin VB.Form frmStationUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "消息用户设置"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   Icon            =   "frmStationUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk 
      Caption         =   "默认用户(&D)"
      Height          =   195
      Left            =   1890
      TabIndex        =   0
      Top             =   765
      Width           =   1530
   End
   Begin VB.TextBox txt 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1875
      Width           =   3045
   End
   Begin VB.TextBox txt 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   8
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1485
      Width           =   3045
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   7
      Left            =   1890
      TabIndex        =   2
      Top             =   1080
      Width           =   3045
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   3840
      TabIndex        =   8
      Top             =   2535
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   345
      Left            =   2670
      TabIndex        =   7
      Top             =   2535
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   60
      TabIndex        =   9
      Top             =   2250
      Width           =   5010
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "确认密码(&R)"
      Height          =   180
      Index           =   0
      Left            =   870
      TabIndex        =   5
      Top             =   1935
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "连接密码(&P)"
      Height          =   180
      Index           =   8
      Left            =   870
      TabIndex        =   3
      Top             =   1545
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "连接用户(&U)"
      Height          =   180
      Index           =   7
      Left            =   870
      TabIndex        =   1
      Top             =   1140
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "工作站“FRCHEN”采用如下用户登录消息集成平台"
      Height          =   180
      Index           =   11
      Left            =   855
      TabIndex        =   10
      Top             =   300
      Width           =   3960
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   1
      Left            =   135
      Picture         =   "frmStationUser.frx":000C
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmStationUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrStation As String
Private mblnDataChanged As Boolean
Private mstrSQL As String

Public Event AfterDataChanged(ByVal strStation As String, ByVal strMipUser As String, ByVal strMipUserPassword As String)

Public Function ShowDialog(ByVal frmParent As Object, ByVal strStation As String, ByVal strMipUser As String, ByVal strMipUserPassword As String) As Boolean
    
    mblnDataChanged = False
    mstrStation = strStation
        
    Me.Caption = "工作站“" & mstrStation & "”消息用户"
    
    lbl(11).Caption = "为工作站配置对应的登录消息集成平台消息用户"
    
    chk.Value = IIf(strMipUser = "", 1, 0)
    
    txt(7).Text = strMipUser
    txt(8).Text = strMipUserPassword
    txt(0).Text = strMipUserPassword
    
    Call chk_Click
    
    Me.Show 1, frmParent
    ShowDialog = mblnDataChanged
    
End Function

Private Sub chk_Click()
    txt(7).Enabled = (chk.Value = 0)
    txt(8).Enabled = (chk.Value = 0)
    txt(0).Enabled = (chk.Value = 0)
    
    lbl(7).Enabled = (chk.Value = 0)
    lbl(8).Enabled = (chk.Value = 0)
    lbl(0).Enabled = (chk.Value = 0)
    
    Call LocationObj(txt(7))
End Sub

Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnDataChanged = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsPara As ADODB.Recordset
    
    If txt(8).Text <> txt(0).Text Then
        ShowSimpleMsg "确认密码和连接密码不一致，请重新输入确认密码！"
        Call LocationObj(txt(0))
    End If
    
    Set rsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(rsPara, "工作站", mstrStation)
    Call zlCommFun.SetParameter(rsPara, "消息用户", IIf(chk.Value = 1, "", txt(7).Text))
    Call zlCommFun.SetParameter(rsPara, "消息密码", IIf(chk.Value = 1, "", txt(8).Text))
    If gclsBusiness.ClientsEdit("UPDATE", rsPara) Then
        mblnDataChanged = True
        
        RaiseEvent AfterDataChanged(mstrStation, IIf(chk.Value = 1, "", txt(7).Text), IIf(chk.Value = 1, "", txt(8).Text))
        
        Unload Me
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0

'        Select Case Index
'        Case 1
'            If zlCommFun.FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
'        End Select
        
    End If
End Sub
