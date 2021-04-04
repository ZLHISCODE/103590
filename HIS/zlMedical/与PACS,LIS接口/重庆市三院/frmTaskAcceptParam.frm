VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTaskAcceptParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "接受参数"
   ClientHeight    =   1845
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5280
   Icon            =   "frmTaskAcceptParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComCtl2.UpDown udn 
      Height          =   300
      Left            =   2851
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1020
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "txt"
      BuddyDispid     =   196609
      OrigLeft        =   3180
      OrigTop         =   1035
      OrigRight       =   3435
      OrigBottom      =   1260
      Max             =   30
      Min             =   5
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt 
      Height          =   300
      Left            =   1905
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1020
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4050
      TabIndex        =   4
      Top             =   525
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4050
      TabIndex        =   3
      Top             =   120
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4110
      Left            =   3930
      TabIndex        =   5
      Top             =   -750
      Width           =   30
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   180
      Top             =   210
      Width           =   465
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检查间隔(&U)"
      Height          =   180
      Index           =   1
      Left            =   885
      TabIndex        =   0
      Top             =   1095
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "设置接受检验数据的时间间隔，若在这个时间间隔里有新的检验数据产生，则进行接受处理。"
      Height          =   570
      Left            =   885
      TabIndex        =   6
      Top             =   165
      Width           =   2880
   End
End
Attribute VB_Name = "frmTaskAcceptParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private mblnOK As Boolean

Public Function ShowParam(ByVal frmMain As Object) As Boolean

    mblnStartUp = True
    mblnOK = False
    
    Call LoadData
    
    Me.Show 1, frmMain
    
    ShowParam = mblnOK
    
End Function

Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  装载数据
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    txt.Text = Val(GetSetting("ZLSOFT", "公共全局\检验接口", "接受间隔", "10"))
    If Val(txt.Text) < 5 Then txt.Text = "5"
    If Val(txt.Text) > 30 Then txt.Text = "30"
    
    LoadData = True
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
End Function

Private Function SaveData() As Boolean

    On Error GoTo errHand
    
    Call SaveSetting("ZLSOFT", "公共全局\检验接口", "接受间隔", Val(txt.Text))
    
    SaveData = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
    
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mblnOK = True
        Unload Me
    End If
End Sub


Private Sub txt_GotFocus()

    TxtSelAll txt

End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim strFilter As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        PressKey vbKeyTab
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
    
    If Val(txt.Text) < udn.Min Or Val(txt.Text) > udn.Max Then
        Cancel = True
    End If
End Sub

