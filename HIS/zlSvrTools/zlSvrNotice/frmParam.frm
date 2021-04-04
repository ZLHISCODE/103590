VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "服务器设置"
   ClientHeight    =   2340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4515
   Icon            =   "frmParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   1470
      Left            =   105
      TabIndex        =   7
      Top             =   135
      Width           =   4245
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1410
         TabIndex        =   1
         Top             =   345
         Width           =   2040
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Left            =   3450
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1024
         BuddyControl    =   "txt(1)"
         BuddyDispid     =   196611
         BuddyIndex      =   1
         OrigLeft        =   2580
         OrigTop         =   645
         OrigRight       =   2820
         OrigBottom      =   885
         Max             =   9999
         Min             =   1024
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1410
         TabIndex        =   3
         Top             =   690
         Width           =   2040
      End
      Begin VB.Label lbl 
         Caption         =   "通讯IP地址(&I)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "通讯端口号(&P)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   735
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3240
      TabIndex        =   6
      Top             =   1770
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   5
      Top             =   1770
      Width           =   1100
   End
End
Attribute VB_Name = "frmParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mfrmMain As Form
Private mlngPort As Long
Private mstrLocalIP As String
Private mbytState As Byte

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngPort As Long, ByRef strLocalIP As String, ByVal bytState As Byte) As Boolean
    
    Dim rs As New ADODB.Recordset
    
    Set mfrmMain = frmMain
    mstrLocalIP = strLocalIP
    mbytState = bytState
    
    mlngPort = lngPort
    txt(1).Text = mlngPort
    txt(0).Text = mstrLocalIP
        
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
    lngPort = mlngPort
    strLocalIP = mstrLocalIP
    
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim lngLoop As Long
    
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    
    strTmp = txt(1).Text
    
    If strTmp = "" Then
        MsgBox "必须输入提醒服务的端口号！", , gstrSysName
        txt(1).SetFocus
        cmdCancel.Enabled = True
        cmdOK.Enabled = True
        Exit Sub
    End If
    
    If strTmp = "0" Then
        MsgBox "提醒服务的端口号必须大于零！", , gstrSysName
        txt(1).SetFocus
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        Exit Sub
    End If
                
    '检查是否为全数字
    For lngLoop = 1 To Len(txt(1).Text)
        If Mid(strTmp, lngLoop, 1) < "0" Or Mid(strTmp, lngLoop, 1) > "9" Then
            MsgBox "提醒服务的端口号必须为数值字符！", , gstrSysName
            txt(1).SetFocus
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            Exit Sub
        End If
    Next
    
    If mlngPort <> Val(txt(1).Text) Or mstrLocalIP <> Trim(txt(0).Text) Then
        
        '测试是否正确
        If mfrmMain.UpdateRefresh(Val(txt(1).Text), Trim(txt(0).Text)) = False Then
            txt(1).SetFocus
            cmdCancel.Enabled = True
            cmdOK.Enabled = True
            Exit Sub
        End If
        
        mlngPort = Val(txt(1).Text)
        mstrLocalIP = Trim(txt(0).Text)
    End If
    
    On Error Resume Next
    
    '填写数据库
    gcnOracle.Execute "update zloptions set 参数值='" & mstrLocalIP & ";" & mlngPort & ";" & mbytState & "' where 参数号=7"
    
    
    mblnOK = True
    
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not cmdCancel.Enabled
End Sub

Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index).Text)
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdOK.SetFocus
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Dim strTmp As String
    Dim lngLoop As Long
    
    If Index <> 1 Then Exit Sub
    
    strTmp = txt(1).Text
  
    If strTmp = "" Then
        MsgBox "必须输入提醒服务的端口号！", , gstrSysName
        txt(1).SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If strTmp = "0" Then
        MsgBox "提醒服务的端口号必须大于零！", , gstrSysName
        txt(1).SetFocus
        Cancel = True
        Exit Sub
    End If
                
    '检查是否为全数字
    For lngLoop = 1 To Len(txt(1).Text)
        If Mid(strTmp, lngLoop, 1) < "0" Or Mid(strTmp, lngLoop, 1) > "9" Then
            MsgBox "提醒服务的端口号必须为数值字符！", , gstrSysName
            txt(1).SetFocus
            Cancel = True
            Exit Sub
        End If
    Next
    
    If Val(strTmp) < udn.Min Or Val(strTmp) > udn.Max Then
        MsgBox "提醒服务的端口号必须大于等于1024小于等9999！", , gstrSysName
        txt(1).SetFocus
        Cancel = True
        Exit Sub
    End If

End Sub
