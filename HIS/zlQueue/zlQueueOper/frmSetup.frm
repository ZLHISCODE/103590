VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "语音设置"
   ClientHeight    =   6930
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6855
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton optCallWay 
      Caption         =   "远端语音播放"
      Height          =   360
      Index           =   0
      Left            =   255
      TabIndex        =   2
      Top             =   5400
      Width           =   1410
   End
   Begin VB.Frame fram远端语音设置 
      Height          =   855
      Left            =   135
      TabIndex        =   3
      Top             =   5445
      Width           =   6615
      Begin VB.ComboBox cboRemotePlaykStation 
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   5205
      End
      Begin VB.Label Label14 
         Caption         =   "远端站点名："
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   400
         Width           =   1215
      End
   End
   Begin VB.OptionButton optCallWay 
      Caption         =   "本地语音播放"
      Height          =   270
      Index           =   1
      Left            =   255
      TabIndex        =   6
      Top             =   105
      Value           =   -1  'True
      Width           =   1410
   End
   Begin VB.Frame frm语音广播设置 
      Height          =   5205
      Left            =   135
      TabIndex        =   19
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox chkHintSound 
         Caption         =   "呼叫前播放提示音"
         Height          =   240
         Left            =   1665
         TabIndex        =   22
         Top             =   375
         Width           =   1860
      End
      Begin VB.CheckBox chkUseSound 
         Caption         =   "启用语音呼叫"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin RichTextLib.RichTextBox rtbVBS 
         Height          =   3255
         Left            =   390
         TabIndex        =   7
         Top             =   1860
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   5741
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         TextRTF         =   $"frmSetup.frx":06EA
      End
      Begin VB.TextBox txt广播时间长度 
         Height          =   270
         Left            =   1800
         TabIndex        =   8
         Top             =   825
         Width           =   615
      End
      Begin VB.TextBox txtPlayCount 
         Height          =   270
         Left            =   4935
         TabIndex        =   9
         Top             =   825
         Width           =   615
      End
      Begin VB.ComboBox cboSoundType 
         Height          =   300
         ItemData        =   "frmSetup.frx":0787
         Left            =   4530
         List            =   "frmSetup.frx":0789
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   330
         Width           =   1995
      End
      Begin VB.TextBox txtSpeed 
         Height          =   270
         Left            =   1425
         TabIndex        =   11
         Text            =   "6"
         Top             =   1170
         Width           =   495
      End
      Begin VB.TextBox txtLoopQueryTime 
         Height          =   270
         Left            =   5685
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "30"
         Top             =   1170
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "自定义呼叫脚本编辑："
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   1620
         Width           =   1860
      End
      Begin VB.Label Label13 
         Caption         =   "语音播放速度为"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1875
      End
      Begin VB.Label Label12 
         Caption         =   "每段语音播放时长为        秒       语音循环播放次数为        次"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   855
         Width           =   5775
      End
      Begin VB.Label Label11 
         Caption         =   "语音类型："
         Height          =   210
         Left            =   3630
         TabIndex        =   15
         Top             =   390
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "(-10到10之间) "
         Height          =   255
         Left            =   1965
         TabIndex        =   16
         Top             =   1215
         Width           =   1260
      End
      Begin VB.Label Label9 
         Caption         =   "秒"
         Height          =   255
         Left            =   6300
         TabIndex        =   17
         Top             =   1215
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "轮询访问语音数据时间间隔为"
         Height          =   255
         Left            =   3285
         TabIndex        =   18
         Top             =   1215
         Width           =   2400
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   5655
      TabIndex        =   1
      Top             =   6405
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&S)"
      Height          =   400
      Left            =   4470
      TabIndex        =   0
      Top             =   6405
      Width           =   1100
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const M_STR_DEFAULT_VBS As String = "sub CusVoicePlay(lngCallId,strCallContext)" & vbCrLf & _
                                            "    Dim i                                          " & vbCrLf & _
                                            "                                                   " & vbCrLf & _
                                            "    SpVoice.Rate = 0                               " & vbCrLf & _
                                            "    SpVoice.Volume = 100                           " & vbCrLf & _
                                            "                                                   " & vbCrLf & _
                                            "    'Lili呼叫中文和英文                            " & vbCrLf & _
                                            "    Set SpVoice.Voice = SpVoice.GetVoices(""" & "Name=Microsoft Lili" & """).Item(0)" & vbCrLf & _
                                            "    SpVoice.Speak strCallContext, 1                " & vbCrLf & _
                                            "                                                   " & vbCrLf & _
                                            "    'Anna只能呼叫英文                              " & vbCrLf & _
                                            "    Set SpVoice.Voice = SpVoice.GetVoices(""" & "Name=Microsoft Anna" & """).Item(0)" & vbCrLf & _
                                            "    SpVoice.Speak strCallContext, 1                " & vbCrLf & _
                                            "End Sub                                            "

Private mlngModule As Long
Private mblnOk As Boolean



Public Function ShowMe(objOwner As Object) As Boolean
    mblnOk = False
    Call Me.Show(1, objOwner)
    
    ShowMe = mblnOk
End Function

Private Sub cbo语速_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cboSoundType_Click()
On Error GoTo errHandle
    If cboSoundType.Text = "自定义脚本呼叫" Then
        rtbVBS.Enabled = True
    Else
        rtbVBS.Enabled = False
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub chkUseSound_Click()
On Error GoTo errHandle
    Dim blnUseLocalPlay As Boolean
    Dim lngBackColor As Long
    
    blnUseLocalPlay = IIf(chkUseSound.value <> 0, True, False)
    lngBackColor = IIf(chkUseSound.value <> 0, &H80000005, Me.BackColor)
    
    Label10.Enabled = blnUseLocalPlay
    Label11.Enabled = blnUseLocalPlay
    Label12.Enabled = blnUseLocalPlay
    Label13.Enabled = blnUseLocalPlay
    Label2.Enabled = blnUseLocalPlay
    Label9.Enabled = blnUseLocalPlay
    
    txt广播时间长度.Enabled = blnUseLocalPlay
    txt广播时间长度.BackColor = lngBackColor
    
    txtPlayCount.Enabled = blnUseLocalPlay
    txtPlayCount.BackColor = lngBackColor
    
    txtSpeed.Enabled = blnUseLocalPlay
    txtSpeed.BackColor = lngBackColor
    
    txtLoopQueryTime.Enabled = blnUseLocalPlay
    txtLoopQueryTime.BackColor = lngBackColor
    
    cboSoundType.Enabled = blnUseLocalPlay
    cboSoundType.BackColor = lngBackColor

    rtbVBS.Enabled = blnUseLocalPlay
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cmdCancel_Click()
    '关闭窗口
    mblnOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    '保存参数设置
    
    If optCallWay(0).value = True Then
        If Trim(cboRemotePlaykStation.Text) = "" Then
            MsgBox "远端站点名不允许为空，请重新选择。", vbOKOnly, Me.Caption
            cboRemotePlaykStation.SetFocus
            Exit Sub
        End If
    End If
    
    
    If Val(txtLoopQueryTime.Text) > 65 Then
        MsgBox "轮询访问语音数据时间间隔不能大于65秒，请重新设置。", vbOKOnly, Me.Caption
        
        txtLoopQueryTime.SetFocus
        Call zlControl.TxtSelAll(txtLoopQueryTime)
    End If

    SaveSetting "ZLSOFT", gstrRegPath, "语音播放时长", Val(txt广播时间长度.Text)            'Call zlDatabase.SetPara("语音广播时间长度", Val(txt广播时间长度.Text), glngSys, glngModul)
    SaveSetting "ZLSOFT", gstrRegPath, "语音播放语速", Val(txtSpeed.Text)                   'Call zlDatabase.SetPara("语音广播语速", Val(txtSpeed.Text), glngSys, glngModul)

    SaveSetting "ZLSOFT", gstrRegPath, "启用语音呼叫", chkUseSound.value                    'Call zlDatabase.SetPara("启用语音呼叫", chkUseSound.value, glngSys, glngModul)
    SaveSetting "ZLSOFT", gstrRegPath, "语音呼叫前播放提示音", chkHintSound.value
    
    SaveSetting "ZLSOFT", gstrRegPath, "远端呼叫站点", cboRemotePlaykStation.Text           'Call zlDatabase.SetPara("远端呼叫站点", cboWorkStation.Text, glngSys, glngModul)
    SaveSetting "ZLSOFT", gstrRegPath, "语音播放次数", IIf(Val(txtPlayCount.Text) <= 0, 0, Val(txtPlayCount.Text))   'Call zlDatabase.SetPara("语音播放次数", Val(txtPlayCount.Text), glngSys, glngModul)

    '语音类型
    SaveSetting "ZLSOFT", gstrRegPath, "语音类型", cboSoundType.Text                        'Call zlDatabase.SetPara("语音类型", cboSoundType.Text, glngSys, glngModul)
    '轮询时间
    SaveSetting "ZLSOFT", gstrRegPath, "轮询间隔时间", IIf(Val(txtLoopQueryTime.Text) <= 0, 30, Val(txtLoopQueryTime.Text))     'Call zlDatabase.SetPara("轮询时间", Val(txtLoopQueryTime.Text), glngSys, glngModul)

    SaveSetting "ZLSOFT", gstrRegPath, "启用VBS自定义呼叫", IIf(Trim(cboSoundType.Text) = "自定义脚本呼叫", 1, 0)
    SaveSetting "ZLSOFT", gstrRegPath, "VBS脚本", rtbVBS.Text
    
    '保存叫号方式
    If optCallWay(0).value Then
        SaveSetting "ZLSOFT", gstrRegPath, "播放方式", 1                                    'Call zlDatabase.SetPara("叫号方式", 1, glngSys, glngModul)
    Else
        SaveSetting "ZLSOFT", gstrRegPath, "播放方式", 0                                    'Call zlDatabase.SetPara("叫号方式", 0, glngSys, glngModul)
    End If

    mblnOk = True
    '关闭窗口
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ReadWorkStationInf()
'*****************************************************
'读取站点信息
'*****************************************************

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset

    strSql = "select 工作站 from zlClients where 禁止使用<>1 order by 工作站"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "读取站点信息")

    cboRemotePlaykStation.Clear
    If rsTemp.EOF Then Exit Sub

    While Not rsTemp.EOF
        Call cboRemotePlaykStation.AddItem(rsTemp("工作站"))
        rsTemp.MoveNext
    Wend
    
End Sub

Private Sub LoadMSSoundType()
    Dim objVoice As Object
    Dim objToken As Object
    
    Set objVoice = CreateObject("SAPI.SPVoice")

    cboSoundType.Clear
    If objVoice Is Nothing Then Exit Sub
    
    For Each objToken In objVoice.GetVoices()
        cboSoundType.AddItem objToken.GetAttribute("Name")
    Next
    
    cboSoundType.AddItem "自定义脚本呼叫"
    
    cboSoundType.ListIndex = 0
End Sub

Private Sub ReadLocalPara()
    Dim i As Integer
    Dim lng广播语速 As Long
    Dim strSoundType As String

   '读取叫号方式
    cboRemotePlaykStation.Text = GetSetting("ZLSOFT", gstrRegPath, "远端呼叫站点", "")     'zlDatabase.GetPara("远端呼叫站点", glngSys, glngModule, "")
    cboRemotePlaykStation.Enabled = optCallWay(0).value

    chkUseSound.value = Val(GetSetting("ZLSOFT", gstrRegPath, "启用语音呼叫", 1))   'zlDatabase.GetPara("启用语音呼叫", glngSys, glngModul, "1")

    txtLoopQueryTime.Text = Val(GetSetting("ZLSOFT", gstrRegPath, "轮询间隔时间", 30))      ' Val(zlDatabase.GetPara("轮询时间", glngSys, glngModul, "30"))
    txt广播时间长度.Text = Val(GetSetting("ZLSOFT", gstrRegPath, "语音播放时长", 15))       'Val(zlDatabase.GetPara("语音广播时间长度", glngSys, glngModul, "15"))
    txtPlayCount.Text = Val(GetSetting("ZLSOFT", gstrRegPath, "语音播放次数", 2))           'Val(zlDatabase.GetPara("语音播放次数", glngSys, glngModul, "3"))

    If cboSoundType.Enabled = True Then                                                     'zlDatabase.GetPara("语音类型", glngSys, glngModul, "系统默认")
        strSoundType = Trim(GetSetting("ZLSOFT", gstrRegPath, "语音类型", ""))
        
        For i = 0 To cboSoundType.ListCount - 1
            If cboSoundType.List(i) = strSoundType Then
                cboSoundType.ListIndex = i
                Exit For
            End If
        Next
        
        If cboSoundType.ListCount > 0 And cboSoundType.ListIndex < 0 Then cboSoundType.ListIndex = 0
    End If
    
    chkHintSound.value = Val(GetSetting("ZLSOFT", gstrRegPath, "语音呼叫前播放提示音", ""))
    
    lng广播语速 = Val(GetSetting("ZLSOFT", gstrRegPath, "语音播放语速", 0))                 'Val(zlDatabase.GetPara("语音广播语速", glngSys, glngModul, "0"))
    txtSpeed.Text = IIf(lng广播语速 <= 10 And lng广播语速 >= -10, lng广播语速, 0)
    
    rtbVBS.Text = GetSetting("ZLSOFT", gstrRegPath, "VBS脚本", M_STR_DEFAULT_VBS)


    rtbVBS.Enabled = IIf(cboSoundType.Text = "自定义脚本呼叫", True, False)

    optCallWay(0).value = Val(GetSetting("ZLSOFT", gstrRegPath, "播放方式", 1))     'Val(zlDatabase.GetPara("叫号方式", glngSys, glngModule, "0"))
    optCallWay(1).value = Not optCallWay(0).value
    
    Call optCallWay_Click(0)
End Sub

Private Sub Form_Load()
    Call LoadMSSoundType
    
    Call ReadWorkStationInf
    
    Call ReadLocalPara
End Sub


Private Sub optCallWay_Click(Index As Integer)
    
    chkUseSound.Enabled = Not optCallWay(0).value
    cboRemotePlaykStation.Enabled = optCallWay(0).value
    
    If optCallWay(0).value Then
        frm语音广播设置.Enabled = False
        
        txt广播时间长度.BackColor = Me.BackColor
        txtPlayCount.BackColor = Me.BackColor
        txtSpeed.BackColor = Me.BackColor
        rtbVBS.BackColor = Me.BackColor
        txtLoopQueryTime.BackColor = Me.BackColor
        cboSoundType.BackColor = Me.BackColor
        
        cboRemotePlaykStation.BackColor = &H80000005
    Else
        frm语音广播设置.Enabled = True
        
        txt广播时间长度.BackColor = &H80000005
        txtPlayCount.BackColor = &H80000005
        txtSpeed.BackColor = &H80000005
        rtbVBS.BackColor = &H80000005
        txtLoopQueryTime.BackColor = &H80000005
        cboSoundType.BackColor = &H80000005
        
        cboRemotePlaykStation.BackColor = Me.BackColor
    End If
End Sub


Private Sub txt广播时间长度_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
