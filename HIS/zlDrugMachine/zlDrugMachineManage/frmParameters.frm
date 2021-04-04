VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParameters 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8430
   Icon            =   "frmParameters.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtToken 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   32
      TabIndex        =   8
      Top             =   1230
      Width           =   3015
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   6
      Top             =   1230
      Width           =   1575
   End
   Begin VB.TextBox txtValidDays 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      MaxLength       =   1
      TabIndex        =   14
      Top             =   2355
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   7080
      TabIndex        =   26
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&O)"
      Height          =   345
      Left            =   5880
      TabIndex        =   25
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame fraLog 
      Caption         =   "日志选项"
      Height          =   855
      Left            =   240
      TabIndex        =   19
      Top             =   4440
      Width           =   7935
      Begin VB.TextBox txtSaveDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   24
         Top             =   345
         Width           =   615
      End
      Begin VB.OptionButton optType 
         Caption         =   "详细(调试)"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optType 
         Caption         =   "简要"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkOutput 
         Appearance      =   0  'Flat
         Caption         =   "日志输出(&L)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblSaveDays 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日志保存(&A)          (3-30)天"
         Height          =   180
         Left            =   4440
         TabIndex        =   23
         Top             =   390
         Width           =   2610
      End
   End
   Begin VB.Frame fraTime 
      Caption         =   "定时数据选项"
      Height          =   2175
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   7935
      Begin MSComctlLib.ListView lvwBusiness 
         Height          =   1335
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.TextBox txtViewLines 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1035
         Width           =   615
      End
      Begin VB.TextBox txtCycle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   16
         Top             =   675
         Width           =   615
      End
      Begin VB.Label lblViewLines 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "显示日志(&L)          (200-2000)行"
         Height          =   180
         Left            =   4440
         TabIndex        =   17
         Top             =   1080
         Width           =   2970
      End
      Begin VB.Label lblCycle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "定时周期(&Y)          (1-999)分钟"
         Height          =   180
         Left            =   4440
         TabIndex        =   15
         Top             =   720
         Width           =   2880
      End
      Begin VB.Label lblValidDays 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "有效天数(&V)          (1-5)天"
         Height          =   180
         Left            =   4440
         TabIndex        =   13
         Top             =   360
         Width           =   2520
      End
      Begin VB.Label lblBusiness 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "业务数据(&B)"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.CheckBox chkTimeTransmit 
      Appearance      =   0  'Flat
      Caption         =   "启用定时数据传送(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdTest 
      Height          =   300
      Left            =   7830
      Picture         =   "frmParameters.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   900
      Width           =   330
   End
   Begin VB.TextBox txtAddrIIP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   4000
      TabIndex        =   3
      Top             =   900
      Width           =   6615
   End
   Begin VB.CheckBox chkEnabledIIP 
      Appearance      =   0  'Flat
      Caption         =   "启用信息交互平台（&I)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.CheckBox chkEnabledMachine 
      Appearance      =   0  'Flat
      Caption         =   "启用药品自动化设备接口(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblToken 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "令牌(&T)"
      Height          =   180
      Left            =   4080
      TabIndex        =   7
      Top             =   1260
      Width           =   630
   End
   Begin VB.Label lblKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密钥(&K)"
      Height          =   180
      Left            =   510
      TabIndex        =   5
      Top             =   1260
      Width           =   630
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地址(&U)"
      Height          =   180
      Left            =   510
      TabIndex        =   2
      Top             =   930
      Width           =   630
   End
End
Attribute VB_Name = "frmParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnShow As Boolean                     '显示状态（Load事件后的过程处理）
Private mblnReturn As Boolean                   '返回值； True确认；False取消
Private mblnEdited As Boolean                   '是否已经编辑；True是；False否
Private mfrmOwner As Form

Public Function ShowMe(ByVal frmOwner As Form) As Boolean
    Set mfrmOwner = frmOwner
    Call mdlMain.VerifyConfigFile(App.Path & "\" & GSTR_CONFIG_FILE)    '检查参数文件
    Me.Show vbModal, frmOwner
    ShowMe = mblnReturn
End Function

Private Sub chkOutput_Click()
    optType(0).Enabled = chkOutput.Value = 1
    optType(1).Enabled = chkOutput.Value = 1
    optType(0).ForeColor = IIf(chkOutput.Value = 1, vbBlack, &H80000010)
    txtSaveDays.Enabled = chkOutput.Value = 1
    lblSaveDays.Enabled = chkOutput.Value = 1
    
    If Visible And mblnShow = False Then
        mblnEdited = True
    End If
End Sub

Private Sub chkEnabledIIP_Click()
    txtAddrIIP.Enabled = chkEnabledIIP.Value = 1
    cmdTest.Enabled = chkEnabledIIP.Value = 1
    txtKey.Enabled = chkEnabledIIP.Value = 1
    txtToken.Enabled = chkEnabledIIP.Value = 1
    
    If chkEnabledIIP.Value = 1 Then
        txtAddrIIP.ToolTipText = "请填写完整的地址！"
    Else
        txtAddrIIP.ToolTipText = ""
    End If
    
    If Visible And mblnShow = False Then
        mblnEdited = True
    End If
End Sub

Private Sub chkEnabledMachine_Click()
    If chkEnabledMachine.Value <> 1 Then
        chkTimeTransmit.Value = 0
        chkOutput.Value = 0
    End If
    
    chkEnabledIIP.Enabled = chkEnabledMachine.Value = 1
    chkTimeTransmit.Enabled = chkEnabledMachine.Value = 1
    chkOutput.Enabled = chkEnabledMachine.Value = 1
    
    If Visible And mblnShow = False Then
        mblnEdited = True
    End If
End Sub

Private Sub chkTimeTransmit_Click()
    fraTime.Enabled = chkTimeTransmit.Value = 1
    lvwBusiness.Enabled = fraTime.Enabled
    txtValidDays.Enabled = fraTime.Enabled
    txtCycle.Enabled = fraTime.Enabled
    txtViewLines.Enabled = fraTime.Enabled
    lblBusiness.Enabled = fraTime.Enabled
    lblValidDays.Enabled = fraTime.Enabled
    lblCycle.Enabled = fraTime.Enabled
    lblViewLines.Enabled = fraTime.Enabled
    lvwBusiness.ForeColor = IIf(fraTime.Enabled, vbBlack, &H80000010)
    
    If Visible And mblnShow = False Then
        mblnEdited = True
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '检查参数
    If VerifyParams = False Then Exit Sub
    
    '保存参数
    Call SaveParams

    Unload Me
    mblnReturn = True
End Sub

Private Sub cmdTest_Click()
    If mfrmOwner.mobjHTTP Is Nothing Then
        MsgBox "实例化“WinHttp”失败，请联系技术人员！", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    On Error Resume Next
    mfrmOwner.mobjHTTP.Open "POST", txtAddrIIP.Text
    If txtAddrIIP.Text = "" Then
        txtAddrIIP.Tag = ""
        MsgBox "请填写WEBService地址！", vbInformation, GSTR_MSG
    ElseIf Err.Number <> -2147012891 Then
        txtAddrIIP.Tag = "1"           '标记连接成功
        If cmdTest.Tag <> "1" Then
            MsgBox "连接成功！", vbInformation, GSTR_MSG
        End If
    Else
        txtAddrIIP.Tag = ""            '标记连接失败
        MsgBox "连接失败！", vbCritical, GSTR_MSG
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    If mblnShow Then
        Screen.MousePointer = vbHourglass
        
        '初始化控件与控件之间的关系
        Call InitControls
        
        mblnShow = False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    mblnReturn = False
    mblnEdited = False
    
    '
    
    mblnShow = True         '本行放最后
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnEdited And UnloadMode = 0 Then
        If MsgBox("是否放弃本次修改？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
End Sub

Private Sub InitControls()
    Dim strFile As String, strTemp As String, strChoose As String
    Dim arrTemp As Variant
    Dim i As Integer
    Dim lsiTemp As ListItem
    
    strFile = App.Path & "\" & GSTR_CONFIG_FILE
    
    '获取模块参数
    chkEnabledMachine.Value = Val(gobjComLib.zlDatabase.GetPara("启用药品自动化设备接口", GLNG_SYSTEM, GLNG_MODULE))
    
    strTemp = gobjComLib.zlDatabase.GetPara("启用信息交互平台", GLNG_SYSTEM, GLNG_MODULE)
    chkEnabledIIP.Value = Val(strTemp)
    If InStr(strTemp, "|") > 0 Then
        txtAddrIIP.Text = Split(strTemp, "|")(1)
    End If
    
    strTemp = gobjComLib.zlDatabase.GetPara("信息交互平台密钥", GLNG_SYSTEM, GLNG_MODULE)
    If strTemp <> "" Then
        If Not gobjEncrypt Is Nothing Then
            txtKey.Text = gobjEncrypt.Base64Decode(strTemp)
        Else
            txtKey.Text = ""
        End If
    Else
        txtKey.Text = ""
    End If
    
    strTemp = gobjComLib.zlDatabase.GetPara("信息交互平台令牌", GLNG_SYSTEM, GLNG_MODULE)
    If strTemp <> "" Then
        If Not gobjEncrypt Is Nothing Then
            txtToken.Text = gobjEncrypt.Base64Decode(strTemp)
        Else
            txtToken.Text = ""
        End If
    Else
        txtToken.Text = ""
    End If
    
    '读取参数文件的信息
    If gobjXML.OpenXMLFile(strFile) = False Then
        MsgBox "管理工具的参数文件不正确！", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    chkTimeTransmit.Value = Val(GetParameter(gobjXML, "enabled", "0"))
    txtCycle.Text = Val(GetParameter(gobjXML, "cycle"))
    txtValidDays.Text = Val(GetParameter(gobjXML, "validdays"))
    txtViewLines.Text = Val(GetParameter(gobjXML, "viewlines"))
    chkOutput.Value = Val(GetParameter(gobjXML, "output"))
    optType(0).Value = Val(GetParameter(gobjXML, "detailed")) = 0
    optType(1).Value = Not optType(0).Value
    txtSaveDays.Text = Val(GetParameter(gobjXML, "savedays"))
    
    strChoose = LCase(GetParameter(gobjXML, "businessdata"))
    strChoose = "|" & strChoose & "|"
    
    With lvwBusiness
        .ColumnHeaders.Add , "Data", "数据", .Width
        .ListItems.Clear
    End With
    
    strTemp = mfrmOwner.SupportData
    arrTemp = Split(strTemp & "|", "|")
    For i = LBound(arrTemp) To UBound(arrTemp)
        If Trim(arrTemp(i)) <> "" Then
            Set lsiTemp = lvwBusiness.ListItems.Add(, "K_" & i, arrTemp(i))
            If strChoose Like "*|" & LCase(Trim(lsiTemp.Text)) & "|*" Then
                lsiTemp.Checked = True
            End If
        End If
    Next
    Erase arrTemp
    
    gobjXML.CloseXMLDocument
    
    '控件关系
    Call chkEnabledMachine_Click
    Call chkEnabledIIP_Click
    Call chkTimeTransmit_Click
    Call chkOutput_Click
    Call txtAddrIIP_Change
    
End Sub

Private Function VerifyParams() As Boolean
    Dim i As Integer
    Dim blnFind As Boolean, blnCancel As Boolean
    
    '信息平台WebService地址
    If chkEnabledIIP.Value = 1 Then
        '地址
        If Val(txtAddrIIP.Tag) <> 1 Then
            cmdTest.Tag = "1"
            Call cmdTest_Click
            cmdTest.Tag = ""
            If Val(txtAddrIIP.Tag) <> 1 Then
                txtAddrIIP.SetFocus
                Exit Function
            End If
        End If
        If LenB(StrConv(txtAddrIIP.Text, vbFromUnicode)) > txtAddrIIP.MaxLength Then
            MsgBox "密钥填写超长（1-4000字符）！", vbInformation, GSTR_MSG
            txtAddrIIP.SetFocus
            Exit Function
        End If
        
        '密钥
        If txtKey.Text = "" Then
            MsgBox "请填写密钥！", vbInformation, GSTR_MSG
            txtKey.SetFocus
            Exit Function
        End If
        If Len(txtKey.Text) > txtKey.MaxLength Then
            MsgBox "密钥填写超长（1-16字符）！", vbInformation, GSTR_MSG
            txtKey.SetFocus
            Exit Function
        End If
        
        '令牌
        If txtToken.Text = "" Then
            MsgBox "请填写令牌！", vbInformation, GSTR_MSG
            txtToken.SetFocus
            Exit Function
        End If
        If Len(txtToken.Text) > txtToken.MaxLength Then
            MsgBox "密钥填写超长（1-32字符）！", vbInformation, GSTR_MSG
            txtToken.SetFocus
            Exit Function
        End If
    End If
    
    '业务数据
    If lvwBusiness.Enabled Then
        For i = 1 To lvwBusiness.ListItems.Count
            If lvwBusiness.ListItems(i).Checked Then
                blnFind = True
            End If
        Next
        If blnFind = False Then
            MsgBox "启用定时数据传送后，“业务数据”至少需要选定一个业务数据。", vbInformation, GSTR_MSG
            lvwBusiness.SetFocus
            Exit Function
        End If
    End If
    
    '有效天数
    If txtValidDays.Enabled Then
        Call txtValidDays_Validate(blnCancel)
        If blnCancel Then Exit Function
    End If
    
    '定时周期
    If txtCycle.Enabled Then
        Call txtCycle_Validate(blnCancel)
        If blnCancel Then Exit Function
    End If
    
    '显示日志行数
    If txtViewLines.Enabled Then
        Call txtViewLines_Validate(blnCancel)
        If blnCancel Then Exit Function
    End If
    
    '日志保存
    If txtSaveDays.Enabled Then
        Call txtSaveDays_Validate(blnCancel)
        If blnCancel Then Exit Function
    End If
    
    VerifyParams = True
    
End Function


Private Sub SaveParams()
    Dim strFile As String, strBusiness As String, strIIP As String, strEncryptKey As String, strToken As String
    Dim i As Integer
    
    strFile = App.Path & "\" & GSTR_CONFIG_FILE
    strIIP = IIf(chkEnabledIIP.Value = 1, "1", "0") & "|" & Trim(txtAddrIIP.Text)
    If txtKey.Text <> "" Then
        strEncryptKey = gobjEncrypt.Base64Encode(txtKey.Text)
    End If
    If txtToken.Text <> "" Then
        strToken = gobjEncrypt.Base64Encode(txtToken.Text)
    End If
    
    '模块参数
    Call gobjComLib.zlDatabase.SetPara("启用药品自动化设备接口", IIf(chkEnabledMachine.Value = 1, "1", "0"), GLNG_SYSTEM, GLNG_MODULE)
    Call gobjComLib.zlDatabase.SetPara("启用信息交互平台", strIIP, GLNG_SYSTEM, GLNG_MODULE)
    Call gobjComLib.zlDatabase.SetPara("信息交互平台密钥", strEncryptKey, GLNG_SYSTEM, GLNG_MODULE)
    Call gobjComLib.zlDatabase.SetPara("信息交互平台令牌", strToken, GLNG_SYSTEM, GLNG_MODULE)
    
    '本地参数
    For i = 1 To lvwBusiness.ListItems.Count
        If lvwBusiness.ListItems(i).Checked Then
            strBusiness = strBusiness & "|" & lvwBusiness.ListItems(i).Text
        End If
    Next
    If Left(strBusiness, 1) = "|" Then strBusiness = Mid(strBusiness, 2)
    
    If gobjXML.OpenXMLFile(strFile) Then
        Call gobjXML.SetSingleNodeValue("enabled", IIf(chkTimeTransmit.Value = 1, "1", "0"))
        Call gobjXML.SetSingleNodeValue("businessdata", strBusiness)
        Call gobjXML.SetSingleNodeValue("cycle", txtCycle.Text)
        Call gobjXML.SetSingleNodeValue("validdays", txtValidDays.Text)
        Call gobjXML.SetSingleNodeValue("viewlines", txtViewLines.Text)
        Call gobjXML.SetSingleNodeValue("output", IIf(chkOutput.Value = 1, "1", "0"))
        Call gobjXML.SetSingleNodeValue("detailed", IIf(optType(0).Value, "0", "1"))
        Call gobjXML.SetSingleNodeValue("savedays", txtSaveDays.Text)
    End If
    gobjXML.SaveXMLFile strFile
    gobjXML.CloseXMLDocument
End Sub

Private Sub txtAddrIIP_Change()
    If Visible And mblnShow = False And mblnEdited = False Then
        mblnEdited = True
    End If
    txtAddrIIP.Tag = ""
    cmdTest.Enabled = Trim(txtAddrIIP.Text) <> "" And txtAddrIIP.Enabled
End Sub

Private Sub txtAddrIIP_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtAddrIIP)
End Sub

Private Sub txtAddrIIP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtCycle_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtCycle)
End Sub

Private Sub txtCycle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtCycle_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtCycle_Validate(Cancel As Boolean)
    If Val(txtCycle.Text) < 1 Or Val(txtCycle.Text) > 999 Then
        MsgBox "“定时周期”填写不正确！", vbInformation, GSTR_MSG
        Cancel = True
    End If
End Sub

Private Sub txtKey_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtKey)
End Sub

Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtSaveDays_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtSaveDays)
End Sub

Private Sub txtSaveDays_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtSaveDays_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtSaveDays_Validate(Cancel As Boolean)
    If Val(txtSaveDays.Text) < 3 Or Val(txtSaveDays.Text) > 30 Then
        MsgBox "“日志保存”填写不正确！", vbInformation, GSTR_MSG
        Cancel = True
    End If
End Sub

Private Sub txtToken_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtToken)
End Sub

Private Sub txtToken_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtValidDays_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtValidDays)
End Sub

Private Sub txtValidDays_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtValidDays_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtValidDays_Validate(Cancel As Boolean)
    If Val(txtValidDays.Text) < 1 Or Val(txtValidDays.Text) > 5 Then
        MsgBox "“有效天数”填写不正确！", vbInformation, GSTR_MSG
        Cancel = True
    End If
End Sub

Private Sub txtViewLines_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtViewLines)
End Sub

Private Sub txtViewLines_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtViewLines_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtViewLines_Validate(Cancel As Boolean)
    If Val(txtViewLines.Text) < 200 Or Val(txtViewLines.Text) > 2000 Then
        MsgBox "“显示日志”填写不正确！", vbInformation, GSTR_MSG
        Cancel = True
    End If
End Sub
