VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.2#0"; "DicomObjects.ocx"
Begin VB.Form frmPacsSendImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发送图像"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   Icon            =   "frmPacsSendImage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Caption         =   "图像预预览"
      Height          =   3915
      Left            =   5160
      TabIndex        =   31
      Top             =   60
      Width           =   4425
      Begin DicomObjects.DicomViewer Viewer 
         Height          =   3135
         Left            =   120
         TabIndex        =   35
         Top             =   570
         Width           =   4155
         _Version        =   262146
         _ExtentX        =   7329
         _ExtentY        =   5530
         _StockProps     =   35
      End
      Begin VB.CheckBox ChkShowImage 
         Caption         =   "预览"
         Height          =   285
         Left            =   150
         TabIndex        =   34
         Top             =   210
         Value           =   1  'Checked
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "查询"
      Height          =   1545
      Left            =   30
      TabIndex        =   25
      Top             =   4020
      Width           =   9555
      Begin VB.TextBox txtPatientID 
         Height          =   300
         Left            =   7140
         MaxLength       =   18
         TabIndex        =   4
         Top             =   420
         Width           =   2130
      End
      Begin VB.CommandButton CmdRefresh 
         Cancel          =   -1  'True
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   8190
         TabIndex        =   9
         Top             =   1020
         Width           =   1100
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   3690
         MaxLength       =   12
         TabIndex        =   6
         Top             =   1050
         Width           =   1590
      End
      Begin VB.CheckBox chk来源 
         Caption         =   "门诊病人"
         Height          =   195
         Index           =   0
         Left            =   5460
         TabIndex        =   7
         Top             =   1110
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chk来源 
         Caption         =   "住院病人"
         Height          =   195
         Index           =   1
         Left            =   6750
         TabIndex        =   8
         Top             =   1110
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1050
         Width           =   3270
      End
      Begin VB.TextBox txtChkNoEnd 
         Height          =   300
         Left            =   5460
         MaxLength       =   18
         TabIndex        =   3
         Top             =   420
         Width           =   1260
      End
      Begin VB.TextBox txtChkNoBegin 
         Height          =   300
         Left            =   3690
         MaxLength       =   18
         TabIndex        =   2
         Top             =   420
         Width           =   1260
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   25165825
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   150
         TabIndex        =   0
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   25165825
         CurrentDate     =   38082
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标识号"
         Height          =   180
         Left            =   7140
         TabIndex        =   36
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Left            =   5130
         TabIndex        =   30
         Top             =   480
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Left            =   1680
         TabIndex        =   29
         Top             =   480
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓  名"
         Height          =   180
         Left            =   3690
         TabIndex        =   23
         Top             =   810
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人来源"
         Height          =   180
         Left            =   5460
         TabIndex        =   24
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人科室"
         Height          =   180
         Left            =   150
         TabIndex        =   22
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查号"
         Height          =   180
         Left            =   3690
         TabIndex        =   21
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查时间"
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   210
         Width           =   720
      End
   End
   Begin MSComctlLib.TreeView tvwImageDate 
      Height          =   3915
      Left            =   30
      TabIndex        =   28
      Top             =   60
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   6906
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8340
      TabIndex        =   18
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "发送(&S)"
      Height          =   350
      Left            =   7005
      TabIndex        =   17
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton CmdSelectAll 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   5685
      TabIndex        =   16
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton CmdSelectClear 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   4350
      TabIndex        =   15
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   180
      TabIndex        =   19
      Top             =   6750
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "发送设置"
      Height          =   945
      Left            =   60
      TabIndex        =   26
      Top             =   5640
      Width           =   9525
      Begin VB.OptionButton ChkImageFormat 
         Caption         =   "JPG"
         Height          =   225
         Index           =   2
         Left            =   5730
         TabIndex        =   13
         Top             =   510
         Width           =   1005
      End
      Begin VB.OptionButton ChkImageFormat 
         Caption         =   "BMP"
         Height          =   225
         Index           =   1
         Left            =   4770
         TabIndex        =   12
         Top             =   510
         Width           =   1005
      End
      Begin VB.OptionButton ChkImageFormat 
         Caption         =   "DICOM"
         Height          =   225
         Index           =   0
         Left            =   3630
         TabIndex        =   11
         Top             =   510
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.CheckBox ChkSendTool 
         Caption         =   "发送独立观片站"
         Height          =   315
         Left            =   7290
         TabIndex        =   14
         Top             =   450
         Width           =   2025
      End
      Begin VB.ComboBox CboPath 
         Height          =   300
         ItemData        =   "frmPacsSendImage.frx":000C
         Left            =   150
         List            =   "frmPacsSendImage.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发送图像格式"
         Height          =   180
         Left            =   3660
         TabIndex        =   33
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发送目录"
         Height          =   180
         Left            =   150
         TabIndex        =   32
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   27
      Top             =   7245
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmPacsSendImage.frx":0010
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12409
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPacsSendImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************API调用*****************************************
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)


Private Sub Check1_Click()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub ShowMe(frmobj As Object)
    
    Me.Show vbModal, frmobj
End Sub

Private Sub CmdRefresh_Click()
    Me.CmdRefresh.Enabled = False
    If zlCommFun.StrIsValid(Me.txtChkNoBegin) = False Then
        With Me.txtChkNoBegin
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    If zlCommFun.StrIsValid(Me.txtChkNoEnd) = False Then
        With Me.txtChkNoEnd
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    If zlCommFun.StrIsValid(Me.txtChkNoEnd) = False Then
        With Me.txtChkNoEnd
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    zl9comlib.zlCommFun.ShowFlash "请等待正在读取数据.....", Me
    zl9comlib.zlCommFun.ShowFlash
    RefreshImageDate
    zl9comlib.zlCommFun.StopFlash
    AllSelectOrAllClear True
    Me.CmdRefresh.Enabled = True
End Sub

Private Sub CmdSelectAll_Click()
    AllSelectOrAllClear True
End Sub

Private Sub CmdSelectClear_Click()
    AllSelectOrAllClear False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdSend_Click()
    '下载图后再发送
    Dim i As Long, j As Long, m As Long, n As Long
    Dim strPath As String                   '发送路径
    Dim strDate As String                   '节点时间
    Dim strName As String                   '节点姓名
    Dim strSq As String                     '节点序列号
    Dim xNodes As Node                      '节点对象
    Dim blWriteSucceed  As Boolean          '写入是否成功
    Dim strSql As String                    '存放SQL语句变量
    Dim strTmp As String                    '临时字串分解变量
    Dim strPas As String                    '远程目录密码
    Dim strUse As String                    '远程用户名
    Dim rsTmp As New ADODB.Recordset        '临时记录集
    Dim duTime As Double                    '记录时间秒用于连接网络超时
    Dim strRemotePath As String             '远程目录路径
    Dim DicomPath As New DicomDataSet       '生成DIR文件
    Dim objFile As New Scripting.FileSystemObject           '复制文件使用
    On Error GoTo SendErr
    
    If Me.CboPath.Text = "" Then
        MsgBox "请选择要发送的目录!", vbInformation, Me.Caption
        Exit Sub
    End If
    '检查目录是否可写
    
    If Me.CboPath.List(Me.CboPath.ListIndex) <> "" Then
        strTmp = Mid(Me.CboPath.Text, 1, InStr(1, Me.CboPath.Text, "_") - 1)
        strSql = "select 设备号,本机目录,用户名,密码 from 影像设备目录 where 类型 = 5 and 设备号 = [1]"
        Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTmp)
        If rsTmp.EOF = True Then
            MsgBox "没有找到要发送目录的信息！", vbInformation, Me.Caption
            Exit Sub
        End If
        zl9comlib.zlCommFun.ShowFlash "正在连接测试请等待.....", Me
        zl9comlib.zlCommFun.ShowFlash
        strRemotePath = rsTmp("本机目录")
        strPas = Nvl(rsTmp("密码"))
        strUse = Nvl(rsTmp("用户名"))
        Shell "net use " & strRemotePath & " " & strPas & " /user:" & strUse, vbHide
            
        duTime = Timer
        Do Until CLng(Timer - duTime) >= 20
            Shell "net use " & strRemotePath & " " & strPas & " /user:" & strUse, vbHide
            If WriteTest(False, strRemotePath) = True Then
                Exit Do
            End If
            DoEvents
        Loop
        zl9comlib.zlCommFun.StopFlash
    End If
    
    If WriteTest(False, strRemotePath) = False Then
        MsgBox "写入测试失败请检查共享目录!", vbQuestion, App.EXEName
        Exit Sub
    End If
    
    With Me.tvwImageDate
        If .Nodes.Count = 0 Then
            MsgBox "没有可以发送的文件!请选择查询条件后点击刷新更新列表!", vbInformation, App.EXEName
            Exit Sub
        End If
        '清除以前的错误日志
        If Dir(App.Path & "\WriteErrLog.txt") <> vbNullString Then
            Kill App.Path & "\WriteErrLog.txt"
        End If
        
        Me.CmdOK.Enabled = False
        Me.cmdCancel.Enabled = False
        Me.CmdSelectAll.Enabled = False
        Me.CmdSelectClear.Enabled = False
        Me.CmdRefresh.Enabled = False
        Me.MousePointer = 11
        zlCommFun.ShowFlash "正在读入文件请等待.....", Me
        zlCommFun.ShowFlash
        For i = 1 To .Nodes.Count
            If .Nodes(i).Checked = True And .Nodes(i).Child Is Nothing Then
                strDate = Mid(.Nodes(i).Parent.Parent.Text, InStr(1, .Nodes(i).Parent.Parent.Text, "[") + 1)
                strDate = Format(Mid(strDate, 1, InStr(1, strDate, "]") - 1), "yyyymmdd")
                strName = Mid(.Nodes(i).Parent.Parent.Text, 1, InStr(1, .Nodes(i).Parent.Parent.Text, "[") - 1)
                strSq = .Nodes(i).Key
                strPath = strDate & "\" & strName & "\" & Mid(strSq, 1, InStr(strSq, "_") - 1)
                '发送
                If SendFilesToDir(strSq, DicomPath, strRemotePath & "\DICOM\IMAGE\", strPath) Then
                    '失败
                    .Nodes(i).Checked = False
                    blWriteSucceed = False
                    m = m + 1
                Else
                    '成功
                    blWriteSucceed = True
                    n = n + 1
                End If
                j = j + 1
            End If
            '新病人时记数清零
            If .Nodes(i).Parent Is Nothing Then
                j = 0
            End If
            DoEvents
            Me.stbThis.Panels(2).Text = "正在发送病人[" & strName & "]的第" & j & "个图像. 已完成" & CInt((i) / .Nodes.Count * 100) & "%."
            .Nodes(i).Checked = blWriteSucceed
            blWriteSucceed = False
        Next
        
        If Me.ChkSendTool.Value = 1 Then
            If Dir(App.Path & "\PacsLite\") <> "" Then
                objFile.CopyFile App.Path & "\PacsLite\*.*", strRemotePath & "\"
            End If
        End If
        If m > 0 Then
            DicomPath.WriteDirectory IIf(Len(strRemotePath) > 3, strRemotePath & "\DICOM\DICOMDIR", strRemotePath & "\DICOM\DICOMDIR")
        End If
        zl9comlib.zlCommFun.StopFlash
        
        Me.stbThis.Panels(2).Text = "发送完成!发送成功" & m & "个图像,失败" & n & "个图像."
        If n > 0 Then
            If MsgBox("发送完成!发送成功" & m & "个图像,失败" & n & "个图像." & _
            vbCrLf & "查看日志请选择[是]", vbYesNo + vbDefaultButton2 + vbInformation, App.EXEName) = vbYes Then
                Shell "Notepad " & App.Path & "\WriteErrLog.txt", vbNormalFocus
            End If
        Else
            MsgBox "发送完成!发送成功" & m & "个图像,失败" & n & "个图像.", vbInformation, App.EXEName
        End If
        Me.MousePointer = 0
    End With
    Shell "net use " & strRemotePath & " /delete "
    Me.CmdOK.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.CmdSelectAll.Enabled = True
    Me.CmdSelectClear.Enabled = True
    Me.CmdRefresh.Enabled = True
    Exit Sub
SendErr:
    Me.MousePointer = 0
    zl9comlib.zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.CmdOK.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.CmdSelectAll.Enabled = True
    Me.CmdSelectClear.Enabled = True
    Me.CmdRefresh.Enabled = True
End Sub



Private Sub dtpBegin_Change()
    Me.dtpBegin.MaxDate = Me.dtpEnd.Value
End Sub

Private Sub dtpEnd_Change()
    Me.dtpEnd.MinDate = Me.dtpBegin.Value
End Sub


Private Function LoadData() As Boolean
'功能：根据病人来源读取病人科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim lngPre As Long
    
    
    If cboDept.ListIndex <> -1 Then
        lngPre = cboDept.ItemData(cboDept.ListIndex)
    End If
    strSql = "Select Distinct A.ID,A.编码,A.名称,B.服务对象" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.工作性质 IN('检查')" & _
        " And B.服务对象 IN(3," & IIf(chk来源(0).Value, 1, -1) & "," & IIf(chk来源(1).Value, 2, -1) & ")" & _
        " And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.编码"
    On Error GoTo errH
    Call OpenRecord(rsTmp, strSql, Me.Caption)
    On Error GoTo 0
    cboDept.Clear
    cboDept.AddItem "所有科室"
    cboDept.ListIndex = 0
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPre Then cboDept.ListIndex = cboDept.NewIndex
        rsTmp.MoveNext
    Next
    strSql = "select 设备号, 设备名,本机目录,用户名,密码  from 影像设备目录 where 类型 = 5"
    
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption)
    Me.CboPath.Clear
    Do Until rsTmp.EOF
        Me.CboPath.AddItem rsTmp("设备号") & "_" & rsTmp("设备名")
        rsTmp.MoveNext
    Loop
    If CboPath.ListCount > 0 Then
        CboPath.ListIndex = 0
    End If
    
    LoadData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chk来源_Click(Index As Integer)
    If chk来源(0).Value = 0 And chk来源(1).Value = 0 Then
        chk来源((Index + 1) Mod 2).Value = 1
    End If
    Call LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "间隔日期", CInt(Me.dtpEnd - Me.dtpBegin)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "结束检查时间", Format(Me.dtpEnd.Value, "yyyy-mm-dd")
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "科室", Me.cboDept.ListIndex
'    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "发送路径", Me.TxtPath
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "门诊病人", Me.chk来源(0).Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "住院病人", Me.chk来源(1).Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "发送目录", Me.CboPath.ListIndex
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "预览", Me.ChkShowImage.Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "格式0", Me.ChkImageFormat(0).Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "格式1", Me.ChkImageFormat(1).Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "格式2", Me.ChkImageFormat(2).Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "观片站", Me.ChkSendTool.Value
    
    Unload Me
End Sub
Private Sub Form_Load()
    Dim intDept As Integer
    Dim intDiffDay As Integer
    Dim intPath As String
    intDiffDay = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "间隔日期", 3)
    Me.dtpBegin = Format(Now - intDiffDay, "yyyy-mm-dd")
    Me.dtpEnd = Format(Now, "yyyy-mm-dd")
    intDept = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "科室", 0)
'    Me.TxtPath = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "发送路径", "")
    Me.chk来源(0).Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "门诊病人", 1)
    Me.chk来源(1).Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "住院病人", 1)
    Me.ChkShowImage.Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "预览", 1)
    Me.ChkImageFormat(0).Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "格式0", True)
    Me.ChkImageFormat(1).Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "格式1", False)
    Me.ChkImageFormat(2).Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "格式2", False)
    Me.ChkSendTool.Value = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "观片站", 1)
    intPath = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "发送目录", 0)
    LoadData
    SendMessage Me.cboDept.Hwnd, CB_SETCURSEL, intDept, 0
    SendMessage Me.CboPath.Hwnd, CB_SETCURSEL, intPath, 0
End Sub
Private Sub RefreshImageDate()
    '刷新病人图像
    Dim rsMain As New ADODB.Recordset
    Dim rsNode As New ADODB.Recordset
    Dim strSql As String
    Dim strDeptNo As String
    Dim intMZ As Integer
    Dim intZY As Integer
    Dim blnMoved As Boolean
    Dim strSQLBak As String
    Dim strSerialUID As String
    Dim strPatientUID As String
    Dim i As Integer
    On Error GoTo RefreshError
    
    blnMoved = MovedByDate(Me.dtpBegin.Value)
    
    strDeptNo = Me.cboDept.Text
    
    If strDeptNo <> "所有科室" Then
        strDeptNo = Mid(strDeptNo, 1, InStr(1, strDeptNo, "-") - 1)
    End If
    strSql = "select a.医嘱ID,a.影像类别,a.检查号,a.姓名,a.检查UID,a.接收日期,b.病人来源,c.编码||'-'||c.名称 as 部门名称, " & _
             " d.首次时间,e.序列UID,f.图像UID,e.序列描述 from 影像检查记录 a , " & _
             " 病人医嘱记录 b , 部门表 c , 病人医嘱发送 d , 影像检查序列 e , 影像检查图象 f , 病人信息 g " & _
             " where a.医嘱id = b.id and b.执行科室id = c.id and b.id = d.医嘱id and  " & _
             " a.检查UID = e.检查UID and e.序列UID = f.序列UID and b.病人ID = g.病人ID and " & _
             " d.首次时间 >= [1] and d.首次时间 <= [2] "
    If strDeptNo = "所有科室" Then
        strSql = strSql & " and [3] = [3] "
    Else
        strSql = strSql & " and c.编码 = [3] "
    End If
    If Trim(txtChkNoBegin) = "" Or Trim(txtChkNoEnd) = "" Then
        strSql = strSql & " and [4] = [4] and [5] = [5] "
    Else
        strSql = strSql & " and a.检查号 >= [4] and a.检查号 <= [5] "
    End If
    If Trim(txt姓名) = "" Then
        strSql = strSql & " and [6]= [6] "
    Else
        strSql = strSql & " and a.姓名 = [6] "
    End If
    If Me.chk来源(0).Value = 1 Then
        intMZ = 1
    Else
        intMZ = 3
    End If
    If Me.chk来源(1).Value = 1 Then
        intZY = 2
    Else
        intZY = 3
    End If
    strSql = strSql & " and  b.病人来源 in (3,4,[7],[8]) "
    If Len(Trim(Me.txtPatientID)) <= 0 Then
        strSql = strSql & " and [9] = [9] "
    Else
        strSql = strSql & " and Decode(B.病人来源,1,G.门诊号,2,G.住院号,NULL)= " & Me.txtPatientID
    End If
    If blnMoved Then
        strSQLBak = strSql
        strSQLBak = Replace(strSQLBak, "影像检查记录", "H影像检查记录")
        strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
        strSql = strSql & " Union ALL " & strSQLBak & " order by 检查UID,序列UID,图像UID"
    Else
        strSql = strSql & " order by 检查UID,序列UID,图像UId"
    End If
    Set rsMain = OpenSQLRecord(strSql, Me.Caption, CDate(Format(Me.dtpBegin.Value, "yyyy-mm-dd")), _
    CDate(Format(Me.dtpEnd.Value, "yyyy-mm-dd 23:59:59")), strDeptNo, _
    IIf(Trim(txtChkNoBegin) = "", 1, txtChkNoBegin), IIf(Trim(txtChkNoEnd) = "", 1, txtChkNoEnd), _
    IIf(Trim(txt姓名) = "", 1, txt姓名), intMZ, intZY, IIf(Trim(txtPatientID) = "", 1, txtPatientID))
    Me.tvwImageDate.Nodes.Clear
        
    Do Until rsMain.EOF
        With Me.tvwImageDate.Nodes
            '病人级
            If strPatientUID <> rsMain("检查UID") Then
                .Add , , "A" & rsMain("检查UID"), rsMain("姓名") & "[" & rsMain("接收日期") & "]"
            End If
            
            '检查序列级
            If strSerialUID <> rsMain("序列UID") Then
                .Add "A" & rsMain("检查UID"), tvwChild, rsMain("检查UID") & "_" & rsMain("序列UID"), "[" & rsMain("序列描述") & "]" & rsMain("序列UID")
            End If
            
            '图像级
            .Add rsMain("检查UID") & "_" & rsMain("序列UID"), tvwChild, rsMain("序列UID") & "_" & rsMain("图像UID"), rsMain("图像UID")
            
            DoEvents
            
            strPatientUID = rsMain("检查UID")
            strSerialUID = rsMain("序列UID")
            rsMain.MoveNext
        End With
    Loop
    
    
'    Do Until rsMain.EOF
'        With Me.tvwImageDate.Nodes
'            .Add , , "A" & rsMain("医嘱ID"), rsMain("姓名") & "[" & rsMain("接收日期") & "]"
'            strSQL = "select 序列UID,序列描述 from 影像检查序列 where 检查UID = [1]"
'            If blnMoved Then
'                strSQLBak = strSQL
'                strSQLBak = Replace(strSQLBak, "影像检查序列", "H影像检查序列")
'                strSQL = strSQL & " Union ALL " & strSQLBak
'            End If
'            Set rsNode = OpenSQLRecord(strSQL, Me.Caption, Nvl(rsMain("检查UID"), 0))
'            Do Until rsNode.EOF
'                .Add "A" & rsMain("医嘱ID"), tvwChild, "A" & rsNode("序列UID"), "[" & rsNode("序列描述") & "]" & rsNode("序列UID")
'                rsNode.MoveNext
'            Loop
'            rsNode.Close
'            Set rsNode = Nothing
'        End With
'        rsMain.MoveNext
'        DoEvents
'    Loop
    rsMain.Close
    Set rsMain = Nothing
    Exit Sub
RefreshError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvwImageDate_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim blSelOrCls As Boolean
    Dim blLoopEOF As Boolean
    Dim objNode As Node

    '病人级
    If Node.Parent Is Nothing Then
        SelectAllChild Node, Node.Checked
        Set objNode = Node.Child.FirstSibling
        objNode.Checked = Node.Checked
        
        For i = 1 To Node.Children
            objNode.Checked = Node.Checked
            SelectAllChild objNode, Node.Checked
            Set objNode = objNode.Next
        Next
    End If
    
    '序列级
    If Not Node.Parent Is Nothing And Not Node.Child Is Nothing Then
        SelectAllChild Node, Node.Checked
    End If
    
    If Node.Checked = True Then
        '选中
        If Not Node.Parent Is Nothing Then
            Node.Parent.Checked = True
            If Not Node.Parent.Parent Is Nothing Then
                Node.Parent.Parent.Checked = True
            End If
        End If
    Else
        '取消
        If Not Node.Parent Is Nothing Then
            Set objNode = Node.Parent.Child.FirstSibling
            '处理上一级
            For i = 1 To Node.Parent.Children
                If objNode.Checked = True Then
                    blLoopEOF = True
                    Exit For
                End If
                Set objNode = objNode.Next
            Next
            '处理上上级
            If blLoopEOF = False Then
                Node.Parent.Checked = False
                If Not Node.Parent.Parent Is Nothing Then
                    Set objNode = Node.Parent.Parent.Child.FirstSibling
                    For i = 1 To Node.Parent.Parent.Children
                        If objNode.Checked = True Then
                            blLoopEOF = True
                            Exit For
                        End If
                        Set objNode = objNode.Next
                    Next
                    If blLoopEOF = False Then
                        Node.Parent.Parent.Checked = False
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub AllSelectOrAllClear(TrueOrFalse As Boolean)
    With Me.tvwImageDate
        For i = 1 To .Nodes.Count
            .Nodes(i).Checked = TrueOrFalse
        Next
    End With
End Sub
'显示保存目录
Private Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    On Error GoTo OpenFileError
    With udtBI
        '设置浏览窗口
        .hWndOwner = lWindowHwnd
        '返回选中的目录
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "请选定开始搜索的文件夹："
        Else
            .lpszTitle = sTitle
        End If
    End With
    '调出浏览窗口
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '获取路径
        SHGetPathFromIDList lpIDList, sPath
        '释放内存
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
    Exit Function
OpenFileError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function WriteTest(ShowErrMsg As Boolean, strPath As String) As Boolean
    Dim strTmpPath As String
    On Error GoTo CopyError
    strTmpPath = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) & "temp.txt"
    Open strTmpPath For Output As #1
    Close #1
    FileCopy strTmpPath, IIf(Len(strPath) > 3, strPath & "\", strPath) & "temp.txt"
    Kill IIf(Len(strPath) > 3, strPath & "\", strPath) & "temp.txt"
    Kill strTmpPath
    WriteTest = True
    Exit Function
CopyError:
    If ShowErrMsg = False Then Exit Function
    If Err.Number = 75 Then
        MsgBox "写入测试失败!请查看[" & strPath & "]是否有写入权限!", vbInformation, App.EXEName
    Else
        MsgBox "发生其他错误！", vbQuestion, App.EXEName
    End If
End Function

Private Function SendFilesToDir(lngSeqUID As String, DicomDirPath As DicomDataSet, DestinationBoot As String, DestinationDir As String) As Boolean
    '功能:从FTP下载文件
    '传入:序列UID
    '返回下载成功后的文件路径
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strCachePath As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strTmpFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim DicomImg As New DicomImages
    Dim DicomImgFormat As String
    Dim ImageUID As String
    Dim SerialUID As String
    
    On Error GoTo WriteFileErr
    SendFilesToDir = True
    strSql = "Select A.图像号,D.用户名 As User1,D.密码 As Pwd1,a.图像UID, " & _
        "D.IP地址 As Host1," & _
        "'/'||D.Ftp目录||'/' As Root1,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL1,d.设备号 as 设备号1, " & _
        "E.用户名 As User2,E.密码 As Pwd2," & _
        "E.IP地址 As Host2," & _
        "'/'||E.Ftp目录||'/' As Root2,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL2 , e.设备号 as 设备号2 " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
        "And A.序列UID= [1] And A.图像UID = [2] Order By A.图像号"
    If mblnMoved Then
        strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
        strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
    SerialUID = Mid(lngSeqUID, 1, InStr(lngSeqUID, "_") - 1)
    ImageUID = Mid(lngSeqUID, InStr(lngSeqUID, "_") + 1)
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, SerialUID, ImageUID)
    strCachePath = App.Path & "\TmpImage\"
    ClearCacheFolder strCachePath
    If rsTmp.EOF <> True Then
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL1")))
    End If
    If Me.ChkImageFormat(1).Value = True Then
        DicomImgFormat = ".BMP"
    ElseIf Me.ChkImageFormat(2).Value = True Then
        DicomImgFormat = ".JPG"
    End If
    
    Do While Not rsTmp.EOF
        
'        If rsTmp("URL1") Is Nothing Then
'            strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
'        Else
'            strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
'        End If
            
'        Inet.strIPAddress = Nvl(rsTmp("Host1")): Inet.strUser = Nvl(rsTmp("User1")): Inet.strPsw = Nvl(rsTmp("Pwd1"))
        strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
    
        If strDeviceNO1 <> rsTmp("设备号1") Then
            strDeviceNO1 = rsTmp("设备号1")
            Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
        End If
        
        If strDeviceNO2 <> rsTmp("设备号2") Then
            strDeviceNO2 = rsTmp("设备号2")
            Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
        End If
        
        If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFile.GetFileName(rsTmp("URL1"))) <> 0 Then
'            Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
            Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFile.GetFileName(rsTmp("URL2")))
        End If

        On Error Resume Next
        MkLocalDir DestinationBoot & DestinationDir
        DicomImg.ReadFile Replace(strTmpFile, "/", "\")
        '保存格式
        If DicomImgFormat <> "" Then
            DicomImg(1).FileExport strTmpFile & DicomImgFormat, Mid(DicomImgFormat, 2)
        End If
        
        DicomDirPath.Name = "ZLPACS"
        DicomDirPath.AddToDirectory DicomImg(1), "IMAGE\" & DestinationDir & "\" & rsTmp("图像UId") & _
                                    DicomImgFormat, "1.2.840.10008.1.2.1", 0
        DicomImg.Clear
        Err.Clear
        
        If Dir(DestinationBoot & DestinationDir & "\" & rsTmp("图像UId")) = vbNullString Then
            FileCopy strTmpFile & DicomImgFormat, DestinationBoot & DestinationDir & "\" & rsTmp("图像UId") & DicomImgFormat
        End If
         
        If Err.Number <> 0 Then
            Open App.Path & "\WriteErrLog.txt" For Append As #1
                Print #1, "复制[" & strCachePath & Nvl(rsTmp("URL1")) & "]到[" & DestinationBoot & DestinationDir & "]" & vbCrLf & _
                "发生" & Err.Description & "错误号:" & Err.Number
            Close #1
            SendFilesToDir = False
        End If
        
        DoEvents
        rsTmp.MoveNext
    Loop
    Exit Function
WriteFileErr:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tvwImageDate_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strImageUID As String
    Dim strSerialUID As String
    
    If Node.Child Is Nothing And Me.ChkShowImage.Value = 1 Then
        strImageUID = Mid(Node.Key, InStr(Node.Key, "_") + 1)
        strSerialUID = Mid(Node.Key, 1, InStr(Node.Key, "_") - 1)
        ShowImage strImageUID, strSerialUID
    End If
End Sub

Private Sub txtChkNoBegin_GotFocus()
    With Me.txtChkNoBegin
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtChkNoBegin_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Function IsStrValib() As Boolean
    '检查字串的合法性
    If zlCommFun.StrIsValid(Me.txtChkNoBegin) = False Then
        MsgBox "开始检查号中包括了非法字符串！", vbInformation, App.EXEName
        With Me.txtChkNoBegin
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    If zlCommFun.StrIsValid(Me.txtChkNoEnd) = False Then
        MsgBox "结束检查号中包括了非法字符串！", vbInformation, App.EXEName
        With Me.txtChkNoEnd
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    If zlCommFun.StrIsValid(Me.txt姓名, 12) = False Then
        MsgBox "姓名中包括了非法字符串！", vbInformation, App.EXEName
        With Me.txt姓名
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Function

Private Sub txtChkNoEnd_GotFocus()
    With Me.txtChkNoEnd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtPath_Change()

End Sub

Private Sub txt姓名_GotFocus()
    With Me.txt姓名
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Function ShowImage(lngImageUID As String, lngSerialUID As String) As Boolean
    '功能:从FTP下载文件
    '传入:序列UID
    '返回下载成功后的文件路径
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strCachePath As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strTmpFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim DicomImg As New DicomImages
    
    On Error GoTo WriteFileErr
    ShowImage = True
    strSql = "Select A.图像号,D.用户名 As User1,D.密码 As Pwd1,a.图像UID, " & _
        "D.IP地址 As Host1," & _
        "'/'||D.Ftp目录||'/' As Root1,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL1,d.设备号 as 设备号1, " & _
        "E.用户名 As User2,E.密码 As Pwd2," & _
        "E.IP地址 As Host2," & _
        "'/'||E.Ftp目录||'/' As Root2,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL2 , e.设备号 as 设备号2 " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
        "And A.图像UID= [1]  and a.序列UID = [2]  Order By A.图像号"
    If mblnMoved Then
        strSql = Replace(strSql, "影像检查图象", "H影像检查图象")
        strSql = Replace(strSql, "影像检查序列", "H影像检查序列")
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
            
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, lngImageUID, lngSerialUID)
    strCachePath = App.Path & "\TmpImage\"
    ClearCacheFolder strCachePath
    If rsTmp.EOF <> True Then
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL1")))
    End If
    Do While Not rsTmp.EOF
        If strDeviceNO1 <> rsTmp("设备号1") Then
            strDeviceNO1 = rsTmp("设备号1")
            Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
        End If
        
        If strDeviceNO2 <> rsTmp("设备号2") Then
            strDeviceNO2 = rsTmp("设备号2")
            Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
        End If
        
        strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
        If Dir(strTmpFile) = "" Then
'            Inet.strIPAddress = Nvl(rsTmp("Host1")): Inet.strUser = Nvl(rsTmp("User1")): Inet.strPsw = Nvl(rsTmp("Pwd1"))
            If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFile.GetFileName(rsTmp("URL1"))) <> 0 Then
                strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
'                Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFile.GetFileName(rsTmp("URL2")))
            End If
        End If
        On Error Resume Next
        Viewer.Images.ReadFile strTmpFile
        Kill strTmpFile
        DoEvents
        rsTmp.MoveNext
    Loop
    Exit Function
WriteFileErr:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SelectAllChild(xNode As Node, blCheck As Boolean)
    '功能           取消下一级中的所有字级
    '参数           xNode    Node对象
    '                blCheck 是否选中
    Dim nNode As Node
    
    If xNode.Children = 0 Then Exit Sub
    Set nNode = xNode.Child
    For i = 1 To xNode.Children
        nNode.Checked = blCheck
        Set nNode = nNode.Next
    Next
End Sub
