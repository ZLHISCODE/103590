VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmFilesSet 
   Caption         =   "文件共享部件升级管理"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9495
   Icon            =   "frmFilesSet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9495
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd升级日期 
      Caption         =   "设置(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8175
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5850
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTP升级日期 
      Height          =   300
      Left            =   8175
      TabIndex        =   29
      Top             =   5505
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   119472129
      CurrentDate     =   40908
   End
   Begin VB.CheckBox chk升级日期 
      Caption         =   "升级日期"
      Height          =   240
      Left            =   8175
      TabIndex        =   28
      Top             =   5220
      Width           =   1020
   End
   Begin VB.OptionButton OptType 
      Caption         =   "收集更新"
      Height          =   180
      Index           =   1
      Left            =   8310
      TabIndex        =   27
      ToolTipText     =   "收集部件时候进行比较部件与服务器是否相同,不相同的进行收集."
      Top             =   3090
      Width           =   1200
   End
   Begin VB.OptionButton OptType 
      Caption         =   "收集所有"
      Height          =   180
      Index           =   0
      Left            =   8310
      TabIndex        =   26
      ToolTipText     =   "收集部件时收集所有的部件."
      Top             =   2685
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.FileListBox FileList 
      Height          =   630
      Left            =   15
      TabIndex        =   22
      Top             =   5115
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton cmdSaveInfo 
      Caption         =   "保存设置"
      Height          =   350
      Left            =   8250
      TabIndex        =   16
      Top             =   255
      Width           =   1100
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "收集(&R)"
      Height          =   350
      Left            =   8250
      TabIndex        =   17
      Top             =   710
      Width           =   1100
   End
   Begin VB.Frame fra文件管理 
      Caption         =   "升级文件清单"
      Height          =   3630
      Left            =   75
      TabIndex        =   14
      Top             =   2580
      Width           =   8055
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   3255
         Left            =   195
         TabIndex        =   15
         Top             =   255
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   5741
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin ZL9BillEdit.BillEdit mshBillShow 
         Height          =   3255
         Left            =   195
         TabIndex        =   25
         Top             =   225
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   5741
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "上传(&O)"
      Height          =   350
      Left            =   8250
      TabIndex        =   18
      Top             =   1165
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   8250
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   8250
      TabIndex        =   19
      Top             =   1620
      Width           =   1100
   End
   Begin VB.Frame fra服务器 
      Caption         =   "文件服务器管理"
      Height          =   2295
      Left            =   90
      TabIndex        =   21
      Top             =   180
      Width           =   8025
      Begin VB.CommandButton cmdAccessDir 
         Caption         =   "…"
         Height          =   270
         Left            =   7650
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   255
      End
      Begin MSComCtl2.UpDown upd编号 
         Height          =   300
         Left            =   7680
         TabIndex        =   9
         Top             =   660
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt服务器编号"
         BuddyDispid     =   196625
         OrigLeft        =   7695
         OrigTop         =   660
         OrigRight       =   7935
         OrigBottom      =   915
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt服务器编号 
         Height          =   300
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   660
         Width           =   345
      End
      Begin VB.TextBox txtUserName 
         Height          =   300
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   4
         Top             =   675
         Width           =   1935
      End
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   6
         Top             =   675
         Width           =   1755
      End
      Begin VB.TextBox txtAccessDir 
         Height          =   300
         Left            =   1200
         MaxLength       =   500
         TabIndex        =   1
         Top             =   300
         Width           =   6705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&N)"
         Height          =   350
         Left            =   150
         Picture         =   "frmFilesSet.frx":058A
         TabIndex        =   10
         Top             =   1110
         Width           =   960
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "修改(&M)"
         Height          =   350
         Left            =   150
         TabIndex        =   11
         Top             =   1455
         Width           =   960
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   150
         TabIndex        =   12
         Top             =   1800
         Width           =   960
      End
      Begin MSComctlLib.ListView lvwFileServer 
         Height          =   1095
         Left            =   1170
         TabIndex        =   13
         Top             =   1080
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1931
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils"
         SmallIcons      =   "ils"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "服务器"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "站点访问目录"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "站点访问用户"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "站点访问密码"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblFileNo 
         Caption         =   "服务器编号"
         Height          =   225
         Left            =   6435
         TabIndex        =   7
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lblPassWord 
         AutoSize        =   -1  'True
         Caption         =   "访问密码"
         Height          =   180
         Left            =   3270
         TabIndex        =   5
         Top             =   735
         Width           =   720
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         Caption         =   "访问用户名"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   735
         Width           =   900
      End
      Begin VB.Label lblAccessDir 
         AutoSize        =   -1  'True
         Caption         =   "访问目录"
         Height          =   180
         Left            =   420
         TabIndex        =   0
         Top             =   360
         Width           =   720
      End
   End
   Begin MSComctlLib.ProgressBar pgbState 
      Height          =   150
      Left            =   3405
      TabIndex        =   23
      Top             =   6390
      Visible         =   0   'False
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   6255
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFilesSet.frx":6DDC
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12753
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "9:52"
            Key             =   "STANUM"
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
   Begin MSComctlLib.ImageList ils 
      Left            =   8730
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesSet.frx":766E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFilesSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum HeadInfor
    序号 = 0
    部件名
    版本号
    修改日期
    信息
    加入日期
    说明
    类型
    安装路径
    MD5
    收集类型
End Enum

Private mblnReturn As Boolean
Private mblnChangeDirectory As Boolean      '是否改变目录
Private mblnAutoSet As Boolean     '自动进行配置(包含自动收集文件、自动保存本次升级的文件清单、自动将所有的客户端默认为要升级)
Private mblnFirst As Boolean
Private mblnSourceCode As Boolean '是源代码执行
Private Const mstrzlAppSoftPath = "C:\AppSoft"
Private mstrSourceFloder As String '临时收集目录
Public mobjFile As New FileSystemObject
Public mblnOptType As Boolean 'False 收集所有 True 收集部分

Private Sub cmdAccessDir_Click()
    Dim strFolderName As String
    
    strFolderName = OpenFolder(Me, "选择最新部件的所在目录")
    If strFolderName = "" Then SetCtlEnable True: Exit Sub
    If Len(strFolderName) = 3 Then
        MsgBox "不能选择根目录(" & strFolderName & ")!", vbInformation + vbDefaultButton1, gstrSysName
        SetCtlEnable True
        Exit Sub
    End If
    err = 0
    Me.txtAccessDir.Tag = Trim(strFolderName)
    
    If InStr(1, strFolderName, "\\") <> 0 Then
        Me.txtAccessDir.Text = strFolderName
    Else
        Me.txtAccessDir.Text = "\\" & GetMyCompterName & Mid(strFolderName, 3)
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim objItem As ListItem
    If CheckFileServer = False Then Exit Sub
    With lvwFileServer
        Set objItem = .ListItems.Add(, "K" & txt服务器编号.Text, txt服务器编号.Text, 1, 1)
        objItem.Selected = True
        objItem.SubItems(1) = Trim(txtAccessDir.Text)
        objItem.SubItems(2) = Trim(txtUserName)
        objItem.SubItems(3) = Trim(txtPassword)
        objItem.Tag = "1"
    End With
    Call SetFileSeverCtrlEnable
End Sub

Private Sub CmdDelete_Click()
    '功能:删除服务器
    Dim lngIndex As Long
    With lvwFileServer
        If .SelectedItem Is Nothing Then Exit Sub
        lngIndex = .SelectedItem.Index
        .ListItems.Remove lngIndex
        If lngIndex >= lvwFileServer.ListItems.Count And lvwFileServer.ListItems.Count <> 0 Then
            .ListItems(.ListItems.Count).Selected = True
            .SelectedItem.EnsureVisible
        ElseIf lvwFileServer.ListItems.Count <> 0 Then
            .ListItems(lngIndex).Selected = True
            .SelectedItem.EnsureVisible
        End If
    End With
    Call SetFileSeverCtrlEnable
End Sub
Private Sub SetFileSeverCtrlEnable()
    '-------------------------------------------------------------------------------------------------
    '功能:设置文件服务器相关控件的属性值
    '-------------------------------------------------------------------------------------------------
    Dim blnSel  As Boolean
    blnSel = Not Me.lvwFileServer.SelectedItem Is Nothing
    cmdModify.Enabled = blnSel
    cmdDelete.Enabled = blnSel
End Sub
Private Sub cmdModify_Click()
    Dim objItem As ListItem
    If CheckFileServer(True) = False Then Exit Sub
    With lvwFileServer
        If .SelectedItem Is Nothing Then Exit Sub
        Set objItem = .SelectedItem
        objItem.Key = "K" & txt服务器编号.Text
        objItem.Text = txt服务器编号.Text
        objItem.SubItems(1) = Trim(txtAccessDir.Text)
        objItem.SubItems(2) = Trim(txtUserName)
        objItem.SubItems(3) = Trim(txtPassword)
        objItem.Tag = "1"
    End With
    Call SetFileSeverCtrlEnable
End Sub
Private Function CopyFileToServer(ByVal strFileServer As String, ByVal strSourcePath As String, ByVal strSharePath As String, _
    Optional ByVal strUserName As String, Optional ByVal strPassword As String, Optional ByRef strErrInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------
    '功能:拷贝文件给指定的服务器
    '参数:strFileServer-文件服务器编号
    '     strSourcePath-源文件目录
    '     strSharePath-服务器的共享目录
    '     strUserName-访问的用户名
    '     strPassWord-密码
    '出参:strErrInfor-返回的错误信息
    '返回:拷贝成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/06
    '---------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    
    '第一步:先检查相关的目录是否存在
    '     1.检查源目录是否存在
    If objFile.FolderExists(strSourcePath) = False Then
        strErrInfor = "源文件目录:" & strSourcePath & "不存在,请检查!"
        Exit Function
    End If
    '     2.检查共享目录是否存在
    If objFile.FolderExists(strSharePath) = False Then
        '试着去连接,看是否可行
        If IsNetServer(strSharePath, strUserName, strPassword) = False Then
            strErrInfor = "文件服务器目录:" & strSharePath & "不存在,请检查!"
            Exit Function
        End If
        If objFile.FolderExists(strSharePath) = False Then
            strErrInfor = "文件服务器目录:" & strSharePath & "不存在,请检查!"
            Exit Function
        End If
    Else
        '检查是否具有写权限
        If funCanWrite(strSharePath) = False Then
            strErrInfor = "文件服务器目录:[" & strSharePath & "]不具有写权限!" & vbCrLf & "请设置写入权限或者手动拷贝文件!"
            Exit Function
        End If
    End If
    '   3.检查是否源文件与文件服务器是否一至
    If UCase(strSharePath) = UCase(strSourcePath) Then
        '用不着再进行文件拷贝处理
        CopyFileToServer = True
        Exit Function
    End If
    
    '第二步:先清除客户端访问服务器的共享目录中的所有内容
    If mblnOptType = False Then
        err = 0: On Error Resume Next
        objFile.DeleteFolder strSharePath & "\*", True
        objFile.DeleteFile strSharePath & "\*.*", True
    End If
    err = 0: On Error GoTo ErrHand:
    '第三步:拷贝相关的文件到指定的文件目录中
    '      1.拷贝公共目录:Common
    If CopyFileToTargetServer(strSourcePath, strSharePath, "正在拷贝部件到服务器目录[" & strFileServer & "]", True, strErrInfor) = False Then
        Exit Function
    End If
    
    CopyFileToServer = True
    Exit Function
ErrHand:
    strErrInfor = "出现无可预知的错误,错误信息如下:" & vbCrLf & "错误号:" & err.Number & vbCrLf & "错误描述:" & err.Description
End Function

Public Function CopyFileToTargetServer(ByVal strSourcePath As String, ByVal strTargetPath As String, _
    ByVal strProcessCaption As String, Optional blnChildFolder As Boolean = False, Optional ByRef strErrInfor As String) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:将数据拷贝到指定的服务器上去
    '参数:strSourcePath-源文件路径
    '     strTargetPath-目标文件路径
    '     strProcessCaption-进度条显示名称
    '     blnChildFolder-是否包含源文件的子目录也一并拷贝到目标路径中,true- 拷贝,false-不拷贝
    '出参:strErrInfor-错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/06
    '--------------------------------------------------------------------------------------------------------
    Dim strSourceFile As String, strTarGetFile As String, strTemp As String
    Dim objFile As New FileSystemObject
    Dim i As Long
    
    stbThis.Panels(2).Text = strProcessCaption
    pgbState.Left = stbThis.Panels(2).Left + TextWidth(strProcessCaption) + 100
    pgbState.Width = stbThis.Panels(3).Left - pgbState.Left - 100
    pgbState.Top = stbThis.Top + stbThis.Height / 3
    pgbState.Visible = True
    '判断common目录是否存在
    If FindFile(strTargetPath) = False Then
        '创建common目录
        err = 0
        On Error Resume Next
        MkDir strTargetPath
        If err <> 0 Then
            strErrInfor = "创建目录:" & strTargetPath & "失败," & vbCrLf & "可能你无相关的权限,请找系统员管理!"
            err.Clear
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    With FileList
        .Refresh
        .Path = strSourcePath
        .FileName = "*.*"
        pgbState.Max = .ListCount + IIf(blnChildFolder, 1, 0)
        pgbState.Min = 0
        For i = 0 To .ListCount - 1
            strTemp = "\" & .List(i)
            strSourceFile = strSourcePath & strTemp
            strTarGetFile = strTargetPath & strTemp
            err = 0: On Error Resume Next
            '文件拷贝,需先判断
            If ISCopyFile(strSourceFile, strTarGetFile) = True Then
                objFile.CopyFile strSourceFile, strTarGetFile, True
            End If
            If err <> 0 Then
                If MsgBox("源文件：" & strSourceFile & vbCrLf & " 不能拷贝到目标文件：" & vbCrLf & strTarGetFile & vbCrLf & "中,是否继续？" & vbNewLine & err.Description, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    pgbState.Visible = False
                    stbThis.Panels(2).Text = ""
                    Exit Function
                End If
            End If
            pgbState.value = i + 1
            DoEvents
        Next
'        If blnChildFolder Then
'            '拷贝源路径下的子目录数据给相关的目标目录中:源路径中可能不存在子目录,因此屏蔽相关的错误
'            Err = 0: On Error Resume Next
'            objFile.CopyFolder strSourcePath & "\*", strTargetPath
'            Err = 0: On Error GoTo 0
'            pgbState.Value = pgbState.Value + 1
'            DoEvents
'        End If
    End With
    pgbState.Visible = False
    CopyFileToTargetServer = True
End Function

Private Sub cmdSaveInfo_Click()
'保存相关的配置到数据库中
    If IsValid = False Then Exit Sub
    If Not SaveFile Then Exit Sub
    mblnReturn = True
    stbThis.Panels(2).Text = "服务器配置信息保存成功!"
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Height < 7035 Then
        Me.Height = 7035
    End If
    If Me.Width < 9615 Then
        Me.Width = 9615
    End If
    
    With cmdSaveInfo
        .Left = ScaleWidth - .Width - 100
    End With
    
    With cmdRefresh
        .Left = ScaleWidth - .Width - 100
    End With
    
    With cmdSave
        .Left = ScaleWidth - .Width - 100
        cmdCancel.Left = .Left
        cmdHelp.Left = .Left
'        cmdHelp.Top = Me.ScaleHeight - cmdHelp.Height - IIf(stbThis.Visible, stbThis.Height, 0) - 50


        chk升级日期.Left = .Left
        DTP升级日期.Left = .Left
        cmd升级日期.Left = .Left
        
        
        cmd升级日期.Top = Me.ScaleHeight - cmdHelp.Height - IIf(stbThis.Visible, stbThis.Height, 0) - 50
        DTP升级日期.Top = cmd升级日期.Top - DTP升级日期.Height - 50
        chk升级日期.Top = DTP升级日期.Top - chk升级日期.Height - 50
        
    End With
    
 
    
    
    With fra服务器
        .Width = cmdSave.Left - .Left - 50
        txtAccessDir.Width = .Width - txtAccessDir.Left - 50
        cmdAccessDir.Left = txtAccessDir.Left + txtAccessDir.Width - cmdAccessDir.Width
        
        upd编号.Left = txtAccessDir.Left + txtAccessDir.Width - upd编号.Width
        txt服务器编号.Left = upd编号.Left - txt服务器编号.Width
        lblFileNo.Left = txt服务器编号.Left - lblFileNo.Width
        lvwFileServer.Width = txtAccessDir.Width
    End With
    
    With fra文件管理
        .Width = fra服务器.Width
        .Left = fra服务器.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        mshBill.Height = .Height - mshBill.Top - 60
        mshBill.Width = .Width - mshBill.Left - 60
        mshBillShow.Left = mshBill.Left
        mshBillShow.Top = mshBill.Top
        mshBillShow.Height = mshBill.Height
        mshBillShow.Width = mshBill.Width
    End With
    
    With OptType(0)
        .Left = cmdCancel.Left
        .Top = fra文件管理.Top + 100
    End With
    
    With OptType(1)
        .Left = cmdCancel.Left
        .Top = OptType(0).Top + OptType(0).Height + 150
    End With
End Sub

Private Sub lvwFileServer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    upd编号.value = Val(Item.Text)
    txtAccessDir.Text = Item.SubItems(1)
    txtUserName.Text = Item.SubItems(2)
    txtPassword.Text = Item.SubItems(3)
    Call SetFileSeverCtrlEnable
End Sub

Private Sub cmdCancel_Click()
'    If mobjFile.FolderExists(mstrSourceFloder) Then
'        mobjFile.DeleteFolder mstrSourceFloder, True
'    End If
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub cmdRefresh_Click()
    '收集文件到目录
    Dim strSourceFile As String
    SetCtlEnable False
    cmdSave.Enabled = False
    mshBill.Visible = True
    mshBillShow.Visible = False
    OptType(0).Enabled = False
    OptType(1).Enabled = False
    If GetFileInforamtion() = True Then
        '清空7z.exe残余系统进程
        Call fun_KillProcess(PROAPPCTION)
        '拷贝收集文件到剪贴板
        Call FloderToClipBoard(mstrSourceFloder)
        '收集排序
        Call BillFileSort
        
        Me.cmdSave.Enabled = True
        
        strSourceFile = mstrSourceFloder & "\"
        If mobjFile.FolderExists(strSourceFile) Then
            With FileList
                .Refresh
                .Path = strSourceFile
                .FileName = "*.*"
                
                If .ListCount = 0 Then
                    MsgBox "没有收集到任何文件," & vbCrLf & "本地文件MD5与服务器上一致!", vbInformation + vbDefaultButton1 + vbOKOnly
                    cmdSave.Enabled = False
                Else
                    MsgBox "文件收集成功,临时收集文件已拷贝到剪贴板中," & vbCrLf & "当服务器共享目录没有写权限时,请直接粘贴收集文件." & vbCrLf & "注意:上传完成和关闭都将删除临时收集目录和剪贴板!", vbInformation + vbDefaultButton1 + vbOKOnly
                    cmdSave.Enabled = True
                End If
            End With
        End If
            
        mblnChangeDirectory = True
    End If
    OptType(0).Enabled = True
    OptType(1).Enabled = True
    SetCtlEnable True
End Sub

Private Sub cmdSave_Click()
    Dim strErrMsg As String
    Dim objItem As ListItem
    Dim strSourcePath As String
    '1.检查数据的合法性
    If IsValid = False Then Exit Sub

    '2.需要将相关的文件分布到相关的服务器上
    strSourcePath = Trim(mstrSourceFloder)
    For Each objItem In lvwFileServer.ListItems
        If CopyFileToServer(objItem.Text, strSourcePath, objItem.SubItems(1), objItem.SubItems(2), objItem.SubItems(3), strErrMsg) = False Then
            MsgBox strErrMsg, vbDefaultButton1 + vbInformation, gstrSysName
            stbThis.Panels(2).Text = strErrMsg
            Exit Sub
        End If
    Next
    
    '3.保存相关的配置到数据库中
    If Not SaveFile Then Exit Sub
    mobjFile.DeleteFolder strSourcePath, True
    mblnReturn = True
    Unload Me
End Sub

Private Function IsValid() As Boolean
    '--------------------------------------------------------------------
    '功能:验证数据的合法性
    '--------------------------------------------------------------------
    Dim objItem As ListItem
    
    IsValid = False
    If mblnChangeDirectory = True Then
        If FindFile(Trim(mstrSourceFloder)) = False Then
            MsgBox "升级文件的指定目录不存在,请重新设置!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If
    
    If InStr(1, mstrSourceFloder, "'") <> 0 Then
        MsgBox "升级文件中不能存在单引号!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If lvwFileServer.ListItems.Count = 0 Then
        MsgBox "没有设置相关的文件服务器,必须设置站点升级的服务器!", vbInformation + vbDefaultButton1, gstrSysName
        If txtAccessDir.Enabled Then txtAccessDir.SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function CheckFileServer(Optional blnModify As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------
    '功能:检查当前的文件服务器是否正确
    '返回:当前的文件服务器的各项正确,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/06
    '-----------------------------------------------------------------------------
    Dim objItem As ListItem
    
    err = 0: On Error GoTo ErrHand:
    CheckFileServer = False
    If Trim(txtAccessDir.Text) = "" Then
        MsgBox "未设置站点访问的服务器目录,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtAccessDir.Enabled Then txtAccessDir.SetFocus
        Exit Function
    End If
    If Trim(txtUserName.Text) = "" Then
        MsgBox "访问用户未设置,请设置访问用用户名!", vbInformation + vbDefaultButton1, gstrSysName
        If txtUserName.Enabled Then txtUserName.SetFocus
        Exit Function
    End If
    If InStr(1, txtUserName.Text, "'") <> 0 Then
        MsgBox "访问用户中不能存在单引号!", vbInformation + vbDefaultButton1, gstrSysName
        If txtUserName.Enabled Then txtUserName.SetFocus
        Exit Function
    End If
    If InStr(1, txtPassword.Text, "'") <> 0 Then
        MsgBox "访问密码中不能存在单引号!", vbInformation + vbDefaultButton1, gstrSysName
        If txtPassword.Enabled Then txtPassword.SetFocus
        Exit Function
    End If
    
    If Right(Trim(txtAccessDir.Text), 1) = "\" Then
        txtAccessDir.Text = Left(txtAccessDir.Text, Len(txtAccessDir.Text) - 1)
    End If
    
    For Each objItem In lvwFileServer.ListItems
        If blnModify = True Then
            If Val(objItem.Text) = Val(txt服务器编号.Text) And objItem.Selected = False Then
                MsgBox "服务器编号为" & txt服务器编号.Text & "已经存在,不能再设置此编号的服务器!", vbInformation + vbDefaultButton1, gstrSysName
                If txt服务器编号.Enabled Then txt服务器编号.SetFocus
                Exit Function
            End If
            If objItem.SubItems(1) = txtAccessDir.Text And objItem.Selected = False Then
                MsgBox "存在相同的访问目录,用不着再设置!", vbInformation + vbDefaultButton1, gstrSysName
                If txtAccessDir.Enabled Then txtAccessDir.SetFocus
                Exit Function
            End If
        Else
            If Val(objItem.Text) = Val(txt服务器编号.Text) Then
                MsgBox "服务器编号为" & txt服务器编号.Text & "已经存在,不能再设置此编号的服务器!", vbInformation + vbDefaultButton1, gstrSysName
                If txt服务器编号.Enabled Then txt服务器编号.SetFocus
                Exit Function
            End If
            If objItem.SubItems(1) = txtAccessDir.Text Then
                MsgBox "存在相同的访问目录,用不着再设置!", vbInformation + vbDefaultButton1, gstrSysName
                If txtAccessDir.Enabled Then txtAccessDir.SetFocus
                Exit Function
            End If
        End If
    Next
    
    If FindFile(Trim(txtAccessDir.Text)) = False Then
        If IsNetServer(Trim(txtAccessDir.Text), Trim(txtUserName.Text), Trim(txtPassword.Text)) = False Then
            MsgBox "升级文件的指定目录不存在,请重新设置!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        Call CancelNetServer(Trim(txtAccessDir.Text))
    End If
    CheckFileServer = True
    Exit Function
ErrHand:
End Function

Private Function SaveFile() As Boolean
    '-----------------------------------------------------------------------------
    '功能:将相关的配置保存到数据库中
    '返回:保存成功,返回true,否则返回False
    '编制:祝庆
    '日期:2010/12/13
    '-----------------------------------------------------------------------------
    Dim strSQL As String, objItem As ListItem

    
    err = 0
    On Error GoTo ErrHand:
    SaveFile = False
    gcnOracle.BeginTrans

    '先清空相关的数据
    strSQL = "Delete zlregInfo where (项目 like '服务器目录%' or 项目 like '访问用户%' or 项目 like '访问密码%') and 项目 not in ('访问用户S','访问密码S','访问用户F','访问密码F') "
    gcnOracle.Execute strSQL
'    strSQL = "Update zlreginfo set 内容=NULL where 项目 in ('服务器目录','访问用户','访问密码')"
'    gcnOracle.Execute strSQL
    '保存新的服务数据
    For Each objItem In lvwFileServer.ListItems
'        If Val(objItem.Text) = 0 Then
'            strSQL = "Update zlreginfo set 内容='" & Trim(objItem.SubItems(1)) & "' where 项目='服务器目录'"
'            gcnOracle.Execute strSQL
'            strSQL = "Update zlreginfo set 内容='" & Trim(objItem.SubItems(2)) & "' where 项目='访问用户'"
'            gcnOracle.Execute strSQL
'            strSQL = "Update zlreginfo set 内容='" & Trim(objItem.SubItems(3)) & "' where 项目='访问密码'"
'            gcnOracle.Execute strSQL
'        Else
            strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('服务器目录" & objItem.Text & "',Null,'" & Trim(objItem.SubItems(1)) & "')"
            gcnOracle.Execute strSQL
            strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('访问用户" & objItem.Text & "',Null,'" & Trim(objItem.SubItems(2)) & "')"
            gcnOracle.Execute strSQL
            strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('访问密码" & objItem.Text & "',Null,'" & Trim(objItem.SubItems(3)) & "')"
            gcnOracle.Execute strSQL
'        End If
    Next
    gcnOracle.CommitTrans
    SaveFile = True
    Exit Function
ErrHand:
    MsgBox "保存最新服务器信息时出现错误,可能存在两个相同的服务器!" & vbNewLine & err.Description, vbInformation + vbDefaultButton1, gstrSysName
    gcnOracle.RollbackTrans
End Function

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    mstrSourceFloder = GetTmpPath & "TEMPGATHER"
    If Load服务器信息() = False Then Unload Me: Exit Sub
    
    mblnSourceCode = IsSourceCode
    Me.cmdSave.Enabled = False
    Me.mshBill.Visible = True
    Me.mshBillShow.Visible = False
    
    '加载头信息
    Call LoadHeadInfor
    '初始部件信息.
    Call intBillInfor
    '比较收集信息
    Call CompareFile
    '判断升级日期
    Call OpinionUpGradeDate
    mshBill.AutoRefresh = False

    If mblnOptType Then
        OptType(1).value = True
    Else
        OptType(0).value = True
    End If
    
    mblnChangeDirectory = False
    mblnReturn = False
    
    '修改为打开窗体时删除临时目录
    If mobjFile.FolderExists(mstrSourceFloder) Then
        mobjFile.DeleteFolder mstrSourceFloder, True
    End If
    
    If mblnAutoSet Then
        '自动进行配置(包含自动收集文件、自动保存本次升级的文件清单、自动将所有的客户端默认为要升级)
        If AutoSet = False Then Exit Sub
        '将所有的站点改为升级
        Call ExecuteProcedure("Zl_Zlclients_Control(4,Null,Null,1)", Me.Caption)
        Call cmdSave_Click
    End If
    Call SetFileSeverCtrlEnable
End Sub

Private Function AutoSet() As Boolean
    '------------------------------------------------------------------------------------------------------------
    '功能:自动配置
    '返回:配置成功,返回true,否则返回False
    '------------------------------------------------------------------------------------------------------------
    SetCtlEnable False

    If GetFileInforamtion = False Then
        SetCtlEnable True: Exit Function
    End If
    
    SetCtlEnable True
    AutoSet = True
End Function

Private Sub Form_Load()
    Call ApplyOEM(stbThis)
    mblnFirst = True
End Sub

Private Sub LoadHeadInfor()
    '------------------------------------------------------------------------------------------------------------
    '功能:加载头信息
    '------------------------------------------------------------------------------------------------------------
    With mshBill
        .Active = True
        .Cols = 11
        .Clear
        .Rows = 2
        '.MsfObj.FixedCols = 1
        .TextMatrix(0, HeadInfor.序号) = "序号"
        .TextMatrix(0, HeadInfor.部件名) = "部件名"
        .TextMatrix(0, HeadInfor.版本号) = "版本号"
        .TextMatrix(0, HeadInfor.修改日期) = "修改日期"
        .TextMatrix(0, HeadInfor.信息) = "信息"
        .TextMatrix(0, HeadInfor.加入日期) = "加入日期"
        .TextMatrix(0, HeadInfor.说明) = "说明"
        .TextMatrix(0, HeadInfor.类型) = "类型"
        .TextMatrix(0, HeadInfor.安装路径) = "安装路径"
        .TextMatrix(0, HeadInfor.收集类型) = "收集类型"
        
        .ColWidth(HeadInfor.序号) = 500
        .ColWidth(HeadInfor.部件名) = 2000
        .ColWidth(HeadInfor.版本号) = 900
        .ColWidth(HeadInfor.修改日期) = 1800
        .ColWidth(HeadInfor.信息) = 2000
        .ColWidth(HeadInfor.加入日期) = 1800
        .ColWidth(HeadInfor.说明) = 2000
        .ColWidth(HeadInfor.类型) = 800
        .ColWidth(HeadInfor.安装路径) = 2000
        .ColWidth(HeadInfor.MD5) = 0
        .ColWidth(HeadInfor.收集类型) = 0
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择
        
        .ColData(HeadInfor.序号) = 5
        .ColData(HeadInfor.部件名) = 5
        .ColData(HeadInfor.版本号) = 5
        .ColData(HeadInfor.修改日期) = 5
        .ColData(HeadInfor.信息) = 5
        .ColData(HeadInfor.加入日期) = 5
        .ColData(HeadInfor.说明) = 5
        .ColData(HeadInfor.类型) = 5
        .ColData(HeadInfor.安装路径) = 5
        .ColData(HeadInfor.MD5) = 5
        .ColData(HeadInfor.收集类型) = 5
        
        .ColAlignment(HeadInfor.序号) = flexAlignCenterCenter
        .ColAlignment(HeadInfor.部件名) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.版本号) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.修改日期) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.信息) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.加入日期) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.说明) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.类型) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.安装路径) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.MD5) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.收集类型) = flexAlignLeftCenter
        
        .Active = False
    End With
End Sub

Private Sub LoadHeadInforShow()
    '------------------------------------------------------------------------------------------------------------
    '功能:加载头信息
    '------------------------------------------------------------------------------------------------------------
    With mshBillShow
        .Active = True
        .Cols = 11
        .Clear
        .Rows = 2
        '.MsfObj.FixedCols = 1
        .TextMatrix(0, HeadInfor.序号) = "序号"
        .TextMatrix(0, HeadInfor.部件名) = "部件名"
        .TextMatrix(0, HeadInfor.版本号) = "版本号"
        .TextMatrix(0, HeadInfor.修改日期) = "修改日期"
        .TextMatrix(0, HeadInfor.信息) = "信息"
        .TextMatrix(0, HeadInfor.加入日期) = "加入日期"
        .TextMatrix(0, HeadInfor.说明) = "说明"
        .TextMatrix(0, HeadInfor.类型) = "类型"
        .TextMatrix(0, HeadInfor.安装路径) = "安装路径"
        .TextMatrix(0, HeadInfor.收集类型) = "收集类型"
        
        .ColWidth(HeadInfor.序号) = 500
        .ColWidth(HeadInfor.部件名) = 2000
        .ColWidth(HeadInfor.版本号) = 900
        .ColWidth(HeadInfor.修改日期) = 1800
        .ColWidth(HeadInfor.信息) = 2000
        .ColWidth(HeadInfor.加入日期) = 1800
        .ColWidth(HeadInfor.说明) = 2000
        .ColWidth(HeadInfor.类型) = 800
        .ColWidth(HeadInfor.安装路径) = 2000
        .ColWidth(HeadInfor.MD5) = 0
        .ColWidth(HeadInfor.收集类型) = 0
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择
        
        .ColData(HeadInfor.序号) = 5
        .ColData(HeadInfor.部件名) = 5
        .ColData(HeadInfor.版本号) = 5
        .ColData(HeadInfor.修改日期) = 5
        .ColData(HeadInfor.信息) = 5
        .ColData(HeadInfor.加入日期) = 5
        .ColData(HeadInfor.说明) = 5
        .ColData(HeadInfor.类型) = 5
        .ColData(HeadInfor.安装路径) = 5
        .ColData(HeadInfor.MD5) = 5
        .ColData(HeadInfor.收集类型) = 5
        
        .ColAlignment(HeadInfor.序号) = flexAlignCenterCenter
        .ColAlignment(HeadInfor.部件名) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.版本号) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.修改日期) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.信息) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.加入日期) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.说明) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.类型) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.安装路径) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.MD5) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.收集类型) = flexAlignLeftCenter
        
        .Active = False
    End With
End Sub

Private Function GetFileInforamtion() As Boolean
        '------------------------------------------------------------------------
        '--功能:获取最新部件信息
        '--返回:加载成功,返回true,否则false
        '------------------------------------------------------------------------
        Dim strCurFileDirectory As String
        Dim lngRow As Long
        Dim lngErr As Long
        
'        Dim strSQL As String
'        Dim rsTmp As New ADODB.Recordset
        Dim strPath As String '程序安装路径
        Dim strFullPath As String '安装路径
        
        Dim strCompTxt  As String '压缩脚本
        Dim strSource   As String '压缩源文件
        Dim strDesc     As String '压缩目标文件
'        Dim RetVal      As Long  '返回值
        Dim objFile As New FileSystemObject
        Dim usrUpList  As UpdateList
        Dim lngSuccess  As Long
        Dim str7zFile   As String
        
        '数据库文件的值
        Dim strFilename As String '文件名
        Dim strFileType As String '文件类型
        Dim strSetupPath As String '安装路径
        Dim strFileMD5   As String '文件MD5值
        
        Dim strLocaFileMD5 As String '本地文件MD5值
        Dim driver As Drive

        err = 0: On Error GoTo ErrHand:
        strCurFileDirectory = Trim(mstrSourceFloder)
        GetFileInforamtion = False
        
        '检查剩余空间
        For Each driver In objFile.Drives
            If driver.IsReady Then
                If driver.DriveLetter = "C" Then
                    If driver.FreeSpace < 204800000 Then '小于200M
                        MsgBox "临时收集目录没有足够的空间!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    Exit For
                End If
            End If
        Next driver

        If FindFile(strCurFileDirectory) = False Then
            On Error Resume Next
            Call mobjFile.CreateFolder(strCurFileDirectory)
            If mobjFile.FolderExists(strCurFileDirectory) = False Then
                MsgBox "临时收集目录不能创建,请检查!" & vbCrLf & strCurFileDirectory, vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
        End If
        
        '2个压缩文件
        str7zFile = GetWinSystemPath & "\7z.exe"
        If FindFile(str7zFile) = False Then
            MsgBox "压缩文件7z.exe不存在,请手动放入系统目录下!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        
        str7zFile = GetWinSystemPath & "\7z.dll"
        If FindFile(str7zFile) = False Then
            MsgBox "压缩文件7z.dll不存在,请手动放入系统目录下!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
       
       
        '先清除临时收集文件目录中的所有内容
        err = 0: On Error Resume Next
        objFile.DeleteFolder strCurFileDirectory & "\*", True
        objFile.DeleteFile strCurFileDirectory & "\*.*", True
        
        If mblnSourceCode Then
            strPath = mstrzlAppSoftPath
        Else
            strPath = App.Path
        End If
        
        
        pgbState.Visible = True
        stbThis.Panels(2).Text = "正在收集和压缩部件"
        pgbState.Left = stbThis.Panels(2).Left + TextWidth("正在收集和压缩部件") + 100
        pgbState.Width = stbThis.Panels(3).Left - pgbState.Left - 100
        pgbState.Top = stbThis.Top + stbThis.Height / 3

        With mshBill
        If .Rows = 0 Then Exit Function
        '        DoEvents
        pgbState.Max = .Rows - 1
        pgbState.Min = 0
        pgbState.value = 0
        
        Erase usrUpList.uFile
        lngSuccess = 0
        lngErr = 0
        
        For lngRow = 1 To .Rows - 1
                strFilename = .TextMatrix(lngRow, HeadInfor.部件名)
                strFileType = .TextMatrix(lngRow, HeadInfor.类型)
                strSetupPath = .TextMatrix(lngRow, HeadInfor.安装路径)
                strFileMD5 = .TextMatrix(lngRow, HeadInfor.MD5)
'                If strFileName = UCase("zl9MediStore.dll") Then
'                    MsgBox 1
'                End If
                '获取完整的路径
                strFullPath = GetSetupPath(Nvl(strFilename, ""), Nvl(strSetupPath, ""), Nvl(strFileType, ""), strPath)
                If strFullPath = "" Then
                    If Nvl(strFilename, "") <> "" Then
                        .TextMatrix(lngRow, HeadInfor.信息) = "未指定路径!"
                        .TextMatrix(lngRow, HeadInfor.收集类型) = "1"
                        .SetRowColor lngRow, vbRed, False
                        lngErr = lngErr + 1
                    End If
                Else
                    '7z进行压缩
                    '4个文件不需要压缩 特殊处理
                    If UCase(Nvl(strFilename, "")) = UCase("zlHisCrust.exe") Or UCase(Nvl(strFilename, "")) = UCase("7z.exe") Or UCase(Nvl(strFilename, "")) = UCase("7z.dll") Or UCase(Nvl(strFilename, "")) = UCase("aamd532.dll") Or UCase(Nvl(strFilename, "")) = UCase("zlRunas.exe") Or UCase(Nvl(strFilename, "")) = UCase("RegCom.dll") Then
                        strDesc = strCurFileDirectory & "\" & strFilename
                        If FindFile(strFullPath) Then
                            strLocaFileMD5 = HashFile(strFullPath, 2 ^ 27)  '记录MD5以便客户端升级时进行文件比较
                            If OptType(1).value Then
                                If UCase(strFileMD5) = UCase(strLocaFileMD5) Then
                                  GoTo Success
                                End If
                            End If
                            ReDim Preserve usrUpList.uFile(lngSuccess)
                            usrUpList.uFile(lngSuccess).FileName = strFilename
                            usrUpList.uFile(lngSuccess).FileVision = GetCommpentVersion(strFullPath)
                            usrUpList.uFile(lngSuccess).FileEditDate = Format(FileDateTime(strFullPath), "yyyy-MM-DD hh:mm:ss")
                            usrUpList.uFile(lngSuccess).FileMD5 = strLocaFileMD5
                            
                            If ISCopyFile(strFullPath, strDesc) = True Then
                                objFile.CopyFile strFullPath, strDesc, True
                            End If
Success:
                            .TextMatrix(lngRow, HeadInfor.信息) = ""
                            .TextMatrix(lngRow, HeadInfor.收集类型) = "0"
                            .SetRowColor lngRow, &HFFFFFF, False
                            lngSuccess = lngSuccess + 1
                        Else
                             .TextMatrix(lngRow, HeadInfor.信息) = "未安装文件!"
                             .TextMatrix(lngRow, HeadInfor.收集类型) = "2"
                             .SetRowColor lngRow, vbRed, False
                             lngErr = lngErr + 1
                        End If
                    Else
                        If FindFile(strFullPath) Then
                            strLocaFileMD5 = HashFile(strFullPath, 2 ^ 27)  '记录MD5以便客户端升级时进行文件比较
                            If OptType(1).value Then
                                If UCase(strFileMD5) = UCase(strLocaFileMD5) Then
                                  GoTo Success1
                                End If
                            End If
                            
                            ReDim Preserve usrUpList.uFile(lngSuccess)
                            usrUpList.uFile(lngSuccess).FileName = strFilename
                            usrUpList.uFile(lngSuccess).FileVision = GetCommpentVersion(strFullPath)
                            usrUpList.uFile(lngSuccess).FileEditDate = Format(FileDateTime(strFullPath), "yyyy-MM-DD hh:mm:ss")
                            usrUpList.uFile(lngSuccess).FileMD5 = strLocaFileMD5
                            
                           
                            strSource = strFullPath
                            strDesc = strCurFileDirectory & "\" & GetCompressName(Nvl(strFilename, ""))
                            strCompTxt = CompressionCmd(strDesc, strSource, COMPRESSIONRATE)
                            If strCompTxt <> "" Then
'                                RetVal = Shell(strCompTxt, vbHide)
                                 Call GetCmdTxt(strCompTxt)
                            End If
Success1:
                            .TextMatrix(lngRow, HeadInfor.信息) = ""
                            .TextMatrix(lngRow, HeadInfor.收集类型) = "0"
                            .SetRowColor lngRow, &HFFFFFF, False
                            lngSuccess = lngSuccess + 1
                        Else
                             .TextMatrix(lngRow, HeadInfor.信息) = "未安装文件!"
                             .TextMatrix(lngRow, HeadInfor.收集类型) = "2"
                             .SetRowColor lngRow, vbRed, False
                             lngErr = lngErr + 1
                        End If
                    End If
                End If

'                .ListCount = lngRow
                DoEvents
                If pgbState.value >= pgbState.Max Then
                    pgbState.value = pgbState.Max
                Else
                    pgbState.value = pgbState.value + 1
                End If
            Next
        End With
        
        '保存升级脚本
        Call SaveUpList(usrUpList)
        
        pgbState.Visible = False
        If lngErr = 0 Then
            stbThis.Panels(2).Text = ""
        Else
            stbThis.Panels(2).Text = lngErr & "个文件未安装"
        End If
        GetFileInforamtion = True
        
        Exit Function
ErrHand:
        MsgBox "在收集文件时,出现了错误,可能是目标文件" & vbCrLf & "已经不存在,错误信息为:" & err.Description, vbInformation + vbDefaultButton1, gstrSysName
        pgbState.Visible = False
        stbThis.Panels(2).Text = ""
        GetFileInforamtion = False
        
End Function

Private Function FindFile(ByVal strFilename As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--功能:查找指定的文件或文夹是否存在
    '--返回: 如果存在此文件为True,否则为Flase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFilename) > 0 Then
        apiOpenFile strFilename, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

Private Function GetWinPath() As String
    '--------------------------------------------------------------------------------------------------------------
    '--功能:获取系统目录
    '--------------------------------------------------------------------------------------------------------------
    Dim Buffer As String
    Dim StrWinPath As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetWindowsDirectory(Buffer, Len(Buffer))
    StrWinPath = Left(Buffer, rtn)
    GetWinPath = StrWinPath
End Function
Private Function GetWinSystemPath() As String
    
    Dim Buffer As String
    Dim strSystem As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetSystemDirectory(Buffer, Len(Buffer))
    strSystem = Left(Buffer, rtn)
    
    GetWinSystemPath = strSystem
End Function
Private Function Load服务器信息() As Boolean
    '---------------------------------------------------------------------------------------------------
    '功能:加载服务器信息
    '参数:
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/06
    '---------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem
    Dim str服务器号 As String
    
    
    err = 0: On Error GoTo ErrHand:
    Set rsTmp = New ADODB.Recordset
    
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 like '服务器目录%'"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Me.lvwFileServer.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            str服务器号 = Replace(Nvl(rsTemp!项目), "服务器目录", "")
                If str服务器号 <> "" Then
                Set objItem = lvwFileServer.ListItems.Add(, "K" & Val(str服务器号), Val(str服务器号), 1, 1)
                objItem.SubItems(1) = Nvl(rsTemp!内容)
                
                gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 ='访问用户" & str服务器号 & "'"
                Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
                If rsTmp.EOF = False Then
                    objItem.SubItems(2) = Nvl(rsTmp!内容)
                Else
                    objItem.SubItems(2) = ""
                End If
                
                gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 ='访问密码" & str服务器号 & "'"
                Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
                If rsTmp.EOF = False Then
                    objItem.SubItems(3) = Nvl(rsTmp!内容)
                Else
                    objItem.SubItems(3) = ""
                End If
                objItem.Tag = ""
            End If
            .MoveNext
        Loop
        .Close
    End With
    If Not lvwFileServer.SelectedItem Is Nothing Then
        lvwFileServer.SelectedItem.Selected = False
        Set lvwFileServer.SelectedItem = Nothing
    End If
    Load服务器信息 = True
    Exit Function
ErrHand:
End Function

Public Sub ShowEdit(ByVal frmMain As Object, ByRef blnRetun As Boolean, Optional blnAutoSet As Boolean)
    '-------------------------------------------------------------------------------
    '--功能：显示和编辑部件信息
    '--参数：frmMain-主窗体
    '       blnAutoSet-自动进行配置(包含自动收集文件、自动保存本次升级的文件清单、自动将所有的客户端默认为要升级)
    '--返回：blnRetun-编辑成功返回true,否则返回false
    '--      strSourceDirectory-返回指定的源文件目录
    '-------------------------------------------------------------------------------
    mblnAutoSet = blnAutoSet
    Me.cmdSave.Enabled = False
    
    Me.Show 1, frmMDIMain
    blnRetun = mblnReturn
End Sub
 

Private Sub lvwFileServer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub mshBill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim i As Integer
    err = 0: On Error Resume Next
    If KeyCode = vbKeyDelete Then
        If mshBill.Rows <> 2 Then
            mshBill.MsfObj.RowPosition(mshBill.MsfObj.Row) = mshBill.MsfObj.Rows - 1
            mshBill.Rows = mshBill.Rows - 1
        Else
            mshBill.ClearBill
        End If
        With mshBill
            .Redraw = False
            For i = 1 To .Rows - 1
                If .TextMatrix(i, HeadInfor.部件名) <> "" Then
                    .TextMatrix(i, HeadInfor.序号) = i
                End If
            Next
            .Redraw = True
        End With
        cmdSave.Enabled = True
        mblnChangeDirectory = True
    End If
    
End Sub

Private Sub mshBill_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub OptType_Click(Index As Integer)
    If OptType(0).value Then
        mblnOptType = False
    Else
        mblnOptType = True
    End If
End Sub

Private Sub txtAccessDir_Change()
    Me.cmdSave.Enabled = True
End Sub

Private Sub txtAccessDir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtFileSource_Change()
    mblnChangeDirectory = True
End Sub

Private Sub txtFileSource_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtPassword_Change()
    Me.cmdSave.Enabled = True
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtUserName_Change()
    Me.cmdSave.Enabled = True
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
Private Sub intBillInfor()
    '--------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化部件信息
    '--------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngRow As Long
    Dim str版本号 As String
    Dim lng版本号 As Long
    Dim str加入日期 As String
    
    err = 0
    On Error Resume Next
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Zlfilesupgrade")
    With rsTmp
        lngRow = 1
        Do While Not .EOF
            mshBill.TextMatrix(lngRow, HeadInfor.序号) = lngRow
            mshBill.TextMatrix(lngRow, HeadInfor.部件名) = IIf(IsNull(!文件名), "", !文件名)
            str版本号 = ""
            If !版本号 > 0 Then
                lng版本号 = !版本号
                str版本号 = Int(lng版本号 / 10 ^ 8)
                lng版本号 = lng版本号 Mod 10 ^ 8
                str版本号 = str版本号 & "." & Int(lng版本号 / 10 ^ 4)
                lng版本号 = lng版本号 Mod 10 ^ 4
                str版本号 = str版本号 & "." & lng版本号
            End If
            
            str加入日期 = IIf(IsNull(!加入日期), "", !加入日期)
            If str加入日期 <> "" Then
                str加入日期 = Format(str加入日期, "yyyy-MM-dd hh:mm:ss")
            End If
            
            mshBill.TextMatrix(lngRow, HeadInfor.版本号) = str版本号
            mshBill.TextMatrix(lngRow, HeadInfor.修改日期) = Format(!修改日期, "yyyy-MM-dd hh:mm:ss")
            mshBill.TextMatrix(lngRow, HeadInfor.加入日期) = str加入日期
            mshBill.TextMatrix(lngRow, HeadInfor.说明) = IIf(IsNull(!说明), "", !说明)
            mshBill.TextMatrix(lngRow, HeadInfor.类型) = IIf(IsNull(!类型), "", !类型)
            mshBill.TextMatrix(lngRow, HeadInfor.安装路径) = IIf(IsNull(!安装路径), "", !安装路径)
            If IIf(IsNull(!MD5), "", !MD5) <> "" Then
                mblnOptType = True
            End If
            mshBill.TextMatrix(lngRow, HeadInfor.MD5) = IIf(IsNull(!MD5), "", !MD5)
            mshBill.TextMatrix(lngRow, HeadInfor.收集类型) = "0"
            
            mshBill.Rows = mshBill.Rows + 1
            lngRow = lngRow + 1
            .MoveNext
        Loop
        If .RecordCount <> 0 Then
            mshBill.Rows = mshBill.Rows - 1
        End If
    End With
End Sub
Private Sub SetCtlEnable(Optional blnEnable As Boolean = False)
    '--------------------------------------------------------------------------------------------
    '功能:设置控件的Enable属性
    '--------------------------------------------------------------------------------------------
    Me.cmdCancel.Enabled = blnEnable
    Me.cmdHelp.Enabled = blnEnable
    Me.txtPassword.Enabled = blnEnable
    Me.txtUserName.Enabled = blnEnable
    Me.mshBill.Enabled = blnEnable
End Sub

Private Function IsSourceCode() As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:确定是否源代码
    '返回:是原代码-true,不是源代码-false
    '-----------------------------------------------------------------------------------------
    err = 0: On Error Resume Next
    Debug.Print 1 / 0
    IsSourceCode = err <> 0
End Function
Public Function ISCopyFile(ByVal strSourceFile As String, ByVal strTarGetFile As String) As Boolean
     '---------------------------------------------------------------------------------------------------------------
    '
    '功能:判断是否需要拷贝文件(比较版本号,修改时间)
    '入参数:
    '   strSourceFile:源文件
    '   strTargetFile:目标文件
    '返回:需要拷贝则返回true,否则返回false
    '---------------------------------------------------------------------------------------------------------------
    Dim strSource As String, strTarget As String
    
    ISCopyFile = False
    err = 0: On Error Resume Next
    If FindFile(strTarGetFile) = False Then
        '没有发现文件，则返回true
        ISCopyFile = True
        Exit Function
    End If
    
    '比较文件版本号
    strTarget = GetCommpentVersion(strTarGetFile)
    strSource = GetCommpentVersion(strSourceFile)
    If RtnVerNum(strTarget) < RtnVerNum(strSource) Then
        ISCopyFile = True
        Exit Function
    End If
    
    '比较文件的最后修改时间
    strTarget = Format(FileDateTime(strTarGetFile), "yyyy-MM-DD hh:mm:ss")
    strSource = Format(FileDateTime(strSourceFile), "yyyy-MM-DD hh:mm:ss")
    If strTarget < strSource Then
        ISCopyFile = True
        Exit Function
    End If
End Function
Private Function RtnVerNum(ByVal strVer As String) As Long
    '--------------------------------------------------------------------------------------------------------------------------------
    '--功能:返回数字版本
    '--------------------------------------------------------------------------------------------------------------------------------
    Dim strArr
    
    If strVer <> "" Then
        strArr = Split(strVer, ".")
        RtnVerNum = strArr(0) * 10 ^ 8 + strArr(1) * 10 ^ 4 + strArr(2)
    Else
        RtnVerNum = 0
    End If
End Function
 
Private Sub txt服务器编号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If

End Sub
Private Function GetCommpentVersion(ByVal strFile As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取指定控件的版本号
    '入参:
    '出参:
    '返回:成功,返回版本号,否则返回空
    '编制:刘兴洪
    '日期:2009-01-16 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    
    err = 0: On Error Resume Next
    '获取文件版本号
    strVer = objFile.GetFileVersion(strFile)
    If err <> 0 Then
        err.Clear: err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
    If Trim(strVer) <> "" Then
        varVersion = Split(strVer, ".")
        If UBound(varVersion) > 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
        ElseIf UBound(varVersion) = 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
        End If
    End If
    GetCommpentVersion = strVer
End Function

Private Function GetSetupPath(ByVal strFilename As String, ByVal strPathSign As String, ByVal strFileType As String, ByVal strPath As String) As String
    '--------------------------------------------------------------------------------------------------------
    '功能:获取收集文件的完整路径
    '返回:返回完整的路径
    '编制:祝庆
    '日期:2010/12/10
    '--------------------------------------------------------------------------------------------------------
    On Error GoTo ErrH
    Dim strTemp As String '临时路径组合
    Dim strSystemDirectory As String '系统system32目录
    Dim strWinDirectory As String  'windows目录
    strSystemDirectory = GetWinSystemPath
    strWinDirectory = GetWinPath
    
    If strFilename = "" Then
        GetSetupPath = ""
        Exit Function
    End If
    
    If Len(strPathSign) = 0 Then
        Select Case strFileType
        Case "0" '公共
            strTemp = strPath & "\Public\" & strFilename
        Case "1" '应用
            strTemp = strPath & "\Apply\" & strFilename
        Case "2" '帮助
            strTemp = strWinDirectory & "\Help\" & strFilename
        Case "3" '其它
            strTemp = strPath & "\" & strFilename
        Case "4" '三方
            strTemp = ""
        Case "5" '系统
            strPathSign = UCase(strPathSign)
            If (InStrRev(strPathSign, "[SYSTEM]", -1) > 0) Or (strPathSign = "") Then
                strTemp = strSystemDirectory & "\" & strFilename
            End If
            
            '新路径
            If InStrRev(strPathSign, "[PUBLIC]", -1) > 0 Then
                strTemp = strPath & "\PUBLIC\" & strFilename
            End If
        End Select
    Else
        strPathSign = UCase(strPathSign)
        If InStrRev(strPathSign, "[APPSOFT]", -1) > 0 Then
            strTemp = Replace(strPathSign, "[APPSOFT]", strPath)
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFilename
            Else
                strTemp = strTemp & strFilename
            End If
        ElseIf InStrRev(strPathSign, "[SYSTEM]", -1) > 0 Then
            strTemp = Replace(strPathSign, "[SYSTEM]", strSystemDirectory)
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFilename
            Else
                strTemp = strTemp & strFilename
            End If
        ElseIf InStrRev(strPathSign, "[PUBLIC]", -1) > 0 Then
            strTemp = Replace(strPathSign, "[PUBLIC]", strPath & "\PUBLIC")
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFilename
            Else
                strTemp = strTemp & strFilename
            End If
        ElseIf InStrRev(strPathSign, "[HELP]", -1) Then
            strTemp = Replace(strPathSign, "[HELP]", strWinDirectory & "\Help")
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFilename
            Else
                strTemp = strTemp & strFilename
            End If
        Else '完整路径
            If Left(strFilename, 2) = "\\" Then
                strTemp = ""
            Else
                strTemp = Left(strPath, 1) & Right(strFilename, Len(strFilename) - 1)
            End If
        End If
    End If
    
    GetSetupPath = strTemp
    Exit Function
ErrH:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function GetCompressName(ByVal strFilename As String) As String
'功能转换为7z后缀的压缩格式名称
    On Error GoTo ErrH
    GetCompressName = strFilename & ".7z"
    Exit Function
ErrH:
    If err Then
         MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Sub CompareFile()
'功能:比较文件是否需要收集
    On Error GoTo ErrH
    
    Dim i      As Long
    Dim strMD5 As String
    Dim lngErr As Long  '未安装的个数
    Dim lngSJ  As Long  '未收集的个数
    Dim strFullPath As String
    Dim strFilename As String
    Dim strSetupPath As String
    Dim strFileType As String
    
    If FindFile(mstrSourceFloder) = False Then
        Exit Sub
    End If

    If mshBill.Rows = 0 Then Exit Sub
    For i = 1 To mshBill.Rows - 1
        strMD5 = mshBill.TextMatrix(i, HeadInfor.MD5)
        
        If Len(strMD5) = 0 Then
            
            strFilename = mshBill.TextMatrix(i, HeadInfor.部件名)
            strSetupPath = mshBill.TextMatrix(i, HeadInfor.安装路径)
            strFileType = mshBill.TextMatrix(i, HeadInfor.类型)
            
            strFullPath = GetSetupPath(Nvl(strFilename, ""), Nvl(strSetupPath, ""), Nvl(strFileType, ""), mstrzlAppSoftPath)
            If FindFile(strFullPath) Then
                mshBill.TextMatrix(i, HeadInfor.信息) = "未收集的文件!"
                mshBill.TextMatrix(i, HeadInfor.收集类型) = "1"
                mshBill.SetRowColor i, vbBlue, False
                lngSJ = lngSJ + 1
            Else
                mshBill.TextMatrix(i, HeadInfor.信息) = "未安装文件!"
                mshBill.TextMatrix(i, HeadInfor.收集类型) = "2"
                mshBill.SetRowColor i, vbRed, False
                lngErr = lngErr + 1
            End If
            
        End If
    Next
    
    If lngErr = 0 Then
        stbThis.Panels(2).Text = ""
    Else
        stbThis.Panels(2).Text = lngErr & "个文件未安装 " & IIf(lngSJ = 0, "", lngSJ & "个文件未收集")
    End If
    Exit Sub
ErrH:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Function GetFileName(ByVal strFile As String) As String
'功能:去掉文件后缀的文件名
    Dim i As Integer
    If strFile = "" Then Exit Function
    i = InStrRev(strFile, ".")
    If i > 0 Then
        GetFileName = Left(strFile, i - 1)
    End If
End Function

Private Sub SaveUpList(upList As UpdateList)
    On Error GoTo ErrH
    Dim strSQL As String
    Dim i As Integer
    Dim strFilename As String
    Dim strMD5      As String 'MD5
    Dim str版本号   As String '版本号
    Dim str修改日期 As String '修改日期
    Dim strVision   As String
    Dim strArr    As Variant
    Dim lng版本号   As Double
    
    If mblnOptType = False Then
        strSQL = "update zlfilesupgrade set MD5= ''"
        gcnOracle.Execute strSQL
    End If
    
    If SafeArrayGetDim(upList.uFile) <> 0 Then
        gcnOracle.BeginTrans
        For i = 0 To UBound(upList.uFile)
            strFilename = upList.uFile(i).FileName
            strMD5 = upList.uFile(i).FileMD5
            str版本号 = upList.uFile(i).FileVision
            strVision = str版本号
            If strVision <> "" Then
                strArr = Split(strVision, ".")
                lng版本号 = strArr(0) * 10 ^ 8 + strArr(1) * 10 ^ 4 + strArr(2)
                strVision = lng版本号
            End If
            
            str修改日期 = upList.uFile(i).FileEditDate
            
            
            If strFilename <> "" And strMD5 <> "" Then
                strSQL = "update zlfilesupgrade set MD5= '" & strMD5 & "',版本号='" & strVision & "',修改日期='" & str修改日期 & "' where upper(文件名)='" & UCase(strFilename) & "'"
                gcnOracle.Execute strSQL
            End If
        Next
        gcnOracle.CommitTrans
    End If
    
    Exit Sub
ErrH:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
        Resume
        gcnOracle.RollbackTrans
    End If
End Sub

Private Function funCanWrite(strWritePath As String) As Boolean
'判断远程目录是否具有写权限
    Dim strDest     As String

    On Error GoTo ErrH
            strDest = strWritePath & "\tmp.txt"
            mobjFile.CreateTextFile strDest
            mobjFile.DeleteFile strDest, True
            funCanWrite = True
    Exit Function
ErrH:
    funCanWrite = False
End Function

Private Function GetTmpPath() As String
    Dim tmpBuffer As String
    tmpBuffer = String(255, Chr(0))
    GetTempPath 256, tmpBuffer
    GetTmpPath = Trim(Left(tmpBuffer, InStr(1, tmpBuffer, Chr(0)) - 1))
End Function

Private Sub FloderToClipBoard(ByVal strSourceFloder As String)
    '拷贝临时收集文件目录的文件到剪贴板中去
    Dim strFile() As String
    Dim strSourceFile As String
    Dim strTemp As String
    Dim i As Integer
    strSourceFile = strSourceFloder & "\"
    Erase strFile
    
    
    If mobjFile.FolderExists(strSourceFile) Then
        With FileList
            .Refresh
            .Path = strSourceFile
            .FileName = "*.*"
            
            For i = 0 To .ListCount - 1
                ReDim Preserve strFile(i)
                strTemp = strSourceFile & .List(i)
                strFile(i) = strTemp
            Next
            
            If .ListCount <> 0 Then
                Call clipCopyFiles(strFile)
            End If
        End With
    End If
End Sub

Private Sub BillFileSort()
    On Error GoTo ErrH
    Dim lngRow As Long
    Dim curRow As Long
    Dim strGradeType As String
    curRow = 1
    LoadHeadInforShow
    'strGradeType 0 正常收集 1 未收集文件 2 未安装路径
    
    '2 为安装的路径
    For lngRow = 1 To mshBill.Rows - 1
        strGradeType = mshBill.TextMatrix(lngRow, HeadInfor.收集类型)
        If strGradeType = "2" Then
            With mshBillShow
                 .TextMatrix(curRow, HeadInfor.序号) = curRow
                 .TextMatrix(curRow, HeadInfor.部件名) = mshBill.TextMatrix(lngRow, HeadInfor.部件名)
                 .TextMatrix(curRow, HeadInfor.版本号) = mshBill.TextMatrix(lngRow, HeadInfor.版本号)
                 .TextMatrix(curRow, HeadInfor.修改日期) = mshBill.TextMatrix(lngRow, HeadInfor.修改日期)
                 .TextMatrix(curRow, HeadInfor.信息) = mshBill.TextMatrix(lngRow, HeadInfor.信息)
                 .TextMatrix(curRow, HeadInfor.加入日期) = mshBill.TextMatrix(lngRow, HeadInfor.加入日期)
                 .TextMatrix(curRow, HeadInfor.说明) = mshBill.TextMatrix(lngRow, HeadInfor.说明)
                 .TextMatrix(curRow, HeadInfor.类型) = mshBill.TextMatrix(lngRow, HeadInfor.类型)
                 .TextMatrix(curRow, HeadInfor.安装路径) = mshBill.TextMatrix(lngRow, HeadInfor.安装路径)
                 .TextMatrix(curRow, HeadInfor.MD5) = mshBill.TextMatrix(lngRow, HeadInfor.MD5)
                 .TextMatrix(curRow, HeadInfor.收集类型) = mshBill.TextMatrix(lngRow, HeadInfor.收集类型)
                 .Rows = .Rows + 1
                 .SetRowColor curRow, vbRed, False
            End With
            curRow = curRow + 1
        End If
    Next
    
    '1 未收集的文件
    For lngRow = 1 To mshBill.Rows - 1
        strGradeType = mshBill.TextMatrix(lngRow, HeadInfor.收集类型)
        If strGradeType = "1" Then
            With mshBillShow
                 .TextMatrix(curRow, HeadInfor.序号) = curRow
                 .TextMatrix(curRow, HeadInfor.部件名) = mshBill.TextMatrix(lngRow, HeadInfor.部件名)
                 .TextMatrix(curRow, HeadInfor.版本号) = mshBill.TextMatrix(lngRow, HeadInfor.版本号)
                 .TextMatrix(curRow, HeadInfor.修改日期) = mshBill.TextMatrix(lngRow, HeadInfor.修改日期)
                 .TextMatrix(curRow, HeadInfor.信息) = mshBill.TextMatrix(lngRow, HeadInfor.信息)
                 .TextMatrix(curRow, HeadInfor.加入日期) = mshBill.TextMatrix(lngRow, HeadInfor.加入日期)
                 .TextMatrix(curRow, HeadInfor.说明) = mshBill.TextMatrix(lngRow, HeadInfor.说明)
                 .TextMatrix(curRow, HeadInfor.类型) = mshBill.TextMatrix(lngRow, HeadInfor.类型)
                 .TextMatrix(curRow, HeadInfor.安装路径) = mshBill.TextMatrix(lngRow, HeadInfor.安装路径)
                 .TextMatrix(curRow, HeadInfor.MD5) = mshBill.TextMatrix(lngRow, HeadInfor.MD5)
                 .TextMatrix(curRow, HeadInfor.收集类型) = mshBill.TextMatrix(lngRow, HeadInfor.收集类型)
                 .Rows = .Rows + 1
                 .SetRowColor curRow, vbBlue, False
            End With
            curRow = curRow + 1
        End If
    Next
    
    '0 未收集的文件
    For lngRow = 1 To mshBill.Rows - 1
        strGradeType = mshBill.TextMatrix(lngRow, HeadInfor.收集类型)
        If strGradeType = "0" Then
            With mshBillShow
                 .TextMatrix(curRow, HeadInfor.序号) = curRow
                 .TextMatrix(curRow, HeadInfor.部件名) = mshBill.TextMatrix(lngRow, HeadInfor.部件名)
                 .TextMatrix(curRow, HeadInfor.版本号) = mshBill.TextMatrix(lngRow, HeadInfor.版本号)
                 .TextMatrix(curRow, HeadInfor.修改日期) = mshBill.TextMatrix(lngRow, HeadInfor.修改日期)
                 .TextMatrix(curRow, HeadInfor.信息) = mshBill.TextMatrix(lngRow, HeadInfor.信息)
                 .TextMatrix(curRow, HeadInfor.加入日期) = mshBill.TextMatrix(lngRow, HeadInfor.加入日期)
                 .TextMatrix(curRow, HeadInfor.说明) = mshBill.TextMatrix(lngRow, HeadInfor.说明)
                 .TextMatrix(curRow, HeadInfor.类型) = mshBill.TextMatrix(lngRow, HeadInfor.类型)
                 .TextMatrix(curRow, HeadInfor.安装路径) = mshBill.TextMatrix(lngRow, HeadInfor.安装路径)
                 .TextMatrix(curRow, HeadInfor.MD5) = mshBill.TextMatrix(lngRow, HeadInfor.MD5)
                 .TextMatrix(curRow, HeadInfor.收集类型) = mshBill.TextMatrix(lngRow, HeadInfor.收集类型)
                 .Rows = .Rows + 1
                 .SetRowColor curRow, &HFFFFFF, False
            End With
            curRow = curRow + 1
        End If
    Next
    
    If mshBillShow.Rows <> 0 Then
        mshBillShow.Rows = mshBillShow.Rows - 1
    End If
    
    Me.mshBill.Visible = False
    Me.mshBillShow.Visible = True
    Exit Sub
ErrH:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub chk升级日期_Click()
    If chk升级日期.value = 1 Then
        DTP升级日期.Enabled = True
        cmd升级日期.Enabled = True
    Else
        DTP升级日期.Enabled = False
        cmd升级日期.Enabled = True
    End If
End Sub

Private Sub cmd升级日期_Click()
    Dim strNow As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ErrHand
    
    If chk升级日期.value = 1 Then
        strNow = Format(CurrentDate(), "yyyy-MM-dd")
        If DTP升级日期 < CDate(strNow) Then
            MsgBox "升级日期不能小于当前服务器日期!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        
        Set rsTmp = New ADODB.Recordset
        gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 = '客户端升级日期'"
        Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        
        If rsTmp.EOF = False Then
            strSQL = "Update zlRegInfo Set 内容='" & Format(DTP升级日期, "yyyy-MM-dd") & "' Where 项目='客户端升级日期'"
            gcnOracle.Execute strSQL
        Else
            strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('客户端升级日期',Null,'" & Format(DTP升级日期, "yyyy-MM-dd") & "')"
            gcnOracle.Execute strSQL
        End If
        
    Else
        Set rsTmp = New ADODB.Recordset
        gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 = '客户端升级日期'"
        Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        
        If rsTmp.EOF = False Then
            strSQL = "Update zlRegInfo Set 内容=Null Where 项目='客户端升级日期'"
            gcnOracle.Execute strSQL
        Else
            strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('客户端升级日期',Null,Null)"
            gcnOracle.Execute strSQL
        End If
    
    End If

  Exit Sub
ErrHand:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Sub OpinionUpGradeDate()
    '判断是否设定了升级日期
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
     
    Set rsTmp = New ADODB.Recordset
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 = '客户端升级日期'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    cmd升级日期.Enabled = True
    
    If rsTmp.EOF = False Then
        If Nvl(rsTmp!内容) = "" Then
            chk升级日期.Enabled = True
            DTP升级日期.Enabled = False
            
            chk升级日期.value = 0
            DTP升级日期.value = Format(CurrentDate(), "yyyy-MM-dd")
        Else
            chk升级日期.Enabled = True
            DTP升级日期.Enabled = True
        
            chk升级日期.value = 1
            DTP升级日期.value = Nvl(rsTmp!内容, Format(CurrentDate(), "yyyy-MM-dd"))
        End If
    Else
        chk升级日期.Enabled = True
        DTP升级日期.Enabled = False
        
        chk升级日期.value = 0
        DTP升级日期.value = Format(CurrentDate(), "yyyy-MM-dd")
    End If
    
    Exit Sub
ErrHand:
End Sub
