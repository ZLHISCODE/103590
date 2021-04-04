VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisPic2Ftp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "检验图片数据转移"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   15255
   Icon            =   "frmLisPic2Ftp.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmLisPic2Ftp.frx":6852
   ScaleHeight     =   10110
   ScaleWidth      =   15255
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame frmType 
      BackColor       =   &H00FFFFFF&
      Caption         =   "数据来源"
      Height          =   645
      Left            =   840
      TabIndex        =   43
      Top             =   5280
      Width           =   10800
      Begin VB.OptionButton optNew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "新版LIS"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton optOld 
         BackColor       =   &H00FFFFFF&
         Caption         =   "老版LIS"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label lblBanner 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3000
         TabIndex        =   46
         Top             =   330
         Width           =   90
      End
   End
   Begin VB.Timer Timer 
      Interval        =   3000
      Left            =   12960
      Top             =   1680
   End
   Begin VB.CommandButton cmdMulti 
      Caption         =   "转存图片(&O)"
      Height          =   350
      Left            =   8880
      TabIndex        =   15
      ToolTipText     =   "将数据库中的图形数据转换为图片保存到本地或FTP服务器"
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Frame frmFtpUp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "处理模式"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   840
      TabIndex        =   37
      Top             =   3600
      Width           =   10815
      Begin VB.TextBox txtProc 
         Height          =   300
         Left            =   840
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "2"
         Top             =   1140
         Width           =   375
      End
      Begin VB.OptionButton optAuto 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 实时上传"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optManu 
         BackColor       =   &H00FFFFFF&
         Caption         =   " 异步上传"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label lblFTP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "使用多进程加快图片转存的速度,为了防止超出计算机资源限制,最大进程数为10。"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   2
         Left            =   1560
         TabIndex        =   42
         Top             =   1200
         Width           =   6480
      End
      Begin VB.Label lblPorc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "进程数"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label lblFTP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "所有图片转存到本地后，使用其他工具批量上传。需要大量本地空间，总体耗时短."
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   1560
         TabIndex        =   39
         Top             =   750
         Width           =   6570
      End
      Begin VB.Label lblFTP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "每张图片转出后立即上传到FTP。所需本地空间小，总体耗时较长。"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   38
         Top             =   330
         Width           =   5310
      End
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "更新信息(&U)"
      Height          =   350
      Left            =   10320
      TabIndex        =   16
      ToolTipText     =   $"frmLisPic2Ftp.frx":E88C
      Top             =   7320
      Width           =   1335
   End
   Begin VB.FileListBox fileList 
      Height          =   450
      Left            =   10560
      Pattern         =   "*.jpg;*.png"
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame frmDownd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "处理范围"
      Height          =   1245
      Left            =   840
      TabIndex        =   23
      Top             =   6000
      Width           =   10800
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "全部数据"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optPart 
         BackColor       =   &H00FFFFFF&
         Caption         =   "部分数据"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   1095
      End
      Begin VB.PictureBox pctTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         ScaleHeight     =   375
         ScaleWidth      =   5415
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   5415
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   345
            Left            =   840
            TabIndex        =   13
            Top             =   15
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   223608835
            CurrentDate     =   43077.4366203704
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   345
            Left            =   3240
            TabIndex        =   14
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   223608835
            CurrentDate     =   43077.4366782407
         End
         Begin VB.Label lblDown 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "开始时间"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   0
            TabIndex        =   35
            Top             =   70
            Width           =   720
         End
         Begin VB.Label lblDown 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "结束时间"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   2400
            TabIndex        =   34
            Top             =   75
            Width           =   720
         End
      End
   End
   Begin VB.Frame fraFTP 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FTP设置"
      Height          =   1290
      Left            =   840
      TabIndex        =   0
      Top             =   2160
      Width           =   10800
      Begin VB.CommandButton cmdFile 
         Caption         =   "…"
         Height          =   300
         Left            =   7800
         TabIndex        =   6
         Top             =   780
         Width           =   300
      End
      Begin VB.TextBox txtTmpPath 
         Height          =   300
         Left            =   4320
         TabIndex        =   5
         ToolTipText     =   "图形数据在转出时会生成图片保存在临时路径下,请设置足够的保存空间"
         Top             =   780
         Width           =   3495
      End
      Begin VB.CommandButton cmdFtp 
         Caption         =   "连接测试"
         Height          =   350
         Left            =   8280
         TabIndex        =   7
         Top             =   755
         Width           =   1215
      End
      Begin VB.TextBox txtFTPPath 
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   780
         Width           =   1500
      End
      Begin VB.TextBox txtFTPIP 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   315
         Width           =   1500
      End
      Begin VB.TextBox txtFTPPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6480
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   315
         Width           =   1620
      End
      Begin VB.TextBox txtFTPUser 
         Height          =   300
         Left            =   4320
         TabIndex        =   2
         Top             =   315
         Width           =   1500
      End
      Begin VB.Label lblDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "本地临时路径"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3195
         TabIndex        =   40
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblFTP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FTP文件路径"
         Height          =   180
         Index           =   6
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblFTP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "地址"
         Height          =   180
         Index           =   5
         Left            =   990
         TabIndex        =   19
         Top             =   375
         Width           =   360
      End
      Begin VB.Label lblFTP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         Height          =   180
         Index           =   4
         Left            =   6075
         TabIndex        =   18
         Top             =   375
         Width           =   360
      End
      Begin VB.Label lblFTP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户"
         Height          =   180
         Index           =   3
         Left            =   3915
         TabIndex        =   17
         Top             =   375
         Width           =   360
      End
   End
   Begin VB.PictureBox pctResult 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   720
      ScaleHeight     =   975
      ScaleWidth      =   10935
      TabIndex        =   28
      Top             =   9240
      Visible         =   0   'False
      Width           =   10935
      Begin VB.Label lblReult 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "共计耗时:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblReult 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "获取数据-9000;上传数据-8500;获取失败-1000;上传失败-500"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   4860
      End
      Begin VB.Label lblReult 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "转出结果:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.PictureBox pctProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   720
      ScaleHeight     =   1215
      ScaleWidth      =   10935
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   10935
      Begin MSComctlLib.ProgressBar pgsBar 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   10000
      End
      Begin VB.Label lblProgress 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "已经转出4000条"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblProgress 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "共有10000条数据待转出"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   1890
      End
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "正在查询待转出数据..."
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   840
      TabIndex        =   32
      Top             =   7320
      Width           =   1890
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmLisPic2Ftp.frx":E922
      Stretch         =   -1  'True
      Top             =   648
      Width           =   480
   End
   Begin VB.Label lblTip 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLisPic2Ftp.frx":F19D
      Height          =   1440
      Left            =   840
      TabIndex        =   22
      Top             =   720
      Width           =   10815
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检验图片数据转移"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   21
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "frmLisPic2Ftp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrTmpPath As String    '下载图片临时路径
Private mstrFtpPath As String    '上传图片FTP路径
Private mblnFtpConncted  As Boolean      '标识FTP是否已经连接
Private mlngImgUp As Long '本地路径下待上传图片数量
Private mlngImgDown As Long '本地路径下待下载图片数量
Private mintCpu As Integer  'CPU建议值
Private mclsFtp As New clsFtp   'FTP类
Private mblnUpload As Boolean   '是否正在转出
Private mdblTime As Double '转出时间
Private mintLisBanner  As Integer 'LIS版本:0=没有安装Lis 1=只有旧版LIS 2=只有新版LIS 3=两者均有

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function
Private Sub cmdCommit_Click()
    '数据转出完毕后,修改源数据
    Dim strSQL As String, strMsg As String
    Dim lngOldNums As Long, lngNewNums As Long, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If CheckTblExist("检验图像结果_Exp_Temp") Then
        strSQL = "Select Count(1) 数量 From 检验图像结果_Exp_Temp"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetNums")
        lngOldNums = rsTmp!数量
    End If
    If CheckTblExist("检验报告图像_Exp_Temp") Then
        strSQL = "Select Count(1) 数量 From 检验报告图像_Exp_Temp"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetNums")
        lngNewNums = rsTmp!数量
    End If
    
    If lngOldNums + lngNewNums = 0 Then
        MsgBox "临时数据表没有数据,请先执行转存图片功能。", , "提示"
        Exit Sub
    End If
    
    strMsg = "本操作将会采用删除后插入新数据的方式，更新原表的图片路径信息,同时清除LOB字段的图像数据和临时数据表。" & vbNewLine & _
                    "当前待修改至数据库共有" & lngOldNums + lngNewNums & "条数据，请在上传完所有图片到FTP并检查确认后执行。你确认要继续吗？"
    If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then
        Exit Sub
    End If
    
    pctProgress.Visible = False
    pctResult.Visible = False
    If lngOldNums > 0 Then UpdatePic 1
    If lngNewNums > 0 Then UpdatePic 2
    MousePointer = vbDefault
    lblState.Caption = "数据更新成功,修改数据" & lngOldNums + lngNewNums & "条。"
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cmdFile_Click()
    Dim strTmp As String, blnTmp As Boolean
    
    strTmp = OpenFolder(Me, "请选择临时路径", mstrTmpPath)
    If strTmp = "" Then
        Exit Sub
    End If
    
    '如果两次选择的临时路径不匹配,会导致上一次上传失败的文件保留,造成转移数据丢失
    blnTmp = True
    If mstrTmpPath <> strTmp And mstrTmpPath <> "" Then
        blnTmp = MsgBox("当前选择的临时路径与上一次的临时路径不同，无法继续上传上次未保存至FTP的图片，是否继续？" & vbNewLine & "注：未上传的图片可以手动上传。" _
        , vbYesNo + vbQuestion + vbDefaultButton2, "确认") = vbYes
    End If
    
    If blnTmp Then
        mstrTmpPath = strTmp
        If Right(mstrTmpPath, 1) = "\" Then
            mstrTmpPath = Mid(mstrTmpPath, 1, Len(mstrTmpPath) - 1)
        End If
        Call SaveSetting("LIS图片转出", "转出设置", "临时路径", mstrTmpPath)
        txtTmpPath.Text = strTmp
        fileList.Path = strTmp
        
        CheckImg
    End If
End Sub

Private Sub cmdFtp_Click()
    If TestFtp() Then lblState.Caption = "FTP连接验证通过"
End Sub


Private Sub cmdMulti_Click()
    Dim i As Integer, strSQL As String, rsTmp As ADODB.Recordset
    Dim rsExp As ADODB.Recordset, strMsg As String
    Dim lngSize As Long, lngTmp As Long, strPath As String
    Dim intProcNum As Integer, lngDays As Long
    
    SetCmdEnable False
    pctProgress.Visible = False
    pctResult.Visible = False
    
    '首先检查FTP连接是否通过
    If Not TestFtp Then
        MousePointer = vbDefault
        SetCmdEnable True
        Exit Sub
    End If
    
    '检查子程序是否存在
    If Right(App.Path, 1) = "\" Then
        strPath = Mid(App.Path, 1, Len(App.Path) - 1)
    Else
        strPath = App.Path
    End If
    If Not gobjFile.FileExists(strPath & "\zlLisPic2FtpSub.exe") Then
        MsgBox "目录" & strPath & "下不存在执行文件:" & "zlLisPic2FtpSub.exe,无法继续操作。", , gstrSysName
        SetCmdEnable True
        Exit Sub
    End If
    
    CreateTable IIf(optOld.Value, 1, 2)
    If optOld.Value = True Then
    '已转出图片大于10000时,就进行提示,是否先提交一次再进行转出操作
        strSQL = "Select Count(1) 数量 From 检验图像结果_EXP_TEMP"
    Else
        strSQL = "Select Count(1) 数量 From 检验报告图像_EXP_TEMP"
    End If
    Set rsExp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetExp")

    If rsExp!数量 > 100000 Then
        strMsg = "当前已经转存了临时图片数据" & rsExp!数量 & "条，是否先更新信息至数据库?" & vbNewLine & _
                        "注:临时表中数据过多,会导致转存图片变慢。点击是更新信息，点击否继续进行图片转存"
        If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbYes Then
            lblState.Caption = "正在提交已转存数据至数据库中..."
            MousePointer = vbArrowHourglass
            UpdatePic IIf(optOld.Value, 1, 2)
            lblState.Caption = ""
            MousePointer = vbDefault
            SetCmdEnable True
            Exit Sub
        End If
    End If
    
    lblState.Caption = "检查转存相关数据..."
    MousePointer = vbArrowHourglass
    lblState.Refresh
    mlngImgDown = GetDownNum(IIf(optOld.Value, 1, 2))
    
    If mlngImgDown = 0 Then
        MsgBox "数据库中所有图片数据都已经转出", , "提示"
        lblState.Caption = ""
        MousePointer = vbDefault
        SetCmdEnable True
        Exit Sub
    Else
        If optManu.Value Then   '手动上传,需要提示占用本地空间
            If optPart.Value Then
                strMsg = "本次数据转出，共需转出图片数据" & mlngImgDown & "条，请检查后确认是否继续？"
            Else
                lngSize = GetLobSize(IIf(optOld.Value, 1, 2))
                strMsg = "本次数据转出，共需转出图片数据" & mlngImgDown & "条，预计占用临时空间" & lngSize & "M，请检查后确认是否继续？"
            End If
        Else
            strMsg = "本次数据转出，共需转出图片数据" & mlngImgDown & "条，请检查后确认是否继续？"
        End If
        
        If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then
            lblState.Caption = ""
            MousePointer = vbDefault
            SetCmdEnable True
            Exit Sub
        End If
    End If
    

    
    '界面刷新
    lblState.Caption = "正在开启转出进程..."
    lblState.Refresh
    lngDays = DateDiff("d", CDate(dtpStart.Tag), CDate(dtpEnd.Tag)) + 1
    pgsBar.Max = lngDays
    pgsBar.Value = 0
    lblProgress(0).Caption = Format(CDate(dtpStart.Tag), "yyyy/mm/dd") & "到" & Format(CDate(dtpEnd.Tag), "yyyy/mm/dd") & "共有" & lngDays & "天的数据待转出"
    lblProgress(1).Caption = "已经转出0天的数据。"

    pctProgress.Visible = True
    pctResult.Visible = False
    MousePointer = vbArrowHourglass
    Me.Refresh
    
    '开启进程
    mblnUpload = True
    mdblTime = 0
    SaveSetting "LIS图片转出", "转出进度", "转出错误", ""   '开始操作时 先把错误清空
    intProcNum = IIf(Val(txtProc.Text) = 0, 1, txtProc.Text)
    For i = 1 To intProcNum
        SaveSetting "LIS图片转出", "转出进度", "进程" & i, 0
        
        '通知进程开启转出
        '命令格式: 转出类型(1-FTP 2-保存本地);数据来源(1-旧版LIS 2-新版LIS);进程号;开始时间;结束时间;临时路径;FTP路径
        If optAuto.Value = True Then
            SaveSetting "LIS图片转出", "转出进度", "进程设置", "1;" & IIf(optOld.Value, 1, 2) & ";" & i & ";" & intProcNum & ";" & Format(dtpStart.Tag, "yyyy/mm/dd") & ";" & Format(dtpEnd.Tag, "yyyy/mm/dd") & ";" & mstrTmpPath & ";" & mstrFtpPath '同步上传
        Else
            SaveSetting "LIS图片转出", "转出进度", "进程设置", "2;" & IIf(optOld.Value, 1, 2) & ";" & i & ";" & intProcNum & ";" & Format(dtpStart.Tag, "yyyy/mm/dd") & ";" & Format(dtpEnd.Tag, "yyyy/mm/dd") & ";" & mstrTmpPath & ";" & mstrFtpPath    '异步上传
        End If
        Shell """" & strPath & "\zlLisPic2FtpSub.exe"" ""zlUserName=" & gstrUserName & "zlPassword=" & gstrPassword & "zlServer=" & gstrServer & " ", vbMaximizedFocus
    Next
    Call Timer_Timer
    

End Sub

Private Sub Form_load()

    On Error GoTo errH
    Call LoadFtpPara
    mstrTmpPath = GetSetting("LIS图片转出", "转出设置", "临时路径")
    txtTmpPath.Text = mstrTmpPath
    dtpStart.Value = Now: dtpEnd.Value = Now
    
    Call CheckImg
    
    mintCpu = GetCpuAdv
    mintLisBanner = CheckLisSys
    If mintLisBanner = 1 Or mintLisBanner = 3 Then
        SetDtpPicker 1
    ElseIf mintLisBanner = 2 Then
        SetDtpPicker 2
    End If
    
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsFtp = Nothing
End Sub

Private Sub optNew_Click()
    Call SetDtpPicker(2)
End Sub

Private Sub optOld_Click()
    Call SetDtpPicker(1)
End Sub

Private Sub optPart_Click()
    pctTime.Visible = True
End Sub

Private Sub optAll_Click()
    pctTime.Visible = False
End Sub

Private Sub LoadFtpPara()
    '功能:获取用户的FTP参数设置
    Dim strFtpParas As String
    
    On Error GoTo errH
    strFtpParas = gclsBase.GetPara("FTP设置", 100, 1208, 1)
    
    If strFtpParas = "" Then
        strFtpParas = GetSetting("LIS图片转出", "转出设置", "FTP路径")
    End If
        
    If strFtpParas <> "" Then
        txtFTPUser.Text = Split(strFtpParas, ";")(0)
        txtFTPPWD.Text = Split(strFtpParas, ";")(1)
        txtFTPIP.Text = Split(strFtpParas, ";")(2)
        txtFTPPath.Text = Split(strFtpParas, ";")(3)
    End If
    Exit Sub
errH:
    MsgBox "FTP设置读取失败", vbExclamation, gstrSysName
End Sub

Private Function TestFtp() As Boolean
    '功能:根据ftp参数验证连接是否通过
    Dim strUser As String, strPwd As String
    Dim strIp As String, strPath As String
    
    On Error GoTo errH
    strUser = Trim(txtFTPUser.Text): strPwd = Trim(txtFTPPWD.Text)
    strIp = Trim(txtFTPIP.Text): strPath = Trim(txtFTPPath.Text)
    mstrFtpPath = strPath
    
    SaveSetting "LIS图片转出", "转出设置", "FTP路径", strUser & ";" & strPwd & ";" & strIp & ";" & strPath
    
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect  '如果已经连接了FTP,那么就断开之前的,防止占用FTP会话量
    mblnFtpConncted = False
    '进行连接测试
    If mclsFtp.FuncFtpConnect(strIp, strUser, strPwd) = 0 Then
        MsgBox "连接FTP服务器失败，请检查服务器地址及帐号。"
        Exit Function
    End If
    
    '创建一个文件进行上传测试
    If Not gobjFile.FolderExists(mstrTmpPath) Then
        MsgBox "临时路径不存在,请重新输入", , "提示"
        Exit Function
    End If
    If Not gobjFile.FileExists(mstrTmpPath & "\tmp") Then
        gobjFile.CreateTextFile mstrTmpPath & "\tmp", True
    End If
    
    If mclsFtp.FuncUploadFile(strPath, mstrTmpPath & "\tmp", "tmp") <> 0 Then
        MsgBox "FTP路径错误,未能通过连接"
        Exit Function
    End If
    

    '删除临时文件
    Kill mstrTmpPath & "\tmp"
    mblnFtpConncted = True
    TestFtp = True
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Function

Private Function CheckImg() As Boolean
    '功能:检查临时目录下是否有图片未上传
    
    If gobjFile.FolderExists(mstrTmpPath) Then
        With fileList
            .Path = mstrTmpPath
            .Refresh
            mlngImgUp = .ListCount
        End With
    Else
        mlngImgUp = 0
    End If
    
    If mlngImgUp > 0 Then
        lblState.Caption = "当前临时路径下共有" & mlngImgUp & "张图片未上传,请使用图片转存功能继续上传"
        CheckImg = True
    Else
        lblState.Caption = ""
        CheckImg = False
    End If
End Function

Private Sub Timer_Timer()
    Dim blnDone As Boolean, strTmp As String
    Dim lngDone As Long, intExit As Integer
    Dim i As Integer, intProcNum As Integer, intActive As Integer
    
    If Not mblnUpload Then Exit Sub

    If mdblTime = 0 Then mdblTime = GetTickCount
    
    blnDone = True
    lngDone = 0: intExit = 0
    intProcNum = IIf(txtProc.Text = 0, 1, txtProc.Text)
    
    '循环每个进程
    For i = 1 To intProcNum
        strTmp = GetSetting("LIS图片转出", "转出进度", "进程" & i, 0)   '转出完毕后,将进度修改为 数量;
        If IsNumeric(strTmp) Then
            lngDone = lngDone + strTmp  '已经转出的数量
        Else
            lngDone = lngDone + Val(Mid(strTmp, 1, Len(strTmp) - 1))
            intExit = intExit + 1   '已经完成转出后退出的进程数量
        End If
        
        If lngDone < pgsBar.Max Then
            blnDone = False
        Else
            blnDone = True
        End If
    Next
    
    '判断进程是否因为特殊原因终止
    intActive = CheckProcExist("zllispic2ftpsub.exe")
    If intActive <> intProcNum Then
        If intProcNum <> intActive + intExit And lngDone <> pgsBar.Max Then    '退出进程+活跃进程<>进程总数 说明有进程以外终止
            SaveSetting "LIS图片转出", "转出进度", "转出错误", "进程被意外终止"
            blnDone = True
        End If
    End If
    '正在下载
    If Not blnDone Then
        lblState.Caption = "正在进行图片转出..."
        lblProgress(1).Caption = "已经转出" & lngDone & "天的数据。"
        lblProgress(1).Refresh
        
        pgsBar.Value = IIf(lngDone > pgsBar.Max, pgsBar.Max, lngDone)
        pgsBar.Refresh
    Else
        '下载完成
        SetCmdEnable True
        MousePointer = vbDefault
        mblnUpload = False
        pctResult.Visible = True
        pctResult.Refresh
        lblProgress(1).Caption = "已经转出" & lngDone & "天的数据。"
        If GetSetting("LIS图片转出", "转出进度", "转出错误") <> "" Then
            lblReult(1).Caption = "转存中发生错误,错误信息已经保存至当前目录下的Lis2FtpErrLog日志文件中"
            lblState.Caption = "转存中发生错误,请检查后重试。"
        Else
            lblReult(1).Caption = "共转存图片:" & mlngImgDown & "张"
            lblState.Caption = "图片转存成功,请使用提交数据功能,将修改后的结果改变至数据库"
            pgsBar.Value = pgsBar.Max   '保证进度条到最大
        End If
        lblReult(2).Caption = "共计耗时:" & Format((GetTickCount - mdblTime) / 1000, "0.00") & "S"
    End If
End Sub

Private Sub txtFTPIP_Change()
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect '如果已经连接了FTP,那么就断开之前的,防止占用FTP会话量
    mblnFtpConncted = False
End Sub

Private Sub txtFTPIP_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtFTPPath_Change()
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect '如果已经连接了FTP,那么就断开之前的,防止占用FTP会话量
    mblnFtpConncted = False
End Sub

Private Sub txtFTPPath_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtFTPPWD_Change()
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect '如果已经连接了FTP,那么就断开之前的,防止占用FTP会话量
    mblnFtpConncted = False
End Sub

Private Sub txtFTPPWD_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtFTPUser_Change()
    If mblnFtpConncted Then mclsFtp.FuncFtpDisConnect '如果已经连接了FTP,那么就断开之前的,防止占用FTP会话量
    mblnFtpConncted = False
End Sub

Private Function GetDownNum(ByVal intType As Integer) As Long
    '功能:获取需要从数据库中进行转换的图片数量
    '参数: intType 1=旧版LIS 2=新版LIS
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngTmp As Long
    
    '1.区分全部数据还是部分数据
    If intType = 1 Then
        strSQL = "Select " & IIf(mintCpu = 0, "", "/*+ parallel(a," & mintCpu & ") parallel(b," & mintCpu & ")*/ ") & vbNewLine & _
                        " Count(1) 数量, Max(a.核收时间) 结束时间, Min(a.核收时间) 开始时间" & vbNewLine & _
                        "From 检验标本记录 A, 检验图像结果 B" & vbNewLine & _
                        "Where a.审核人 Is Not Null And a.Id = b.标本id And b.图像位置 Is Null And b.图像点 Is Not Null"
    Else
        strSQL = "Select " & IIf(mintCpu = 0, "", "/*+ parallel(a," & mintCpu & ") parallel(b," & mintCpu & ")*/ ") & vbNewLine & _
                " Count(1) 数量, Max(a.核收时间) 结束时间, Min(a.核收时间) 开始时间" & vbNewLine & _
                "From 检验报告记录 A, 检验报告图像 B" & vbNewLine & _
                "Where a.审核人 Is Not Null And a.Id = b.标本id And b.图像位置 Is Null And b.图像点 Is Not Null"
    End If
    strSQL = strSQL & IIf(optPart = True, " And a.核收时间 Between [1] And [2]", "")
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetDownNum", CDate(Format(dtpStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")))
    lngTmp = rsTmp!数量
    
    If lngTmp = 0 Then Exit Function
    
    dtpStart.Tag = rsTmp!开始时间
    dtpEnd.Tag = rsTmp!结束时间
    
    '2.从检验图像结果_EXP_TEMP中计算已转出的数据
    If Not optPart.Value Then '如果勾选了全部数据
        If intType = 1 Then
            strSQL = "Select count(1) 数量 From 检验图像结果_EXP_TEMP"
        Else
            strSQL = "Select count(1) 数量 From 检验报告图像_EXP_TEMP"
        End If
    Else
        If intType = 1 Then
            strSQL = "Select Count(1) 数量" & vbNewLine & _
                            "From 检验标本记录 A, 检验图像结果_EXP_TEMP B" & vbNewLine & _
                            "Where a.审核人 Is Not Null And a.Id = b.标本id" & vbNewLine & _
                            "And a.核收时间 Between [1] And [2]"
        Else
            strSQL = "Select Count(1) 数量" & vbNewLine & _
                            "From 检验报告记录 A, 检验报告图像_EXP_TEMP B" & vbNewLine & _
                            "Where a.审核人 Is Not Null And a.Id = b.标本id" & vbNewLine & _
                            "And a.核收时间 Between [1] And [2]"
        End If
    End If
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetDownNum", CDate(Format(dtpStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")))
    lngTmp = lngTmp - rsTmp!数量
    
    GetDownNum = lngTmp
End Function

Private Sub txtFTPUser_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtProc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtProc_LostFocus()
    txtProc.Text = IIf(Val(txtProc.Text) = 0, 1, Val(txtProc.Text))
End Sub

Private Sub txtTmpPath_Change()
    mstrTmpPath = Trim(txtTmpPath.Text)
    
    If Right(mstrTmpPath, 1) = "\" Then
        mstrTmpPath = Mid(mstrTmpPath, 1, Len(mstrTmpPath) - 1)
    End If
End Sub

Private Function GetLobSize(ByVal intType As Integer) As Long
    '功能:估算占用空间
    '参数: intType 1=旧版LIS 2=新版LIS
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If intType = 1 Then
        strSQL = "Select a.Segment_Name, Round(a.Bytes / 1024 / 1024)  As Lobsize" & vbNewLine & _
                        "From User_Segments A, User_Lobs B" & vbNewLine & _
                        "Where b.Table_Name  = '检验图像结果' And a.Segment_Name = b.Segment_Name"
    Else
        strSQL = "Select a.Segment_Name, Round(a.Bytes / 1024 / 1024)  As Lobsize" & vbNewLine & _
                "From User_Segments A, User_Lobs B" & vbNewLine & _
                "Where b.Table_Name  = '检验报告图像' And a.Segment_Name = b.Segment_Name"
    End If
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetLobSize")
    If rsTmp.RecordCount = 0 Then Exit Function
    
    GetLobSize = rsTmp!Lobsize & ""
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Function

Private Function UpdatePic(ByVal intType As Integer) As Boolean
    '功能:更新数据至数据库
    '参数: intType 1=旧版LIS 2=新版LIS
    Dim strSQL As String
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    
    If intType = 1 Then
        strSQL = "Delete 检验图像结果 Where ID In (Select ID From 检验图像结果_Exp_Temp)"
        gcnOracle.Execute strSQL
        
        strSQL = "Insert Into /*+ append */ 检验图像结果" & vbNewLine & _
                        "  (ID, 标本id, 图像类型, 图像点, 图像位置, 待转出)" & vbNewLine & _
                        "  Select ID, 标本id, 图像类型, Null, 图像位置, Null From 检验图像结果_Exp_Temp"
        gcnOracle.Execute strSQL
    Else
        strSQL = "Delete 检验报告图像 Where ID In (Select ID From 检验报告图像_Exp_Temp)"
        gcnOracle.Execute strSQL
        
        strSQL = "Insert Into /*+ append */ 检验报告图像" & vbNewLine & _
                        "  (ID, 标本id, 图像类型, 图像点, 图像位置)" & vbNewLine & _
                        "  Select ID, 标本id, 图像类型, Null, 图像位置 From 检验报告图像_Exp_Temp"
        gcnOracle.Execute strSQL
    End If
    
    gcnOracle.CommitTrans
    
    '删除临时表和过程
    If intType = 1 Then
        strSQL = "Drop Procedure Zl_检验图像结果_Temp_Insert"
        gcnOracle.Execute strSQL
        strSQL = "Drop Table 检验图像结果_EXP_TEMP"
        gcnOracle.Execute strSQL
    Else
        strSQL = "Drop Procedure Zl_检验报告图像_Temp_Insert"
        gcnOracle.Execute strSQL
        strSQL = "Drop Table 检验报告图像_EXP_TEMP"
        gcnOracle.Execute strSQL
    End If
        
    Exit Function
errH:
    If InStr(1, UCase(Err.Description), "ORA") Then
        gcnOracle.RollbackTrans
    End If
    MsgBox Err.Description, vbExclamation, gstrSysName
End Function

Private Sub txtTmpPath_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Function GetCpuAdv() As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intAdvise As Integer
    
    On Error GoTo errH
    strSQL = "Select Nvl(Max(Value),0) CPU From V$parameter Where Name = 'cpu_count'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取可用CUP数")
    
    If rsTmp!cpu <= 4 Then
        intAdvise = 1
    ElseIf rsTmp!cpu <= 8 Then
        intAdvise = 4
    ElseIf rsTmp!cpu <= 12 Then
        intAdvise = 8
    Else
        intAdvise = 12
    End If
    
    GetCpuAdv = intAdvise
    Exit Function
errH:
    GetCpuAdv = 0
End Function

Private Function CheckLisSys() As Integer
    '功能:检查当前LIS系统的版本
    '返回 0=没有安装Lis 1=只有旧版LIS 2=只有新版LIS 3=两者均有
    Dim blnOld As Boolean, blnNew As Boolean
    
    blnOld = CheckTblExist("检验图像结果")
    blnNew = CheckTblExist("检验报告图像")
    
    If blnOld And blnNew Then
        CheckLisSys = 3    '均有
        lblBanner.Caption = ""
    ElseIf blnOld And Not blnNew Then
        lblBanner.Caption = "当前只安装了旧版LIS系统，无法选择数据来源。"
        optNew.Enabled = False
        CheckLisSys = 1    '只有旧版
    ElseIf Not blnOld And blnNew Then
        lblBanner.Caption = "当前只安装了新版LIS系统，无法选择数据来源。"
        optOld.Enabled = False
        CheckLisSys = 2    '只有新版
    ElseIf Not blnOld And Not blnNew Then
        lblBanner.Caption = "当前没有安装LIS系统，无法进行转出操作。"
        SetCmdEnable False
        CheckLisSys = 0    '没有LIS系统
    End If
End Function

Private Sub SetDtpPicker(ByVal intType As Integer)
    '功能:设置dtpPicker的值
    '参数:intType 1=老版LIS 2=新版LIS
    Dim strSQL As String, rsTmp As ADODB.Recordset
      
    On Error GoTo errH
    '将待转存图片的最大开始时间和结束时间存在时间控件中
    If intType = 1 Then
        strSQL = "Select a.核收时间" & vbNewLine & _
                        "From 检验标本记录 A, 检验图像结果 B" & vbNewLine & _
                        "Where a.Id = b.标本id And (b.Id = (Select Max(ID) As ID From 检验图像结果) Or b.Id = (Select Min(ID) As ID From 检验图像结果))" & vbNewLine & _
                        "Order By a.核收时间"
    ElseIf intType = 2 Then
        strSQL = "Select a.核收时间" & vbNewLine & _
                        "From 检验报告记录 A, 检验报告图像 B" & vbNewLine & _
                        "Where a.Id = b.标本id And (b.Id = (Select Max(ID) As ID From 检验报告图像) Or b.Id = (Select Min(ID) As ID From 检验报告图像))" & vbNewLine & _
                        "Order By a.核收时间"
    Else
        Exit Sub
    End If
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "1")
    If rsTmp.RecordCount = 0 Then Exit Sub
    dtpStart.Value = CDate(Trim(rsTmp!核收时间)) - 1
    rsTmp.MoveLast
    dtpEnd.Value = CDate(Trim(rsTmp!核收时间)) + 1
    
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub


Private Sub CreateTable(ByVal intType As Integer)
    '功能:根据传入的类型创建不同的临时表及过程
    '参数:  intType 1=旧版Lis数据 2=新版Lis数据
    Dim strSQL As String
    
    If intType = 1 Then
        '没有临时转出表,就创建
        If Not CheckTblExist("检验图像结果_EXP_TEMP") Then
            strSQL = "Create Table 检验图像结果_EXP_TEMP As Select id,标本id,图像类型,图像位置 From 检验图像结果 Where 1=0"
            gcnOracle.Execute strSQL
            strSQL = "Create Or Replace Procedure Zl_检验图像结果_Temp_Insert" & vbNewLine & _
                            "(" & vbNewLine & _
                            "  Id_In       In 检验图像结果_Exp_Temp.Id%Type," & vbNewLine & _
                            "  标本id_In   In 检验图像结果_Exp_Temp.标本id%Type," & vbNewLine & _
                            "  图像类型_In In 检验图像结果_Exp_Temp.图像类型%Type," & vbNewLine & _
                            "  图像位置_In In 检验图像结果_Exp_Temp.图像位置%Type" & vbNewLine & _
                            ") Is" & vbNewLine & _
                            "Begin" & vbNewLine & _
                            "  Insert Into 检验图像结果_Exp_Temp Values (Id_In, 标本id_In, 图像类型_In, 图像位置_In);" & vbNewLine & _
                            "Exception" & vbNewLine & _
                            "  When Others Then" & vbNewLine & _
                            "    zl_ErrorCenter(SQLCode, SQLErrM);" & vbNewLine & _
                            "End Zl_检验图像结果_Temp_Insert;"
            gcnOracle.Execute strSQL
        End If
    Else
        If Not CheckTblExist("检验报告图像_EXP_TEMP") Then
            strSQL = "Create Table 检验报告图像_EXP_TEMP As Select id,标本id,图像类型,图像位置 From 检验报告图像 Where 1=0"
            gcnOracle.Execute strSQL
            strSQL = "Create Or Replace Procedure Zl_检验报告图像_Temp_Insert" & vbNewLine & _
                            "(" & vbNewLine & _
                            "  Id_In       In 检验报告图像_Exp_Temp.Id%Type," & vbNewLine & _
                            "  标本id_In   In 检验报告图像_Exp_Temp.标本id%Type," & vbNewLine & _
                            "  图像类型_In In 检验报告图像_Exp_Temp.图像类型%Type," & vbNewLine & _
                            "  图像位置_In In 检验报告图像_Exp_Temp.图像位置%Type" & vbNewLine & _
                            ") Is" & vbNewLine & _
                            "Begin" & vbNewLine & _
                            "  Insert Into 检验报告图像_Exp_Temp Values (Id_In, 标本id_In, 图像类型_In, 图像位置_In);" & vbNewLine & _
                            "Exception" & vbNewLine & _
                            "  When Others Then" & vbNewLine & _
                            "    zl_ErrorCenter(SQLCode, SQLErrM);" & vbNewLine & _
                            "End Zl_检验报告图像_Temp_Insert;"
            gcnOracle.Execute strSQL
        End If
    End If
    
End Sub

Private Sub SetCmdEnable(ByVal blnEnable As Boolean)
    cmdFtp.Enabled = blnEnable: cmdFile.Enabled = blnEnable
    cmdMulti.Enabled = blnEnable: cmdCommit.Enabled = blnEnable
End Sub
