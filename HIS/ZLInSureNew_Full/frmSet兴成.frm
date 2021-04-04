VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSet兴成 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   3270
      Left            =   120
      TabIndex        =   4
      Top             =   915
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   5768
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "医院端前置机(&0)"
      TabPicture(0)   =   "frmSet兴成.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra医保服务器(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "医保端前置机(&1)"
      TabPicture(1)   =   "frmSet兴成.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdZXTest"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdODBC"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fra医保服务器(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "其他(&2)"
      TabPicture(2)   =   "frmSet兴成.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lbl(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lbl(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lbl(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lbl(3)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lbl(4)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lbl(5)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtPath(0)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtPath(1)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtPath(2)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtPath(3)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdSel(0)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdSel(1)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmdSel(3)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cmdSel(2)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "cmdSel(4)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtPath(4)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "chk读卡器"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "cbo医院级别"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Chk启用政策审核"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).ControlCount=   19
      Begin VB.CheckBox Chk启用政策审核 
         Caption         =   "启用政策审核(&Q)"
         Height          =   285
         Left            =   3795
         TabIndex        =   41
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.ComboBox cbo医院级别 
         Height          =   300
         Left            =   1320
         TabIndex        =   40
         Text            =   "Combo1"
         Top             =   2610
         Width           =   2340
      End
      Begin VB.CheckBox chk读卡器 
         Caption         =   "本站点存在读卡器(&R)"
         Height          =   285
         Left            =   3795
         TabIndex        =   38
         Top             =   2580
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.CommandButton cmdZXTest 
         Caption         =   "测试(&T)"
         Height          =   350
         Left            =   -70485
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1100
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   4
         Left            =   1305
         TabIndex        =   28
         Text            =   "C:\"
         Top             =   2235
         Width           =   4020
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   300
         Index           =   4
         Left            =   5370
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2250
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   300
         Index           =   2
         Left            =   5370
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1440
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   300
         Index           =   3
         Left            =   5370
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1860
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   300
         Index           =   1
         Left            =   5370
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1020
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   300
         Index           =   0
         Left            =   5370
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   285
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   3
         Left            =   1305
         TabIndex        =   22
         Text            =   "C:\Out"
         Top             =   1845
         Width           =   4020
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   2
         Left            =   1305
         TabIndex        =   20
         Text            =   "C:\IN"
         Top             =   1440
         Width           =   4020
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   1
         Left            =   1305
         TabIndex        =   18
         Text            =   "C:\xcyb\put"
         Top             =   1035
         Width           =   4020
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   0
         Left            =   1305
         TabIndex        =   16
         Text            =   "C:\xcyb\get"
         Top             =   600
         Width           =   4020
      End
      Begin VB.CommandButton cmdODBC 
         Caption         =   "数据源(&D)"
         Height          =   350
         Left            =   -70485
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1100
      End
      Begin VB.Frame fra医保服务器 
         Height          =   1875
         Index           =   0
         Left            =   -74820
         TabIndex        =   5
         Top             =   660
         Width           =   5595
         Begin VB.CommandButton cmdTest 
            Caption         =   "测试(&T)"
            Height          =   1095
            Left            =   4515
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   555
            Width           =   1005
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   8
            Top             =   1335
            Width           =   3075
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1200
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   945
            Width           =   3075
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   6
            Top             =   555
            Width           =   3075
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "服务器(&S)"
            Height          =   180
            Index           =   2
            Left            =   330
            TabIndex        =   12
            Top             =   1395
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "密码(&P)"
            Height          =   180
            Index           =   1
            Left            =   510
            TabIndex        =   11
            Top             =   1005
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "用户名(&U)"
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   10
            Top             =   615
            Width           =   810
         End
      End
      Begin VB.Frame fra医保服务器 
         Height          =   2070
         Index           =   1
         Left            =   -74850
         TabIndex        =   30
         Top             =   705
         Width           =   5595
         Begin VB.TextBox txtODBC 
            Height          =   300
            Index           =   0
            Left            =   1710
            MaxLength       =   40
            TabIndex        =   33
            Top             =   555
            Width           =   2565
         End
         Begin VB.TextBox txtODBC 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1710
            MaxLength       =   40
            TabIndex        =   32
            Top             =   930
            Width           =   2565
         End
         Begin VB.TextBox txtODBC 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1710
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   31
            Top             =   1275
            Width           =   2565
         End
         Begin VB.Label lblODBC 
            AutoSize        =   -1  'True
            Caption         =   "ODBC数据源名(&U)"
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   36
            Top             =   615
            Width           =   1350
         End
         Begin VB.Label lblODBC 
            AutoSize        =   -1  'True
            Caption         =   "ODBC数据源用户(&U)"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   35
            Top             =   975
            Width           =   1530
         End
         Begin VB.Label lblODBC 
            AutoSize        =   -1  'True
            Caption         =   "ODBC数据源密码(&P)"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   34
            Top             =   1380
            Width           =   1530
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "医院级别"
         Height          =   180
         Index           =   5
         Left            =   570
         TabIndex        =   39
         Top             =   2670
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "操作系统目录"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   29
         Top             =   2280
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "接口出参目录"
         Height          =   180
         Index           =   3
         Left            =   225
         TabIndex        =   21
         Top             =   1890
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "接口入参目录"
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   19
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label lbl 
         Caption         =   "上传临时目录"
         Height          =   240
         Index           =   1
         Left            =   165
         TabIndex        =   17
         Top             =   1110
         Width           =   1140
      End
      Begin VB.Label lbl 
         Caption         =   "下载临时目录"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   15
         Top             =   675
         Width           =   1140
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   3
      Top             =   4290
      Width           =   7665
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   660
      Width           =   7665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4065
      TabIndex        =   0
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5250
      TabIndex        =   1
      Top             =   4440
      Width           =   1100
   End
   Begin VB.Label lblNote 
      Caption         =   "    设置到医疗保险数据服务器的连接串；为保证设置有效，这时医疗保险数据服务器必须可用。"
      Height          =   390
      Left            =   810
      TabIndex        =   14
      Top             =   240
      Width           =   5475
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmSet兴成.frx":0054
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmSet兴成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
End Enum
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private m医院级别 As String

Public Function 参数设置() As Boolean
    mblnChange = False
    Dim rsTemp As New ADODB.Recordset
    frmSet兴成.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdODBC_Click()
    On Error Resume Next
    Shell "ODBCAD32", vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "不能进入ODBC数据源管理器，请检查系统是否正确安装！", vbInformation, gstrSysName
    End If
    Err.Clear
End Sub

Private Sub cmdSel_Click(Index As Integer)
    Dim strPath As String
    strPath = OpenDire(Me, "请指定目录：")
    If strPath = "" Then Exit Sub
    txtPath(Index).Text = strPath
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub

Private Sub cmdZXTest_Click()
   Dim cnInsure As New ADODB.Connection
    Err = 0
    On Error Resume Next
    With cnInsure
            If .State = adStateOpen Then .Close
            .ConnectionString = "dsn=" & txtODBC(0).Text & ";uid=" & txtODBC(1).Text & ";pwd=" & txtODBC(2).Text & ""
            .Open
            If Err <> 0 Then
                MsgBox "测试不成功，请检查医保数据服务器是否可用，以及数据源是否正确配置！", vbExclamation, gstrSysName
                Exit Sub
            End If
            .Close
            MsgBox "测试成功，与医保数据服务器正常连接！", vbInformation, gstrSysName
    End With
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    If mblnFirst = False Then Exit Sub
    
    mblnFirst = False
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_兴成核工业
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    m医院级别 = "01"
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!参数名)
            Case "医保用户名"
                  txtEdit(text医保用户).Text = Nvl(!参数值)
            Case "医保用户密码"
                  txtEdit(Text医保密码).Text = Nvl(!参数值)
            Case "医保服务器"
                  txtEdit(Text医保服务器).Text = Nvl(!参数值)
            Case "医院级别"
                  m医院级别 = Nvl(!参数值, "01")
            End Select
            .MoveNext
        Loop
    End With
    txtODBC(0).Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_NAME"), "")
    txtODBC(1).Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_USERNAME"), "")
    txtODBC(2).Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_PASSWORD"), "")
    
    txtPath(0).Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Get"), "C:\xcyb\get")
    txtPath(1).Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Put"), "C:\xcyb\Put")
    txtPath(2).Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_In"), "C:\xcyb\In")
    txtPath(3).Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Out"), "C:\xcyb\Out")
    txtPath(4).Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("strPath_System"), "C:\")
    
    chk读卡器.Value = IIf(Val(GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("读卡器"), "1")) = 1, 1, 0)
    
    '陈宏悦于20050408增加
    
    Chk启用政策审核.Value = IIf(Val(GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("启用政策审核"), "1")) = 1, 1, 0)
    
    Call Load医院级别
 
 
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text医保密码 Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    If Index = Text医保服务器 Or Index = Text医保密码 Or Index = text医保用户 Then
        '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag, False) = False Then
            If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_兴成核工业 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_兴成核工业 & ",null,'医保用户名','" & txtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_兴成核工业 & ",null,'医保用户密码','" & txtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_兴成核工业 & ",null,'医保服务器','" & txtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_兴成核工业 & ",null,'医院级别','" & Split(cbo医院级别.Text, " ")(0) & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Get"), Trim(txtPath(0).Text)
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Put"), Trim(txtPath(1).Text)
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("strPath_In"), Trim(txtPath(2).Text)
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("strPath_Out"), Trim(txtPath(3).Text)
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("strPath_System"), Trim(txtPath(4).Text)
    
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_NAME"), Trim(txtODBC(0).Text)
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_USERNAME"), Trim(txtODBC(1).Text)
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("ODBC_PASSWORD"), Trim(txtODBC(2).Text)
    
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("读卡器"), IIf(chk读卡器.Value = 1, 1, 0)
    
    '陈宏悦于20050408日增加修改,由于住院记帐时调用医保部件,但是该部件在动态库初始化了一个特殊端口
    
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("启用政策审核"), IIf(Chk启用政策审核.Value = 1, 1, 0)

    
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Function OpenDire(odtvOwner As Form, Optional odtvTitle As String) As String
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = odtvTitle
   With tBrowseInfo
      .hwndOwner = odtvOwner.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      OpenDire = sBuffer
   End If
End Function

Private Function CreatePath(ByVal strPath As String) As Boolean
    '功能:创建文件路径
    Dim objPath As New FileSystemObject
    Dim strArr As Variant
    Dim strTmpPath As String
    CreatePath = False
    Dim i As Long
    strTmpPath = strPath
    If InStr(strTmpPath, "\") = 0 Or InStr(strTmpPath, "\\") <> 0 Then
         MsgBox "路径不正确!", vbInformation + vbDefaultButton1, gstrSysName
         Exit Function
    End If
    strArr = Split(strTmpPath, "\")
    If InStr(1, "A:B:C:D:E:F:G:H:I:J:K:L:M:N:O:P:Q:R:S:T:U:V:W:X:Y:Z:", strArr(0)) = 0 Then
         MsgBox "路径不正确!", vbInformation + vbDefaultButton1, gstrSysName
         Exit Function
    End If
    
    strTmpPath = strArr(0)
    For i = 1 To UBound(strArr)
        Err = 0
        On Error Resume Next
        strTmpPath = strTmpPath & "\" & strArr(i)
        
        If objPath.FolderExists(strTmpPath) = False Then
            objPath.CreateFolder strTmpPath
            If Err <> 0 Then
                MsgBox "创建路径失败(" & strTmpPath & ")" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & " 错误描述:" & Err.Description, vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
        End If
    Next
    CreatePath = True
End Function


Private Sub txtODBC_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtPath_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtPath_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr(1, ";-`@#$%^&**()#@+-|", Chr(KeyAscii)) <> 0 Then
        KeyAscii = 0
    End If
End Sub
Private Function Load医院级别()
    Dim i As Integer
    cbo医院级别.Clear
    
    With cbo医院级别
        .AddItem "01 一级医院"
        .AddItem "02 二级医院"
        .AddItem "03 三级医院"
        .AddItem "11 转外地市内一级医院"
        .AddItem "12 转外地省内一级医院"
        .AddItem "13 转外地省外一级医院"
        .AddItem "14 转外地市内二级医院"
        .AddItem "15 转外地省内二级医院"
        .AddItem "16 转外地省外二级医院"
        .AddItem "17 转外地市内三级医院"
        .AddItem "18 转外地省内三级医院"
        .AddItem "19 转外地省外三级医院"
        
        For i = 0 To .ListCount - 1
            If Split(.List(i), " ")(0) = m医院级别 Then
                .ListIndex = i
                Exit For
            End If
        Next
        If .ListIndex < 0 Then
            .ListIndex = 0
        End If
    End With
End Function

