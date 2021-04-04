VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAppReplant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "应用系统再植"
   ClientHeight    =   4416
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   7140
   Icon            =   "frmAppReplant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4416
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSetup 
      Height          =   3675
      Index           =   0
      Left            =   1305
      TabIndex        =   4
      Top             =   -120
      Width           =   6075
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   465
         Width           =   5800
      End
      Begin VB.Frame fraSys 
         Height          =   1005
         Left            =   765
         TabIndex        =   26
         Top             =   2340
         Width           =   3930
         Begin VB.Label lblVersion 
            AutoSize        =   -1  'True
            Caption         =   "版本号："
            Height          =   180
            Left            =   210
            TabIndex        =   28
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lblSysName 
            AutoSize        =   -1  'True
            Caption         =   "系统名："
            Height          =   180
            Left            =   210
            TabIndex        =   27
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdSetupFile 
         Caption         =   "选择(&S)…"
         Height          =   350
         Left            =   765
         TabIndex        =   5
         Top             =   1980
         Width           =   1260
      End
      Begin VB.Label lblSetupFile 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   765
         TabIndex        =   18
         Top             =   1650
         Width           =   3930
      End
      Begin VB.Label lbliniFile 
         AutoSize        =   -1  'True
         Caption         =   "应用安装配置文件"
         Height          =   180
         Left            =   765
         TabIndex        =   17
         Top             =   1410
         Width           =   1440
      End
      Begin VB.Label lblNote 
         Caption         =   "    应用系统的再植依赖于配置文件和与之相关的服务器创建脚本文件，请正确指定安装配置文件。"
         Height          =   450
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   720
         Width           =   5250
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第一步 指定安装配置文件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.4
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Width           =   2595
      End
   End
   Begin VB.Frame fraSetup 
      Height          =   3675
      Index           =   1
      Left            =   1305
      TabIndex        =   9
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.Frame fraOwner 
         Caption         =   "植入系统所有者"
         Height          =   2010
         Left            =   600
         TabIndex        =   19
         Top             =   1320
         Width           =   4530
         Begin VB.ComboBox cboOwnerUsr 
            Height          =   300
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   360
            Width           =   2160
         End
         Begin VB.CheckBox chkDBA 
            Caption         =   "授予DBA角色"
            Height          =   255
            Left            =   405
            TabIndex        =   30
            Top             =   1545
            Width           =   1320
         End
         Begin VB.TextBox txtOwnerPwd 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   825
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   20
            Top             =   750
            Width           =   2160
         End
         Begin VB.Label lblDBA 
            AutoSize        =   -1  'True
            Caption         =   "可以根据管理习惯决定是否授予DBA角色。"
            Height          =   180
            Left            =   405
            TabIndex        =   29
            Top             =   1290
            Width           =   3330
         End
         Begin VB.Label lblNewUser 
            AutoSize        =   -1  'True
            Caption         =   "用户"
            Height          =   180
            Left            =   405
            TabIndex        =   22
            Top             =   420
            Width           =   360
         End
         Begin VB.Label lblNewPwd 
            AutoSize        =   -1  'True
            Caption         =   "口令"
            Height          =   180
            Left            =   405
            TabIndex        =   21
            Top             =   810
            Width           =   360
         End
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第二步 指定植入系统所有者"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.4
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   225
         Width           =   2820
      End
      Begin VB.Label lblNote 
         Caption         =   "    植入系统必然有一个数据库用户作为所有者，同时你必须知道植入系统所有者的密码，以便检查植入系统的正确性。"
         Height          =   585
         Index           =   1
         Left            =   225
         TabIndex        =   11
         Top             =   720
         Width           =   5250
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   1545
      TabIndex        =   25
      Top             =   3645
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pgbState 
      Height          =   150
      Left            =   3180
      TabIndex        =   24
      Top             =   4185
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5652
      _ExtentY        =   275
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox PicSetup 
      Align           =   3  'Align Left
      Height          =   4044
      Left            =   0
      ScaleHeight     =   3996
      ScaleWidth      =   1284
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      Begin VB.Image imgSetup 
         Height          =   3315
         Left            =   60
         Picture         =   "frmAppReplant.frx":058A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一步(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4695
      TabIndex        =   3
      Top             =   3645
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3585
      TabIndex        =   1
      Top             =   3645
      Width           =   1100
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步(&N)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5790
      TabIndex        =   0
      Top             =   3645
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   23
      Top             =   4044
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "frmAppReplant.frx":5B70
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8975
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1185
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "11:37"
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
   Begin VB.Frame fraSetup 
      Height          =   3675
      Index           =   2
      Left            =   1305
      TabIndex        =   13
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblRegAudit 
         AutoSize        =   -1  'True
         Caption         =   "    由于还不具备该系统应用授权，虽然可以继续装载，但无法正常使用。"
         Height          =   360
         Left            =   225
         TabIndex        =   33
         Top             =   1335
         Width           =   5580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNextDo 
         AutoSize        =   -1  'True
         Caption         =   "    点击""完成""开始自动装载系统，或者""取消""终止系统装载，或""上一步""重新调整应用系统装载配置。"
         Height          =   360
         Left            =   225
         TabIndex        =   32
         Top             =   2025
         Width           =   5580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "第三步 完成"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.4
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   165
         TabIndex        =   16
         Top             =   225
         Width           =   1245
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "    已经完成了对该系统再植的全部设置。"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   15
         Top             =   720
         Width           =   3420
      End
   End
End
Attribute VB_Name = "frmAppReplant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strIniPath      As String                 '安装配置文件目录
Dim intDefSysCode   As String                 '系统编号
Dim strDefSysName   As String                 '系统名称
Dim strDefVersion   As String                 '版本号
Dim strDefSpace   As String                   '表空间定义串
Dim strDefUser      As String                 '新的缺省用户名
Dim strDefData      As String                 '用户可选的数据

Dim mstrExtSysCode  As String                  '要进行扩展的主系统的编号
Dim mstrExtVersion  As String                  '要进行扩展的主系统的版本

Dim strTbsPath As String                        '缺省表空间路径名称，根据历史表空间产生

Dim objText As TextStream

Dim mbln帐套 As Boolean    '本次安装是否是属于帐套安装
Dim mlng帐套 As Long       '帐套号
Dim mlst标准 As ListItem   '相对于要安装的帐套，这是提供标准管理数据的系统

Dim intStep As Integer

Dim mcnOwner As New ADODB.Connection
Dim mlngEnjoy As Long
Dim strSQL As String, strTemp As String
Dim intCount As Integer, intItems As Integer
        
Dim aryRow() As String
Dim aryVal() As String

Private Sub cboOwnerUsr_Click()
    Dim rsTemp As New ADODB.Recordset
    
On Error GoTo errHandle
    txtOwnerPwd.Text = ""
    
    If mstrExtSysCode = "" Then
        '非扩展系统
        With rsTemp
            If .State = adStateOpen Then .Close
            strSQL = "select 编号,名称" & _
                    " from zlSystems" & _
                    " where 所有者='" & cboOwnerUsr.Text & "'" & _
                    " start with 共享号 is null" & _
                    " connect by prior 编号=共享号" & _
                    " order by level"
            .Open strSQL, gcnOracle, adOpenKeyset
            mlngEnjoy = 0
            If .EOF Or .BOF Then Exit Sub
            .MoveLast
            mlngEnjoy = .Fields(0).value
            strSQL = "该所有者是" & .Fields(1).value & "的所有者，" & _
                vbCr & "选择该用户表示两个系统是共享安装。"
            MsgBox strSQL, vbExclamation, gstrSysName
        End With
    Else
        '扩展系统，只能使用共享方式
        mlngEnjoy = cboOwnerUsr.ItemData(cboOwnerUsr.ListIndex)
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("安装未完成，真的取消吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub cmdSetupFile_Click()
    With frmMDIMain.dlgMain
        .FileName = ""
        .DialogTitle = "选择应用安装配置文件"
        .Filter = "(应用安装配置文件)|zlSetup.ini"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblSetupFile.Caption = .FileName
        End If
    End With
    If ChkSetupFile(True) = False Then
        lblSetupFile.Caption = ""
        cmdSetupFile.SetFocus
    End If

End Sub

Private Sub cmdNext_Click()
    Dim strErr As String
    
    If fraSetup(0).Visible Then
        '------------------------------------------------------------
        '第一步：
        '------------------------------------------------------------
        If Trim(lblSetupFile.Caption) = "" Then
            MsgBox "未正确选择服务器安装配置文件，不能继续", vbExclamation, gstrSysName
            cmdSetupFile.SetFocus
            Exit Sub
        End If
        
        '------------------------------
        fraSetup(0).Visible = False
        fraSetup(1).Visible = True
        cmdPrevious.Enabled = True
    
    ElseIf fraSetup(1).Visible Then
        '------------------------------------------------------------
        '第二步：
        '------------------------------------------------------------
        Set mcnOwner = gobjRegister.GetConnection(gstrServer, Trim(cboOwnerUsr.Text), Trim(txtOwnerPwd.Text), False, MSODBC, "", False)
        If mcnOwner.State = adStateClosed Then
            Set mcnOwner = gobjRegister.GetConnection(gstrServer, Trim(cboOwnerUsr.Text), Trim(txtOwnerPwd.Text), True, MSODBC, "", False)
            If mcnOwner.State = adStateClosed Then
                MsgBox "所有者密码错误，不能继续" & vbNewLine & strErr, vbExclamation, gstrSysName
                txtOwnerPwd.SetFocus
                Exit Sub
            End If
        End If
        Call SetSQLTrace(gstrServer, Trim(cboOwnerUsr.Text), mcnOwner)
        
        On Error Resume Next
        Dim strErrInfo As String
        MousePointer = 11
        cmdSetupFile.Enabled = False
        cmdCancel.Enabled = False
        cmdNext.Enabled = False
        stbThis.Panels(2).Text = "检查植入系统"
        strErrInfo = CheckTable(strIniPath & "zlTable.sql")
        MousePointer = 0
        cmdSetupFile.Enabled = True
        cmdCancel.Enabled = True
        cmdNext.Enabled = True
        stbThis.Panels(2).Text = ""
                
        If strErrInfo <> "" Then
            If InStr(strErrInfo, "是否继续？") > 0 Then
                If MsgBox(strErrInfo, vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox strErrInfo, vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
        '------------------------------
        fraSetup(1).Visible = False
        fraSetup(2).Visible = True
        cmdNext.Caption = "完成(&F)"
    ElseIf fraSetup(2).Visible Then
        '------------------------------------------------------------
        '第三步：
        '------------------------------------------------------------
        If mlngEnjoy = 0 Then
            Set gcnTools = GetConnection("ZLTOOLS")
            If gcnTools Is Nothing Then Exit Sub
        End If
        
        strSQL = "    已经完成了所有的再植设置，系统将进入自动再植过程。" & vbCr & vbCr _
                & "    再植过程可能运行较长时间，请不要随意强行中断；否则，" & vbCr _
                & "将可能产生数据垃圾，影响系统运行。" & vbCr & vbCr _
                & "   继续再植吗？"
        If MsgBox(strSQL, vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
        cmdCancel.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = False
        fraSetup(2).Enabled = False
        
        If SysInstall() Then
            MsgBox "再植成功，可以在完成应用程序再植后正常使用该系统。", vbInformation, gstrSysName
            
            
            
        Else
            gcnOracle.Execute "delete zlSystems where 编号=" & intDefSysCode * 100 + mlng帐套
            MsgBox "再植失败，请检查安装文件的正确性。", vbInformation, gstrSysName
        End If
        cmdNext.Enabled = True
        Unload Me
    End If

End Sub

Private Sub cmdPrevious_Click()
    If fraSetup(2).Visible Then
        fraSetup(2).Visible = False
        fraSetup(1).Visible = True
        cmdNext.Caption = "下一步(&N)"
    ElseIf fraSetup(1).Visible Then
        fraSetup(1).Visible = False
        fraSetup(0).Visible = True
        cmdPrevious.Enabled = False
    End If

End Sub

Private Sub Form_Load()
    Call ApplyOEM(stbThis)
    Dim objItem As ListItem
    With imgSetup
        .Top = PicSetup.ScaleTop
        .Left = PicSetup.ScaleLeft
        .Height = PicSetup.ScaleHeight
        .Width = PicSetup.ScaleWidth
    End With
    pgbState.Left = stbThis.Panels(2).Left + TextWidth("正在创建数据表")
    pgbState.Width = stbThis.Panels(3).Left - pgbState.Left - 100
    pgbState.Top = stbThis.Top + stbThis.Height / 3
    
    '如果发现当前目录存在安装培植文件，则直接填写
    If Dir(App.Path & "\zlSetup.ini") <> "" Then
        lblSetupFile.Caption = App.Path & "\zlSetup.ini"
        If ChkSetupFile() = False Then
            lblSetupFile.Caption = ""
        End If
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdNext.Enabled = False Then
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Function ChkSetupFile(Optional blnMsg As Boolean) As Boolean
    
    '-------------------------------------
    '检查解释安装配置文件的正确性
    '-------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim varVersion As Variant, varExtVersin As Variant
    Dim i As Long
    
    strIniPath = Mid(lblSetupFile.Caption, 1, Len(lblSetupFile.Caption) - 11)
    '相关文件匹配性检查
    strTemp = ""
    If Dir(strIniPath & "zlSequence.sql") = "" Then strTemp = strTemp & vbCr & "序列文件" & strIniPath & "zlSequence.sql"
    If Dir(strIniPath & "zlTable.sql") = "" Then strTemp = strTemp & vbCr & "数据表文件" & strIniPath & "zlTable.sql"
    If Dir(strIniPath & "zlConstraint.sql") = "" Then strTemp = strTemp & vbCr & "约束文件" & strIniPath & "zlConstraint.sql"
    If Dir(strIniPath & "zlIndex.sql") = "" Then strTemp = strTemp & vbCr & "索引文件" & strIniPath & "zlIndex.sql"
    If Dir(strIniPath & "zlView.sql") = "" Then strTemp = strTemp & vbCr & "视图文件" & strIniPath & "zlView.sql"
    If Dir(strIniPath & "zlProgram.sql") = "" Then strTemp = strTemp & vbCr & "函数过程文件" & strIniPath & "zlProgram.sql"
    If Dir(strIniPath & "zlManData.sql") = "" Then strTemp = strTemp & vbCr & "管理数据文件" & strIniPath & "zlManData.sql"
    If Dir(strIniPath & "zlAppData.sql") = "" Then strTemp = strTemp & vbCr & "应用数据文件" & strIniPath & "zlAppData.sql"
    If strTemp <> "" Then
        If blnMsg Then MsgBox "以下服务器安装的相关文件丢失，不能继续，包括：" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '安装配置文件解释
    err = 0
    On Error Resume Next
    Set objText = gobjFile.OpenTextFile(lblSetupFile.Caption)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[系统号]" Then
        intDefSysCode = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[系统名]" Then
        strDefSysName = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[版本号]" Then
        strDefVersion = Trim(Mid(strTemp, 6))
    
        '判断是否应该把本次安装作为帐套安装
        Dim lngTemp As Long
        Dim lngMax As Long        '最大的帐套号
        Dim blnHase  As Boolean   '是否有同系统存在
        Dim lstTemp As ListItem

        
        mbln帐套 = False
        mlng帐套 = 0
        For Each lstTemp In frmAppStart.lvwSys.ListItems
            lngTemp = Mid(lstTemp.Key, 2)
            If lngTemp \ 100 = intDefSysCode Then
                '系统相同
                blnHase = True
                If lngMax < lngTemp Mod 100 Then
                    lngMax = lngTemp Mod 100 '保存最大的帐套号
                End If
                
                If strDefVersion = lstTemp.SubItems(1) Then
                    '版本也相同，那就可以了
                    mbln帐套 = True
                    Set mlst标准 = lstTemp
                End If
            End If
        Next
        If blnHase = True Then
            '有同系统的安装
            If mbln帐套 = False Then
                If blnMsg Then MsgBox "当前数据库中也有相同类型的系统存在，但由于版本不符，不能再植。", vbInformation, gstrSysName
                Exit Function
            Else
                If blnMsg = False Then
                    Exit Function
                Else
                    If lngMax >= 99 Then
                        MsgBox "当前数据库中也有相同类型的系统存在，且数量足够多，不能再植。", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If MsgBox("当前数据库中已有" & strDefSysName & "系统存在，你是否要再再植一个新的？", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                    mlng帐套 = lngMax + 1
                End If
            End If
        End If
    Else
        err.Raise 10
    End If
    Caption = "应用系统安装" & " - " & strDefSysName & " V" & strDefVersion
    lblSysName.Caption = "系统名：" & strDefSysName
    lblVersion.Caption = "版本号：" & strDefVersion
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[表空间]" Then
        strDefSpace = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[用户名]" Then
        strDefUser = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[数据组]" Then
        strDefData = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    mstrExtSysCode = ""
    mstrExtVersion = ""
    If Not objText.AtEndOfStream Then
        '还有扩展系统的设置
        strTemp = Trim(objText.ReadLine)
        If Left(strTemp, 5) = "[主系统]" Then
            mstrExtSysCode = Trim(Mid(strTemp, 6))
            
            strTemp = Trim(objText.ReadLine)
            If Left(strTemp, 5) = "[主版本]" Then
                mstrExtVersion = Trim(Mid(strTemp, 6))
            Else
                mstrExtSysCode = ""
            End If
        End If
    End If
    
    If err <> 0 Then
        If blnMsg Then MsgBox "安装配置文件丢失或不正确", vbExclamation, gstrSysName
        Exit Function
    End If
    objText.Close
    
    '查找用户清单
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    cboOwnerUsr.Clear
    If mstrExtSysCode = "" Then
        '非扩展系统
        With rsTemp
            gstrSQL = "select username " & _
                    " from dba_users U" & _
                    " where username not in ('SYS','SYSTEM','ZLTOOLS')" & _
                    "       and exists (select 1 from dba_Tables T where T.owner=U.username)" & _
                    "       and username not in (select 所有者 from zlsystems where FLOOR(编号/100)=" & intDefSysCode & ")"
            .Open gstrSQL, gcnOracle, adOpenKeyset
            Do While Not .EOF
                cboOwnerUsr.AddItem .Fields(0).value
                .MoveNext
            Loop
            If cboOwnerUsr.ListCount > 1 Then cboOwnerUsr.ListIndex = 0
        End With
    Else
        '是扩展系统，那必须要完成三方面的判断
        '1)系统号相符
        '2)没被其它的相同系统扩展
        '3)版本不能低于要求
        gstrSQL = "select A.所有者,A.版本号,A.编号 from zlsystems A " & _
                  "  Where floor(A.编号 / 100) = " & mstrExtSysCode & _
                  "        and not exists (select B.编号 from zlsystems B where B.共享号=A.编号 and floor(B.编号/100)=" & intDefSysCode & ")"
        
        If Not rsTemp Is Nothing Then
            If rsTemp.State = 1 Then rsTemp.Close
        End If
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        varExtVersin = Split(mstrExtVersion, ".")
        Do Until rsTemp.EOF
            '判断版本
            varVersion = Split(rsTemp("版本号"), ".")
            For i = LBound(varExtVersin) To UBound(varExtVersin)
                If varExtVersin(i) > varVersion(i) Then
                    '从数据库中读出的版本更低
                    Exit For
                End If
            Next
            If i > UBound(varExtVersin) Then
                '符合条件
                cboOwnerUsr.AddItem rsTemp("所有者")
                cboOwnerUsr.ItemData(cboOwnerUsr.NewIndex) = rsTemp("编号")
            End If
            rsTemp.MoveNext
        Loop
        
    End If
    
    For intCount = 0 To cboOwnerUsr.ListCount - 1
        If cboOwnerUsr.List(intCount) = UCase(strDefUser) Then
            cboOwnerUsr.ListIndex = intCount
            Exit For
        End If
    Next
    If cboOwnerUsr.ListCount = 0 Then
        If blnMsg Then MsgBox "没有合适的可再植的用户。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If cboOwnerUsr.ListIndex < 0 Then cboOwnerUsr.ListIndex = 0
    
    
    '顺便把注册文件也一并检查了
    Call ChkRegFile
    
    ChkSetupFile = True
End Function

Private Sub ChkRegFile()
    '判断系统授权
    Dim rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo errHand
    gstrSQL = "Select Count(*) From Zlregfunc f, Zlreginfo r, zlRegCheck t Where r.项目 = '授权证章' And f.系统 = " & intDefSysCode
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.Fields(0).value > 0 Then
        Me.lblRegAudit.Caption = "    已经具备该系统应用授权，可以在装载后正常授权使用。"
        Exit Sub
    End If
errHand:
    Me.lblRegAudit.Caption = "    由于还不具备该系统应用授权，虽然可以继续装载，但无法正常授权使用！"
End Sub

Private Function CheckTable(FileName As String) As String
    '--------------------------------------------
    '功能：检查数据表，同时判断数据表的列是否正确
    '--------------------------------------------
    Dim arySql() As String, strObjName As String, strTables As String
    Dim rsTemp As New ADODB.Recordset
    
    CheckTable = ""
    pgbState.value = 0
    pgbState.Visible = True
    With rsTemp
        .Filter = 0
        If gblnDBA Then
            strSQL = "select TABLE_NAME from DBA_TABLES where OWNER='" & cboOwnerUsr.Text & "'"
        Else
            strSQL = "select TABLE_NAME from USER_TABLES"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        
        err = 0
        On Error Resume Next
        Set objText = gobjFile.OpenTextFile(FileName)
        If err.Number <> 0 Then
            CheckTable = "由于安装脚本不正确，无法验证植入系统的正确性。不能继续！"
            .Filter = 0
            Exit Function
        End If
        intCount = GetFileLineCount(objText)
        objText.Close
        Set objText = gobjFile.OpenTextFile(FileName)
        
        strTables = ""
        On Error GoTo 0
        strSQL = ""
        Do While Not objText.AtEndOfStream
            strTemp = Trim(objText.ReadLine)
            If Left(strTemp, 2) <> "--" Then
                If Right(strTemp, 1) = ";" Then
                    strSQL = strSQL & vbCrLf & Left(strTemp, Len(strTemp) - 1)
                Else
                    strSQL = strSQL & vbCrLf & strTemp
                End If
                If Left(strSQL, 2) = vbCrLf Then
                    If Len(strSQL) = 2 Then
                        strSQL = ""
                    Else
                        strSQL = Mid(strSQL, 3)
                    End If
                End If
            End If
            If (Right(strTemp, 1) = ";" Or objText.AtEndOfStream) And Len(strSQL) <> 0 Then
                strSQL = UCase(Replace(Replace(Trim(strSQL), vbTab, " "), vbCrLf, " "))
                arySql = Split(strSQL, " TABLE ")
                If InStr(1, arySql(1), " ") > 0 And InStr(1, arySql(1), " ") < InStr(1, arySql(1), "(") Then
                    strObjName = Trim(Left(arySql(1), InStr(1, arySql(1), " ")))
                Else
                    strObjName = Trim(Left(arySql(1), InStr(1, arySql(1), "(") - 1))
                End If
                .Filter = "TABLE_NAME='" & strObjName & "'"
                If .EOF Then
                    strTables = strTables & vbCr & "    " & strObjName
                    If UBound(Split(strTables, vbCr)) > 16 Then Exit Do
                End If
                strSQL = ""
            End If
            
            pgbState.value = objText.Line / intCount * 100
        Loop
        .Filter = 0
    End With
    pgbState.value = 0
    pgbState.Visible = False
    If strTables <> "" Then
        CheckTable = "    由于该用户不具有该系统要求的下列数据表，" & _
              vbCr & "不能判断是正确的系统所有者，是否继续？" & _
              vbCr & "    缺少的数据表包括：" & strTables
    End If
End Function

Private Function SysInstall() As Boolean
    '----------------------------------
    '功能：完成系统的安装处理
    '---------安装算法-----------------
    '    创建本系统数据表空间
    '    If not 共享已经安装的系统 Then
    '        创建本系统所有者
    '        由工具所有者授予必要的工具数据对象权限
    '    End If
    '    创建本系统数据对象
    '    必须数据及可选数据安装
    '----------------------------------
    Dim strTmpSpace As String
    Dim rsTemp As New ADODB.Recordset, cnCtxsys As New ADODB.Connection
    
    
    err = 0
    On Error GoTo errHand
    gcnOracle.Execute "Grant Select on sys.v_$session to Public"
    gcnOracle.Execute "Grant Select on sys.v_$parameter to Public"
        
    With rsTemp
        If .State = adStateOpen Then .Close
        strSQL = "SELECT TEMPORARY_TABLESPACE FROM DBA_USERS WHERE USERNAME='ZLTOOLS'"
        .Open strSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then SysInstall = False: Exit Function
        strTmpSpace = .Fields(0).value
    End With
    
    'SYS向新系统授权
    gstrSQL = "Grant Connect,Resource," & IIf(chkDBA.value = 1, "DBA,", "") & _
            " Create Table,UNLIMITED TABLESPACE,Create Role,Create User,Drop User,Create Public Synonym,Drop Public Synonym" & _
            " to " & cboOwnerUsr.Text & " With Admin Option"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Grant Select on sys.dba_role_privs to " & cboOwnerUsr.Text & " With Grant Option"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Grant Select on sys.dba_roles to " & cboOwnerUsr.Text
    gcnOracle.Execute gstrSQL
    gstrSQL = "Grant Execute on sys.dbms_sql to " & cboOwnerUsr.Text & " With Grant Option"
    gcnOracle.Execute gstrSQL
    ' 2007-8-07 仅向所有者授权
    gstrSQL = "Grant Select on sys.gv_$session to " & cboOwnerUsr.Text & " With Grant Option"
    gcnOracle.Execute gstrSQL

    
    On Error Resume Next '创建全文检索的参数，有可能没有该用户，所以把错误屏蔽
    gstrSQL = "Grant CTXAPP to " & cboOwnerUsr.Text & " With Admin Option"
    gcnOracle.Execute gstrSQL
    gcnOracle.Execute "alter user ctxsys identified by ctxsys"
    cnCtxsys.Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrServer, "ctxsys", "ctxsys"
    cnCtxsys.Execute "Grant Execute on ctx_ddl to " & cboOwnerUsr.Text & " With Grant Option" '为了在过程中执行包函数
    
    On Error GoTo errHand
    '如果不共享已经安装的系统
    If mlngEnjoy = 0 Then
        '创建本系统所有者
        stbThis.Panels(2).Text = "向植入所有者" & cboOwnerUsr.Text & "授权"
        '由工具所有者授予必要的工具数据对象权限
        With rsTemp
            If .State = adStateOpen Then .Close
            strSQL = "select OBJECT_NAME,OBJECT_TYPE from user_objects where OBJECT_TYPE in('FUNCTION','PROCEDURE','SEQUENCE','TABLE','VIEW')  And Instr(OBJECT_NAME,'BIN$')<=0"
            .Open strSQL, gcnTools, adOpenKeyset
            Do While Not .EOF
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                Select Case !Object_Type
                Case "FUNCTION", "PROCEDURE"
                    gcnTools.Execute "grant execute on " & !Object_Name & " to " & cboOwnerUsr.Text & " With GRANT Option"
                Case "VIEW"
                    gcnTools.Execute "grant select on " & !Object_Name & " to " & cboOwnerUsr.Text & " With GRANT Option"
                Case "SEQUENCE"
                    gcnTools.Execute "grant select,alter on " & !Object_Name & " to " & cboOwnerUsr.Text & " With GRANT Option"
                Case "TABLE"
                    gcnTools.Execute "grant select,insert,update,delete on " & !Object_Name & " to " & cboOwnerUsr.Text & " With GRANT Option"
                End Select
                DoEvents
                .MoveNext
            Loop
        End With
        pgbState.value = 0
        pgbState.Visible = False
    End If
    
    '填写安装系统清单
    strSQL = "insert into zlSystems(编号,共享号,名称,所有者,安装日期,正常安装,版本号)" & _
            " values(" & intDefSysCode * 100 + mlng帐套
    If mlngEnjoy <> 0 Then
        strSQL = strSQL & "," & mlngEnjoy
    Else
        strSQL = strSQL & ",null"
    End If
    strSQL = strSQL & ",'" & strDefSysName & "'"
    strSQL = strSQL & ",'" & Trim(cboOwnerUsr.Text) & "'"
    strSQL = strSQL & ",sysdate,0,'" & strDefVersion & "')"
    gcnOracle.Execute strSQL
    
    '创建本系统数据对象(第二步已创建并打开mcnOwner)
    
    '必须数据
    stbThis.Panels(2).Text = "管理数据安装"
    If mbln帐套 = False Then
        Call RunSetupFile(mcnOwner, strIniPath & "zlManData.sql", ";", True)
    Else
        '通过数据库中拷贝得到
        If CopyManageData(mcnOwner) = False Then GoTo errHand
    End If
    
    '安装报表
    stbThis.Panels(2).Text = "固定报表安装"
    
    If mbln帐套 = False Then
        If RunSetupFile(mcnOwner, strIniPath & "zlReport.sql", ";", False) = 3 Then GoTo errHand
    Else
        '通过数据库中拷贝得到
        If CopyReport(mcnOwner, Mid(mlst标准.Key, 2), intDefSysCode * 100 + mlng帐套) = False Then GoTo errHand
    End If
    
    '调整安装导致的序列与实际数值的匹配
    stbThis.Panels(2).Text = "序列检查"
    DoEvents
    Call ChkSequence
    
    '填写安装记录为正常安装
    strSQL = "update zlSystems set 正常安装=1 where 编号=" & intDefSysCode * 100 + mlng帐套
    gcnOracle.Execute strSQL
    strSQL = "insert into zlSysFiles(系统,操作,文件名,日期,操作人)" & _
            " values (" & intDefSysCode * 100 + mlng帐套 & ",1,'" & lblSetupFile.Caption & "',sysdate,user)"
    gcnOracle.Execute strSQL
     
    '刘兴宏加入历史数据空间的判断
    gstrSQL = "Select 表名 from zltools.zlBakTables where rownum<=1 and  系统=" & intDefSysCode * 100 + mlng帐套
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    If Not rsTemp.EOF Then
        '需要检查是否存在H表
        gstrSQL = "Select 1 From User_Tables where table_name='H" & Nvl(rsTemp!表名) & "'"
        OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOwner
        If Not rsTemp.EOF Then
            MsgBox "该系统的数据结构太旧，请手工升级后再进行再植!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        
        Dim strMsg As String, lngCount As Long
        lngCount = 0
ResumeGo:
        lngCount = lngCount + 1
        strMsg = "被再植的系统必需存在历史数据空间,是否创建历史数据空间？" & vbCrLf & _
             "选择【是】：新创建一个历史数据空间。" & vbCrLf & _
             "选择【否】：再植一个已经存在的历史数据空间。"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
        '注意*****,由于调用gobjRegister.GetPassword函数，要求本窗体的标题固定为"应用系统再植",否则不能返回转换后的密码
            '创建一个历史数据空间
            If frmHistorySpaceSet.ShowInstall(Me, mcnOwner, cboOwnerUsr.Text, gobjRegister.GetPassword, intDefSysCode * 100 + mlng帐套, 0, 0) = False Then
                '需选择三次
                If lngCount > 3 Then
                        MsgBox "该系统的再植失败!", vbInformation + vbDefaultButton1, gstrSysName
                        GoTo errHand:
                Else
                    GoTo ResumeGo:
                End If
            End If
        Else
            '创建一个历史数据空间
            If frmHistorySpaceSet.ShowInstall(Me, mcnOwner, cboOwnerUsr.Text, gobjRegister.GetPassword, intDefSysCode * 100 + mlng帐套, 2, 0) = False Then
                '需选择三次
                If lngCount > 3 Then
                        MsgBox "该系统的再植失败!", vbInformation + vbDefaultButton1, gstrSysName
                        GoTo errHand:
                Else
                    GoTo ResumeGo:
                End If
            End If
        End If
    End If
    
   
    If mcnOwner.State = adStateOpen Then mcnOwner.Close
    
    SysInstall = True
    Exit Function

errHand:
    If mcnOwner.State = adStateOpen Then mcnOwner.Close
    pgbState.Visible = False
    SysInstall = False
    MsgBox err.Description, vbExclamation, "提醒"
End Function

Private Function CopyManageData(ByVal cnExecuter As ADODB.Connection) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    Dim lngNewSystem As Long
    Dim lngOldSystem As Long
    Dim strOldOwner As String
    
    pgbState.value = 0
    pgbState.Visible = True
    
    lngNewSystem = intDefSysCode * 100 + mlng帐套
    lngOldSystem = Mid(mlst标准.Key, 2)
    
    strOldOwner = GetOwnerName(lngOldSystem, gcnOracle)
    
    On Error GoTo errHandle
    'zlComponent数据
    gstrSQL = "insert into zlComponent(部件,名称,主版本,次版本,附版本,系统) " & _
                "select 部件,名称,主版本,次版本,附版本," & lngNewSystem & " from zlComponent where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 5
    
    'zlPrograms数据
    gstrSQL = "insert into zlPrograms(序号,标题,说明,部件,系统) " & _
                "select 序号,标题,说明,部件," & lngNewSystem & " from zlPrograms where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 15
    
    'zlProgFuncs数据
    gstrSQL = "insert into zlProgFuncs(序号,功能,系统) " & _
                "select 序号,功能," & lngNewSystem & " from zlProgFuncs where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 35
    
    'zlProgPrivs数据
    gstrSQL = "insert into zlProgPrivs(序号,功能,所有者,对象,权限,系统) " & _
                "select 序号,功能,decode(所有者,'" & strOldOwner & "',user,所有者),对象,权限," & lngNewSystem & " from zlProgPrivs where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 65
    
    'zlMenus数据
    '清理无效菜单
    With rsTemp
        Do
            If .State = adStateOpen Then .Close
            gstrSQL = "select 1 from zlMenus A where 模块 is null and not exists(select 1 from zlMenus B where B.上级ID=A.ID)"
            .Open gstrSQL, cnExecuter
            If .EOF Then Exit Do
            strSQL = "delete from zlMenus A where 模块 is null and not exists(select 1 from zlMenus B where B.上级ID=A.ID)"
            cnExecuter.Execute gstrSQL
        Loop
    End With
    CopyMenu gcnOracle, lngOldSystem, lngNewSystem
    pgbState.value = 85
    
    'zlBaseCode数据
    gstrSQL = "insert into zlBaseCode(表名,固定,说明,分类,系统) " & _
                "select 表名,固定,说明,分类," & lngNewSystem & " from zlBaseCode where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 90
    
    'zlDataMove数据
    gstrSQL = "insert into zlDataMove(组号,组名,说明,日期字段,转出描述,上次日期,系统,状态) " & _
                "select 组号,组名,说明,日期字段,转出描述,上次日期," & lngNewSystem & ",状态 from zlDataMove where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 95
    
    'zlAutoJobs数据
    gstrSQL = "insert into zlAutoJobs(类型,序号,名称,说明,内容,参数,执行时间,间隔时间,系统) " & _
                "select 类型,序号,名称,说明,内容,参数,执行时间,间隔时间," & lngNewSystem & " from zlAutoJobs where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 97
    
    'zlParameters数据
    gstrSQL = "Insert Into zlParameters(ID,系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明) " & _
            " Select zlParameters_ID.Nextval," & lngNewSystem & ",模块,私有,参数号,参数名,参数值,缺省值,参数说明 From zlParameters Where 系统=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 99
    
    pgbState.value = 0
    pgbState.Visible = True
    CopyManageData = True
    Exit Function
errHandle:
    If MsgBox("出现下列错误，是否继续？" & vbCrLf & "    " & err.Description, vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
    pgbState.value = 0
    pgbState.Visible = True
    
End Function

Private Function RunSetupFile(cnThisDB As ADODB.Connection, FileName As String, Optional DeLimiter As String = ";", Optional ResumeNext As Boolean) As Byte
    '----------------------------------------------
    '功能：执行安装脚本文件
    '参数：
    '   cnThisDB:   指定的数据库连接
    '   FileName:   脚本文件
    '   Delimiter:  脚本文件语句分隔符号
    '   ResumeNext: 是否错误继续
    '返回：1-执行成功；2-存在错误但继续执行完毕；3-错误而中断
    '----------------------------------------------
    Dim lngLines As Long
    err = 0
    On Error Resume Next
    Set objText = gobjFile.OpenTextFile(FileName)
    If err <> 0 Then
        MsgBox "无法打开脚本文件" & FileName & ",执行中断" & vbNewLine & err.Description, vbExclamation, gstrSysName
        RunSetupFile = 3
        Exit Function
    End If
    
    lngLines = GetFileLineCount(objText)
    objText.Close
    Set objText = gobjFile.OpenTextFile(FileName)
        
    pgbState.value = 0
    pgbState.Visible = True
    
    RunSetupFile = 1
    err = 0
    On Error GoTo 0
    strSQL = ""
    Do While Not objText.AtEndOfStream
        strTemp = Trim(objText.ReadLine)
        If Left(strTemp, 2) <> "--" Then
            If Right(strTemp, 1) = DeLimiter Then
                strSQL = strSQL & vbCrLf & Left(strTemp, Len(strTemp) - 1)
            Else
                strSQL = strSQL & vbCrLf & strTemp
            End If
            If Left(strSQL, 2) = vbCrLf Then
                If Len(strSQL) = 2 Then
                    strSQL = ""
                Else
                    strSQL = Mid(strSQL, 3)
                End If
            End If
        End If
        If (Right(strTemp, 1) = DeLimiter Or objText.AtEndOfStream) And Len(strSQL) <> 0 Then
            err = 0
            On Error Resume Next
            cnThisDB.Execute strSQL
            If err <> 0 Then
                If ResumeNext Then
                    RunSetupFile = 2
                Else
                    MsgBox "由于文件" & FileName & "中存在下面错误导致执行中断：" & vbCr & strSQL, vbExclamation, gstrSysName
                    RunSetupFile = 3
                    Exit Function
                End If
            End If
            err = 0
            On Error GoTo 0
            strSQL = ""
        End If
        pgbState.value = objText.Line / lngLines * 100
        DoEvents
    Loop
    pgbState.value = 0
    pgbState.Visible = False

End Function

Private Sub ChkSequence()
    '----------------------------------------------
    '功能：整理序列的当前号码
    '----------------------------------------------
    Dim rsLst As ADODB.Recordset
    
    pgbState.value = 0
    pgbState.Visible = True
    Set rsLst = GetSequence("", mcnOwner)
    With rsLst
        Do While Not .EOF
            DoEvents
            pgbState.value = .AbsolutePosition / .RecordCount * 100
            Call AdjustNameSequece(!Owner & "." & !Table_Name, mcnOwner, !Column_Name)
            .MoveNext
        Loop
        
        Call Adjust结帐ID(mcnOwner)
    End With
    pgbState.value = 0
    pgbState.Visible = False
End Sub
