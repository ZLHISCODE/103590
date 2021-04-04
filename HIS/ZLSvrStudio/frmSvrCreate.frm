VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSvrCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "创建管理工具"
   ClientHeight    =   5115
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "frmSvrCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7950
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraTbs 
      Height          =   1035
      Left            =   2310
      TabIndex        =   4
      Top             =   1380
      Width           =   5145
      Begin VB.TextBox txtTbsFile 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   630
         TabIndex        =   6
         Top             =   600
         Width           =   4185
      End
      Begin VB.TextBox txtTmpFile 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   630
         TabIndex        =   7
         Top             =   600
         Width           =   4185
      End
      Begin VB.TextBox txtTbsSize 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "200"
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtTmpSize 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "50"
         Top             =   180
         Width           =   555
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "名称"
         Height          =   180
         Left            =   225
         TabIndex        =   11
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblTbsFile 
         AutoSize        =   -1  'True
         Caption         =   "文件"
         Height          =   180
         Left            =   225
         TabIndex        =   10
         Top             =   660
         Width           =   360
      End
      Begin VB.Label lblTbsSize 
         AutoSize        =   -1  'True
         Caption         =   "大小        M"
         Height          =   180
         Left            =   3645
         TabIndex        =   9
         Top             =   255
         Width           =   1170
      End
      Begin VB.Label lblTbsName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "zlToolsTbs"
         Height          =   300
         Left            =   630
         TabIndex        =   14
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label lblTmpName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "zlToolsTmp"
         Height          =   300
         Left            =   630
         TabIndex        =   15
         Top             =   210
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdRegFile 
      Caption         =   "选择(&R)…"
      Height          =   350
      Left            =   6585
      TabIndex        =   26
      Top             =   3330
      Width           =   1100
   End
   Begin VB.TextBox txtPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2205
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "ZLSOFT"
      Top             =   360
      Width           =   2160
   End
   Begin MSComctlLib.ProgressBar pgbState 
      Height          =   150
      Left            =   2595
      TabIndex        =   20
      Top             =   4875
      Visible         =   0   'False
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   1995
      TabIndex        =   18
      Top             =   4245
      Width           =   1100
   End
   Begin VB.CommandButton cmdSqlFile 
      Caption         =   "选择(&S)…"
      Height          =   350
      Left            =   6570
      TabIndex        =   12
      Top             =   2625
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6585
      TabIndex        =   2
      Top             =   4245
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5385
      TabIndex        =   1
      Top             =   4245
      Width           =   1100
   End
   Begin VB.PictureBox PicSetup 
      Align           =   3  'Align Left
      Height          =   4740
      Left            =   0
      ScaleHeight     =   4680
      ScaleWidth      =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   1740
      Begin VB.FileListBox fltFile 
         Appearance      =   0  'Flat
         Height          =   1104
         Left            =   135
         Pattern         =   "*.zcr"
         TabIndex        =   27
         Top             =   3405
         Visible         =   0   'False
         Width           =   1410
      End
      Begin MSComDlg.CommonDialog dlgMain 
         Left            =   255
         Top             =   2835
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgSetup 
         Height          =   2625
         Left            =   60
         Picture         =   "frmSvrCreate.frx":058A
         Stretch         =   -1  'True
         Top             =   -105
         Width           =   945
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   19
      Top             =   4740
      Width           =   7944
      _ExtentX        =   14023
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1746
            MinWidth        =   882
            Text            =   "安装进度 "
            TextSave        =   "安装进度 "
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8493
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "16:59"
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
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   855
      TabIndex        =   21
      Top             =   4095
      Width           =   6975
   End
   Begin MSComctlLib.TabStrip tbsTbs 
      Height          =   1440
      Left            =   2205
      TabIndex        =   28
      ToolTipText     =   "本地管理,自动分配区尺寸(AUTOALLOCATE),如果是临时表空间,则统一区尺寸1M"
      Top             =   1080
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   2540
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "缺省空间"
            Key             =   "Tbs"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "临时空间"
            Key             =   "Tmp"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label lblRegFile 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2190
      TabIndex        =   25
      Top             =   3645
      Width           =   5490
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "4)系统运行需要注册授权文件，请按指南选择："
      Height          =   180
      Index           =   3
      Left            =   1995
      TabIndex        =   24
      Top             =   3375
      Width           =   3780
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "(默认密码为""ZLSOFT"")"
      Height          =   180
      Left            =   4410
      TabIndex        =   23
      Top             =   435
      Width           =   1800
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "1)管理工具所有者用户固定为""zlTools""，建议重新设置其密码："
      Height          =   180
      Index           =   0
      Left            =   1995
      TabIndex        =   22
      Top             =   90
      Width           =   5130
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "3)管理工具创建依赖于脚本文件，请按指南选择："
      Height          =   180
      Index           =   2
      Left            =   1995
      TabIndex        =   17
      Top             =   2670
      Width           =   3960
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "2)管理工具需存储一定管理数据，请确定其表空间的位置与大小："
      Height          =   180
      Index           =   1
      Left            =   1995
      TabIndex        =   16
      Top             =   840
      Width           =   5220
   End
   Begin VB.Label lblSqlFile 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2190
      TabIndex        =   13
      Top             =   2940
      Width           =   5475
   End
End
Attribute VB_Name = "frmSvrCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrTbsPath As String                        '缺省表空间路径名称，根据历史表空间产生
Private mobjFiles As New FileSystemObject
Private mobjText As TextStream

Private cnTools As New ADODB.Connection

Private mclsRunScript As clsRunScript
'临时变量
Dim rsTemp As New ADODB.Recordset
Dim strSQL As String, strTemp As String
Dim lngCount As Long

Private Sub cmdCancel_Click()
    If MsgBox("尚未创建管理工具，真的取消吗？", vbQuestion + vbYesNo, "提示") = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub cmdRegFile_Click()
    With Me.dlgMain
        .FileName = lblRegFile.Caption
        .DialogTitle = "选择注册授权文件"
        .Filter = "(注册授权文件)|*.zcr"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblRegFile.Caption = .FileName
        End If
    End With
End Sub

Private Sub cmdSqlFile_Click()
    With Me.dlgMain
        .FileName = lblSqlFile.Caption
        .DialogTitle = "选择管理工具脚本文件"
        .Filter = "(管理工具脚本文件)|zlServer.sql;*.plb"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            Me.fltFile.Path = Mid(.FileName, 1, Len(.FileName) - InStr(1, StrReverse(.FileName), "\") + 1)
            Me.fltFile.Pattern = IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")
            If Me.fltFile.ListCount = 0 Then
                lblSqlFile.Caption = ""
                MsgBox "该位置未包含授权验证文件！", vbExclamation, gstrSysName
            Else
                lblSqlFile.Caption = .FileName
            End If
        End If
    End With
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub cmdOK_Click()
    
    If Trim(Me.txtPwd.Text) = "" Then
        MsgBox "没有设置管理工具所有者密码，不能继续！", vbExclamation, "提示"
        Me.txtPwd.SetFocus
        Exit Sub
    End If

    If Val(txtTbsSize.Text) < 100 Then
        MsgBox "缺省空间大小设置错误！", vbExclamation, "提示"
        txtTbsSize.Text = 100
        If txtTbsSize.Visible Then txtTbsSize.SetFocus
        Exit Sub
    End If
    If Val(txtTmpSize.Text) < 50 Then
        MsgBox "临时空间大小设置错误！", vbExclamation, "提示"
        txtTmpSize.Text = 50
        If txtTmpSize.Visible Then txtTmpSize.SetFocus
        Exit Sub
    End If

    If Trim(lblSqlFile.Caption) = "" Then
        MsgBox "未指定管理工具脚本文件，不能继续！", vbExclamation, "提示"
        cmdSqlFile.SetFocus
        Exit Sub
    End If

    If Trim(lblRegFile.Caption) = "" Then
        MsgBox "未执行注册授权文件，不能继续！", vbExclamation, "提示"
        cmdRegFile.SetFocus
        Exit Sub
    End If
    
    If MsgBox("工具创建过程将持续较长的时间，" & vbCr & "请不要随意中断程序的运行。" & vbCr & vbCr & "继续吗？", vbQuestion + vbYesNo, "提示") = vbNo Then Exit Sub
    
    Me.txtPwd.Enabled = False
    fraTbs.Enabled = False
    cmdSqlFile.Enabled = False
    cmdRegFile.Enabled = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    If Not CheckCBOPars Then Exit Sub
    If svrCreate(lblSqlFile.Caption) = True Then
        MsgBox "管理工具成功创建！", vbExclamation, "提示"
        If Not sysRegist(Me.lblRegFile.Caption) Then
            MsgBox "系统注册授权出错，请重新登录进行注册授权。", vbInformation, "提示"
        End If
        Me.txtPwd.Enabled = True
        fraTbs.Enabled = True
        cmdSqlFile.Enabled = True
        cmdRegFile.Enabled = True
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        Unload Me
    Else
        MsgBox "创建过程发生错误，系统将自动清除已经执行的操作！", vbExclamation, "提示"
        Call svrRemove
        Me.txtPwd.Enabled = True
        fraTbs.Enabled = True
        cmdSqlFile.Enabled = True
        cmdRegFile.Enabled = True
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
    End If

End Sub

Private Sub Form_Load()
    
    
    imgSetup.Top = PicSetup.ScaleTop
    imgSetup.Left = PicSetup.ScaleLeft
    imgSetup.Height = PicSetup.ScaleHeight
    imgSetup.Width = PicSetup.ScaleWidth
    
    pgbState.Left = stbThis.Panels(3).Left + 90
    pgbState.Width = stbThis.Panels(4).Left - pgbState.Left - 90
    pgbState.Top = stbThis.Top + stbThis.Height / 3
    
    With rsTemp
        .Filter = 0
        If .State = adStateOpen Then .Close
        strSQL = "SELECT NAME from V$DATAFILE where ROWNUM<2 order by CREATION_TIME"
        .Open strSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then
            mstrTbsPath = "C:\"
        Else
            For lngCount = Len(!name) To 2 Step -1
                If Mid(!name, lngCount, 1) = "\" Or Mid(!name, lngCount, 1) = "/" Then
                    mstrTbsPath = Left(!name, lngCount)
                    Exit For
                End If
            Next
        End If
    End With
    
    txtTbsFile.Text = mstrTbsPath & lblTbsName.Caption & ".DBF"
    txtTmpFile.Text = mstrTbsPath & lblTmpName.Caption & ".DBF"
    
    If Dir(App.Path & "\Tools\" & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")) <> "" And Dir(App.Path & "\Tools\zlServer.Sql") <> "" Then
        lblSqlFile.Caption = App.Path & "\Tools\zlServer.Sql"
    End If
    
    Me.fltFile.Path = App.Path
    Me.fltFile.Pattern = "*.zcr"
    If Me.fltFile.ListCount > 0 Then
        Me.lblRegFile.Caption = App.Path & "\" & Me.fltFile.List(0)
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdOK.Enabled = False Then
        Cancel = 1
        Exit Sub
    End If
    Set mclsRunScript = Nothing
    Set mobjFiles = Nothing
    Set mobjText = Nothing
End Sub

Private Sub tbsTbs_Click()
    If tbsTbs.Tabs(1).Selected Then
        lblTbsName.Visible = True
        txtTbsFile.Visible = True
        txtTbsSize.Visible = True
        
        lblTmpName.Visible = False
        txtTmpFile.Visible = False
        txtTmpSize.Visible = False
        
    ElseIf tbsTbs.Tabs(2).Selected Then
        lblTbsName.Visible = False
        txtTbsFile.Visible = False
        txtTbsSize.Visible = False
        
        lblTmpName.Visible = True
        txtTmpFile.Visible = True
        txtTmpSize.Visible = True
        
    End If

End Sub

Private Sub txtPWD_GotFocus()
    Me.txtPwd.SelStart = 0: Me.txtPwd.SelLength = 100
End Sub


'-----------------------------------------------------
'以下为内部通用函数方法：
'-----------------------------------------------------
Private Function svrCreate(strSqlFile As String) As Boolean
    '----------------------------------
    '功能：完成系统的安装处理
    '    创建工具数据空间
    '    创建工具所有者
    '    创建工具数据对象
    '    建立公共同义词，授予public权限
    '----------------------------------
    Dim intVer As Integer
    Dim strRegFunFile As String
    Dim blnJSONRemain As Boolean
    
    '创建表空间及回滚段
    stbThis.Panels(2).Text = "创建工具缺省空间…"
    If CreateTbs(lblTbsName.Caption, txtTbsFile.Text, txtTbsSize.Text, True, False, False, 1) = 2 Then GoTo ErrHand
    
    intVer = GetOracleVersion(, True)
    If intVer >= 9 Then
        'Oracle9i版本,用户的临时空间只能是本地管理临时表空间；且不需要创建公共回滚段
        DoEvents
        stbThis.Panels(2).Text = "创建工具临时空间…"
        If CreateTbs(lblTmpName.Caption, txtTmpFile.Text, txtTmpSize.Text, True, True, False, 1) = 2 Then GoTo ErrHand
    Else
        'Oracle8i以下版本
        DoEvents
        stbThis.Panels(2).Text = "创建工具临时空间…"
        If CreateTbs(lblTmpName.Caption, txtTmpFile.Text, txtTmpSize.Text, True, True, False, 1) = 2 Then GoTo ErrHand
    
        DoEvents
        stbThis.Panels(2).Text = "创建公共回滚段…"
        err = 0
        On Error Resume Next
        strSQL = "create public rollback segment rbs_ZLTOOLS tablespace RBS"
        gcnOracle.Execute strSQL
        
        If err <> 0 Then
            err = 0
            On Error GoTo ErrHand
            '本地管理表空间,无需再指定段的存储参数
            strSQL = "create public rollback segment rbs_ZLTOOLS tablespace " & lblTbsName.Caption
            gcnOracle.Execute strSQL
        End If
        strSQL = "alter rollback segment rbs_ZLTOOLS online"
        gcnOracle.Execute strSQL
    End If
    
    '----------------------------------------------
    '创建工具所有者
    stbThis.Panels(2).Text = "创建工具所有者…"
    err = 0
    On Error Resume Next
    gcnOracle.Execute "create user ZLTOOLS identified by " & txtPwd.Text
    If err <> 0 Then
        MsgBox "无法创建创建工具所有者，错误：" & vbNewLine & err.Description, vbExclamation, "提示"
        
        gcnOracle.Execute "drop tablespace " & Trim(lblTbsName.Caption) & " including contents and datafiles cascade constraints"
        gcnOracle.Execute "drop tablespace " & Trim(lblTmpName.Caption) & " including contents and datafiles cascade constraints"
        GoTo ErrHand
    End If
    
    gcnOracle.Execute "alter user ZLTOOLS DEFAULT TABLESPACE " & Trim(lblTbsName.Caption)
    gcnOracle.Execute "alter user ZLTOOLS TEMPORARY TABLESPACE " & Trim(lblTmpName.Caption)
    gcnOracle.Execute "grant Connect,Resource,UNLIMITED TABLESPACE,Create Public Synonym,Drop Public Synonym,Alter Session,Create Session,Create Synonym,Create Table,Create View,Create Sequence,Create Database Link,Create Cluster to ZLTOOLS"
    gcnOracle.Execute "grant select on Sys.v_$session to ZLTOOLS"
    gcnOracle.Execute "grant select on Sys.gv_$session to ZLTOOLS"
    gcnOracle.Execute "grant select on Sys.dba_role_privs to ZLTOOLS"

    
    If err <> 0 Then
        MsgBox "无法创建创建工具所有者，请检查数据库系统的正确性" & vbNewLine & err.Description, vbExclamation, "提示"
        GoTo ErrHand
    End If

    '----------------------------------------------
    '创建工具数据对象
    stbThis.Panels(2).Text = "创建对象:"
    err = 0
    On Error GoTo ErrHand
    With cnTools
        If .State = adStateOpen Then .Close
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & Trim(gstrServer), "ZLTOOLS", txtPwd.Text
    End With
    
    Call SetSQLTrace(gstrServer, "ZLTOOLS", cnTools)

    
    Set mclsRunScript = New clsRunScript
    With mclsRunScript
        Set .Connection = cnTools: .ConnectType = 1
        Call .InitGlobalPara(Me)
        Call .InitUserList(, , txtPwd.Text)
        If IsCanInstallPLJson(gobjFSO.GetParentFolderName(strSqlFile), blnJSONRemain) Then
            Call InstallPLJSON(gcnOracle, gobjFSO.GetParentFolderName(strSqlFile), mclsRunScript, blnJSONRemain)
        End If
        On Error Resume Next
        If .OpenFile(strSqlFile) = False Then
            GoTo ErrHand
        End If
        
        pgbState.value = 0
        pgbState.Visible = True
        err = 0
        On Error GoTo ErrHand
        Do While Not mclsRunScript.EOF
            pgbState.value = Int(.ProcessValue)
            err = 0
            On Error Resume Next
            If pgbState.value > 90 Then
                Debug.Print ""
            End If
            cnTools.Execute .SQLInfo.SQL
            If err <> 0 Then
                MsgBox "由于文件" & strSqlFile & "中存在下面错误导致执行中断：" & vbCr & .SQLInfo.SQL & vbNewLine & err.Description, vbExclamation, "提示"
                GoTo ErrHand
            End If
            err = 0
            On Error GoTo ErrHand
            DoEvents
            .ReadNextSQL
        Loop
    End With
        '----------------------------------------------
        '通过Shell方式，调入授权验证函数
        
    stbThis.Panels(2).Text = "命令执行脚本…"
    strRegFunFile = Mid(strSqlFile, 1, Len(strSqlFile) - InStr(1, StrReverse(strSqlFile), "\") + 1) & IIf(GetOracleVersion(True, True) > 11, "ZLREGIST12C.PLB", "zlRegist.plb")
    
    If Not RunRegistFile(Me, cnTools, Trim(txtPwd.Text), gstrServer, strRegFunFile) Then
        GoTo ErrHand
    End If
    
    With rsTemp
        If .State = adStateOpen Then .Close
        strSQL = "Select 1 From User_Objects Where Object_Type = 'FUNCTION' And Object_Name = '" & UCase("f_Reg_Audit") & "' and status='VALID'"
        .Open strSQL, cnTools
        If .RecordCount = 0 Then GoTo ErrHand
    End With
    
    '----------------------------------------------
    '建立公共同义词，授予public权限
    stbThis.Panels(2).Text = "授权处理:"
    pgbState.Visible = False
    Call ReGrantForTools(cnTools)
    cnTools.Close
    svrCreate = True
    Exit Function

ErrHand:
    If cnTools.State = adStateOpen Then cnTools.Close
    pgbState.Visible = False
    svrCreate = False
End Function

Private Function svrRemove() As Boolean
    '----------------------------------
    '功能：删除已经的安装处理
    '----------------------------------
    Dim strSpaces As String, strFiles As String, aryFile() As String, strErrInfo As String
    Dim strStep As String, aryStep() As String
    Dim lngRowH As Long, intVer As Integer
    
    strFiles = ""
    With rsTemp
        .Filter = 0
        If .State = adStateOpen Then .Close
        strSQL = "select F.NAME " & _
                " from V$TABLESPACE T,V$DATAFILE F " & _
                " where T.TS#=F.TS# " & _
                "       and T.NAME in('" & UCase(lblTbsName.Caption) & "','" & UCase(lblTbsName.Caption) & "')"
        .Open strSQL, gcnOracle
        Do While Not .EOF
            strFiles = strFiles & ";" & .Fields(0).value
            DoEvents
            .MoveNext
        Loop
    End With
    
    err = 0
    On Error Resume Next
    stbThis.Panels(2).Text = "删除所有者…"
    Do
        gcnOracle.Execute "drop user ZLTOOLS cascade"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open "select * from all_users where username='ZLTOOLS'", gcnOracle
        If rsTemp.EOF Then Exit Do
        lngCount = lngCount + 1
        DoEvents
        If lngCount > 10000 Then
            strErrInfo = strErrInfo & vbCr & "用户:ZLTOOLS"
            Exit Do
        End If
    Loop
    
    stbThis.Panels(2).Text = "删除工具空间及其文件…"
    intVer = GetOracleVersion(, True)
    If intVer < 9 Then
        gcnOracle.Execute "alter rollback segment rbs_ZLTOOLS offline"
        gcnOracle.Execute "drop rollback segment rbs_ZLTOOLS"
    End If
    
    gcnOracle.Execute "alter tablespace " & lblTmpName.Caption & " offline"
    gcnOracle.Execute "alter tablespace " & lblTbsName.Caption & " offline"
    
    gcnOracle.Execute "drop tablespace " & lblTmpName.Caption & " including contents and datafiles cascade constraints"
    gcnOracle.Execute "drop tablespace " & lblTbsName.Caption & " including contents and datafiles cascade constraints"
    
    
    aryFile = Split(Mid(strFiles, 2), ";")
    For lngCount = 0 To UBound(aryFile)
        err = 0
        mobjFiles.DeleteFile aryFile(lngCount), True
        If err <> 0 Then
            strErrInfo = strErrInfo & vbCr & "文件：" & aryFile(lngCount)
        End If
    Next
    
    stbThis.Panels(2).Text = ""
    If strErrInfo <> "" Then
        MsgBox "请重启动数据库后,手工删除以下内容：" & strErrInfo, vbExclamation, "提示"
    Else
        MsgBox "请检查硬盘空间和数据库系统，确认无误后重新操作", vbExclamation, "提示"
    End If
End Function

Private Function CreateTbs(TbsName As String, TbsFile As String, TbsSize As Integer, Optional AutoExtend As Boolean, _
     Optional Temp As Boolean, Optional AutoAllocate As Boolean, Optional ExtentSize As Integer) As Byte
    '----------------------------------------------
    '功能：系统用户,根据参数创建表空间,固定为本地管理类型(8i以前不支持,那时只能创建字典管理类型)
    '       因可能涉及LOB字段等原因,不创建ASSM表空间(仅9i以上支持,SEGMENT SPACE MANAGEMENT AUTO)
    '参数：
    '   TbsName:表空间名称
    '   TbsFile:表空间文件
    '   TbsSize:表空间大小(M为单位)
    '   Extend:是否自动管理区,否则统一范围尺寸
    '   ExtentSize:统一区尺寸,临时表空间必须指定尺寸(Oracle缺省为1M)
    '   Temp:是否为临时表空间
    '返回：1-创建成功；2-表空间已经存在；3-创建失败
    '----------------------------------------------
    DoEvents
    If Temp Then
        gstrSQL = "CREATE TEMPORARY TABLESPACE " & TbsName & " TEMPFILE '" & TbsFile & "'"
    Else
        gstrSQL = "CREATE TABLESPACE " & TbsName & " DATAFILE '" & TbsFile & "'"
    End If
    gstrSQL = gstrSQL & _
            " SIZE " & TbsSize & "M REUSE " & _
             IIf(AutoExtend, "AUTOEXTEND ON NEXT " & IIf(TbsSize \ 10 = 0, 1, TbsSize \ 10) & "M", "") & _
            " EXTENT MANAGEMENT LOCAL " & _
                IIf(AutoAllocate And Not Temp, " AUTOALLOCATE", " UNIFORM SIZE " & IIf(ExtentSize = 0, "1", ExtentSize) & "M")
    
    err = 0
    On Error Resume Next
    gcnOracle.Execute gstrSQL
    DoEvents
    If err = 0 Then
        CreateTbs = 1
    ElseIf gcnOracle.Errors.Count > 0 Then
        MsgBox gcnOracle.Errors(0).Description & _
            IIf(InStr(1, gcnOracle.Errors(0).Description, "00406") > 0, vbCrLf & "请修改Oracle的配置文件的Compatible参数为8.1.5以上", ""), _
            vbExclamation, "提示"
        CreateTbs = 2
    Else
        MsgBox "表空间" & TbsName & "无法创建，请检查磁盘大小等", vbExclamation, "提示"
        CreateTbs = 2
    End If

End Function

Private Function sysRegist(strRegFile As String) As Boolean
    '----------------------------------
    '功能：系统注册
    '----------------------------------
    stbThis.Panels(2).Text = "系统注册授权…"
    
    '写入临时表，并验证
    err = 0: On Error GoTo ErrHand
    Me.MousePointer = vbHourglass
    
    If gobjRegister.zlRegBuild(strRegFile, pgbState) = False Then GoTo ErrHand
    
    Me.MousePointer = vbDefault
    
    If gobjRegister.zlRegCheck(True) <> "" Then GoTo ErrHand
    
    '正式写入
    gcnOracle.Execute "call zltools.p_Reg_Apply()", , adCmdText
    
    sysRegist = True
    Exit Function
ErrHand:
    Me.MousePointer = vbDefault
    sysRegist = False
End Function

