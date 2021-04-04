VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "图像校对"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8160
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picDept 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   8280
      ScaleHeight     =   2745
      ScaleWidth      =   3675
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   3705
      Begin VB.CommandButton cmdFindCancle 
         Caption         =   "取消"
         Height          =   270
         Left            =   3000
         TabIndex        =   18
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdFindOk 
         Caption         =   "确定"
         Height          =   270
         Left            =   2280
         TabIndex        =   17
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox ChkSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2400
         Width           =   675
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2280
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fraValid 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "检查文件大小是否为0"
      Top             =   120
      Width           =   7935
      Begin VB.CheckBox chkPassive 
         Caption         =   "启用被动传输"
         Height          =   375
         Left            =   6360
         TabIndex        =   21
         Top             =   2340
         Width           =   1455
      End
      Begin VB.TextBox txtDept 
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton cmdDept 
         Caption         =   "…"
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdValid 
         Caption         =   "校对"
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkValid 
         Caption         =   "是否校对已审核数据"
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CheckBox chkRoadValid 
         Caption         =   "是否检查文件路径"
         Height          =   495
         Left            =   3480
         TabIndex        =   5
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkSizeVlid 
         Caption         =   "是否校对文件大小"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkReadValid 
         Caption         =   "是否校对文件读取(耗时较长)"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "检查文件是否能正常读取"
         Top             =   3120
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "历史查阅"
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   167903235
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   167903235
         CurrentDate     =   38082.9993055556
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "校对科室"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         Caption         =   "结束日期"
         Height          =   180
         Left            =   3840
         TabIndex        =   13
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "开始日期"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1530
         Width           =   720
      End
   End
   Begin VB.PictureBox picHint 
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   4000
      Width           =   8055
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   11040
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6852
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DEC
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7386
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staPane 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   4125
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14340
            MinWidth        =   1587
            Text            =   "准备检测"
            TextSave        =   "准备检测"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmResult As frmResult
Attribute mfrmResult.VB_VarHelpID = -1
Private mlngPassive As Long
Private mstrCurValid As String

Public Sub zlShowMe(Optional strCmdLine As String)
    On Error GoTo errHandle
    
    Set gobjComlib = DynamicCreate("zl9ComLib.clsComLib", "zl9ComLib.dll")
    
    Call gobjComlib.InitCommon(gcnOracle)
    
    mstrCurValid = ""
    mlngPassive = Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用被动传输", 0))
    chkPassive.Value = mlngPassive

    picHint.BackColor = &H8000000D
    picHint.Width = 0
    picHint.Left = -15
    
    Call InitLvwList
    Call LoadDept
    
    Call InitPara(strCmdLine)
    
    If glngState <> 2 Then
        Me.Show
    Else
        If Len(txtDept) = 0 Then Exit Sub
        Call ImageValid(dtpBegin.Value, dtpEnd.Value)
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Function ImageValid(dtBegin As Date, dtEnd As Date) As Boolean
'图像检测
    Dim strSql As String
    Dim rsRecord As New ADODB.Recordset
    Dim objFile As New Scripting.FileSystemObject
    Dim strCachePath As String
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim strTmpFile As String
    Dim lngResult As emResult
    Dim dcmImage As DicomImage
    Dim blnMoved As Boolean
    Dim lngCount As Long
    Dim lngCurIndex As Long
    Dim lngDefult As Long
    Dim lngUnValid As Long
    Dim strWhere As String
    Dim i As Long
    Dim strFtpDef As String
    Dim strFtpConnErr As String
    Dim strDept As String
    
    Call SetState(False)
    
    mstrCurValid = ""
    strFtpConnErr = ""
    cmdValid.Caption = "正在校对"
    lngDefult = 0

    strWhere = strWhere & " c.接收日期 >= [1] and c.接收日期 <= [2] and "
    strDept = GetDept
    strWhere = strWhere & "f.名称 in " & strDept
    
    If chkValid.Value <> 1 Then
        strWhere = strWhere & " and a.校对结果 is null"
    End If

    blnMoved = MovedByDate(dtpBegin.Value)
    strSql = "Select Rownum As 顺序号,c.医嘱ID,c.姓名, c.性别, c.年龄,c.影像类别,c.检查号, a.图像号, a.采集时间,c.接收日期, d.Ftp用户名 As User1, d.Ftp密码 As Pwd1, d.Ip地址 As Host1," & vbNewLine & _
                "       '/' || d.Ftp目录 || '/' As Root1, d.共享目录 As 共享目录1, d.共享目录用户名 As 共享目录用户名1, d.共享目录密码 As 共享目录密码1," & vbNewLine & _
                "       Decode(c.接收日期, Null, '', To_Char(c.接收日期, 'YYYYMMDD') || '/') || c.检查uid || '/' || a.图像uid As Url, d.设备号 As 设备号1," & vbNewLine & _
                "       d.设备名 As 设备名1, e.Ftp用户名 As User2, e.Ftp密码 As Pwd2, e.Ip地址 As Host2, '/' || e.Ftp目录 || '/' As Root2," & vbNewLine & _
                "       e.共享目录 As 共享目录2, e.共享目录用户名 As 共享目录用户名2, e.共享目录密码 As 共享目录密码2, e.设备号 As 设备号2, e.设备名 As 设备名2, a.图像uid, c.检查uid,f.名称,g.执行间," & vbNewLine & _
                "       b.序列uid, a.动态图, a.编码名称, a.录制长度, c.校对日期, a.校对结果" & vbNewLine & _
                "From 影像检查图象 a, 影像检查序列 b, 影像检查记录 c, 影像设备目录 d, 影像设备目录 e ,部门表 f,病人医嘱发送 g" & vbNewLine & _
                "Where a.序列uid = b.序列uid And b.检查uid = c.检查uid And c.位置一 = d.设备号(+) And c.位置二 = e.设备号(+)  and c.执行科室id = f.id and c.医嘱id = g.医嘱id and nvl(a.动态图,0) = 0 and " & strWhere & vbNewLine & _
                "Order by a.采集时间"
    
    Set rsRecord = gobjComlib.zlDatabase.OpenSQLRecord(strSql, "根据采集日期查询图像", dtBegin, dtEnd)
    
    lngCurIndex = 0
    lngCount = rsRecord.RecordCount
    
    If rsRecord.RecordCount > 0 Then
        Do While Not rsRecord.EOF
            strFtpDef = ""
            lngResult = etUndetected
            lngCurIndex = lngCurIndex + 1
            
'            staPane.Panels(1).Text = "正在校对："  & "。已发现" & lngDefult & "个校对失败文件。"
            staPane.Panels(1).Text = "正在校对(" & lngCurIndex & "/" & lngCount & ")：" & NVL(IIf(Len(NVL(rsRecord("设备名1"))) = 0, NVL(rsRecord("Root2")), NVL(rsRecord("Root1"))) & NVL(rsRecord("URL")))
''            lblHint.Refresh
            picHint.Width = 8055 / lngCount * lngCurIndex
            picHint.Refresh
            staPane.Refresh
            If InStr(strFtpConnErr, "[" & IIf(Len(rsRecord!Host1) = 0, rsRecord!Host2, rsRecord!Host1) & "]") = 0 Then

                lngResult = DoValid(rsRecord, lngDefult, strTmpFile, strFtpDef)
                
                If Len(strFtpDef) > 0 Then
                    strFtpConnErr = strFtpConnErr & "[" & strFtpDef & "]"
                    lngUnValid = lngUnValid + 1
                End If
'                '将校对失败的图像显示到表格中，并将结果记录到数据库中
                If lngResult <> etSucceed And lngResult <> etUndetected Then
                    lngDefult = lngDefult + 1
'                    If mfrmResult Is Nothing Then
'                        Set mfrmResult = New frmResult
'                    End If
'
'                    mfrmResult.AddNew rsRecord, lngResult, strTmpFile
                    If InStr(mstrCurValid, "[" & rsRecord("医嘱ID") & "]") = 0 Then
                        mstrCurValid = mstrCurValid & "[" & rsRecord("医嘱ID") & "]"
                    End If
'
                End If
            Else
                lngUnValid = lngUnValid + 1
            End If
            rsRecord.MoveNext
        Loop
    End If
    
'    lblHint.Caption = ""
    picHint.Width = 0
    staPane.Panels(1).Text = "校对完成。本次共" & lngCount & "个文件。" & lngDefult & "个校对失败" & IIf(lngUnValid > 0, "，" & lngUnValid & "个未校对(FTP连接失败)。", "。")
    cmdValid.Caption = "校对"
    
    Call SetState(True)
    
    
    strSql = "Select a.医嘱id, a.影像类别, a.姓名, 性别, a.年龄, a.检查uid, a.检查号, b.名称" & vbNewLine & _
            "From 影像检查记录 a, 部门表 b" & vbNewLine & _
            "Where a.执行科室id = b.Id And 校对状态 = [1]" & IIf(Len(strDept) > 0, " and b.名称 in " & strDept, "")

    Set rsRecord = gobjComlib.zlDatabase.OpenSQLRecord(strSql, "获取校对失败的检查信息", 2)
    
    
    If rsRecord.RecordCount > 0 Or lngDefult > 0 Then
        If mfrmResult Is Nothing Then
            Set mfrmResult = New frmResult
        End If
        
        mfrmResult.ShowMe GetDept, mstrCurValid
    Else
        If glngState = 2 Then
            Unload Me
        End If
    End If
End Function


Private Function DoValid(rsRecord As Recordset, ByRef lngDefult As Long, ByRef strTmpFile As String, ByRef strFtp As String, Optional ByVal blnRedo As Boolean) As emResult
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim objFile As New Scripting.FileSystemObject
    Dim strCachePath As String
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim lngResult As emResult
    Dim dcmImage As DicomImage
    Dim strSql As String
    
    '创建本地图像缓存目录
    strFtp = ""
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsRecord("URL")))
    strImgInstanceUid = Trim(NVL(rsRecord!图像UID))
    
    strTmpFile = strCachePath & NVL(rsRecord("URL"))
    
    
    strTmpFile = Replace(Trim(strTmpFile), "/", "\")
    
    '建立FTP连接
    If NVL(rsRecord("设备号1")) <> vbNullString And Inet1.hConnection = 0 Then
        If Inet1.FuncFtpConnect(NVL(rsRecord("Host1")), NVL(rsRecord("User1")), NVL(rsRecord("Pwd1"))) = 0 Then
            If NVL(rsRecord("设备号2")) <> vbNullString Then
                If Inet2.FuncFtpConnect(NVL(rsRecord("Host2")), NVL(rsRecord("User2")), NVL(rsRecord("Pwd2"))) = 0 Then
                    If glngState <> 2 And Not blnRedo Then
                        MsgBox "FTP【" & rsRecord("Host2") & "】不能正常连接，请检查网络设置。", vbOKOnly, CON_STR_HINT_TITLE
                    End If
                    strFtp = rsRecord("Host2")
                    DoValid = etUndetected
                    Inet1.FuncFtpDisConnect
                    Inet2.FuncFtpDisConnect
                    Exit Function
                End If
            Else
                If glngState <> 2 And Not blnRedo Then
                    MsgBox "FTP【" & rsRecord("Host1") & "】不能正常连接，请检查网络设置。", vbOKOnly, CON_STR_HINT_TITLE
                End If
                strFtp = rsRecord("Host1")
                DoValid = etUndetected
                Inet1.FuncFtpDisConnect
                Inet2.FuncFtpDisConnect
                Exit Function
            End If
        End If
    End If
    
    '检查文件是否存在
    If Not Inet1.FuncFtpFileExists(objFile.GetParentFolderName(NVL(rsRecord("Root1")) & rsRecord("URL")), objFile.GetFileName(rsRecord("URL"))) Then
        If NVL(rsRecord("设备号2")) <> vbNullString Then
            If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsRecord("Host2")), NVL(rsRecord("User2")), NVL(rsRecord("Pwd2"))
            If Not Inet2.FuncFtpFileExists(objFile.GetParentFolderName(NVL(rsRecord("Root2")) & rsRecord("URL")), strTmpFile) Then
                lngResult = etFileMiss
            End If
        Else
            lngResult = etFileMiss
        End If
    End If
    
    '检查文件大小
    If chkSizeVlid.Value = 1 And lngResult = etUndetected Then
        If Inet1.FuncFtpGetFileSize(objFile.GetParentFolderName(NVL(rsRecord("Root1")) & rsRecord("URL")), objFile.GetFileName(rsRecord("URL"))) = 0 Then
            If NVL(rsRecord("设备号2")) <> vbNullString Then
                If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsRecord("Host2")), NVL(rsRecord("User2")), NVL(rsRecord("Pwd2"))
                If Inet2.FuncFtpGetFileSize(objFile.GetParentFolderName(NVL(rsRecord("Root2")) & rsRecord("URL")), strTmpFile) = 0 Then
                    lngResult = etFileNull
                End If
            Else
                lngResult = etFileNull
            End If
        End If
    End If
    
    '文件不存在，在根目录中判断是否存在，若存在便是路径错误
    If chkRoadValid.Value = 1 And lngResult = etFileMiss Then
        If Not Inet1.FuncFtpFileExists(objFile.GetParentFolderName(rsRecord("URL")), objFile.GetFileName(rsRecord("URL"))) Then
            If NVL(rsRecord("设备号2")) <> vbNullString Then
                If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsRecord("Host2")), NVL(rsRecord("User2")), NVL(rsRecord("Pwd2"))
                If Not Inet2.FuncFtpFileExists(objFile.GetParentFolderName(rsRecord("URL")), strTmpFile) Then
                    lngResult = etRoadError
                End If
            End If
        Else
            lngResult = etRoadError
        End If
    End If
    
    '检查读取
    If chkReadValid.Value = 1 And lngResult = etUndetected Then
        '从FTP下载到本地
        If Inet1.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsRecord("Root1")) & rsRecord("URL")), strTmpFile & ".001", objFile.GetFileName(rsRecord("URL")), , hwnd) <> 0 Then
            '从设备号1提取图像失败，则从设备号2提取图像
            If NVL(rsRecord("设备号2")) <> vbNullString Then
                If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsRecord("Host2")), NVL(rsRecord("User2")), NVL(rsRecord("Pwd2"))
                If Inet2.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsRecord("Root2")) & rsRecord("URL")), strTmpFile & ".001", objFile.GetFileName(rsRecord("URL")), , hwnd) <> 0 Then
                    lngResult = etReadError
                End If
            Else
                lngResult = etReadError
            End If
        End If
        
        '从本地读取
        If lngResult = etUndetected Then
            Set dcmImage = ReadViewImage(strTmpFile & ".001")
            
            Kill strTmpFile & ".001"
            
            If dcmImage Is Nothing Then
                lngResult = etReadError
            End If
        End If
    End If
    
    If lngResult = etUndetected Then lngResult = etSucceed
    
    ' 记录到数据库
    strSql = "zl_影像检查图象_校对('" & rsRecord("医嘱ID") & "','" & rsRecord("图像UID") & "',to_date('" & gobjComlib.zlDatabase.Currentdate & "','yyyy-mm-dd hh24:mi:ss')," & lngResult & ")"
    Call gobjComlib.zlDatabase.ExecuteProcedure(strSql, "保存校对结果")
    
    DoValid = lngResult
    
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
End Function

Private Sub InitPara(strPara As String)
    Dim arrPara() As String
    Dim intDate As Integer
    Dim strSql As String
    Dim rsTmp As Recordset
    
    If Len(strPara) > 0 Then
        arrPara = Split(strPara, "||")
        
        strSql = "select 名称 from 部门表 where id = [1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSql, "获取科室名称", Val(arrPara(3)))
        If rsTmp.RecordCount > 0 Then
            txtDept.Text = NVL(rsTmp!名称)
            Call CheckDept(NVL(rsTmp!名称))
        End If
    End If
    dtpEnd.Value = CDate(Format(gobjComlib.zlDatabase.Currentdate, "yyyy-mm-dd 23:59")) - 1
    dtpBegin.Value = CDate(Format(gobjComlib.zlDatabase.Currentdate, "yyyy-mm-dd 00:00")) - 1
End Sub

Private Sub chkPassive_Click()
    On Error GoTo errHandle
    
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用被动传输", IIf(chkPassive.Value, 1, 0))
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub ChkSelect_Click()
    Dim i As Long
    
    On Error GoTo errHandle
    
    For i = 1 To lvwItems.ListItems.Count
        lvwItems.ListItems(i).Checked = IIf(ChkSelect.Value = 0, False, True)
    Next
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub


Private Sub cmdDept_Click()
    On Error GoTo errHandle
    
    Me.picDept.Visible = Not Me.picDept.Visible
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdFindCancle_Click()
    On Error GoTo errHandle
    
    picDept.Visible = False
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdFindOk_Click()
    Dim i As Long
    
    On Error GoTo errHandle
    
    txtDept.Text = ""
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Checked Then
            txtDept.Text = txtDept.Text & IIf(Len(txtDept.Text) = 0, "", ";") & lvwItems.ListItems(i).Text
        End If
    Next
    
    picDept.Visible = False
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdValid_Click()
    On Error GoTo errHandle
    
    If Len(txtDept.Text) = 0 Then
        MsgBox "请先选择科室。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    
'    vsfResult.Rows = 1
'    vsfResult.Refresh
    Call ImageValid(dtpBegin.Value, dtpEnd.Value)
    
    Exit Sub
errHandle:
    cmdValid.Enabled = True
    cmdValid.Caption = "校对"
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub


Private Sub cmdView_Click()
    On Error GoTo errHandle
    
    If Len(txtDept.Text) = 0 Then
        MsgBox "请先选择科室。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mfrmResult Is Nothing Then
        Set mfrmResult = New frmResult
    End If
    
    mfrmResult.ShowMe GetDept, mstrCurValid
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub


Private Sub InitLvwList()
    Me.lvwItems.ListItems.Clear
    
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2475
        .Add , "编码", "编码", 900
        
    End With
    
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
End Sub

Private Sub LoadDept()
    Dim strSql As String
    Dim rsDept As Recordset
    Dim objItem As ListItem
    Dim arrDept() As String
    Dim i As Long
    
    strSql = "Select a.编码, a.Id,a.名称 From 部门表 a, 部门性质说明 b Where a.Id = b.部门id And b.工作性质 = '检查'"
    
    Set rsDept = gobjComlib.zlDatabase.OpenSQLRecord(strSql, "查询检查部门")

    Do While Not rsDept.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsDept!ID, rsDept!名称)
        objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
        objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = rsDept!编码
        objItem.Checked = False
        rsDept.MoveNext
    Loop
    
    Me.lvwItems.ListItems(1).Selected = True
    
    arrDept = Split(txtDept.Text, ";")
    
    For i = 0 To UBound(arrDept)
        If Len(arrDept(i)) > 0 Then
            Call CheckDept(arrDept(i))
        End If
    Next
    
End Sub

Private Sub CheckDept(strDept As String)
'根据科室名选中列表中的科室
    Dim i As Long
    
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Text = strDept Then
            lvwItems.ListItems(i).Checked = True
        End If
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picDept.Left = txtDept.Left + fraValid.Left
    picDept.Top = txtDept.Top + txtDept.Height + fraValid.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmResult = Nothing
    Set gobjComlib = Nothing
    Set gobjLogin = Nothing
    
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用被动传输", mlngPassive)
End Sub


Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo errHandle
    
    Me.lvwItems.ListItems(ColumnHeader.Index - 1).Selected = True
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub lvwItems_DblClick()
    On Error GoTo errHandle
    
'    txtDept.Text = ""
'    txtDept.Text = lvwItems.SelectedItem.Text
'    picDept.Visible = False
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub lvwItems_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errHandle
    
    Item.Selected = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errHandle
    
    Item.Selected = True
    Item.Checked = Not Item.Checked
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub SetState(BlnState As Boolean)
'    txtDept.Enabled = blnState
'    cmdDept.Enabled = blnState
'    chkReadValid.Enabled = blnState
    cmdValid.Enabled = BlnState
End Sub



Private Sub mfrmResult_OnUnload()
    Set mfrmResult = Nothing
    
    If glngState = 2 Then
        Unload Me
    End If
End Sub

Private Sub mfrmResult_OnValid(rsResult As ADODB.Recordset, lngResult As emResult, strFtpDef As String)
    Dim lngCount As Long
    Dim strTmpFile As String
    
    lngResult = DoValid(rsResult, lngCount, strTmpFile, strFtpDef, True)
End Sub

Private Function GetDept() As String
    Dim i As Long
    Dim strWhere As String
    Dim arrDept() As String
    
    arrDept = Split(txtDept.Text, ";")
    For i = 0 To UBound(arrDept)
        strWhere = strWhere & IIf(Len(strWhere) > 0, ",", "(") & "'" & arrDept(i) & "'"
    Next
    strWhere = strWhere & ")"
    
    GetDept = strWhere
End Function

