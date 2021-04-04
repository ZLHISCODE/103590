VERSION 5.00
Begin VB.Form frmSet北京 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "frmSet北京.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame frame其它参数 
      Caption         =   "其它参数"
      Height          =   2745
      Left            =   150
      TabIndex        =   8
      Top             =   1920
      Width           =   4515
      Begin VB.TextBox txt下载目录 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   27
         Tag             =   "6"
         Top             =   1890
         Width           =   2235
      End
      Begin VB.CommandButton cmd下载目录 
         Caption         =   "…"
         Height          =   300
         Left            =   3945
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1890
         Width           =   285
      End
      Begin VB.CommandButton cmd医保项目目录 
         Caption         =   "…"
         Height          =   300
         Left            =   3945
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2280
         Width           =   285
      End
      Begin VB.TextBox txt医保项目目录 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   22
         Tag             =   "6"
         Top             =   2280
         Width           =   2235
      End
      Begin VB.CommandButton cmd医院名称 
         Caption         =   "…"
         Height          =   300
         Left            =   3945
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txt医院名称 
         Height          =   300
         Left            =   1710
         MaxLength       =   40
         TabIndex        =   10
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox Txt入参目录 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   13
         Tag             =   "6"
         Top             =   660
         Width           =   2235
      End
      Begin VB.CommandButton Cmd入参目录 
         Caption         =   "…"
         Height          =   300
         Left            =   3945
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   660
         Width           =   285
      End
      Begin VB.CommandButton cmd上传目录 
         Caption         =   "…"
         Height          =   300
         Left            =   3945
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1500
         Width           =   285
      End
      Begin VB.TextBox txt上传目录 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   19
         Tag             =   "6"
         Top             =   1500
         Width           =   2235
      End
      Begin VB.CommandButton cmd出参目录 
         Caption         =   "…"
         Height          =   300
         Left            =   3945
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1080
         Width           =   285
      End
      Begin VB.TextBox txt出参目录 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "6"
         Top             =   1080
         Width           =   2235
      End
      Begin VB.Label lbl下载目录 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "下载目录(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   28
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lbl医保项目目录 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保项目目录(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   21
         Top             =   2340
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "医院名称(&N)"
         Height          =   180
         Index           =   3
         Left            =   660
         TabIndex        =   9
         Top             =   300
         Width           =   990
      End
      Begin VB.Label Lbl入参目录 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "入参目录(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   12
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lbl上传目录 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "上传目录(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   18
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label lbl出参目录 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "出参目录(&O)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   15
         Top             =   1140
         Width           =   990
      End
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1605
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4515
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   3330
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4800
      TabIndex        =   25
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4800
      TabIndex        =   24
      Top             =   450
      Width           =   1100
   End
End
Attribute VB_Name = "frmSet北京"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
    Text医院名称 = 3
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度
Dim mblnTest As Boolean
Dim mcnTest As New ADODB.Connection

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    If Not mblnTest Then MsgBox "连接成功！", vbInformation, gstrSysName
End Sub

Private Sub Cmd入参目录_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "请指定入参目录：")
    If strPath = "" Then Exit Sub
    Txt入参目录.Text = strPath
End Sub

Private Sub Cmd出参目录_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "请指定出参目录：")
    If strPath = "" Then Exit Sub
    txt出参目录.Text = strPath
End Sub

Private Sub cmd上传目录_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "请指定入参目录：")
    If strPath = "" Then Exit Sub
    txt上传目录.Text = strPath
End Sub

Private Sub cmd下载目录_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "请指定下载目录：")
    If strPath = "" Then Exit Sub
    txt下载目录.Text = strPath
End Sub

Private Sub cmd医保项目目录_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "请指定医保项目目录：")
    If strPath = "" Then Exit Sub
    txt医保项目目录.Text = strPath
End Sub

Private Sub cmd医院名称_Click()
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If mcnTest.State = 0 Then
        mblnTest = True
        Call cmdTest_Click
        mblnTest = False
        If mcnTest.State = 0 Then Exit Sub
    End If
    
    gstrSQL = "" & _
        " SELECT A.医院编码,A.医院名称,zlSpellcode(A.医院名称) As 简码,B.编码||'-'||B.名称 AS 医院等级,C.编码||'-'||C.名称 AS 医院类型" & _
        " FROM 医院等级 A," & _
        "     (SELECT B.编码,B.名称" & _
        "     FROM 指标主表 A,指标体系对照表 B" & _
        "     WHERE A.类别=B.类别 AND A.名称='医院等级') B," & _
        "     (SELECT B.编码,B.名称" & _
        "     FROM 指标主表 A,指标体系对照表 B" & _
        "     WHERE A.类别=B.类别 AND A.名称='医院类型') C" & _
        " WHERE A.医院等级=B.编码(+) AND A.医院类型=C.编码(+) AND A.生效日期<=SYSDATE"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\保险参数设置", gstrSQL): rsTemp.Open gstrSQL, mcnTest: Call SQLTest
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有找到该医院信息，请重输！", vbInformation, gstrSysName
        txt医院名称.SetFocus
        zlControl.TxtSelAll txt医院名称
        Exit Sub
    Else
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_北京, rsTemp, "医院编码", "医院等级选择", "请选择医院等级：")
        Else
            blnReturn = True
        End If
    End If
    If blnReturn Then
        txt医院名称.Text = rsTemp!医院名称
        txt医院名称.Tag = rsTemp!医院编码
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    ElseIf KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
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
    gstrSQL = "zl_保险参数_Delete(" & TYPE_北京 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京 & ",null,'医保用户名','" & txtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京 & ",null,'医保用户密码','" & txtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京 & ",null,'医保服务器','" & txtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京 & ",null,'医院名称','" & txt医院名称.Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京 & ",null,'入参目录','" & Txt入参目录.Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京 & ",null,'出参目录','" & txt出参目录.Text & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京 & ",null,'上传目录','" & txt上传目录.Text & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京 & ",null,'下载目录','" & txt下载目录.Text & "',8)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京 & ",null,'医保项目目录','" & txt医保项目目录.Text & "',9)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '更新医院编号
    gstrSQL = "Select 名称,说明,是否禁止 From 保险类别 Where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_北京)
    '调试重庆医保银海版 204-04-07
    gstrSQL = "zl_保险类别_Update(" & TYPE_北京 & ",'" & rsTemp!名称 & "','" & IIf(IsNull(rsTemp!说明), "", rsTemp!说明) & "','" & Me.txt医院名称.Tag & "'," & IIf(IsNull(rsTemp!是否禁止), 0, rsTemp!是否禁止) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    
    gComInfo_北京.医院编码 = txt医院名称.Tag
    gComInfo_北京.入参目录 = Txt入参目录.Text
    gComInfo_北京.出参目录 = txt出参目录.Text
    
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

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

Public Function 参数设置() As Boolean
'功能：设置与东大阿尔派的医保接口
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    
    On Error GoTo errHandle
    
    '取保险参数
    gstrSQL = "select 参数名,参数值 from 保险参数 " & _
              " where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_北京)
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保用户名"
                txtEdit(text医保用户) = str参数值
            Case "医保服务器"
                txtEdit(Text医保服务器) = str参数值
            Case "医保用户密码"
                txtEdit(Text医保密码).Text = "        "    '假密码
                txtEdit(Text医保密码).Tag = str参数值
            Case "医院名称"
                txt医院名称.Text = str参数值
            Case "入参目录"
                Txt入参目录.Text = str参数值
            Case "出参目录"
                txt出参目录.Text = str参数值
            Case "上传目录"
                txt上传目录.Text = str参数值
            Case "下载目录"
                txt下载目录.Text = str参数值
            Case "医保项目目录"
                txt医保项目目录.Text = str参数值
        End Select
        rsTemp.MoveNext
    Loop
    
    '取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_北京)
    If Not rsTemp.EOF Then txt医院名称.Tag = Nvl(rsTemp!医院编码)
    
    mblnChange = False
    frmSet北京.Show vbModal, frm医保类别
    
    参数设置 = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt医院名称_GotFocus()
    Call zlControl.TxtSelAll(txt医院名称)
End Sub

Private Sub txt医院名称_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrInput As String
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    StrInput = UCase(Trim(txt医院名称.Text))
    If Trim(StrInput) = "" Then Exit Sub
    
    If mcnTest.State = 0 Then
        mblnTest = True
        Call cmdTest_Click
        mblnTest = False
        If mcnTest.State = 0 Then Exit Sub
    End If
    
    gstrSQL = "SELECT * FROM (" & _
        " SELECT A.医院编码,A.医院名称,zlSpellcode(A.医院名称) As 简码,B.编码||'-'||B.名称 AS 医院等级,C.编码||'-'||C.名称 AS 医院类型" & _
        " FROM 医院等级 A," & _
        "     (SELECT B.编码,B.名称" & _
        "     FROM 指标主表 A,指标体系对照表 B" & _
        "     WHERE A.类别=B.类别 AND A.名称='医院等级') B," & _
        "     (SELECT B.编码,B.名称" & _
        "     FROM 指标主表 A,指标体系对照表 B" & _
        "     WHERE A.类别=B.类别 AND A.名称='医院类型') C" & _
        " WHERE A.医院等级=B.编码(+) AND A.医院类型=C.编码(+) AND A.生效日期<=SYSDATE) A" & _
        " WHERE (A.医院编码 Like '" & StrInput & "%' Or A.医院名称 Like '" & StrInput & "%' Or A.简码 Like '" & StrInput & "%')"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\保险参数设置", gstrSQL): rsTemp.Open gstrSQL, mcnTest: Call SQLTest
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有找到该医院信息，请重输！", vbInformation, gstrSysName
        txt医院名称.SetFocus
        zlControl.TxtSelAll txt医院名称
        Exit Sub
    Else
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_北京, rsTemp, "医院编码", "医院等级选择", "请选择医院等级：")
        Else
            blnReturn = True
        End If
    End If
    If blnReturn Then
        txt医院名称.Text = rsTemp!医院名称
        txt医院名称.Tag = rsTemp!医院编码
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

