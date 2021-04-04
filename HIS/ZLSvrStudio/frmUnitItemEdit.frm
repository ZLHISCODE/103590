VERSION 5.00
Begin VB.Form frmUnitItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医院信息项目维护"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4845
   Icon            =   "frmUnitItemEdit.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton optType 
      Caption         =   "图片"
      Height          =   180
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.OptionButton optType 
      Caption         =   "文本"
      Height          =   180
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Frame fraInfo 
      Height          =   120
      Left            =   0
      TabIndex        =   11
      Top             =   2100
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3525
      TabIndex        =   10
      Top             =   2275
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2340
      TabIndex        =   9
      Top             =   2275
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1290
      Width           =   2625
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   2
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lblNoteNo 
      AutoSize        =   -1  'True
      Caption         =   "可选编号:003"
      Height          =   180
      Left            =   2880
      TabIndex        =   3
      Top             =   900
      Width           =   1080
   End
   Begin VB.Label lblMarks 
      BackStyle       =   0  'Transparent
      Caption         =   "该功能用于医院信息项目的定义、调整。项目名称与项目类型的调整会影响数据的适用。"
      Height          =   390
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   4590
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   180
      Picture         =   "frmUnitItemEdit.frx":6852
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "项目类型"
      Height          =   180
      Left            =   1080
      TabIndex        =   6
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "项目名称"
      Height          =   180
      Left            =   1080
      TabIndex        =   4
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "项目编码"
      Height          =   180
      Left            =   1080
      TabIndex        =   1
      Top             =   900
      Width           =   720
   End
End
Attribute VB_Name = "frmUnitItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mstrNO As String '项目编码
Private mstrName As String '项目名称
Private mintType As Integer '项目类型
Private mblnChange As Boolean
Private mstrRemarks As String
'系统固定适用的项目
Private Const SYS_ITEMS = "版本号,服务器目录,访问用户,访问密码,收集目录,收集类型,站点编号,站点数量,站点类型,消息集成平台客户端," & _
                          "客户端升级日期,收集目录S,访问用户S,访问密码S,收集目录F,访问用户F,访问密码F,访问端口F,收集方式," & _
                          "管理员账号,管理员密码,客户端预升级时间点,管理员,验证码,注册码,发行码,升级类型,授权证章,授权工具," & _
                          "授权邮戳,站点编号,产品标题,支持商简名,产品简名,授权性质,单位名称,产品开发商,技术支持商,支持商简名," & _
                          "支持商MAIL,支持商URL,支持商BBS,授权站点,使用期限,授权日期,影像DICOM设备数量,影像视频设备数量," & _
                          "影像胶片打印机数量,影像观片站数量,检验仪器数量"
'系统适用的多级项目
Private Const SYS_ITEMS_EXTEND = "服务器目录[n],访问用户[n],访问密码[n],FTP服务器[n],FTP用户[n],FTP密码[n],FTP端口[n]"

Private Enum UnitCol
    Col_编码 = 0
    Col_项目 = 1
    Col_是否图片 = 2
    Col_内容 = 3
    Col_Edit = 4
    Col_Del = 5
    Col_是否改变 = 6
End Enum

'===========================================================================
'==公共接口
'===========================================================================
Public Function ShowMe(Optional ByRef strNo As String, Optional ByRef strName As String, Optional ByRef intType As Integer) As Boolean
'功能：项目编辑设置设置
'     strNo=编辑的编码，为空表示新增
'返回：是否产生了编辑
'     strNo=编辑后的编码
'     strName=编辑后的名称
'     intType=编辑后的类型
    mblnOk = False
    mstrNO = strNo
    mstrName = strName
    mintType = intType
    Me.Show vbModal
    strNo = mstrNO
    strName = mstrName
    intType = mintType
    ShowMe = mblnOk
End Function
'===========================================================================
'==事件
'===========================================================================
Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mstrNO <> "" Then
        '验证身份并输入操作说明
        If Not CheckAuditStatus("0312", "调整项目", mstrRemarks) Then Exit Sub
    End If
    If ValiData() Then
        If mblnChange Then
            If Not SaveData() Then
                Exit Sub
            End If
            mblnOk = True
        End If
    Else
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Call LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrRemarks = ""
End Sub

Private Sub OptType_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtName_Change()
    mblnChange = True
End Sub

Private Sub txtName_GotFocus()
    SelAll txtName
End Sub

Private Sub txtNO_Change()
    mblnChange = True
End Sub

Private Sub txtNO_GotFocus()
 SelAll txtNO
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

'===========================================================================
'==私有方法
'===========================================================================
Private Sub LoadData()
'功能：数据加载
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngMax As Long, strTmp As String
    Dim strNote As String
    
    On Error GoTo errH
    If mstrNO <> "" Then
        txtNO.Enabled = False
        txtNO.BackColor = Me.BackColor
        '新增项目
        strSQL = "Select 编码, 名称, 是否图片 From Zlunitinfoitem Where 编码 = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mstrNO)
        txtNO.Text = rsTmp!编码
        txtNO.Tag = rsTmp!编码
        txtName.Text = rsTmp!名称
        txtName.Tag = rsTmp!名称
        optType(Val(rsTmp!是否图片 & "")).value = True
        lblType.Tag = Val(rsTmp!是否图片 & "")
        lblNoteNo.Visible = False
    Else
        txtNO.Enabled = True
        txtNO.BackColor = &H80000005
        strSQL = "Select Max(Lpad(编码, 3, '0')) 最大编码 From Zlunitinfoitem"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mstrNO)
        lngMax = Val(rsTmp!最大编码 & "")
        '已经达到最大编码，寻找编码空隙
        If lngMax = 999 Then
            strSQL = "Select b.编码" & vbNewLine & _
                    "From (Select Lpad(编码, 3, '0') 编码 From Zlunitinfoitem) a," & vbNewLine & _
                    "     (Select Lpad(Rownum || '', 3, '0') 编码 From Dual Connect By Rownum < [1]) b" & vbNewLine & _
                    "Where a.编码(+) = b.编码 And a.编码 Is Null"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, lngMax)
            If Not rsTmp.EOF Then
                If rsTmp.RecordCount <> 0 Then
                    strTmp = rsTmp!编码
                    txtNO.Text = rsTmp!编码
                    If rsTmp.RecordCount > 0 Then
                        rsTmp.MoveNext
                        strTmp = strTmp & "," & rsTmp!编码
                    End If
                End If
            End If
            lblNoteNo.Visible = True
            If strTmp <> "" Then
                lblNoteNo.Caption = "可选编号:" & strTmp
            Else
                lblNoteNo.Caption = "无法产生编号,请手工指定二位或一位编码"
            End If
        Else
            lblNoteNo.Caption = "可选编号:大于" & Lpad(lngMax & "", 3, "0")
            txtNO.Text = Lpad((lngMax + 1) & "", 3, "0")
        End If
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function ValiData() As Boolean
'功能：进行数据校验
    Dim intType As Integer, strName As String, strNo As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intTmp As Integer
    
    On Error GoTo errH
    intType = IIf(optType(0).value, 0, 1)
    strName = Trim(txtName.Text)
    strNo = Trim(txtNO.Text)
    If mstrNO = "" Then
        If strNo = "" Then
            MsgBox "请输入项目编码。", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        If Not IsNumeric(strNo) Then
            MsgBox "项目编码必须为数值类型，请重新输入。", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        If ActualLen(strNo) > txtNO.MaxLength Then
            MsgBox "项目编码超过" & txtNO.MaxLength & "位长度，请重新输入。", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        
    Else
        '数据未发生改变
        If intType = Val(lblType.Tag) And txtName.Tag = strName And txtNO.Tag = strNo Then
            mblnChange = False
            ValiData = True
            Exit Function
        End If
    End If
    If Trim(txtName.Text) = "" Then
        MsgBox "请输入项目名称。", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Function
    End If
    
    If txtName.Tag <> strName Then
        If ActualLen(strName) > txtName.MaxLength Then
            MsgBox "项目名称超过" & txtName.MaxLength & "位长度，请重新输入。", vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
    
        '是系统项目
        If InStr("," & SYS_ITEMS & ",", "," & strName & ",") > 0 Then
            MsgBox "该名称是系统固定项目，请换用其他名称。" & vbNewLine & "系统项目：" & SYS_ITEMS, vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
        '名称加编号方式，判读你是否是系统扩展项目
        If Not IsNumeric(strName) Then
            intTmp = ValEx(strName)
            If strName Like "*" & intTmp Then
                If InStr("," & SYS_ITEMS_EXTEND & ",", "," & Mid(strName, 1, Len(strName) - Len(intTmp & "")) & "[n]" & ",") > 0 Then
                    MsgBox "该名称是系统固定项目，请换用其他名称。" & vbNewLine & "系统项目：" & SYS_ITEMS_EXTEND, vbInformation, gstrSysName
                    txtName.SetFocus
                    Exit Function
                End If
            End If
        End If
        '是否该名称已经存在
        strSQL = "Select 1 From Zlunitinfoitem Where 名称 ='" & strName & "'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            MsgBox "该名称已经被使用，请换用其他名称。", vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
        
        '是否是系统级使用的项目
        strSQL = "Select 1 From Zlreginfo Where 项目 = '" & strName & "' And Not Exists (Select 1 From Zlunitinfoitem Where 名称 = '" & strName & "')"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            MsgBox "该名称是系统固定项目，请换用其他名称。", vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
    End If
    If mstrNO = "" Then
        '是否该编码已经存在
        strSQL = "Select 1 From Zlunitinfoitem Where 名称 ='" & strNo & "'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            MsgBox "该编码已经被使用，请换用其他名称。", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
    Else
        '判断是否存在数据
        If Val(lblType.Tag) = 0 Then
            strSQL = "Select 1 From Zlreginfo Where 项目 = '" & txtName.Tag & "'"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        Else
            strSQL = "Select 1 From Zlunitinfoimage Where 项目 = '" & txtName.Tag & "'"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        End If
        If Not rsTmp.EOF Then
            '类型改变
            If intType <> Val(lblType.Tag) Then
                If MsgBox("项目类型发生改变，以前的数据会被清空。是否继续？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                    txtNO.SetFocus
                    Exit Function
                End If
            ElseIf txtName.Tag <> strName Then
                If MsgBox("项目名称发生改变，会对该项目的适用产生影响。是否继续？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                    txtNO.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    mblnChange = True
    ValiData = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Function SaveData() As Boolean
'功能：进行数据保存
    Dim intType As Integer, strName As String, strNo As String
    Dim strSQL As String
    
    On Error GoTo errH
    intType = IIf(optType(0).value, 0, 1)
    strName = Trim(txtName.Text)
    strNo = Trim(txtNO.Text)
    strSQL = "Zltools.b_Public.Zlunitinfoitemchange(" & IIf(mstrNO = "", 0, 1) & ",'" & strNo & "','" & strName & "'," & intType & ")"
    Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
    If mstrNO = "" Then
        '插入重要操作日志
        Call SaveAuditLog(1, "新增项目", strName)
    Else
        '插入重要操作日志
        Call SaveAuditLog(2, "调整项目", "由“" & mstrName & "”调整为“" & strName & "”", mstrRemarks)
    End If
    
    mstrNO = strNo
    mstrName = strName
    mintType = intType
    SaveData = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function



