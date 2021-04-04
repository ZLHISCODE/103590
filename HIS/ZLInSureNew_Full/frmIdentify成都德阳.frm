VERSION 5.00
Begin VB.Form frmIdentify成都德阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   285
      Left            =   6360
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3540
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Left            =   4380
      MaxLength       =   25
      TabIndex        =   3
      Tag             =   "社会保障号"
      Top             =   1005
      Width           =   2265
   End
   Begin VB.CommandButton cmd验卡 
      Caption         =   "重新获取(&R)"
      Height          =   350
      Left            =   120
      TabIndex        =   22
      Top             =   4245
      Width           =   1305
   End
   Begin VB.ComboBox cbo社保 
      Height          =   300
      Left            =   855
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1005
      Width           =   2310
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5580
      TabIndex        =   24
      Top             =   4245
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   26
      Top             =   615
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -510
      TabIndex        =   25
      Top             =   3960
      Width           =   8340
   End
   Begin VB.TextBox txt病种 
      Height          =   315
      Left            =   825
      TabIndex        =   29
      Top             =   3525
      Width           =   5820
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4290
      TabIndex        =   23
      Top             =   4245
      Width           =   1100
   End
   Begin VB.Label lbl病种 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "病种(&F)"
      Height          =   180
      Left            =   165
      TabIndex        =   30
      Top             =   3630
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗性质"
      Height          =   180
      Index           =   4
      Left            =   3645
      TabIndex        =   14
      Top             =   2355
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   855
      TabIndex        =   17
      Top             =   2760
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   855
      TabIndex        =   5
      Top             =   1440
      Width           =   2310
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "医保病人基本信息显示，可以通过[重新获取]按钮重新进行读取病人基本信息。"
      Height          =   180
      Left            =   630
      TabIndex        =   27
      Top             =   360
      Width           =   6300
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify成都德阳.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "记录号"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   4
      Top             =   1485
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "社会保障号"
      Height          =   180
      Index           =   1
      Left            =   3465
      TabIndex        =   2
      Top             =   1065
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   2
      Left            =   4005
      TabIndex        =   6
      Top             =   1485
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   1905
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出身日期"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   2355
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "退休管理"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2805
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "社保机构"
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   0
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   8
      Left            =   4005
      TabIndex        =   10
      Top             =   1905
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位名称"
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3210
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗标志"
      Height          =   180
      Index           =   12
      Left            =   3645
      TabIndex        =   18
      Top             =   2805
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   4380
      TabIndex        =   7
      Top             =   1425
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   855
      TabIndex        =   9
      Top             =   1860
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   4380
      TabIndex        =   11
      Top             =   1845
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   855
      TabIndex        =   13
      Top             =   2310
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   4380
      TabIndex        =   15
      Top             =   2295
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   855
      TabIndex        =   21
      Top             =   3165
      Width           =   5775
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   4380
      TabIndex        =   19
      Top             =   2745
      Width           =   2265
   End
End
Attribute VB_Name = "frmIdentify成都德阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐

Private mlng病人ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '第一次起动系统时调用
Private mblnChange As Boolean
Private Sub cbo社保_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd验卡_Click()
   If 获取参保人员信息 = False Then
        cmd确定.Enabled = False
        Call ClearData
        Exit Sub
    End If
    Call LoadCtrlData
    cmd确定.Enabled = True
End Sub

Private Sub Form_Activate()
    '
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    
'    If 获取参保人员信息 = False Then
'        cmd确定.Enabled = False
'        Exit Sub
'    End If
'    Call LoadCtrlData
    cmd确定.Enabled = False
End Sub

Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmd确定.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim lng状态 As Long
    
    IsValid = False
    If Trim(g病人身份_成都德阳.姓名) = "" Then
        MsgBox "还没进行身份验证！", vbInformation, gstrSysName
        If cmd验卡.Enabled Then cmd验卡.SetFocus
        Exit Function
    End If
    
     If cbo社保.Text = "" Then
        ShowMsgbox "社保机构还未选择"
        Exit Function
    End If
    If g病人身份_成都德阳.保障号 = "" Then
        ShowMsgbox "请输入社会保障号!"
        Exit Function
    End If
      
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '不检查录前着态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_成都德阳, g病人身份_成都德阳.保障号)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=[1] and  医保号=[2]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_成都德阳, g病人身份_成都德阳.保障号)
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
        End If
        rsTemp.Close
        mstrReturn = mlng病人ID
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim lng疾病ID As Long
    Dim StrInput  As String, strOutput As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str社保 As String
    Dim int当前状态 As Integer
    Dim lng状态 As Long
    
    
    lng疾病ID = IIf(Val(Me.txt病种.Tag) = 0, 0, Val(Me.txt病种.Tag))
    g病人身份_成都德阳.机构编码 = Split(cbo社保.Text, "-")(0)
    
    If IsValid = False Then Exit Sub
    
    If lng疾病ID <> 0 Then
        gstrSQL = "Select * From 保险病种 where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病种", lng疾病ID)
        g病人身份_成都德阳.病种编码 = Nvl(rsTemp!编码)
        g病人身份_成都德阳.病种名称 = Nvl(rsTemp!名称)
    Else
        g病人身份_成都德阳.病种编码 = ""
        g病人身份_成都德阳.病种名称 = ""
    End If
    
    int当前状态 = 0
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=[1] and  医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", TYPE_成都德阳, g病人身份_成都德阳.保障号)
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
        End If
        rsTemp.Close
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_成都德阳
        
        strIdentify = .记录号                                '0卡号
        strIdentify = strIdentify & ";" & .保障号            '1医保号
        strIdentify = strIdentify & ";"                    '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & Decode(.性别, "1", "男", "2", "女", .性别)              '4性别
        strIdentify = strIdentify & ";" & .出生日期                '5出生日期
        strIdentify = strIdentify & ";" & ""           '6身份证
        strIdentify = strIdentify & ";" & .单位名称 & IIf(.单位编码 = 0, "", "(" & .单位编码 & ")")          '7.单位名称(编码)
        strAddition = ";" & cbo社保.ItemData(cbo社保.ListIndex)                                           '8.中心代码
        strAddition = strAddition & ";" & .医疗性质                              '9.顺序号
        strAddition = strAddition & ";"                                '10人员身份
        strAddition = strAddition & ";" & ""                 '11帐户余额
        
        strAddition = strAddition & ";" & int当前状态                            '12当前状态
        strAddition = strAddition & ";" & lng疾病ID             '13病种ID
        strAddition = strAddition & ";1"                        '14在职(1,2,3)
        strAddition = strAddition & ";" & .退休管理           '15退休证号
        strAddition = strAddition & ";" & .年龄                     '16年龄段
        strAddition = strAddition & ";"                         '17灰度级
        strAddition = strAddition & ";"                         '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_成都德阳)
    If mlng病人ID = 0 Then Exit Sub
    
    If mbytType = 0 Or mbytType = 3 Then
    Else
        
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_成都德阳 & ",'医疗标志','''" & g病人身份_成都德阳.医疗标志 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "医疗标志")
        
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_成都德阳 & ",'社保编号','''" & g病人身份_成都德阳.机构编码 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "社保编号")
    End If
    g病人身份_成都德阳.病人ID = mlng病人ID
    
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    DebugTool "进入身份验证,并开始加入基本信息"
    If Load社保机构 = False Then
        DebugTool "加入失败(身份验证)"
        Exit Function
    End If
    DebugTool "加入成功(身份验证)"
    
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    With g病人身份_成都德阳
        lblEdit(0).Caption = .记录号
        lblEdit(1).Caption = .姓名
        lblEdit(2).Caption = Decode(.性别, "1", "男", "2", "女", .性别)
        lblEdit(3).Caption = .年龄
        
        lblEdit(4).Caption = .出生日期
        lblEdit(5).Caption = .医疗性质
        
        lblEdit(6).Caption = .退休管理
        
        lblEdit(7).Caption = .医疗标志
        lblEdit(8).Caption = .单位名称 & IIf(.单位编码 <> "", "(" & .单位编码 & ")", "")
    End With
End Sub
Private Sub Form_Load()
        mblnFirst = True
End Sub

Private Function Load社保机构() As Boolean
    Dim rsTemp As New ADODB.Recordset
On Error GoTo errHand:
    gstrSQL = "Select * From 保险中心目录 where 险类=[1] and 序号<>0 order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_成都德阳)
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "不存社保机构目录，请在参数中下载机构!"
        Exit Function
    End If
    
    With rsTemp
        cbo社保.Clear
        Do While Not .EOF
            cbo社保.AddItem Nvl(!编码) & "--" & Nvl(!名称)
            cbo社保.ItemData(cbo社保.NewIndex) = Nvl(!序号, 0)
            .MoveNext
        Loop
    End With
    cbo社保.ListIndex = 0
    SetDefaultSel
'    cbo社保.Enabled = False
    Load社保机构 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SetDefaultSel() As Boolean
    Dim strReg As String
    Dim i As Integer
    
    SetDefaultSel = False
    Err = 0: On Error GoTo errHand:
    Call GetRegInFor(g公共模块, "医保", "社保机构代码", strReg)
    If cbo社保.ListCount = 0 Then Exit Function
    For i = 0 To cbo社保.ListCount
        If Split(cbo社保.List(i), "--")(0) = strReg Then
            cbo社保.ListIndex = i
            Exit For
        End If
    Next
    If cbo社保.ListIndex < 0 Then
        cbo社保.ListIndex = 0
    End If
    SetDefaultSel = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function 获取参保人员信息() As Boolean
    '获取参保人员信息
    Dim StrInput As String
    Dim strOutput As String
    Dim strArr
    
    获取参保人员信息 = False
    
    
    Err = 0
    On Error GoTo errHand:
    g病人身份_成都德阳.机构编码 = Split(cbo社保.Text, "-")(0)
    g病人身份_成都德阳.保障号 = txtEdit.Text
   
   If g病人身份_成都德阳.保障号 = "" Then
        ShowMsgbox "请输入社会保障号!"
        Exit Function
    End If
    'ASBBH   PCHAR   参保人员的社会保障号
    'ABXJGBH PCHAR   参保人员所在的保险机构编号
    
    StrInput = g病人身份_成都德阳.保障号
    StrInput = StrInput & "||" & g病人身份_成都德阳.机构编码
    
    If 业务请求_成都德阳(获得参保人员资料, StrInput, strOutput) = False Then
        Call ClearData
        Exit Function
    End If
    
    strArr = Split(strOutput, "||")
    '返回: 个人记录号||医疗性质||退休管理||单位名称||姓名||性别||出生日期（格式：YYYY-MM-DD）||单位编号||参加基本医疗标志
    
    With g病人身份_成都德阳
        .记录号 = strArr(0)
        .医疗性质 = strArr(1)
        .退休管理 = strArr(2)
        .单位名称 = strArr(3)
        
        .姓名 = strArr(4)
        .性别 = strArr(5)
        .出生日期 = strArr(6)
        .单位编码 = strArr(7)
        .医疗标志 = strArr(8)
        .年龄 = Get年龄(.出生日期)
    End With
    获取参保人员信息 = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function Get年龄(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as 年龄 from dual "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If Not rsTemp.EOF Then
        Get年龄 = Int(Nvl(rsTemp!年龄, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function
Private Sub ClearData()
    Dim i As Long
    '清除相关信息
    With g病人身份_成都德阳
        .姓名 = ""
        .性别 = ""
        .出生日期 = ""
        .机构编码 = ""
        .单位名称 = ""
        .单位编码 = ""
        .保障号 = ""
        .记录号 = ""
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    cmd验卡_Click
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m文本式
End Sub
Private Sub cmd病种_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_成都德阳
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , txt病种.Text)
    If rsTemp.State = 0 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt病种.Text = rsTemp("名称")
        txt病种.Tag = rsTemp("ID")
        zlControl.TxtSelAll txt病种
    End If
    txt病种.SetFocus
End Sub

Private Sub txt病种_Change()
    txt病种.Tag = ""
    txt病种.ForeColor = &HC0&
End Sub

Private Sub txt病种_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt病种.Text = "" Or txt病种.Tag <> "" Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strText = txt病种.Text
    gstrSQL = "Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特殊病','普通病') 类别 " & _
             "   FROM 保险病种 A WHERE A.险类=[1] And " & _
             "(A.编码 like [2] || '%' or A.编码 like [2] || '%' or A.编码 like [2] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_成都德阳, strText)
    
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_成都德阳, rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '肯定是有记录集的
        txt病种.Text = rsTemp("名称")
        txt病种.Tag = rsTemp("ID")
        SendKeys "{TAB}"
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub txt病种_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt病种.Text = ""
        txt病种.Tag = ""
    End If
End Sub

