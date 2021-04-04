VERSION 5.00
Begin VB.Form frmIdentify大连 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmIdentify大连.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Txt诊断摘要 
      Height          =   300
      Left            =   4635
      TabIndex        =   49
      Top             =   3990
      Width           =   2500
   End
   Begin VB.TextBox Txt最高限额 
      Height          =   300
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   29
      Top             =   2910
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   285
      Left            =   6885
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3660
      Width           =   255
   End
   Begin VB.TextBox txt起付线 
      Enabled         =   0   'False
      Height          =   300
      Left            =   5730
      TabIndex        =   33
      Top             =   3285
      Width           =   1425
   End
   Begin VB.TextBox Txt疾病 
      Height          =   300
      Left            =   1080
      TabIndex        =   38
      Top             =   4035
      Width           =   2715
   End
   Begin VB.CommandButton cmd验卡 
      Caption         =   "读卡(&R)"
      Height          =   350
      Left            =   180
      TabIndex        =   41
      Top             =   4605
      Width           =   1100
   End
   Begin VB.ComboBox cbo中心 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   4035
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox Txt转诊单号 
      Height          =   300
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   31
      Top             =   3270
      Width           =   3195
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -30
      TabIndex        =   47
      Top             =   630
      Width           =   8340
   End
   Begin VB.ComboBox cbo就诊分类 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   2910
      Width           =   6075
   End
   Begin VB.CheckBox chk工伤 
      Caption         =   "工伤可用(&G)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5820
      TabIndex        =   45
      Top             =   2490
      Width           =   1305
   End
   Begin VB.CheckBox chk生育 
      Caption         =   "生育可用(&S)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4470
      TabIndex        =   44
      Top             =   2490
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6030
      TabIndex        =   43
      Top             =   4605
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4650
      TabIndex        =   42
      Top             =   4605
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -120
      TabIndex        =   46
      Top             =   4425
      Width           =   8340
   End
   Begin VB.TextBox txt病种 
      Height          =   300
      Left            =   1080
      TabIndex        =   35
      Top             =   3645
      Width           =   6060
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   13
      Left            =   6675
      TabIndex        =   52
      Top             =   960
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   18
      Left            =   6255
      TabIndex        =   51
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "就诊摘要"
      Height          =   375
      Left            =   3870
      TabIndex        =   50
      Top             =   4065
      Width           =   750
   End
   Begin VB.Label lbl病种 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "病种(&F)"
      Height          =   180
      Left            =   420
      TabIndex        =   34
      Top             =   3690
      Width           =   630
   End
   Begin VB.Label lbl起付线 
      AutoSize        =   -1  'True
      Caption         =   "起付线(&Q)"
      Enabled         =   0   'False
      Height          =   180
      Left            =   4920
      TabIndex        =   32
      Top             =   3345
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "诊断情况(&M)"
      Height          =   180
      Index           =   16
      Left            =   60
      TabIndex        =   37
      Top             =   4095
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保中心(&Z)"
      Height          =   180
      Index           =   15
      Left            =   60
      TabIndex        =   39
      Top             =   4095
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "转诊单号(&N)"
      Height          =   180
      Index           =   14
      Left            =   60
      TabIndex        =   30
      Top             =   3330
      Width           =   990
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "通过IC卡验证人员身份，并将验证结果信息显示出来，同时可对就诊分类进行选择。"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   390
      Width           =   6540
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   30
      Picture         =   "frmIdentify大连.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "就诊分类(&J)"
      Height          =   180
      Index           =   13
      Left            =   60
      TabIndex        =   27
      Top             =   2970
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户状态"
      Height          =   180
      Index           =   12
      Left            =   2190
      TabIndex        =   25
      Top             =   2587
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   12
      Left            =   2940
      TabIndex        =   26
      Top             =   2527
      Width           =   1335
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   11
      Left            =   1080
      TabIndex        =   24
      Top             =   2527
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "参保类别"
      Height          =   180
      Index           =   11
      Left            =   330
      TabIndex        =   23
      Top             =   2587
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   10
      Left            =   5820
      TabIndex        =   22
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "补助个人帐户余额"
      Height          =   180
      Index           =   10
      Left            =   4350
      TabIndex        =   21
      Top             =   2190
      Width           =   1440
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   9
      Left            =   2940
      TabIndex        =   20
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "月缴基数"
      Height          =   180
      Index           =   9
      Left            =   2190
      TabIndex        =   19
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   8
      Left            =   1080
      TabIndex        =   18
      Top             =   2130
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "统筹累计"
      Height          =   180
      Index           =   8
      Left            =   330
      TabIndex        =   17
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   5820
      TabIndex        =   16
      Top             =   1740
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "基本个人帐户余额"
      Height          =   180
      Index           =   7
      Left            =   4350
      TabIndex        =   15
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   2940
      TabIndex        =   14
      Top             =   1740
      Width           =   1335
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   2940
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "就医类别"
      Height          =   180
      Index           =   6
      Left            =   2190
      TabIndex        =   13
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   1080
      TabIndex        =   12
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "治疗序号"
      Height          =   180
      Index           =   5
      Left            =   330
      TabIndex        =   11
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   4
      Left            =   5820
      TabIndex        =   10
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "IC卡号"
      Height          =   180
      Index           =   4
      Left            =   5250
      TabIndex        =   9
      Top             =   1410
      Width           =   540
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   1080
      TabIndex        =   8
      Top             =   1350
      Width           =   3195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   3
      Left            =   330
      TabIndex        =   7
      Top             =   1410
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   5400
      TabIndex        =   6
      Top             =   960
      Width           =   435
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   2
      Left            =   4980
      TabIndex        =   5
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   1
      Left            =   2550
      TabIndex        =   3
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "个人编号"
      Height          =   180
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "最高限额(&K)"
      Height          =   180
      Index           =   17
      Left            =   45
      TabIndex        =   48
      Top             =   2970
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "frmIdentify大连"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99

Dim mblnFirst  As Boolean
Dim mstrReturn As String    '返回信息串
Dim mlng病人ID As Long
'mbytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
Dim mbytType As Byte
Dim mblnOK As Boolean
Dim mlng性质 As Long
Dim mlng记录ID As Long
Dim mbytCallType As Byte  '(1-结帐处调用;0-病人费用查询调用的)对虚拟结算有效即byttype=4的情况
Dim mint险类 As Integer

Private Sub cbo就诊分类_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:数据验证
    '--入参数:
    '--出参数:
    '--返  回:验证成功返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    IsValid = False
    
    If LenB(StrConv(Trim(Txt转诊单号.Text), vbFromUnicode)) > 6 Then
        ShowMsgbox "转诊单号超长了,最多能输入6个字符!"
        If Txt转诊单号.Enabled Then Txt转诊单号.SetFocus
        Exit Function
    End If
    If Me.cbo就诊分类.ListIndex < 0 Then
        ShowMsgbox "就诊分类必需选择!"
        If cbo就诊分类.Enabled Then cbo就诊分类.SetFocus
        Exit Function
    End If
    If Me.cbo中心.ListIndex < 0 Then
        ShowMsgbox "中心必需选择!"
        If cbo中心.Enabled Then cbo中心.SetFocus
        Exit Function
    End If
    If Txt转诊单号.Text <> "" And mbytType <> 0 And mbytType <> 3 Then
        If Val(txt起付线.Text) = 0 Then
            Dim blnYes As Boolean
            
            ShowMsgbox "起付线未输入,是否忽略此项?", True, blnYes
            
            If blnYes = False Then
                If txt起付线.Enabled Then txt起付线.SetFocus
                Exit Function
            End If
        End If
    End If
    Dim lng分类 As Long
     lng分类 = cbo就诊分类.ItemData(cbo就诊分类.ListIndex)
     If (lng分类 = 3 Or lng分类 = 4) And Trim(Txt疾病.Text) = "" Then
        ShowMsgbox "大病或慢病必须输入诊断情况!"
        Exit Function
     End If
    '检查用户状态
    '   A正常、B半止付、C全止付、D销户
    'mbytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    Select Case g病人身份_大连.帐户状态
        Case "A"
        Case "B"
            If mbytType = 4 Then
            ShowMsgbox "该病人状态为“半止付”状态,只能在门诊使用!"
            End If
        Case "C"
            ShowMsgbox "该病人状态为“全止付”状态,只能用现金结算!"
        Case "D"
            ShowMsgbox "该病人已医保中心销户,不能继续!"
            Exit Function
    End Select
    
    '检查病人状态
    Dim lng病人ID As Long
    gstrSQL = "select 病人id,nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类, g病人身份_大连.个人编号)
    If mbytType <> 4 Then   '不是住院结算时，需验证当前状态
        If rsTemp.RecordCount > 0 Then
            If rsTemp("状态") > 0 Then
                MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        '2005-01-04 ZHQ修改
        '功能：如果医保病人未挂号，不允许收费
        If rsTemp.RecordCount > 0 Then
            lng病人ID = Nvl(rsTemp.Fields("病人ID").Value, 0)
        Else
            lng病人ID = 0
        End If
        If mbytType = 0 Then
            Dim lngRegDay As Long   '病人挂号允许的天数
            #If gverControl >= 4 Then
                lngRegDay = Val(zlDatabase.GetPara(21, glngSys, , "0"))
            #Else
                lngRegDay = Val(GetPara(21, glngSys, , , "0"))
            #End If
            
            If lngRegDay <> 0 Then  '=0表示不进行判断
                #If gverControl >= 5 Then
                    gstrSQL = "Select No,门诊号,病人ID From 病人挂号记录 " & _
                            "  Where sysdate-登记时间<=" & lngRegDay & _
                            "  And 记录性质=1 And 记录状态=1 And 病人ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "求实际挂号记录", lng病人ID)
                #Else
                    gstrSQL = "Select No,门诊号,病人ID From 病人挂号记录 " & _
                            "  Where sysdate-登记时间<=[1] And 病人ID=[2]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "求实际挂号记录", lngRegDay, lng病人ID)
                #End If
                If rsTemp.RecordCount <= 0 Then
                    MsgBox "此病人 " & lngRegDay & " 天内未挂号，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '住院结算时,需处理是否为同一个人
        If rsTemp.EOF Then
            ShowMsgbox "在保险帐户中不存在当前病人!"
            Exit Function
        Else
            lng病人ID = Nvl(rsTemp!病人ID, 0)
            If mlng病人ID <> lng病人ID Then
                ShowMsgbox "虚假结帐的当前病人与身份验证的病人不一致!"
                Exit Function
            End If
        End If
    End If
    If Txt疾病.Tag = "" And Txt疾病.Text <> "" And mbytType <> 3 Then
        ShowMsgbox "选择的诊断情况有误,不能继续!"
        Exit Function
    End If
    
    
    If cbo就诊分类.ItemData(cbo就诊分类.ListIndex) = 3 Then
        '当就诊分类为“门诊大病”时，需输入病种
        '20040621刘兴宏加入
        If Val(txt病种.Tag) = 0 Then
            ShowMsgbox "门诊大病必须输入病种!"
            If txt病种.Enabled And txt病种.Visible Then txt病种.SetFocus
            Exit Function
        End If
    End If
    IsValid = True
End Function

Private Sub cbo中心_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mstrReturn = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strTmp1 As String
    
    If mlng病人ID <> 0 And mbytType > 4 Then
        With g病人身份_大连
             '保存附加信息
             '参数:
             '   性制_IN,记录id_IN,最高限额_IN
             
             gstrSQL = "zl_保险结算记录限额_Update(" & _
                 mlng性质 & "," & _
                 mlng记录ID & "," & _
                 Val(Txt最高限额.Text) & ")" & _
                 ""
                 Err = 0
                 On Error Resume Next
                 gcnOracle.Execute gstrSQL, Me.Caption
                 If Err <> 0 Then
                     ShowMsgbox "保险结算记录的最高限额保存失败!"
                     Exit Sub
                 End If
         End With
         mstrReturn = mlng病人ID
        Unload Me
        Exit Sub
    End If
    '验证数据
    If IsValid = False Then Exit Sub
    
    With g病人身份_大连
        .转诊单号 = Trim(Txt转诊单号)
        .就诊分类 = cbo就诊分类.ItemData(cbo就诊分类.ListIndex)
        
        If Txt疾病.Tag = "" Then
            .诊断编码 = ""
            .诊断名称 = ""
        Else
            .诊断编码 = Split(Txt疾病.Tag, "|||")(0)
            .诊断名称 = Split(Txt疾病.Tag, "|||")(1)
        End If
        If mbytType = 0 Then
            .起付线 = 0
        Else
            If .转诊单号 = "" Then
                '读取起付线
                If Val(txt起付线.Text) = 0 Then
                    .起付线 = Get起付线(.职工就医类别, .年龄, mint险类)
                Else
                    .起付线 = Val(txt起付线.Text)
                End If
            Else
                .起付线 = Val(txt起付线.Text)
            End If
            
        End If
    End With
    
        
    '确定相关返回串
    
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计

    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8《病人ID》
    '9中心;10.顺序号;11人员身份;12帐户余额;13当前状态;14病种ID;15在职(0,1);16退休证号;17年龄段;18灰度级
    '19帐户增加累计,20帐户支出累计,21进入统筹累计,22统筹报销累计,23住院次数累计;24就诊类型 (1、急诊门诊);25开单科室
    Dim int当前状态 As Integer, strUnitName As String
    int当前状态 = 0
    If mbytType = 3 Or mbytCallType = 0 Then
        '如果是挂号,则需确定该用户是否存在
        'mbytCallType (1-结帐处调用;0-病人费用查询调用的)对虚拟结算有效即byttype=4的情况
        '需确定当前状态,因为当前状态是不能改变的.(周顺利提出:2004/06/11)
        gstrSQL = "Select a.病人ID,a.当前状态,b.工作单位 from 保险帐户 a,病人信息 b " & _
                "  Where a.病人ID=b.病人ID And a.险类=" & mint险类 & " and  a.医保号='" & g病人身份_大连.个人编号 & "'"
        Dim rsTemp As New ADODB.Recordset
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
            strUnitName = Nvl(rsTemp!工作单位, "")
        End If
    End If
     
     mstrReturn = ""
    With g病人身份_大连
        strTmp = .IC卡号                    '0卡号
        strTmp = strTmp & ";" & .个人编号   '1医保号
        strTmp = strTmp & ";"               '2密码
        strTmp = strTmp & ";" & .姓名       '3姓名
        strTmp = strTmp & ";" & .性别       '4性别
        strTmp = strTmp & ";" & .出生日期   '5出生日期
        strTmp = strTmp & ";" & .身份证号   '6身份证
        strTmp = strTmp & ";" & strUnitName '7单位名称(编码)
        
        strTmp1 = ""
        strTmp1 = strTmp1 & ";"    '8中心代码
        strTmp1 = strTmp1 & ";" & .治疗序号   '9顺序号
        strTmp1 = strTmp1 & ";" & .转诊单号  '10人员身份,存的是转诊单号
        strTmp1 = strTmp1 & ";" & .基本个人帐户余额       '11帐户余额
        strTmp1 = strTmp1 & ";" & int当前状态               '12当前状态
        strTmp1 = strTmp1 & ";" & IIf(Val(Me.txt病种.Tag) = 0, "", Me.txt病种.Tag)             '13病种ID
        '刘兴宏:20040911,加入了退老
        '医保中心为,A在职、B退休、L离休、T特诊,Q 企业公费,E退老
        strTmp1 = strTmp1 & ";" & Decode(.职工就医类别, "A", 1, "B", 2, "L", 3, "T", 4, "Q", 5, "E", 6, 1) '.就诊分类  '14在职(0,1)
        strTmp1 = strTmp1 & ";" & .补助个人帐户余额 '15退休证号,目前我存的是补助个人帐户余额
        strTmp1 = strTmp1 & ";" & IIf(.年龄 = 0, "", .年龄) '16年龄段
        strTmp1 = strTmp1 & ";" & .就诊分类       '17灰度级,存的就诊分类编码
        strTmp1 = strTmp1 & ";" & .基本个人帐户余额         '18帐户增加累计
        strTmp1 = strTmp1 & ";0"        '19帐户支出累计
        strTmp1 = strTmp1 & ";" & .统筹累计  '20进入统筹累计
        strTmp1 = strTmp1 & ";" & .起付线          '21统筹报销累计
        strTmp1 = strTmp1 & ";0"        '22住院次数累计
    End With
    
    '--------------------------------------------------------------------------
    '2004-06-08,取消挂号项目的限制,需保存相关的病人信息.
    'If mlng病人ID <> 0 And mbytType = 3 Then
        
    'Else
        mlng病人ID = BuildPatiInfo(0, strTmp & strTmp1, mlng病人ID, mint险类)
   ' End If
    
    With g病人身份_大连
        '保存附加信息
        '参数:
        '    险类_IN,病人id_IN,参保类别1_IN,参保类别2_IN,参保类别3_IN,参保类别4_IN,参保类别5_IN,
        gstrSQL = "zl_保险帐户附加_Update(" & _
            mint险类 & "," & _
            mlng病人ID & "," & _
            Val(.参保类别1) & "," & _
            Val(.参保类别2) & "," & _
            Val(.参保类别3) & "," & _
            Val(.参保类别4) & "," & _
            Val(.参保类别5) & ")" & _
            ""
            Err = 0
            On Error Resume Next
            gcnOracle.Execute gstrSQL, Me.Caption
            If Err <> 0 Then
                ShowMsgbox "保险帐户附加信息保存失败,可能有些信息不能正常使用!"
                Exit Sub
            End If
            Dim str医疗付款方式 As String
            
            If InStr(1, "ABE", .职工就医类别) <> 0 Then
                'A在职,B.退休,E（新增的)
                str医疗付款方式 = "社会基本医疗保险"
            End If
            
            If .医保中心 = 1 Then
                If Val(.参保类别4) = 1 Then
                    '0生育不可用、1生育可用
                    str医疗付款方式 = "生育保险"
                End If
                If Val(.参保类别5) = 1 Then
                    '0工伤不可用、1工伤可用
                    str医疗付款方式 = "工伤保险"
                End If
                If InStr(1, "LT", .职工就医类别) <> 0 Then
                    str医疗付款方式 = "公费医疗"
                End If
                If InStr(1, "LT", .职工就医类别) <> 0 Then
                    'T.特诊,L.离休
                    str医疗付款方式 = "公费医疗"
                End If
                If .职工就医类别 = "Q" Then
                    '企业公费
                    str医疗付款方式 = "企业离休"
                End If
            Else
                If mint险类 = TYPE_大连开发区 Then
                    str医疗付款方式 = "社会基本医疗保险"
                End If
            End If
            Err = 0
            On Error GoTo errHand:
            '更新病人信息的医疗付款方式
            gstrSQL = "zl_病人信息医疗付款_Update(" & mlng病人ID & ",'" & _
                str医疗付款方式 & "')"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    '返回格式:中间插入病人ID
   '--身份验证后强制对病人信息的险类进行一次更新,防止险类出现空
   gstrSQL = "update 病人信息 A set A.险类=(select 险类 from 保险帐户 where 病人id=A.病人id) where 病人id=" & mlng病人ID
   gcnOracle.Execute gstrSQL
    
    If mlng病人ID > 0 Then
        mstrReturn = strTmp & ";" & mlng病人ID & strTmp1
    End If
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd病种_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & mint险类
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , txt病种.Text)
    If rsTemp.State = 0 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt病种.Text = rsTemp("名称")
        txt病种.Tag = rsTemp("ID")
        zlControl.TxtSelAll txt病种
    End If
    txt病种.SetFocus
End Sub

Private Sub cmd验卡_Click()
    SetCtlEn False
    mblnOK = ReadCard
    
    '加入年龄计算
    '周海全  2005-02-17
    Dim dtAge As Variant
    If Len(Trim(lblEdit(3).Caption)) <> 0 Then
        Select Case Len(Trim(lblEdit(3).Caption))
        Case 15
            dtAge = "19" & Substr(Trim(lblEdit(3).Caption), 7, 6)
        Case 18
            dtAge = Substr(Trim(lblEdit(3).Caption), 7, 8)
        Case Else
            dtAge = Null
        End Select
        If IsNull(dtAge) Then
            lblEdit(13).Caption = ""
        Else
            lblEdit(13).Caption = CInt(Format(zlDatabase.Currentdate, "yyyy")) - CInt(Substr(dtAge, 1, 4))
            If Format(zlDatabase.Currentdate, "MMdd") < Substr(dtAge, 5, 4) Then
                lblEdit(13).Caption = CInt(lblEdit(13).Caption) - 1
            End If
            If CInt(lblEdit(13).Caption) < 0 Then lblEdit(13).Caption = 0
        End If
    End If
    
    If Txt转诊单号 = "" And mbytType <> 0 Then
        Me.txt起付线 = Format(Get起付线(g病人身份_大连.职工就医类别, g病人身份_大连.年龄, mint险类), "###,###0.00;-###,###0.00; ;")
    End If
    SetCtlEn True
    If Txt疾病.Enabled Then
        Txt疾病.SetFocus
    ElseIf cmdOK.Enabled Then
        cmdOK.SetFocus
    End If
End Sub
Private Sub SetCtlEn(ByVal blnTrue As Boolean)
    cmd验卡.Enabled = blnTrue
    cmdOK.Enabled = blnTrue And mblnOK
    txt起付线.Enabled = blnTrue And mbytType = 4
    lbl起付线.Enabled = blnTrue And mbytType = 4
    cmdCancel.Enabled = blnTrue
    Txt转诊单号.Enabled = blnTrue And mbytType <> 3
    cbo就诊分类.Enabled = blnTrue And mbytType <> 3
    txt病种.Enabled = blnTrue And mbytType <> 3
    cmd病种.Enabled = blnTrue And mbytType <> 3
    Txt疾病.Enabled = blnTrue And mbytType = 0 And mblnOK
    lbl(16).Enabled = blnTrue And mbytType = 0 And mblnOK
    cbo中心.Enabled = blnTrue
    Txt诊断摘要.Enabled = blnTrue And mblnOK
    Txt诊断摘要.Locked = True
        
End Sub
Private Sub Form_Activate()
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnOK = False
    Txt疾病.Tag = ""
    Txt疾病.Text = ""
    txt起付线.Text = ""
    SetCtlEn True
    If cbo就诊分类.Enabled And mlng病人ID = 0 Then
        cbo就诊分类.SetFocus
    ElseIf Txt最高限额.Visible Then
        Txt最高限额.SetFocus
    End If
End Sub
Private Function ReadCard() As Boolean
    
    ReadCard = False
   '验证用户身份
    If 读取病人身份_大连(IIf(mint险类 = TYPE_大连开发区, 2, 1), mint险类) = False Then
        Exit Function
    End If
    Call SetCtlData
    ReadCard = True
End Function
Private Function SetCtlData()
    '功能:设置控件数据
    Dim int性别 As Integer
    Dim int划价天数 As Integer
    Dim rsTemp As New ADODB.Recordset
        
    Txt诊断摘要 = ""
    
    Err = 0
    On Error Resume Next
    '给窗体的相关信息赋值
    With g病人身份_大连
        lblEdit(0).Caption = .个人编号
        lblEdit(1).Caption = .姓名
        lblEdit(3).Caption = Trim(.身份证号)
        int性别 = Val(IIf(Len(lblEdit(3)) = 18, Mid(lblEdit(3), 17, 1), Right(lblEdit(3), 1))) Mod 2
        '根据身份证取出相应的性别
        lblEdit(2).Caption = IIf(int性别 = 0, "女", "男")
        .出生日期 = zlCommFun.GetIDCardDate(Trim(.身份证号))
        '计算年龄
        If IsDate(.出生日期) And .出生日期 <> "" Then
            .年龄 = Abs(Int((zlDatabase.Currentdate - CDate(.出生日期)) / 365))
        Else
            .年龄 = 0
        End If
        
        .性别 = lblEdit(2).Caption
        lblEdit(4).Caption = .IC卡号
        lblEdit(5).Caption = .治疗序号
        '2004/09/11:加入退老
        lblEdit(6).Caption = Decode(.职工就医类别, "A", "在职", "B", "退休", "L", "离休", "T", "特诊", "Q", "企业公费", "E", "退老", "未知")
        lblEdit(7).Caption = Format(.基本个人帐户余额, "###,###0.00;-###,###0.00; ;")
        lblEdit(8).Caption = Format(.统筹累计, "###,###0.00;-###,###0.00; ;")
        lblEdit(9).Caption = Format(.月缴费基数, "###,###0.00;-###,###0.00; ;")
        lblEdit(10).Caption = Format(.补助个人帐户余额, "###,###0.00;-###,###0.00; ;")
        
        lblEdit(11).Caption = Decode(.参保类别3, "0", "企保", "1", "事保", "被征地人员")
        lblEdit(12).Caption = Decode(.帐户状态, "A", "正常", "B", "半止付", "C", "全止付", "D", "销户", "不能确定")
        chk生育.Value = IIf(.参保类别4 = 1, 1, 0)
        chk工伤.Value = IIf(.参保类别5 = 1, 1, 0)
        
        If mbytType <> 4 Then
            int划价天数 = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9OutExse\", "搜寻划价单据", "")
            gstrSQL = "Select Distinct(结论) as 结论 From 门诊费用记录 Where 记录性质 = 4 " & _
                      "And No=(select max(挂号单) from 病人医嘱记录 where id=(select max(id) from 病人医嘱记录  " & _
                      " Where 病人id=(select distinct(病人id)  from 保险帐户 where 医保号='" & .个人编号 & _
                      "') And 开嘱时间>trunc(Sysdate)-" & int划价天数 & "))"
        Else
            '2005-10-14 ZHQ
            '结帐时直接提取出院诊断
            gstrSQL = "Select 描述信息 as 结论 From 诊断情况 " & _
                    "   Where 诊断类型=3 And 诊断次序=1 And 病人ID In (Select 病人ID From 保险帐户 where 医保号='" & .个人编号 & "')" & _
                    "   Order by 主页ID Desc"
        End If
        rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
        If Not rsTemp.EOF Then
            Txt诊断摘要 = Nvl(rsTemp!结论, "")
        Else
            Txt诊断摘要 = ""
        End If
        
    End With
End Function

Private Function LoadCobData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:加载就诊分类数据和中心数据
    '--入参数:
    '--出参数:
    '--返  回:加载成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    Me.cbo就诊分类.Clear
    '   bytType-类型(0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐)
    With cbo就诊分类
        If mbytType = 0 Or mbytType = 2 Or mbytType = 3 Then
        
           '  1-1,2-3,3-5,4-"S"
            .AddItem "普通门诊"
            .ItemData(.NewIndex) = 1
            .AddItem "急诊门诊"
            .ItemData(.NewIndex) = 2
            .AddItem "门诊大病"
            .ItemData(.NewIndex) = 3
            .AddItem "门诊慢病补助"
            .ItemData(.NewIndex) = 4
        End If
        If mbytType = 1 Or mbytType = 2 Or mbytType = 4 Then
            '5-2,6-4,7-"O",8-"Q"
            
            .AddItem "普通住院"
            .ItemData(.NewIndex) = 5
            .AddItem "家庭病床住院"
            .ItemData(.NewIndex) = 6
            .AddItem "生育保险住院"
            .ItemData(.NewIndex) = 7
            .AddItem "工伤保险住院"
            .ItemData(.NewIndex) = 8
        End If
        If .ListCount <> 0 Then .ListIndex = 0
    End With
    
    '加载医保中心数据
    strSQL = "Select * From 保险中心目录 where 险类=" & mint险类 & " Order by 序号"
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    If rsTmp.RecordCount = 0 Then
        ShowMsgbox "医保中心未设置,请在保险类型中设置中心!"
        Exit Function
    End If
    With rsTmp
        cbo中心.Clear
        Do While Not .EOF
            cbo中心.AddItem Nvl(!编码) & "-" & Nvl(!名称)
            cbo中心.ItemData(cbo中心.NewIndex) = Nvl(!序号, 0)
            If Nvl(!序号, 0) = 2 And gblnKFQCom_大连 Then
                cbo中心.ListIndex = cbo中心.NewIndex
            End If
            .MoveNext
        Loop
        If cbo中心.ListCount <> 0 Then
            If cbo中心.ListIndex < 0 Then
               cbo中心.ListIndex = 0
            End If
        End If
    End With
    
    LoadCobData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetPatient(ByVal intinsure As Integer, ByVal bytType As Byte, Optional ByVal lng病人ID As Long = 0, _
                Optional lng性质 As Long, Optional lng记录ID As Long, Optional bytCallType As Byte = 1) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取病人的相关信息
    '--入参数:bytType-类型(mbytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐)
    '         lng病人ID-病人ID
    '         lng性质-结算记录中的性质
    '         lng记录id-结逄记录中的id
    '         bytCallType(1-结帐处调用;0-病人费用查询调用的)对虚拟结算有效即byttype=4的情况
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    mstrReturn = ""
    mlng病人ID = lng病人ID
    mint险类 = intinsure
    mbytType = bytType
    
    mlng性质 = lng性质
    mlng记录ID = lng记录ID
    
    If lng病人ID <> 0 And mbytType > 4 Then
        '需确定相关病人信息

        gstrSQL = "select b.医保号,a.姓名,a.性别,a.年龄,a.出生日期,a.身份证号, " & _
                 "        b.卡号,b.灰度级 as 就诊分类,b.顺序号,b.退休证号 as 补助个人帐户余额,b.病种id, " & _
                 "        b.参保类别1,b.参保类别2,b.参保类别3,b.参保类别4,b.参保类别5, " & _
                 "        b.人员身份 as 转诊单号,b.帐户余额 as 基本帐户余额,b.在职 as 职工就医类别 " & _
                 " from 病人信息 a,保险帐户 b " & _
                 " where a.病人id=b.病人id and b.险类=" & mint险类 & " and a.病人id=" & lng病人ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        With rsTemp
            If Not .EOF Then
                g病人身份_大连.IC卡号 = Nvl(!卡号)
                g病人身份_大连.补助个人帐户余额 = Nvl(!补助个人帐户余额, 0)
                g病人身份_大连.补助帐户当前值 = 0
                g病人身份_大连.补助帐户原始值 = 0
                g病人身份_大连.参保类别1 = Nvl(!参保类别1)
                g病人身份_大连.参保类别2 = Nvl(!参保类别2)
                g病人身份_大连.参保类别3 = Nvl(!参保类别3)
                g病人身份_大连.参保类别4 = Nvl(!参保类别4)
                g病人身份_大连.参保类别5 = Nvl(!参保类别5)
                g病人身份_大连.出生日期 = Format(!出生日期, "yyyy-mm-dd")
                g病人身份_大连.个人编号 = Nvl(!医保号)
                g病人身份_大连.基本个人帐户余额 = Nvl(!基本帐户余额, 0)
                g病人身份_大连.就诊分类 = Decode(Nvl(!就诊分类), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, "S", 4, "T", 4, "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
                g病人身份_大连.慢病帐户状态 = 0
                
                g病人身份_大连.年龄 = Val(Nvl(!年龄))
                g病人身份_大连.起付线 = 0
                g病人身份_大连.身份证号 = Nvl(!身份证号)
                g病人身份_大连.统筹累计 = 0
                g病人身份_大连.姓名 = Nvl(!姓名)
                g病人身份_大连.性别 = Nvl(!性别)
                g病人身份_大连.医保中心 = IIf(mint险类 = 82, 1, 2)
                g病人身份_大连.月缴费基数 = 0
                g病人身份_大连.帐户状态 = ""
                g病人身份_大连.诊断编码 = ""
                g病人身份_大连.诊断名称 = ""
                g病人身份_大连.支付金额 = 0
                g病人身份_大连.职工就医类别 = Nvl(!职工就医类别)
                
                g病人身份_大连.治疗序号 = 0
                g病人身份_大连.转诊单号 = Nvl(!转诊单号)
                
                gstrSQL = "Select 最高限额 from 保险结算记录 where 性质=" & mlng性质 & " and 记录id=" & mlng记录ID
                zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
                If .EOF Then
                    Txt最高限额.Text = ""
                Else
                    Txt最高限额.Text = Format(!最高限额, "####0.00;-#####0.00; ;")
                End If
                
                '设置数据
                Call SetCtlData
                '
                SetCtlVisible
                Me.Height = 4185
                fra(0).Top = 3255
                cmdOK.Top = fra(0).Top + fra(0).Height + 40
                cmdCancel.Top = cmdOK.Top
                Me.Caption = "病人最高限额录入"
                lblInfor.Caption = "输入病人的最高限额。"
            End If
        End With
    End If
    
    
    Me.Show 1
    GetPatient = mstrReturn
End Function
Private Sub SetCtlVisible()
    '设置控件的Vizible
    cbo就诊分类.Visible = False
    Txt转诊单号.Visible = False
    txt起付线.Visible = False
    txt病种.Visible = False
    cmd病种.Visible = False
    Txt疾病.Visible = False
    Txt最高限额.Visible = True
    lbl(17).Visible = True
    lbl(16).Visible = False
    lbl(14).Visible = False
    lbl(13).Visible = False
    lbl起付线.Visible = False
    lbl病种.Visible = False
    cmdOK.Enabled = False
    cmd验卡.Visible = False
    Txt诊断摘要.Visible = False
    
End Sub
Private Sub Form_Load()
    mblnFirst = True
    
    '加载分诊类别和医保中心
    Call LoadCobData
End Sub

Private Sub txt病种_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt病种.Text = ""
        txt病种.Tag = ""
    End If
End Sub

Private Sub Txt疾病_Change()
    Txt疾病.Tag = ""
End Sub

Private Sub Txt疾病_GotFocus()
    zlControl.TxtSelAll Txt疾病
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub Txt疾病_KeyPress(KeyAscii As Integer)
  Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str性别 As String
    Dim StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt疾病.Text = "" Then
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            strLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
            StrInput = UCase(Txt疾病.Text)
            str性别 = g病人身份_大连.性别
            If str性别 = "男" Then
                str性别 = " And (A.性别限制='男' Or A.性别限制 is NULL)"
            ElseIf str性别 = "女" Then
                str性别 = " And (A.性别限制='女' Or A.性别限制 is NULL)"
            End If

            strSQL = "Select A.ID,A.编码,A.附码,A.名称,A.简码,A.说明,A.性别限制,B.类别" & _
                " From 疾病编码目录 A,疾病编码类别 B" & _
                " Where A.类别=B.编码 And A.类别 Not IN('B','Z')" & _
                " And (A.编码 Like '" & StrInput & "%'" & _
                " Or Upper(A.名称) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.简码) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.附码) Like '" & strLike & StrInput & "%')" & _
                " And Rownum<=100" & str性别 & _
                " Order by A.类别,A.编码"

            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "疾病编码Input", , , , , , True, _
                Txt疾病.Left + Me.Left, _
                Txt疾病.Top + Me.Top, Txt疾病.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt疾病.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                Txt疾病.Tag = rsTmp!编码 & "|||" & rsTmp!名称
                If cmdOK.Enabled Then
                    cmdOK.SetFocus
                Else
                    Call zlCommFun.PressKey(vbKeyTab)
                End If
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                Call Txt疾病_GotFocus
                Txt疾病.SetFocus
            End If
        End If
    Else
        zlControl.TxtCheckKeyPress Txt疾病, KeyAscii, m文本式
    End If
End Sub

Private Sub Txt疾病_LostFocus()
    '--2004-12-28   ZHQ
    '医保要求门诊诊断必须输入
    If mbytType = 0 Then
        If Len(Trim(Txt疾病)) = 0 Then
            Txt疾病.SetFocus
        End If
    End If
End Sub

Private Sub txt起付线_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt起付线_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt起付线, KeyAscii, m金额式
End Sub

Private Sub Txt转诊单号_Change()
    txt起付线.Enabled = Txt转诊单号 <> "" And mbytType <> 0
    lbl起付线.Enabled = txt起付线.Enabled
End Sub

Private Sub Txt转诊单号_GotFocus()
   zlControl.TxtSelAll Txt转诊单号
End Sub

Private Sub Txt转诊单号_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
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
             "   FROM 保险病种 A WHERE A.险类=[1] And (A.编码 like [1] || '%' or A.名称 like [1] || '%' or A.简码 like [1] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类, strText)
    
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(mint险类, rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：")
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
        txt病种.ForeColor = Txt转诊单号.ForeColor
        SendKeys "{TAB}"
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Txt转诊单号_KeyPress(KeyAscii As Integer)
    KeyAscii = asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt最高限额_Change()
     cmdOK.Enabled = Val(Txt最高限额.Text) <> 0
End Sub

Private Sub Txt最高限额_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Txt最高限额_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt最高限额, KeyAscii, m金额式
End Sub

