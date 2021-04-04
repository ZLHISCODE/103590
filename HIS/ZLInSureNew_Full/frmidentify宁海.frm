VERSION 5.00
Begin VB.Form frmIdentify宁海 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "frmidentify宁海.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk无卡病人 
      Caption         =   "无卡病人"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      TabIndex        =   0
      Top             =   180
      Width           =   3585
   End
   Begin VB.ComboBox cbo业务类型 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5025
      Width           =   3690
   End
   Begin VB.TextBox txt出生日期 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   12
      Top             =   2760
      Width           =   3705
   End
   Begin VB.CommandButton cmd疾病信息 
      Caption         =   "…"
      Height          =   345
      Left            =   5580
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5490
      Width           =   330
   End
   Begin VB.TextBox txt疾病信息 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   25
      Top             =   5490
      Width           =   3360
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6360
      TabIndex        =   29
      Top             =   1200
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6360
      TabIndex        =   28
      Top             =   675
      Width           =   1320
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "读卡(&R)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6360
      TabIndex        =   21
      Top             =   5310
      Width           =   1320
   End
   Begin VB.Frame frame1 
      Height          =   6255
      Left            =   6135
      TabIndex        =   27
      Top             =   -105
      Width           =   45
   End
   Begin VB.TextBox txt医疗证号 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   2
      Top             =   570
      Width           =   3705
   End
   Begin VB.TextBox txt本年个帐余额 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   20
      Top             =   4575
      Width           =   3705
   End
   Begin VB.TextBox txt上年个帐余额 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   18
      Top             =   4125
      Width           =   3705
   End
   Begin VB.TextBox txt单位名称 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   16
      Top             =   3660
      Width           =   3705
   End
   Begin VB.TextBox txt身份证号 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   14
      Top             =   3210
      Width           =   3705
   End
   Begin VB.TextBox txt性别 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   10
      Top             =   2310
      Width           =   1290
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   8
      Top             =   1875
      Width           =   3705
   End
   Begin VB.TextBox txt公务员 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   6
      Top             =   1440
      Width           =   3705
   End
   Begin VB.TextBox txt个人帐号 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2235
      TabIndex        =   4
      Top             =   1005
      Width           =   3705
   End
   Begin VB.Label lbl业务类型 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "业务类型(&U)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   22
      Top             =   5085
      Width           =   1425
   End
   Begin VB.Label lbl出生日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期(&H)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   11
      Top             =   2820
      Width           =   1425
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "病种(&I)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1260
      TabIndex        =   24
      Top             =   5550
      Width           =   915
   End
   Begin VB.Label lbl医疗证号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医疗证号(&A)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   1
      Top             =   630
      Width           =   1425
   End
   Begin VB.Label lbl本年个帐余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "本年个帐余额(&Y)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   19
      Top             =   4635
      Width           =   1935
   End
   Begin VB.Label lbl上年个帐余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "上年个帐余额(&L)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   17
      Top             =   4185
      Width           =   1935
   End
   Begin VB.Label lbl单位名称 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "单位名称(&K)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   15
      Top             =   3720
      Width           =   1425
   End
   Begin VB.Label lbl身份证号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号(&J)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   13
      Top             =   3270
      Width           =   1425
   End
   Begin VB.Label lbl性别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别(&G)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1260
      TabIndex        =   9
      Top             =   2370
      Width           =   915
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&F)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1260
      TabIndex        =   7
      Top             =   1935
      Width           =   915
   End
   Begin VB.Label lbl公务员 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "公务员(&D)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   990
      TabIndex        =   5
      Top             =   1500
      Width           =   1170
   End
   Begin VB.Label lbl个人帐号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "个人帐号(&S)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   3
      Top             =   1065
      Width           =   1425
   End
End
Attribute VB_Name = "frmIdentify宁海"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mlng病人ID As Long
Private mstrReturn As String

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng病人ID As Long = 0) As String
    mstrReturn = ""
    mlng病人ID = lng病人ID
    mbytType = bytType
    Me.Show 1
    
    GetPatient = mstrReturn
End Function

Private Sub chk无卡病人_Click()
    txt医疗证号.Enabled = (chk无卡病人.Value = 1)
End Sub

Private Sub cmdOK_Click()
    Dim lngPatient As Long
    Dim strIdentify As String
    Dim strAddition As String
    Dim str退休证号 As String
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txt医疗证号.Text) = "" Then
        MsgBox "还未读卡！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'Modified by ZYB 2006-04-12，根据宁海县医保中心2006-04-06下发的文件要求修改，门诊住院都必须上传疾病信息
'    If mbytType = 1 Then
        If Val(txt疾病信息.Tag) = 0 Then
            MsgBox "必须选择入院病种！", vbInformation, gstrSysName
            txt疾病信息.SetFocus
            Exit Sub
        End If
'    End If
    
    '检查病人状态
    gstrSQL = "select 病人ID,nvl(当前状态,0) as 状态,顺序号,灰度级,退休证号 from 保险帐户 where 险类=[1] and 医保号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_宁海, txt医疗证号.Text)
    If rsTemp.RecordCount > 0 Then
        If rsTemp!状态 = 1 Then
            MsgBox "当前病人在院，不能再次通过身份验证！", vbInformation, gstrSysName
            Exit Sub
        End If
        str退休证号 = Nvl(rsTemp!退休证号)
    End If
    
    '产生病人信息
    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    
    '灰度级，记录补充入院的次数，缺省第一次为空，因为文档要求，每次医保入院的住院号不能相同，我们只有通过补充入院次数解决，正常的入院没有影响
    strIdentify = txt医疗证号.Text                              '0卡号
    strIdentify = strIdentify & ";" & txt个人帐号.Text          '1医保号
    strIdentify = strIdentify & ";"                             '2密码
    strIdentify = strIdentify & ";" & txt姓名.Text              '3姓名
    strIdentify = strIdentify & ";" & txt性别.Text              '4性别
    strIdentify = strIdentify & ";" & txt出生日期.Text          '5出生日期
    strIdentify = strIdentify & ";"                             '6身份证
    strIdentify = strIdentify & ";" & txt单位名称               '7.单位名称(编码)
    strAddition = ";0"                                          '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";"                             '10人员身份
    strAddition = strAddition & ";" & Val(txt上年个帐余额.Text) + Val(txt本年个帐余额.Text)    '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & Val(txt疾病信息.Tag)      '13病种ID
    strAddition = strAddition & ";1"                            '14在职(1,2,3)
    strAddition = strAddition & ";" & str退休证号               '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";" & chk无卡病人.Value         '17灰度级
    strAddition = strAddition & ";" & Val(txt上年个帐余额.Text) + Val(txt本年个帐余额.Text)     '18帐户增加累计
    strAddition = strAddition & ";0"                            '19帐户支出累计
    strAddition = strAddition & ";"                             '20上年工资总额
    strAddition = strAddition & ";" & Val(IC_Data_宁海.结算次数)      '21住院次数累计

    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_宁海)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    Else
        Exit Sub
    End If
    
    IC_Data_宁海.mstr业务类型 = cbo业务类型.ItemData(cbo业务类型.ListIndex)
    If Val(IC_Data_宁海.mstr业务类型) = 0 Then        '普通时需特殊处理，普通门诊为11，普通住院为21
        If mbytType = 0 Then
            IC_Data_宁海.mstr业务类型 = "11"
        Else
            IC_Data_宁海.mstr业务类型 = "21"
        End If
    End If
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_宁海 & ",'业务类型','''" & IC_Data_宁海.mstr业务类型 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存业务类型")
    
    '保存用户IC卡内所有数据
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_宁海 & ",'IC','''" & IC_Data_宁海.IC卡数据 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存用户IC卡内所有数据")
    '保存上年个帐余额与本年个帐余额
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_宁海 & ",'往年帐户余额','''" & Val(txt上年个帐余额.Text) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存用户IC卡内所有数据")
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_宁海 & ",'本年帐户余额','''" & Val(txt本年个帐余额.Text) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存用户IC卡内所有数据")
    
    IC_Data_宁海.mlng疾病ID = Val(txt疾病信息.Tag)
    Unload Me
End Sub

Private Sub cmdRead_Click()
    '读卡正确完成后，自动跳到确定处
    If Not ReadIC_宁海(IIf(chk无卡病人.Value = 1, txt医疗证号.Text, "")) Then Exit Sub
    
    txt医疗证号.Text = IC_Data_宁海.医疗证号
    txt个人帐号.Text = IC_Data_宁海.帐号
    txt公务员.Text = IIf(IC_Data_宁海.公务员 = "1", "是", "否")
    txt姓名.Text = IC_Data_宁海.姓名
    txt性别.Text = IC_Data_宁海.性别
    txt出生日期.Text = IC_Data_宁海.出生日期
    txt身份证号.Text = IC_Data_宁海.身份号
    txt单位名称.Text = IC_Data_宁海.单位名称
    txt上年个帐余额.Text = Format(IC_Data_宁海.结转余额 - IC_Data_宁海.个帐历年使用累计, "#0.00")
    txt本年个帐余额.Text = Format(IC_Data_宁海.当年实际拨付 - IC_Data_宁海.个帐当年使用累计, "#0.00")
    txt疾病信息.SetFocus
End Sub

Private Sub cmd疾病信息_Click()
    Dim rs病种 As ADODB.Recordset
        
    gstrSQL = " Select A.JBBM AS ID,A.JBBZDM AS 编码,A.JBMC AS 名称,A.PYJM AS 简码 " & _
            " From SIM_JBDA A "
    Set rs病种 = New ADODB.Recordset
    rs病种.Open gstrSQL, gcn宁海
    If rs病种.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_宁海, rs病种, "ID", "医保病种选择", "请选择医保病种：") = True Then
            txt疾病信息.Tag = rs病种!ID
            txt疾病信息.Text = "(" & rs病种!编码 & ")" & rs病种!名称
            lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
        End If
    End If
    cmdOK.SetFocus
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.cbo业务类型.Clear
    Me.cbo业务类型.AddItem "普通"           '门诊为11，住院为21
    Me.cbo业务类型.ItemData(Me.cbo业务类型.NewIndex) = 0
    If mbytType = 0 Then
        Me.cbo业务类型.AddItem "特殊病种门诊"
        Me.cbo业务类型.ItemData(Me.cbo业务类型.NewIndex) = 12
    ElseIf mbytType = 1 Then
        Me.cbo业务类型.AddItem "特殊病种住院"
        Me.cbo业务类型.ItemData(Me.cbo业务类型.NewIndex) = 32
        Me.cbo业务类型.AddItem "家庭病床"
        Me.cbo业务类型.ItemData(Me.cbo业务类型.NewIndex) = 31
    End If
    Me.cbo业务类型.ListIndex = 0
    Me.cbo业务类型.Enabled = (mbytType = 0 Or mbytType = 1)
End Sub

Private Sub txt疾病信息_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt疾病信息.Text = "" And txt疾病信息.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    strText = txt疾病信息.Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
        
    gstrSQL = "Select A.JBBM AS ID,A.JBBZDM AS 编码,A.JBMC AS 名称,A.PYJM AS 简码" & _
             "   FROM SIM_JBDA A WHERE 1=1 And (" & _
                zlCommFun.GetLike("A", "JBBZDM", strText) & " or " & zlCommFun.GetLike("A", "JBMC", strText) & " or " & zlCommFun.GetLike("A", "PYJM", strText) & ")"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn宁海
    If rsTemp.RecordCount = 0 Then
        MsgBox "不存在该病种，请重新输入！", vbInformation, gstrSysName
        txt疾病信息.Text = lbl疾病信息.Tag
        zlControl.TxtSelAll txt疾病信息
        Exit Sub
    Else
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_宁海, rsTemp, "ID", "医保病种选择", "请选择医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        txt疾病信息.Text = lbl疾病信息.Tag
        zlControl.TxtSelAll txt疾病信息
        Exit Sub
    Else
        '肯定是有记录集的
        txt疾病信息.Tag = rsTemp!ID
        txt疾病信息.Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
        lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
    End If
    
    cmdOK.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
