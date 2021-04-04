VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm转院申请 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "转院申请"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   Icon            =   "frm转院申请.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9000
      TabIndex        =   43
      Top             =   5850
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7710
      TabIndex        =   42
      Top             =   5850
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "申请内容"
      Height          =   4305
      Left            =   120
      TabIndex        =   19
      Top             =   1410
      Width           =   10365
      Begin VB.TextBox txt备注 
         Height          =   300
         Left            =   1140
         TabIndex        =   41
         Top             =   3840
         Width           =   8835
      End
      Begin VB.TextBox txt单位意见 
         Height          =   600
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   3150
         Width           =   8835
      End
      Begin VB.TextBox txt医院意见 
         Height          =   600
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2460
         Width           =   8835
      End
      Begin VB.TextBox txt申请意见 
         Height          =   600
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   1770
         Width           =   8835
      End
      Begin MSComCtl2.DTPicker Dtp有效时限 
         Height          =   285
         Left            =   8700
         TabIndex        =   33
         Top             =   1380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   86966275
         CurrentDate     =   38063
      End
      Begin VB.TextBox txt科主任 
         Height          =   300
         Left            =   5070
         TabIndex        =   31
         Top             =   1380
         Width           =   2115
      End
      Begin VB.TextBox txt主治医生 
         Height          =   300
         Left            =   1140
         TabIndex        =   29
         Top             =   1380
         Width           =   2115
      End
      Begin VB.TextBox txt病情摘要 
         Height          =   600
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   690
         Width           =   8835
      End
      Begin VB.CommandButton cmd医疗机构 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9660
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   300
         Width           =   285
      End
      Begin VB.TextBox txt医院名称 
         Height          =   300
         Left            =   5970
         TabIndex        =   24
         Top             =   300
         Width           =   3675
      End
      Begin VB.CommandButton cmd疾病信息 
         Caption         =   "…"
         Height          =   300
         Left            =   4080
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   300
         Width           =   285
      End
      Begin VB.TextBox txt疾病信息 
         Height          =   300
         Left            =   1140
         TabIndex        =   21
         Top             =   300
         Width           =   2955
      End
      Begin VB.CheckBox chk定点医院 
         Caption         =   "定点医院"
         Height          =   255
         Left            =   4890
         TabIndex        =   23
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label lbl备注 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   690
         TabIndex        =   40
         Top             =   3900
         Width           =   360
      End
      Begin VB.Label lbl单位意见 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位意见"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   38
         Top             =   3210
         Width           =   720
      End
      Begin VB.Label lbl医院意见 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医院意见"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   36
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label lbl申请意见 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申请意见"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   34
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label lbl有效时限 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "有效时限"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7890
         TabIndex        =   32
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lbl科主任 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病区(科)主任"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3900
         TabIndex        =   30
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label lbl主治医生 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "主治医生"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   28
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lbl病情摘要 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病情摘要"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   26
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl疾病信息 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "疾病编码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   20
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.TextBox txt单位名称 
      Enabled         =   0   'False
      Height          =   300
      Left            =   7410
      TabIndex        =   18
      Top             =   1020
      Width           =   2985
   End
   Begin VB.TextBox txt医保号 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3630
      TabIndex        =   16
      Top             =   1020
      Width           =   2835
   End
   Begin VB.TextBox txt住院号 
      Enabled         =   0   'False
      Height          =   300
      Left            =   990
      TabIndex        =   14
      Top             =   1020
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -60
      TabIndex        =   44
      Top             =   540
      Width           =   10935
   End
   Begin VB.TextBox txt人员类别 
      Enabled         =   0   'False
      Height          =   300
      Left            =   9090
      TabIndex        =   12
      Top             =   630
      Width           =   1305
   End
   Begin VB.TextBox txt出生日期 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6990
      TabIndex        =   10
      Top             =   630
      Width           =   1035
   End
   Begin VB.TextBox txt性别 
      Enabled         =   0   'False
      Height          =   300
      Left            =   5340
      TabIndex        =   8
      Top             =   630
      Width           =   585
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3630
      TabIndex        =   6
      Top             =   630
      Width           =   1035
   End
   Begin VB.TextBox txt卡号 
      Enabled         =   0   'False
      Height          =   300
      Left            =   990
      MaxLength       =   20
      TabIndex        =   4
      Top             =   630
      Width           =   1935
   End
   Begin VB.ComboBox cbo申请类型 
      Height          =   300
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   1965
   End
   Begin VB.CommandButton cmd读卡 
      Caption         =   "读卡(&R)"
      Height          =   350
      Left            =   3000
      TabIndex        =   2
      Top             =   150
      Width           =   1100
   End
   Begin VB.Label lbl单位名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "单位名称"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6600
      TabIndex        =   17
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lbl医保号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保号"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3060
      TabIndex        =   15
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label lbl住院号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "住院号"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   13
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label lbl人员类别 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "人员类别"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   8250
      TabIndex        =   11
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl出生日期 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6180
      TabIndex        =   9
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl性别 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   7
      Top             =   690
      Width           =   360
   End
   Begin VB.Label lbl姓名 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3210
      TabIndex        =   5
      Top             =   690
      Width           =   360
   End
   Begin VB.Label lbl卡号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IC卡号"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   3
      Top             =   690
      Width           =   540
   End
   Begin VB.Label lbl申请类型 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "申请类型"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frm转院申请"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intMode As Integer          '编辑模式
Private blnOK As Boolean
Private blnEnable As Boolean        '是否连接医保
Private rsHospital As New ADODB.Recordset
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度
'转院申请分为普通转院申请和转外申请
'普通转院申请的业务类型有11和12
'如果转外申请才允许业务序列号为空时继续交易，业务类型为11

Public Function ShowME(ByVal int编辑模式 As Integer, ByVal frmParent As Object) As Boolean
    On Error Resume Next
    blnOK = False
    intMode = int编辑模式
    Me.Show 1, frmParent
    
    ShowME = blnOK
End Function

Private Sub InitFace()
    With cbo申请类型
        .Clear
        .AddItem "门诊"
        .ItemData(.NewIndex) = 11
        .AddItem "住院"
        .ItemData(.NewIndex) = 12
        .AddItem "转外就诊"
        .ItemData(.NewIndex) = 11
        .ListIndex = 1
    End With
    Dtp有效时限.Value = Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-MM-dd")
End Sub

Private Sub InitInsure()
    Dim lngRecord As Long
    Dim str编号 As String, str名称 As String
    Dim strFields As String, strValues As String
    
    '初始化医保接口
    If Not 医保初始化_沈阳市 Then Exit Sub
    
    '获取医院编码清单
    strFields = "ID" & "," & adDouble & "," & "10" & "|" & _
                "编号" & "," & adVarChar & "," & "20" & "|" & _
                "名称" & "," & adLongVarChar & "," & "50"
    Call Record_Init(rsHospital, strFields)
    If Not 调用接口_准备_沈阳市(Function_沈阳市.获取医院信息) Then Exit Sub
    If Not 调用接口_执行_沈阳市 Then Exit Sub
    strFields = "ID|编号|名称"
    If 调用接口_记录数_沈阳市 Then
        Do While True
            lngRecord = lngRecord + 1
            Call 调用接口_读取数据_沈阳市("hospital_id", str编号)
            Call 调用接口_读取数据_沈阳市("hospital_name", str名称)
            
            strValues = lngRecord & "|" & str编号 & "|" & str名称
            Call Record_Add(rsHospital, strFields, strValues)
            'todo 此处是调试代码
'            strValues = "1|001|001"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "2|002|002"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "3|003|003"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "4|004|004"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "5|005|005"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "6|006|006"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "7|007|007"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "8|008|008"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "9|009|009"
'            Call Record_Add(rsHospital, strFields, strValues)
'            strValues = "10|010|010"
            Call Record_Add(rsHospital, strFields, strValues)
            
            blnEnable = True
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
End Sub

Private Sub cbo申请类型_Click()
    gCominfo_沈阳市.业务类型 = cbo申请类型.ItemData(cbo申请类型.ListIndex)
End Sub

Private Sub cbo申请类型_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.cmd读卡.SetFocus
End Sub

Private Sub chk定点医院_Click()
    cmd医疗机构.Enabled = (chk定点医院.Value = 1)
End Sub

Private Sub cmd读卡_Click()
    Dim str卡号 As String
    Dim strReturn As String
    Dim lngReturn As Long
    '--读IC卡
    '读出IC卡中的信息
    Me.txt卡号 = ""
    If Not 调用接口_准备_沈阳市(Function_沈阳市.其他_读卡) Then Exit Sub
    If Not 调用接口_执行_沈阳市 Then Exit Sub
    '取返回的记录集
    'If Not 调用接口_指定记录集_沈阳市("ICInfo") Then Exit Sub
    'Modified By 朱玉宝 地区：长沙 原因：将传入参数由卡号改为身份证号
    If Not 调用接口_读取数据_沈阳市("card_no", str卡号) Then Exit Sub
    
    '获取病人的基本信息
    gstrField_沈阳市 = "hospital_id||iccardno"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & str卡号
    If Not 调用接口_准备_沈阳市(Function_沈阳市.转院申请_病人信息) Then Exit Sub
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Sub
    If Not 调用接口_执行_沈阳市 Then Exit Sub
    lngReturn = glngReturn_沈阳市           '保存执行结果
    
    '将返回数据显示在界面上
    'indi_id,insr_code,name,sex,pers_name,corp_name,patient_id,idcard
    If Not 调用接口_指定记录集_沈阳市("PersonInfo") Then Exit Sub
    Call 调用接口_读取数据_沈阳市("indi_id", strReturn)
    gCominfo_沈阳市.个人编号 = strReturn
    Call 调用接口_读取数据_沈阳市("name", strReturn)
    Me.txt姓名 = strReturn
    Call 调用接口_读取数据_沈阳市("sex", strReturn)
    Me.txt性别 = strReturn
    Call 调用接口_读取数据_沈阳市("idcard", strReturn)
    If strReturn <> "" Then
        If Len(strReturn) > 15 Then
            strReturn = Mid(strReturn, 7, 8)
        Else
            strReturn = "19" & Mid(strReturn, 7, 6)
        End If
        Me.txt出生日期 = Mid(strReturn, 1, 4) & "-" & Mid(strReturn, 5, 2) & "-" & Mid(strReturn, 7)
    End If
    Call 调用接口_读取数据_沈阳市("insr_code", strReturn)
    Me.txt医保号 = strReturn
    Call 调用接口_读取数据_沈阳市("corp_name", strReturn)
    Me.txt单位名称 = strReturn
    If lngReturn = 1 Then
        Call 调用接口_读取数据_沈阳市("patient_id", strReturn)
        Me.txt住院号 = strReturn
    End If
    
    '校验转院申请，入口参数如下
'    1   hospital_id 医疗机构编码   20  否
'    2   indi_id 个人编号           8   否
    gstrField_沈阳市 = "hospital_id||indi_id"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & gCominfo_沈阳市.个人编号
    If Not 调用接口_准备_沈阳市(Function_沈阳市.转院申请_校验信息) Then Exit Sub
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Sub
    If Not 调用接口_执行_沈阳市 Then Exit Sub
    If Not 调用接口_指定记录集_沈阳市("BizInfo") Then Exit Sub
    '取到一条就行（只可能有一条），返回字段如下
'    1   serial_no   业务序列号 12  有效的住院业务的业务序列号
'    2   patient_id  住院号     20
    Call 调用接口_读取数据_沈阳市("patient_id", strReturn)
    Me.txt住院号 = strReturn
    If cbo申请类型.ListIndex <> 2 Then  '转外就诊不需要读取业务序列号
        Call 调用接口_读取数据_沈阳市("serial_no", strReturn)
        gCominfo_沈阳市.业务序列号 = strReturn
    End If
    
    '全部执行成功才将卡号填写在界面上
    Me.txt卡号 = str卡号
End Sub

Private Sub cmd疾病信息_Click()
    Dim rs病种 As ADODB.Recordset
    
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码 " & _
            " From 保险病种 A where A.险类=[1]"
    Set rs病种 = zlDatabase.OpenSQLRecord(gstrSQL, "身份验证", TYPE_沈阳市)
    If rs病种.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_沈阳市, rs病种, "ID", "医保病种选择", "请选择医保病种：") = True Then
            txt疾病信息.Tag = rs病种!编码
            txt疾病信息.Text = "(" & rs病种!编码 & ")" & rs病种!名称
            lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
        End If
    End If
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    '检测数据正确性
    If Trim(txt卡号) = "" Then
        MsgBox "还没获取该参保病人的基本信息，请先点击读卡按钮！", vbInformation, gstrSysName
        cmd读卡.SetFocus
        Exit Sub
    End If
    If Trim(txt疾病信息.Tag) = "" Then
        MsgBox "请输入或选择该参保人的疾病信息！", vbInformation, gstrSysName
        txt疾病信息.SetFocus
        Exit Sub
    End If
    If chk定点医院.Value = 1 Then
        If Trim(chk定点医院.Tag) = "" Then
            MsgBox "请输入或选择定点医疗机构！", vbInformation, gstrSysName
            txt医院名称.SetFocus
            Exit Sub
        End If
    Else
        If Trim(txt医院名称.Text) = "" Then
            MsgBox "请输入医疗机构名称！", vbInformation, gstrSysName
            txt医院名称.SetFocus
            Exit Sub
        End If
    End If
    If Trim(txt主治医生.Text) = "" Then
        MsgBox "请输入主治医生姓名！", vbInformation, gstrSysName
        txt主治医生.SetFocus
        Exit Sub
    End If
    
    '调用保存申请功能
'    1   busi_type      业务类型    2   否  "12"：住院
'    2   indi_id        个人编号    8   否
'    3   apply_hospital 申请医疗机构编码    20  否
'    4   hospital_id    转入非定点医疗机构编码  20
'    5   hospital_name  转入非定点医疗机构名称  60      hospital_id或hospital_name与to_hospital不能同时为空
'    6   to_hospital    转入定点医疗机构编码            20
'    7   apply_content  申请内容    1   否  "4"：转院住院申请    "5"：转外就医申请
'    8   serial_no      业务序列号  12  否
'    9   icd            疾病编码    20  是  医保中心目录编码
'    10  disease_desc   病情摘要    500 是
'    11  doctor_name    申请医师    10  是
'    12  apply_opinion  申请理由    500 是
'    13  hosp_opinion   医院相关意见    500 是
'    14  corp_opinion   病人单位意见    500 是
'    15  apply_date     申请有效时限        是  格式：YYYY-MM-DD
'    16  input_man      录入人      10  否
'    17  input_date     录入时间        是  格式：YYYY-MM-DD
'    18  note           备注        500 是
    gstrField_沈阳市 = "busi_type||indi_id||apply_hospital||hospital_id||hospital_name||to_hospital||" & _
                       "apply_content||serial_no||icd||disease_desc||doctor_name||apply_opinion||" & _
                       "hosp_opinion||corp_opinion||apply_date||input_man||input_date||note"
    gstrValue_沈阳市 = gCominfo_沈阳市.业务类型 & "||" & gCominfo_沈阳市.个人编号 & "||" & gCominfo_沈阳市.医院编码 & "||" & _
                       "||" & IIf(chk定点医院.Value = 1, "", txt医院名称.Text) & "||" & IIf(chk定点医院.Value = 1, chk定点医院.Tag, "") & "||" & _
                       IIf(cbo申请类型.ListIndex = 2, "5", "4") & "||" & gCominfo_沈阳市.业务序列号 & "||" & _
                       txt疾病信息.Tag & "||" & txt病情摘要.Text & "||" & txt主治医生.Text & "||" & txt申请意见.Text & "||" & _
                       txt医院意见.Text & "||" & txt单位意见.Text & "||" & Format(Me.Dtp有效时限.Value, "yyyy-MM-dd") & "||" & _
                       gstrUserName & "||" & Format(zlDatabase.Currentdate, "yyyy-MM-dd") & "||" & txt备注.Text
    If Not 调用接口_准备_沈阳市(Function_沈阳市.转院申请_保存转院申请) Then Exit Sub
    If Not 调用接口_写入口参数_沈阳市(1) Then Exit Sub
    If Not 调用接口_执行_沈阳市 Then Exit Sub
    
    blnOK = True
    Unload Me
    Exit Sub
End Sub

Private Sub cmd医疗机构_Click()
    If chk定点医院.Value = 0 Then Exit Sub
    If frmListSel.ShowSelect(TYPE_沈阳市, rsHospital, "ID", "选择定点医疗机构", "请选择一家定点医疗机构：") = True Then
        chk定点医院.Tag = rsHospital!编号
        txt医院名称.Text = "(" & rsHospital!编号 & ")" & rsHospital!名称
        lbl医院意见.Tag = txt医院名称.Text '用于恢复显示
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If InStr(1, "txt医院名称,txt疾病信息", ActiveControl.Name) <> 0 Then Exit Sub
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call InitFace
    Call InitInsure
    If Not blnEnable Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub txt备注_GotFocus()
    Call zlControl.TxtSelAll(txt备注)
End Sub

Private Sub txt病情摘要_GotFocus()
    Call zlControl.TxtSelAll(txt病情摘要)
End Sub

Private Sub txt单位名称_GotFocus()
    Call zlControl.TxtSelAll(txt单位意见)
End Sub

Private Sub txt单位意见_GotFocus()
    Call zlControl.TxtSelAll(txt单位意见)
End Sub

Private Sub txt疾病信息_GotFocus()
    Call zlControl.TxtSelAll(txt疾病信息)
End Sub

Private Sub txt疾病信息_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt疾病信息.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = txt疾病信息.Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        Else
            strText = Mid(strText, 2)
        End If
    End If
    gstrSQL = "Select A.ID,A.编码,A.名称,A.简码" & _
             "   FROM 保险病种 A WHERE A.险类=" & TYPE_沈阳市 & " And " & _
             "(A.编码 like [2] || '%' or A.名称 like [2] || '%' or A.简码 like [2] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_沈阳市, strText)
    
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_沈阳市, rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        txt疾病信息.SetFocus
        Call zlControl.TxtSelAll(txt疾病信息)
        Exit Sub
    Else
        '肯定是有记录集的
        txt疾病信息.Tag = rsTemp!编码
        txt疾病信息.Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
        lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt科主任_GotFocus()
    Call zlControl.TxtSelAll(txt科主任)
End Sub

Private Sub txt申请意见_GotFocus()
    Call zlControl.TxtSelAll(txt申请意见)
End Sub

Private Sub txt医院名称_GotFocus()
    Call zlControl.TxtSelAll(txt医院名称)
End Sub

Private Sub txt医院名称_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrInput As String
    Dim blnReturn As Boolean
    If KeyCode <> vbKeyReturn Then Exit Sub
    If chk定点医院.Value = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    StrInput = Trim(UCase(txt医院名称.Text))
    If InStr(1, StrInput, "(") <> 0 Then
        If InStr(1, StrInput, ")") <> 0 Then
            StrInput = Mid(StrInput, 2, InStr(1, StrInput, ")") - 2)
        Else
            StrInput = Mid(StrInput, 2)
        End If
    End If
    If IsNumeric(StrInput) Then
        rsHospital.Filter = "编号 Like '" & StrInput & "*'"
    Else
        rsHospital.Filter = "名称 Like '" & StrInput & "*'"
    End If
    If rsHospital.RecordCount = 0 Then
        MsgBox "没找到该编号的定点医疗机构信息，请重新输入！", vbInformation, gstrSysName
        rsHospital.Filter = 0
        txt医院名称.SetFocus
        GoTo ExitSub
    Else
        If rsHospital.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_沈阳市, rsHospital, "ID", "选择定点医疗机构", "请选择一家定点医疗机构：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        Call zlControl.TxtSelAll(txt医院名称)
        GoTo ExitSub
    Else
        '肯定是有记录集的
        chk定点医院.Tag = rsHospital!编号
        txt医院名称.Text = "(" & rsHospital!编号 & ")" & rsHospital!名称
        lbl医院意见.Tag = txt医院名称.Text '用于恢复显示
        Call zlCommFun.PressKey(vbKeyTab)
    End If

ExitSub:
    rsHospital.Filter = 0
End Sub

Private Sub txt医院意见_GotFocus()
    Call zlControl.TxtSelAll(txt医院意见)
End Sub

Private Sub txt主治医生_GotFocus()
    Call zlControl.TxtSelAll(txt主治医生)
End Sub
