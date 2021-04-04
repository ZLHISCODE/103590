VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIdentify奉庆 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk个人帐户 
      Caption         =   "下个人帐户(&D)"
      Height          =   210
      Left            =   5100
      TabIndex        =   13
      Top             =   1305
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.ComboBox cbo审批类别 
      Height          =   300
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   5190
      Visible         =   0   'False
      Width           =   2085
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3525
      Left            =   -225
      TabIndex        =   36
      Top             =   5160
      Visible         =   0   'False
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   6218
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmd验卡 
      Caption         =   "读卡(&R)"
      Height          =   350
      Left            =   165
      TabIndex        =   2
      Top             =   4065
      Width           =   1100
   End
   Begin VB.ComboBox cbo类别 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   4365
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3180
      Width           =   2310
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5625
      TabIndex        =   4
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4335
      TabIndex        =   3
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   270
      Left            =   6420
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3600
      Width           =   255
   End
   Begin VB.TextBox txt病种 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   870
      TabIndex        =   1
      Top             =   3585
      Width           =   5820
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   34
      Top             =   615
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -465
      TabIndex        =   33
      Top             =   3915
      Width           =   8340
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "统筹区号"
      Height          =   180
      Index           =   11
      Left            =   135
      TabIndex        =   28
      Top             =   3225
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   4365
      TabIndex        =   21
      Top             =   2025
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4365
      TabIndex        =   8
      Top             =   855
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   870
      TabIndex        =   6
      Top             =   885
      Width           =   2310
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "医保病人基本信息显示，可以通过读卡按钮重新进行读取IC卡信息。"
      Height          =   180
      Left            =   630
      TabIndex        =   35
      Top             =   360
      Width           =   5400
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify奉庆.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡号"
      Height          =   180
      Index           =   0
      Left            =   495
      TabIndex        =   5
      Top             =   952
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保证号"
      Height          =   180
      Index           =   1
      Left            =   3615
      TabIndex        =   7
      Top             =   945
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   2
      Left            =   495
      TabIndex        =   9
      Top             =   1312
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   3975
      TabIndex        =   11
      Top             =   1305
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   4
      Left            =   135
      TabIndex        =   18
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出身日期"
      Height          =   180
      Index           =   5
      Left            =   3615
      TabIndex        =   16
      Top             =   1695
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "类别编码"
      Height          =   180
      Index           =   6
      Left            =   3615
      TabIndex        =   20
      Top             =   2085
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗类别(&T)"
      Height          =   180
      Index           =   7
      Left            =   3345
      TabIndex        =   30
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   8
      Left            =   495
      TabIndex        =   14
      Top             =   1725
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "类别名称"
      Height          =   180
      Index           =   9
      Left            =   135
      TabIndex        =   22
      Top             =   2505
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位名称"
      Height          =   180
      Index           =   10
      Left            =   135
      TabIndex        =   26
      Top             =   2857
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   12
      Left            =   3615
      TabIndex        =   24
      Top             =   2475
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   870
      TabIndex        =   10
      Top             =   1260
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4365
      TabIndex        =   12
      Top             =   1260
      Width           =   525
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   870
      TabIndex        =   15
      Top             =   1650
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4365
      TabIndex        =   17
      Top             =   1650
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   870
      TabIndex        =   19
      Top             =   2040
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   870
      TabIndex        =   23
      Top             =   2430
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   855
      TabIndex        =   29
      Top             =   3150
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   870
      TabIndex        =   27
      Top             =   2805
      Width           =   5805
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4365
      TabIndex        =   25
      Top             =   2430
      Width           =   2310
   End
   Begin VB.Label lbl病种 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "病种(&F)"
      Height          =   180
      Left            =   225
      TabIndex        =   31
      Top             =   3645
      Width           =   630
   End
End
Attribute VB_Name = "frmIdentify奉庆"
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
Private mstrArr             '获得通过的审批信息 0 审批编码,1 编号,2 名称 ....
Private Sub cbo类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk个人帐户_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd病种_Click()
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "" & _
            "   Select id, 编码, 名称, 助记码,to_char(变更时间,'yyyy-mm-dd hh24:mi:ss') as 变更时间" & _
            "   From 医保病种目录"
                
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        With rsTemp
            If .EOF Then
                MsgBox "不存在任何病种,请下载！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = txt病种.Top - .Height
                    .Left = txt病种.Left + txt病种.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 2000
                    .ColWidth(3) = 1400
                    .ColWidth(4) = .Width - .ColWidth(1) - .ColWidth(2) - .ColWidth(3) - .ColWidth(4)
                    .Row = 1
                    .COL = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt病种 = "[" & Nvl(!编码) & "]" & IIf(IsNull(!名称), "", !名称)
                txt病种.Tag = Nvl(!ID)
                zlCommFun.PressKey vbKeyTab
            End If
        End With
End Sub

Private Sub cmd验卡_Click()
    
   If 获取参保人员信息_奉庆 = False Then
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
    
    If 获取参保人员信息_奉庆 = False Then
        Call ClearData
        cmd确定.Enabled = False
        Exit Sub
    End If
    Call LoadCtrlData
    cmd确定.Enabled = True
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
    Dim lng疾病ID  As Long
    
    lng疾病ID = Val(txt病种.Tag)
    IsValid = False
    If mbytType = 4 Or mbytType = 1 Then
        '2008/04/01:金之裕要求，入院必需选择病种
        If lng疾病ID = 0 Then
            ShowMsgbox "未输入病种"
            If txt病种.Enabled Then txt病种.SetFocus
            Exit Function
        End If
    End If
    If Trim(g病人身份_奉庆.姓名) = "" Then
        MsgBox "还没进行身份验证！", vbInformation, gstrSysName
        If cmd验卡.Enabled Then cmd验卡.SetFocus
        Exit Function
    End If
    
    If Trim(txt病种) <> "" And Val(txt病种.Tag) = 0 Then
        ShowMsgbox "病种选择错误,请重新选择!"
        txt病种.SetFocus
        Exit Function
    End If
    If cbo类别.Text = "" Then
        ShowMsgbox "支付类别未选择"
        Exit Function
    End If
    
    If 参保资格审查_奉庆 = False Then
        Exit Function
    End If
 
    '住院状态判断 1 在住院  2 正常（非住院）  0 初始化值（同2）  4 转院
    If 获取住院状态_奉庆(lng状态) = False Then
        Exit Function
    End If
    
    If lng状态 = 1 And mbytType = 0 Then
        ShowMsgbox "当前病人已经在院，不能看门诊!"
        Exit Function
    End If
    If lng状态 = 1 And mbytType = 1 Then
        ShowMsgbox "当前病人已经在院,不能入院登记!"
        Exit Function
    End If
    
    If lng状态 = 4 And InStr(1, cbo类别.Text, "转院") = 0 Then
        ShowMsgbox "当前病人为转院，请选择医疗类别为转院!"
        Exit Function
    End If
    
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '不检查录前着态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_临沧奉庆, g病人身份_奉庆.医保证号)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
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
    Dim str类别 As String
    Dim int当前状态 As Integer
    Dim lng状态 As Long
    
    g病人身份_奉庆.医疗类别 = Split(cbo类别.Text, "-")(0)
    If cbo审批类别.ListIndex >= 0 Then
        g病人身份_奉庆.审批类别 = Split(cbo审批类别.Text, "-")(0)
    Else
        g病人身份_奉庆.审批类别 = ""
    End If
    lng疾病ID = Val(txt病种.Tag)
    
    If IsValid = False Then Exit Sub
    
    Err = 0: On Error GoTo errHand:
    DebugTool "1.进行病种确定"
    If lng疾病ID <> 0 And txt病种.Text <> "" Then
        g病人身份_奉庆.病种编码 = Mid(txt病种.Text, 2, InStr(1, txt病种.Text, "]") - 2)
        g病人身份_奉庆.病种名称 = Mid(txt病种.Text, InStr(1, txt病种.Text, "]") + 1)
    Else
        g病人身份_奉庆.病种编码 = "000000"
    End If
    DebugTool "2.进行病种确定成功"
    
    g病人身份_奉庆.下个人帐户 = IIf(chk个人帐户.Value = 1, True, False)
    g病人身份_奉庆.病种ID = lng疾病ID
    
    g病人身份_奉庆.医疗类别 = Mid(cbo类别.Text, 1, InStr(1, cbo类别.Text, "-") - 1)
    int当前状态 = 0
    '确定病人是否慢病
    If (mbytType = 0 Or mbytType = 3) And g病人身份_奉庆.医疗类别 = "13" Then
        If 获得通过的审批信息() = False Then Exit Sub
    End If
    
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_临沧奉庆 & " and  医保号='" & g病人身份_奉庆.医保证号 & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            If mlng病人ID <> Nvl(rsTemp!病人ID, 0) And mlng病人ID <> 0 Then
                ShowMsgbox "不是当前所要结算的用户"
                Exit Sub
            End If
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
        End If
        rsTemp.Close
    End If
    DebugTool "3.开始构建病人信息串"
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_奉庆
        
        strIdentify = .卡号                               '0卡号
        strIdentify = strIdentify & ";" & .医保证号           '1医保号
        strIdentify = strIdentify & ";" & .医疗类别                   '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & Decode(.性别, "1", "男", "2", "女", .性别)              '4性别
        strIdentify = strIdentify & ";" & .出生日期                '5出生日期
        strIdentify = strIdentify & ";" & .身份证号           '6身份证
        strIdentify = strIdentify & ";" & .单位名称 & IIf(.单位编码 = 0, "", "(" & .单位编码 & ")")          '7.单位名称(编码)
        strAddition = ";0"                                          '8.中心代码
        strAddition = strAddition & ";"                              '9.顺序号
        strAddition = strAddition & ";"                                '10人员身份
        strAddition = strAddition & ";" & .帐户余额                 '11帐户余额
        
        strAddition = strAddition & ";" & int当前状态                            '12当前状态
        strAddition = strAddition & ";" & IIf(lng疾病ID = 0, "", lng疾病ID)             '13病种ID
        strAddition = strAddition & ";1"                            '14在职(1,2,3)
        strAddition = strAddition & ";" & .审批编号 & "|" & .项目编码 & "|" & .项目名称                               '15退休证号
        strAddition = strAddition & ";" & .年龄                     '16年龄段
        strAddition = strAddition & ";" & .审批类别             '17灰度级
        strAddition = strAddition & ";" & .帐户余额                 '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    DebugTool "5.开始保存病人"
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_临沧奉庆)
    If mlng病人ID = 0 Then Exit Sub
    DebugTool "保存成功"
    
    If mbytType = 0 Or mbytType = 3 Then
        '获取门诊号
        Dim lng就诊次数 As Long
        Dim str交易流水号 As String
        gstrSQL = "Select nvl(就诊次数,0)+1 as 就诊次数 From 保险帐户 where 病人ID=" & mlng病人ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        lng就诊次数 = Nvl(rsTemp!就诊次数, 1)
        g病人身份_奉庆.门诊号 = mlng病人ID & "_" & lng就诊次数
        '更新保险帐户
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_临沧奉庆 & ",'就诊次数','" & lng就诊次数 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊次数")
        DebugTool "6.进行登记处理"
        '先进行登记处理
        If 病人登记处理(str交易流水号) = False Then
            Exit Sub
        End If
        
        DebugTool "登记成功"
        '保存将交易流水号
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_临沧奉庆 & ",'顺序号','''" & str交易流水号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存交易流水号")
        DebugTool "7.更新就诊信息:" & g病人身份_奉庆.帐户余额
        
        If 更新就诊信息_奉庆(0, strOutput) = False Then Exit Sub
        DebugTool "8.更新就诊信息"
    Else
        gstrSQL = "Select * From 保险帐户 where 病人ID=" & mlng病人ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        With g病人身份_奉庆
            .门诊号 = mlng病人ID & "_" & Nvl(rsTemp!就诊次数, 0)
        End With
    End If
    g病人身份_奉庆.病人ID = mlng病人ID
    
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    Unload Me
    Exit Sub
errHand:
    DebugTool "Err:" & Err.Description
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function 获得通过的审批信息() As Boolean
    Dim StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long, strArr
    Err = 0
    On Error GoTo errHand:
    With g病人身份_奉庆
'        strInput = .医保证号 & "|"
'        strInput = strInput & "|" & Split(cbo审批类别.Text, "-")(0)
'        If 业务请求_奉庆(奉庆_获得通过的审批信息, strInput, strOutput) = False Then Exit Function
'        strArr = Split(strOutput, "|")
'        If UBound(mstrArr) <= 1 Then
'            ShowMsgbox "没有相关的审批信息！“"
'            Exit Function
'        End If
'
'        '构建数据
'        If InitTable(rsTemp, "审批编号|C|50||项目编号|C|50||项目名称|C|100") = False Then Exit Function
'        With rsTemp
'            For i = 0 To UBound(strArr) Step 3
'                .AddNew
'                !审批编号 = strArr(i)
'                !项目编码 = strArr(i + 1)
'                !项目名称 = strArr(i + 2)
'                .Update
'            Next
'        End With
'        '选择一个有效的审批项目
'        If frmListSel.ShowSelect(rsTemp, "审批编号", "请选择一个有效的审批项目编码", "所有", False) = False Then
'            .审批编号 = ""
'            .项目编码 = ""
'            .项目名称 = ""
'            Exit Function
'        End If
'        .审批编号 = Nvl(rsTemp!审批编号)
'        .项目编码 = Nvl(rsTemp!项目编码)
'        .项目名称 = Nvl(rsTemp!项目名称)
    End With
    获得通过的审批信息 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function InitTable(ByRef rsTemp As ADODB.Recordset, Optional ByVal strFields As String = "编码|C|30||名称|C|50") As Boolean
    Dim strArr, strArr1
    Dim i As Long
    Err = 0
    On Error GoTo errHand:
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = 1 Then .Close
        strArr = Split(strFields, "||")
        For i = 0 To UBound(strArr)
            strArr1 = Split(strArr(i), "|")
            Select Case strArr1(1)
            Case "C"
                .Fields.Append strArr1(0), adLongVarChar, Val(strArr1(2))
            Case "N"
                .Fields.Append strArr1(0), adDouble, Val(strArr1(2)), adFldIsNullable
            Case Else
            End Select
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    InitTable = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function 病人登记处理(ByRef str交易流水号 As String) As Boolean
    '进行门诊登记
    Dim StrInput As String, strOutput As String
    '交易特定输入数据：住院（门诊）号|医保证号码|IC卡号|入院日期|入院疾病名称|经办人
    '                  住院（门诊）号|个人编号|IC卡号|医疗类别|入院日期|入院疾病名称|科室|经办人
    With g病人身份_奉庆
        StrInput = .门诊号 & "|"
        StrInput = StrInput & .医保证号 & "|"
        StrInput = StrInput & .卡号 & "|"
        StrInput = StrInput & .医疗类别 & "|"
        StrInput = StrInput & "" & "|"
        StrInput = StrInput & "" & "|"
        StrInput = StrInput & "" & "|"
        StrInput = StrInput & gstrUserName
    End With
    DebugTool "进入病人登记"
    Err = 0
    On Error GoTo errHand:
    If 业务请求_奉庆(奉庆_病人登记, StrInput, strOutput) = False Then Exit Function
    DebugTool "病人登记成功"
    str交易流水号 = Split(strOutput, "|")(2)
    
    DebugTool "病人登记成功"
    病人登记处理 = True
    Exit Function
errHand:
    DebugTool "Err:" & Err.Description
    If ErrCenter = 1 Then Resume
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    DebugTool "进入身份验证,并开始加入基本信息"
    If Load医疗类别 = False Then
        DebugTool "加入失败(身份验证)"
        Exit Function
    End If
    Call Load审批类别
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
    With g病人身份_奉庆
        lblEdit(0).Caption = .卡号
        lblEdit(1).Caption = .医保证号
        lblEdit(2).Caption = .姓名
        lblEdit(3).Caption = Decode(.性别, "1", "男", "2", "女", .性别)
        
        lblEdit(4).Caption = .年龄
        lblEdit(5).Caption = .出生日期
        
        lblEdit(6).Caption = .身份证号
        lblEdit(7).Caption = .类别编码
        lblEdit(8).Caption = .类别名称
        lblEdit(9).Caption = Format(.帐户余额, "####0.00;#####0.00; ;")
        lblEdit(10).Caption = .单位名称 & IIf(.单位编码 <> "", "(" & .单位编码 & ")", "")
        lblEdit(11).Caption = .统筹区号
    End With
    
    gstrSQL = "Select 病种ID,密码 from 保险帐户 where 医保号='" & g病人身份_奉庆.医保证号 & "' and 险类=" & TYPE_临沧奉庆
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取相关病种"
    If rsTemp.EOF Then Exit Sub
    g病人身份_奉庆.医疗类别 = Nvl(rsTemp!密码)
    
    gstrSQL = "Select * From 医保病种目录 where ID=" & Nvl(rsTemp!病种ID, 0)
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.ProductName, "获取病种信息", gstrSQL)
    rsTemp.Open gstrSQL, gcnOracle_奉庆
    Call SQLTest
    If rsTemp.EOF Then
        Exit Sub
    End If
    txt病种.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
    txt病种.Tag = Nvl(rsTemp!ID, 0)
    Dim i As Long
    For i = 0 To cbo类别.ListCount - 1
        If InStr(1, cbo类别.List(i), g病人身份_奉庆.医疗类别 & "-") <> 0 Then
            cbo类别.ListIndex = i
            Exit For
        End If
    Next
    
End Sub
Private Sub Form_Load()
        'mblnFirst
        mblnFirst = True
End Sub

Private Sub mshSelect_Click()
    With mshSelect
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            SetColumnSort mshSelect, mintPreCol, mintsort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshSelect_DblClick()
    With mshSelect
        If .Row > 0 And .TextMatrix(.Row, 0) <> "" Then
            mshSelect_KeyPress 13
        End If
    End With
End Sub

Private Sub mshSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshSelect
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .COL = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .COL = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .COL = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .COL = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sngWidth = sngWidth + .ColWidth(i)
                    If sngWidth > .Width Then
                        .LeftCol = i + 1
                        .COL = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
End Sub


'对列头进行排序
Private Sub SetColumnSort(ByVal mshFilter As MSHFlexGrid, ByRef intPreCol As Integer, ByRef intPreSort As Integer)
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshFilter
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .COL = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
            If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
               .Sort = flexSortStringNoCaseAscending
               intPreSort = flexSortStringNoCaseAscending
            Else
               .Sort = flexSortStringNoCaseDescending
               intPreSort = flexSortStringNoCaseDescending
            End If
            intPreCol = intCol
            .Row = FindRow(mshFilter, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .COL = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub


Private Sub txt病种_Change()
    txt病种.Tag = ""
End Sub

Private Sub txt病种_GotFocus()
    OpenIme GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "")
    zlControl.TxtSelAll txt病种
End Sub

Private Sub txt病种_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    If KeyCode = vbKeyReturn Then
        If Me.txt病种 = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        If Trim(txt病种) = "" Then Exit Sub
        If Trim(txt病种.Tag) <> "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        txt病种 = UCase(txt病种)
    
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "" & _
            "   Select id, 编码, 名称, 助记码, to_char(变更时间,'yyyy-mm-dd hh24:mi:ss') as 变更时间" & _
            "   From 医保病种目录" & _
            "   Where " & zlCommFun.GetLike("", "编码", Me.txt病种) & " Or " & _
                        zlCommFun.GetLike("", "名称", Me.txt病种) & " Or " & _
                        zlCommFun.GetLike("", "助记码", Me.txt病种)
                       
        
        With rsTemp
            .CursorLocation = adUseClient
            .Open gstrSQL, gcnOracle_奉庆
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = txt病种.Top - .Height
                    .Left = txt病种.Left + txt病种.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 2000
                    .ColWidth(3) = 1400
                    .ColWidth(4) = .Width - .ColWidth(1) - .ColWidth(2) - .ColWidth(3) - .ColWidth(4)
                    .Row = 1
                    .COL = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt病种 = "[" & Nvl(!编码) & "]" & IIf(IsNull(!名称), "", !名称)
                txt病种.Tag = Nvl(!ID)
                zlCommFun.PressKey vbKeyTab
            End If
        End With
    End If
End Sub

Private Sub txt病种_LostFocus()
    OpenIme ""
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            txt病种.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
            txt病种.Tag = .TextMatrix(.Row, 0)
            If cmd确定.Enabled Then cmd确定.SetFocus
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub
'寻找与某一单元值相等的行
Private Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .Rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function

Private Function Load医疗类别() As Boolean
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "" & _
        "   Select * From 医保医疗类别 " & _
        "   where " & IIf(mbytType = 0, "  nvl(标志,0)=0", "  nvl(标志,0)=1") & _
        "   Order by 编码"
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption & "医保医疗类别"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "不存医疗医疗类别，请与HIS供应商联系!"
        Exit Function
    End If
    
    With rsTemp
        cbo类别.Clear
        Do While Not .EOF
            cbo类别.AddItem Nvl(!编码) & "--" & Nvl(!名称)
            .MoveNext
        Loop
    End With
    cbo类别.ListIndex = 0
    Load医疗类别 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
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
    With g病人身份_奉庆
        .姓名 = ""
        .性别 = ""
        .身份证号 = ""
        .出生日期 = ""
        .类别编码 = ""
        .类别名称 = ""
        .单位名称 = ""
        .单位编码 = ""
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub

Private Sub Load审批类别()
        With cbo审批类别
            .Clear
            .AddItem "1-特殊检查"
            .ListIndex = .NewIndex
            .AddItem "2-特殊治疗"
            .AddItem "3-急诊抢救"
            .AddItem "4-门诊慢性病"
            .AddItem "5-转诊"
            .AddItem "6-转院"
            .AddItem "7-贵重药品"
        End With
End Sub
