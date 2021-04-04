VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIdentify重庆渝北 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmIdentify重庆渝北.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Caption         =   "其他"
      Height          =   1215
      Index           =   2
      Left            =   90
      TabIndex        =   37
      Top             =   3510
      Width           =   7305
      Begin VB.CommandButton cmd病种 
         Caption         =   "…"
         Height          =   285
         Index           =   2
         Left            =   6915
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   870
         Width           =   255
      End
      Begin VB.CommandButton cmd病种 
         Caption         =   "…"
         Height          =   285
         Index           =   1
         Left            =   6915
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   525
         Width           =   255
      End
      Begin VB.CommandButton cmd病种 
         Caption         =   "…"
         Height          =   285
         Index           =   0
         Left            =   6915
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
      End
      Begin VB.TextBox txt病种 
         Height          =   300
         Index           =   0
         Left            =   810
         TabIndex        =   39
         Top             =   165
         Width           =   6390
      End
      Begin VB.TextBox txt病种 
         Height          =   300
         Index           =   1
         Left            =   810
         TabIndex        =   42
         Top             =   510
         Width           =   6390
      End
      Begin VB.TextBox txt病种 
         Height          =   300
         Index           =   2
         Left            =   810
         TabIndex        =   45
         Top             =   855
         Width           =   6390
      End
      Begin VB.Label lbl病种 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病种2(&3)"
         Height          =   180
         Index           =   2
         Left            =   75
         TabIndex        =   46
         Top             =   915
         Width           =   720
      End
      Begin VB.Label lbl病种 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病种1(&2)"
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   43
         Top             =   570
         Width           =   720
      End
      Begin VB.Label lbl病种 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病种(&1)"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   40
         Top             =   225
         Width           =   630
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3525
      Left            =   -405
      TabIndex        =   36
      Top             =   5580
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
   Begin VB.CommandButton cmd修改密码 
      Caption         =   "修改密码"
      Height          =   350
      Left            =   300
      TabIndex        =   35
      Top             =   4980
      Width           =   1100
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5220
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   675
      Width           =   2175
   End
   Begin VB.TextBox TxtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   930
      MaxLength       =   20
      TabIndex        =   1
      Top             =   675
      Width           =   2385
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -435
      TabIndex        =   22
      Top             =   4815
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   510
      Width           =   8340
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4995
      TabIndex        =   6
      Top             =   4980
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6285
      TabIndex        =   7
      Top             =   4980
      Width           =   1100
   End
   Begin VB.ComboBox cbo类别 
      Height          =   300
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1020
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "支付类别"
      Height          =   180
      Index           =   14
      Left            =   195
      TabIndex        =   4
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Index           =   13
      Left            =   4830
      TabIndex        =   2
      Top             =   720
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   5220
      TabIndex        =   34
      Top             =   2745
      Width           =   2175
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   5220
      TabIndex        =   33
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   930
      TabIndex        =   32
      Top             =   3105
      Width           =   6480
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   5220
      TabIndex        =   31
      Top             =   2055
      Width           =   2175
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   930
      TabIndex        =   30
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   930
      TabIndex        =   29
      Top             =   2745
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   5220
      TabIndex        =   28
      Top             =   1710
      Width           =   2175
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   930
      TabIndex        =   27
      Top             =   2055
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   930
      TabIndex        =   26
      Top             =   1710
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5220
      TabIndex        =   25
      Top             =   1365
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   930
      TabIndex        =   24
      Top             =   1365
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5235
      TabIndex        =   23
      Top             =   1035
      Width           =   2175
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   12
      Left            =   4470
      TabIndex        =   19
      Top             =   2790
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗补助类别"
      Height          =   180
      Index           =   11
      Left            =   4110
      TabIndex        =   16
      Top             =   2445
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位名称"
      Height          =   180
      Index           =   10
      Left            =   210
      TabIndex        =   18
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗照顾类别"
      Height          =   180
      Index           =   9
      Left            =   4110
      TabIndex        =   14
      Top             =   2100
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   8
      Left            =   570
      TabIndex        =   15
      Top             =   2445
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位编码"
      Height          =   180
      Index           =   7
      Left            =   210
      TabIndex        =   17
      Top             =   2790
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗人员类别"
      Height          =   180
      Index           =   6
      Left            =   4110
      TabIndex        =   12
      Top             =   1755
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出身日期"
      Height          =   180
      Index           =   5
      Left            =   210
      TabIndex        =   13
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   4
      Left            =   210
      TabIndex        =   11
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   4830
      TabIndex        =   10
      Top             =   1410
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   2
      Left            =   570
      TabIndex        =   9
      Top             =   1410
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "个人编号"
      Height          =   180
      Index           =   1
      Left            =   4485
      TabIndex        =   8
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保卡号"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   720
      Width           =   720
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify重庆渝北.frx":000C
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "通过IC卡验证人员身份，并将验证结果信息显示出来。"
      Height          =   180
      Left            =   630
      TabIndex        =   21
      Top             =   270
      Width           =   4320
   End
End
Attribute VB_Name = "frmIdentify重庆渝北"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐

Private mlng病人ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
'API的医保接口声明
Private Type Struct
    lngAppCode  As Long   '标志服务执行状态代码。等于1时表示服务执行正常结束，小于0时表示服务执行异常或错误。
    strErrMsg  As String  '当服务执行状态代码AppCod小于0时，描述服务执行的异常或错误信息。
End Type
'获取就诊编号
Private Declare Function GetAKC190 Lib "YHMdcrAsistntSvr.dll" Alias "_GetAKC190@12" (ByVal strYab003 As String, ByRef strAkc190 As String, ByRef tmpStrut As Struct) As Boolean

Dim mblnChange As Boolean
Private Sub cbo类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
 
Private Sub cmd病种_Click(Index As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    If mbytType = 0 Or mbytType = 3 Then
        strSQL = " and 性质=1"
    ElseIf mbytType = 2 Then
        strSQL = ""
    Else
        strSQL = " And 性质=2"
    End If

    gstrSQL = "" & _
        "   Select id, 编码, 名称, 支付类别, 助记码, 病种结算办法, 经办构构代码 " & _
        "   From 医保病种目录" & _
        "   where rownum<=2000" & strSQL & _
        "   Order by 编码 "
    
    With rsTemp
        If .State = 1 Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
        .Open gstrSQL, gcnOracle_CQYB
        Call SQLTest
        
        If .EOF Then
            MsgBox "不存在任何病种,请下载！", vbInformation, gstrSysName
            Exit Sub
        End If
        If .RecordCount > 1 Then
            Set mshSelect.Recordset = rsTemp
            mshSelect.Tag = Index
            With mshSelect
                .Top = txt病种(Index).Top - .Height
                .Left = txt病种(Index).Left + txt病种(Index).Width - .Width
                .Visible = True
                .SetFocus
                .ColWidth(0) = 0
                .ColWidth(1) = 800
                .ColWidth(2) = 2000
                .ColWidth(3) = 1400
                .ColWidth(4) = 1000
                .ColWidth(5) = 1400
                .ColWidth(6) = 2000
                .Row = 1
                .COL = 0
                .ColSel = .Cols - 1
                Exit Sub
                
            End With
        Else
            txt病种(Index) = "[" & Nvl(!编码) & "]" & IIf(IsNull(!名称), "", !名称)
            txt病种(Index).Tag = Nvl(!ID)
            zlCommFun.PressKey vbKeyTab
        End If
    End With
End Sub

Private Sub cmd修改密码_Click()
    Dim strOldPassWord As String
    Dim strNewPassWord As String
    
    strNewPassWord = frm修改密码.ChangePassword(strOldPassWord, strOldPassWord)
    If strOldPassWord = strNewPassWord Then Exit Sub
    If strNewPassWord = "" Then Exit Sub
      
    If 修改密码_重庆渝北(strOldPassWord, strNewPassWord) = True Then
        g病人身份_重庆渝北.密码 = strNewPassWord
        cmd确定_Click
        Unload Me
        Exit Sub
    End If
End Sub



Private Sub txtEdit_Change(Index As Integer)
    If Index = 1 Then
        txtEdit(Index).Tag = ""
    End If
    If Index = 0 And mblnChange = False Then
        g病人身份_重庆渝北.个人编号 = ""
        g病人身份_重庆渝北.卡号 = ""
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strCurrDate As String
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    mblnChange = True
    If Index = 0 Then
        SetOKCtrl False
        mblnChange = True
    ElseIf Index = 1 Then
        '密码输入完毕
        '需获取病人信息
         SetOKCtrl False
        
        '需解析卡内数据
        If 解析卡_重庆渝北 = False Then
            Exit Sub
        End If
         If Trim(txtEdit(Index)) = "" Then
            If mbytType = 0 Then
                '如果是门诊,需检查是否当前就诊的,且已经存在该帐户时,不重新输入密码.
                                
                '取密码
                 gstrSQL = "Select 密码,就诊时间 From 保险帐户  where 险类=" & TYPE_重庆渝北 & " and 医保号='" & g病人身份_重庆渝北.个人编号 & "'"
                 zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
                 
                 If rsTemp.RecordCount = 0 Then
                     ShowMsgbox "请输入密码!"
                    txtEdit(Index).SetFocus
                     Exit Sub
                 End If
                 If Format(rsTemp!就诊时间, "yyyy-mm-dd") <> Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
                    ShowMsgbox "请输入密码!"
                    txtEdit(Index).SetFocus
                    Exit Sub
                 End If
                 txtEdit(Index) = Trim(Nvl(rsTemp!密码))
                 If txtEdit(Index) = "" Then
                    ShowMsgbox "请输入密码!"
                    txtEdit(Index).SetFocus
                    Exit Sub
                 End If
            Else
                ShowMsgbox "请输入密码!"
                txtEdit(Index).SetFocus
                Exit Sub
            End If
         End If
         
        txtEdit(0).Text = g病人身份_重庆渝北.卡号
        lblEdit(1).Caption = g病人身份_重庆渝北.个人编号
         
         g病人身份_重庆渝北.密码 = Trim(txtEdit(Index))
        If g病人身份_重庆渝北.卡号 = "" Then
            g病人身份_重庆渝北.卡号 = Trim(txtEdit(0).Text)
        End If
        If 身份鉴别_重庆渝北 = False Then
            Exit Sub
        End If
        
        If g病人身份_重庆渝北.姓名 = "" Then
            ShowMsgbox "无效的用户验证,请核查!"
            Exit Sub
        End If
        
        
        '如果是门诊,需先进行挂号处理,否则是不能进行相应的处理的.
        If mbytType = 0 Then
            strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            gstrSQL = "Select 1 From 门诊费用记录 " & _
                    "   where 记录状态=1 and 记录性质=4  and rownum<=1 and 登记时间 between to_date('" & strCurrDate & " 00:00:00','yyyy-mm-dd hh24:mi:ss') and to_date('" & strCurrDate & " 23:59:59','yyyy-mm-dd hh24:mi:ss') and 病人id in (select 病人id From 保险帐户  where 险类=" & TYPE_重庆渝北 & " and 医保号='" & g病人身份_重庆渝北.个人编号 & "')"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
            If rsTemp.RecordCount = 0 Then
                ShowMsgbox "该医保病人未进行挂号,不能进行门诊结算!"
                Exit Sub
            End If
        End If
        '初始值
        Call LoadCtrlData
        SetOKCtrl True
    End If
    zlCommFun.PressKey vbKeyTab
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
    IsValid = False
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "还没有输入医保卡号！", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    
    If Trim(g病人身份_重庆渝北.姓名) = "" Then
        MsgBox "还没进行身份验证！", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Trim(txt病种(0)) <> "" And Val(txt病种(0).Tag) = 0 Then
        ShowMsgbox "病种选择错误,请重新选择!"
        txt病种(0).SetFocus
        Exit Function
    End If
    If cbo类别.Text = "" Then
        ShowMsgbox "支付类别未选择"
        Exit Function
    End If
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '不检查录前着态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重庆渝北, g病人身份_重庆渝北.个人编号)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If mbytType = 0 Or mbytType = 3 Then
            '设置
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
    
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str类别 As String
    Dim int当前状态 As Integer
    
        
    If IsValid = False Then Exit Sub
    
    lng疾病ID = Val(txt病种(0).Tag)
    
    If lng疾病ID <> 0 And txt病种(0).Text <> "" Then
        g病人身份_重庆渝北.病种编码 = Mid(txt病种(0).Text, 2, InStr(1, txt病种(0).Text, "]") - 2)
    Else
        g病人身份_重庆渝北.病种编码 = "000000"
    End If
    g病人身份_重庆渝北.病种ID = lng疾病ID
    
    g病人身份_重庆渝北.支付类别 = Mid(cbo类别.Text, 1, InStr(1, cbo类别.Text, "-") - 1)
    int当前状态 = 0
    If mbytType = 1 Then
        '入院:
        If lng疾病ID = 0 Then
            ShowMsgbox "必需输入病种,请检查!"
            If txt病种(0).Enabled Then txt病种(0).SetFocus
            Exit Sub
        End If
    
    End If
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_重庆渝北 & " and  医保号='" & g病人身份_重庆渝北.个人编号 & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
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
    With g病人身份_重庆渝北
        
        strIdentify = .卡号                               '0卡号
        strIdentify = strIdentify & ";" & .个人编号           '1医保号
        strIdentify = strIdentify & ";" & .密码                 '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & Decode(.性别, "1", "男", "2", "女", "未知")              '4性别
        strIdentify = strIdentify & ";" & .出生日期                '5出生日期
        strIdentify = strIdentify & ";" & .身份证号           '6身份证
        strIdentify = strIdentify & ";" & .单位名称 & IIf(.单位编码 = 0, "", "[" & .单位编码 & "]")          '7.单位名称(编码)
        strAddition = ";0"                                          '8.中心代码
        strAddition = strAddition & ";"                             '9.顺序号
        strAddition = strAddition & ";" & .社保经办构构代码          '10人员身份
        strAddition = strAddition & ";" & .帐户余额                 '11帐户余额
        
        strAddition = strAddition & ";" & int当前状态                            '12当前状态
        strAddition = strAddition & ";" & IIf(lng疾病ID = 0, "", lng疾病ID)             '13病种ID
        strAddition = strAddition & ";1"                            '14在职(1,2,3)
        strAddition = strAddition & ";" & .医疗人员类别 & "|" & .医疗照顾类别 & "|" & .医疗补助类别 & "|" & .累计缴费月数     '15退休证号
        strAddition = strAddition & ";" & .年龄                     '16年龄段
        strAddition = strAddition & ";"                             '17灰度级
        strAddition = strAddition & ";" & .帐户余额                             '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_重庆渝北)
    
    If mbytType = 3 Or mbytType = 1 Then
        '如果是挂号或入院登记,需确定新的就诊编号
        g病人身份_重庆渝北.就诊编号 = Get就诊编号_重庆渝北
        If g病人身份_重庆渝北.就诊编号 = "" Then
            ShowMsgbox "在获取就诊编号时为空了,请检查"
            Exit Sub
        End If
        
        '更新保险帐户的相关信息
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'就诊编号','''" & g病人身份_重庆渝北.就诊编号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊编号")
        
        If mbytType = 1 Then
            '为了保证先按普通入院再进行补充入院的就诊时间需更改.
             gstrSQL = "Select 入院日期 From 病案主页 where 病人id=" & mlng病人ID & " And 出院日期 is null"
             zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
             If Not rsTemp.EOF Then
                    '应该是补充登记
                    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'就诊时间','" & Format(rsTemp!入院日期, "yyyy-mm-dd HH:MM:SS") & "',1)"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊时间")
             End If
        End If
    Else
        '就诊时间还原
        '更新保险帐户的相关信息
        If g病人身份_重庆渝北.就诊时间 <> "" Then
            gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'就诊时间','" & g病人身份_重庆渝北.就诊时间 & "',1)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊时间")
        End If
    End If
    
    '取保险帐户中的就诊编号
     gstrSQL = "Select 就诊编号,就诊时间 From 保险帐户  where 病人id=" & mlng病人ID & " and 险类=" & TYPE_重庆渝北
     zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
     If rsTemp.RecordCount = 0 Then
         ShowMsgbox "在保险帐户中不存在该病人"
         Exit Sub
     End If
    g病人身份_重庆渝北.就诊编号 = Nvl(rsTemp!就诊编号)
    g病人身份_重庆渝北.就诊时间 = Format(rsTemp!就诊时间, "yyyy-MM-dd HH:mm:ss")
    g病人身份_重庆渝北.lng病人ID = mlng病人ID
    
    '更新保险帐户的相关信息
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'支付类别','''" & g病人身份_重庆渝北.支付类别 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊编号")
    
    '保存病种信息
    Call Save病情TO保险帐户
    
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
    If LoadBaseData = False Then
        DebugTool "加入失败(身份验证)"
        Exit Function
    End If
    DebugTool "加入成功(身份验证)"
    
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function
Private Function LoadBaseData() As Boolean
    '加载基础数据
    Dim rsTemp As New ADODB.Recordset
    LoadBaseData = False
    On Error GoTo errHand:
    
    With rsTemp
    
        .Open "Select * From 支付类别 where 标志=2 or 标志=" & IIf(mbytType = 3, 0, IIf(mbytType = 4, 1, mbytType)) & " order by 编码", gcnOracle_CQYB
        Do While Not .EOF
            cbo类别.AddItem Nvl(!编码) & "-" & Nvl(!名称)
            If !缺省 = 1 Then
                cbo类别.ListIndex = cbo类别.NewIndex
            End If
            .MoveNext
        Loop
        If cbo类别.ListIndex < 0 Then
            If cbo类别.ListCount <> 0 Then
                cbo类别.ListIndex = 0
            End If
        End If
    End With
    If cbo类别.ListCount = 0 Then
        ShowMsgbox "支付类别未初始化,请与系统管理员联系!"
        Exit Function
    End If
    LoadBaseData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    With g病人身份_重庆渝北
        lblEdit(2).Caption = .姓名
        lblEdit(3).Caption = Decode(.性别, "1", "男", "2", "女", "未知")
        lblEdit(4).Caption = .身份证号
        lblEdit(5).Caption = .出生日期
        lblEdit(6).Caption = Get代码数据_重庆渝北(医疗人员类别, .医疗人员类别)
        lblEdit(7).Caption = .单位编码
        lblEdit(8).Caption = .年龄
        '目前没有类别
        lblEdit(9).Caption = ""          'Get代码数据_重庆渝北(医疗照顾类别, .医疗照顾类别)
        lblEdit(10).Caption = .单位名称
        lblEdit(11).Caption = Get代码数据_重庆渝北(医疗补助类别, .医疗补助类别)
        lblEdit(12).Caption = Format(.帐户余额, "####0.00;#####0.00; ;")
    End With
    
    gstrSQL = "Select 病种ID,支付类别,就诊时间 from 保险帐户 where 医保号='" & g病人身份_重庆渝北.个人编号 & "' and 险类=" & TYPE_重庆渝北
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取相关病种"
    If rsTemp.EOF Then Exit Sub
    g病人身份_重庆渝北.支付类别 = Nvl(rsTemp!支付类别)
    g病人身份_重庆渝北.就诊时间 = Format(rsTemp!就诊时间, "yyyy-MM-dd HH:mm:ss")
    
    Call Load历史看病信息
    
    'gstrSQL = "Select * From 医保病种目录 where ID=" & Nvl(rsTemp!病种ID, 0)
    'If rsTemp.State = 1 Then rsTemp.Close
    'Call SQLTest(App.ProductName, "获取病种信息", gstrSQL)
    'rsTemp.Open gstrSQL, gcnOracle_CQYB
    'Call SQLTest
   ' If rsTemp.EOF Then
   '     Exit Sub
   ' End If
   ' txt病种(0).Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
   ' txt病种(0).Tag = Nvl(rsTemp!ID, 0)
    Dim i As Long
    For i = 0 To cbo类别.ListCount - 1
        If InStr(1, cbo类别.List(i), g病人身份_重庆渝北.支付类别 & "-") <> 0 Then
            cbo类别.ListIndex = i
            Exit For
        End If
    Next
    
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


Private Sub txt病种_Change(Index As Integer)
    txt病种(Index).Tag = ""
End Sub

Private Sub txt病种_GotFocus(Index As Integer)
    OpenIme GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "")
    zlControl.TxtSelAll txt病种(Index)
End Sub
 

Private Sub txt病种_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    Dim strKey As String
    
    If KeyCode = vbKeyReturn Then
        If Me.txt病种(Index) = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        If Trim(txt病种(Index)) = "" Then Exit Sub
        If Trim(txt病种(Index).Tag) <> "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        txt病种(Index) = UCase(txt病种(Index))
        strKey = txt病种(Index)
        
        Dim rsTemp As New ADODB.Recordset
        If mbytType = 0 Or mbytType = 3 Then
            strSQL = " and 性质=1"
        
        ElseIf mbytType = 2 Then
            strSQL = ""
        Else
            strSQL = " And 性质=2"
        End If
        gstrSQL = "" & _
            "   Select id, 编码, 名称, 支付类别, 助记码, 病种结算办法, 经办构构代码 " & _
            "   From 医保病种目录" & _
            "   Where (" & zlCommFun.GetLike("", "编码", strKey) & " Or " & _
                        zlCommFun.GetLike("", "名称", strKey) & " Or " & _
                        zlCommFun.GetLike("", "助记码", strKey) & ") " & strSQL
        With rsTemp
            If .State = 1 Then .Close
            Call SQLTest(App.ProductName, Me.Caption, strSQL)
            .Open gstrSQL, gcnOracle_CQYB
            Call SQLTest
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                mshSelect.Tag = Index
                With mshSelect
                    .Top = txt病种(Index).Top + fra(2).Top - .Height
                    .Left = txt病种(Index).Left + txt病种(Index).Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 2000
                    .ColWidth(3) = 1400
                    .ColWidth(4) = 1000
                    .ColWidth(5) = 1400
                    .ColWidth(6) = 1400
                    .Row = 1
                    .COL = 0
                    .ColSel = .Cols - 1
                    .ZOrder 0
                    Exit Sub
                    
                End With
            Else
                txt病种(Index) = "[" & Nvl(!编码) & "]" & IIf(IsNull(!名称), "", !名称)
                txt病种(Index).Tag = Nvl(!ID)
                zlCommFun.PressKey vbKeyTab
            End If
        End With
    End If
End Sub

Private Sub txt病种_LostFocus(Index As Integer)
    OpenIme ""
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    Dim Index As Integer
    With mshSelect
        Index = Val(mshSelect.Tag)
        If KeyAscii = 13 Then
            txt病种(Index).Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
            txt病种(Index).Tag = .TextMatrix(.Row, 0)
            
            If Index < 2 Then
                Index = Index + 1
                If txt病种(Index).Enabled Then txt病种(Index).SetFocus
            Else
                If cmd确定.Enabled Then cmd确定.SetFocus
            End If
            
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
Public Function Save病情TO保险帐户() As Boolean
    '--------------------------------------------------------------------------------------------------
    '功能:将病情信息保存到保险帐户中
    '--------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    '功能:保存
    Err = 0: On Error GoTo errHand:
    
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'病情id','''" & Val(txt病种(0).Tag) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "病情ID")
    
    
    
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'病情1id','" & IIf(Val(txt病种(1).Tag) = 0, "NULL", Val(txt病种(1).Tag)) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "病情ID")
    
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'病情2id','" & IIf(Val(txt病种(2).Tag) = 0, "NULL", Val(txt病种(2).Tag)) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "病情ID")
    
    '保存疾病
    If Val(txt病种(0).Tag) <> 0 Then
        gstrSQL = "select 编码,名称 From 医保病种目录 where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Val(txt病种(0).Tag)))
        
        g病人身份_重庆渝北.病情编码 = Nvl(rsTemp!编码)
        g病人身份_重庆渝北.病情名称 = Nvl(rsTemp!名称)
        g病人身份_重庆渝北.病情ID = Val(txt病种(0).Tag)
    Else
        g病人身份_重庆渝北.病情编码 = ""
        g病人身份_重庆渝北.病情ID = 0
        g病人身份_重庆渝北.病情名称 = Trim(txt病种(0).Text)
        gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'病情1','''" & Trim(txt病种(0).Text) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "病情")
    End If
    If Val(txt病种(1).Tag) <> 0 Then
        gstrSQL = "select 编码,名称 From 医保病种目录 where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Val(txt病种(1).Tag)))
        
        g病人身份_重庆渝北.病情编码1 = Nvl(rsTemp!编码)
        g病人身份_重庆渝北.病情名称1 = Nvl(rsTemp!名称)
        g病人身份_重庆渝北.病情1ID = Val(txt病种(1).Tag)
        
    Else
        g病人身份_重庆渝北.病情编码1 = ""
        g病人身份_重庆渝北.病情名称1 = Trim(txt病种(1).Text)
        g病人身份_重庆渝北.病情1ID = 0
        
        gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'病情2','''" & Trim(txt病种(1).Text) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "病情")
    End If
    If Val(txt病种(2).Tag) <> 0 Then
        gstrSQL = "select 编码,名称 From 医保病种目录 where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Val(txt病种(2).Tag)))
        g病人身份_重庆渝北.病情编码2 = Nvl(rsTemp!编码)
        g病人身份_重庆渝北.病情名称2 = Nvl(rsTemp!名称)
        g病人身份_重庆渝北.病情2ID = Val(txt病种(2).Tag)
    Else
        g病人身份_重庆渝北.病情编码2 = ""
        g病人身份_重庆渝北.病情名称2 = Trim(txt病种(2).Text)
        g病人身份_重庆渝北.病情2ID = 0
        gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_重庆渝北 & ",'病情3','''" & Trim(txt病种(2).Text) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "病情")
    End If
    Save病情TO保险帐户 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function Load历史看病信息() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载历史医保病人的看病信息
    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    If g病人身份_重庆渝北.个人编号 = "" Then Exit Function
    
    Err = 0: On Error GoTo errHand:
    
    gstrSQL = "" & _
        "   Select  a.病情id,a.病情1id,a.病情2id,a.病情1,a.病情2,a.病情3" & _
        "   From 保险帐户 a" & _
        "   where a.险类=[1] and a.医保号=[2]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保其他信息", TYPE_重庆渝北, g病人身份_重庆渝北.个人编号)
    
    If rsTemp.EOF Then
        Exit Function
    End If
    
    txt病种(0).Text = Nvl(rsTemp!病情1)
    txt病种(0).Tag = Nvl(rsTemp!病情ID)
    
    txt病种(1).Text = Nvl(rsTemp!病情2)
    txt病种(1).Tag = Nvl(rsTemp!病情1ID)
    txt病种(2).Text = Nvl(rsTemp!病情3)
    txt病种(2).Tag = Nvl(rsTemp!病情2ID)
    
    Dim lng病种ID As Long
    If Val(txt病种(0).Tag) <> 0 Then
        lng病种ID = Val(txt病种(0).Tag)
        gstrSQL = "select 编码,名称 From 医保病种目录 where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病种ID)
        txt病种(0).Text = IIf(Nvl(rsTemp!编码) = "", "", "[" & Nvl(rsTemp!编码) & "]") & Nvl(rsTemp!名称)
        txt病种(0).Tag = lng病种ID
    End If
  
    If Val(txt病种(1).Tag) <> 0 Then
        lng病种ID = Val(txt病种(1).Tag)
        gstrSQL = "select 编码,名称 From 医保病种目录 where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病种ID)
        txt病种(1).Text = IIf(Nvl(rsTemp!编码) = "", "", "[" & Nvl(rsTemp!编码) & "]") & Nvl(rsTemp!名称)
        txt病种(1).Tag = lng病种ID
    End If
  
    If Val(txt病种(2).Tag) <> 0 Then
        lng病种ID = Val(txt病种(1).Tag)
        gstrSQL = "select 编码,名称 From 医保病种目录 where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病种ID)
        txt病种(2).Text = IIf(Nvl(rsTemp!编码) = "", "", "[" & Nvl(rsTemp!编码) & "]") & Nvl(rsTemp!名称)
        txt病种(2).Tag = lng病种ID
    End If
  
    Load历史看病信息 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function



