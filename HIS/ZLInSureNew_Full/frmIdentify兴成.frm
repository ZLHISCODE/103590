VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIdentify兴成 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo出院类别 
      Height          =   300
      Left            =   7110
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   4275
      Width           =   1155
   End
   Begin VB.ComboBox cbo住院类别 
      Height          =   300
      Left            =   5145
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   4275
      Width           =   1170
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   4410
      Left            =   -300
      TabIndex        =   59
      Top             =   5715
      Visible         =   0   'False
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   7779
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
   Begin VB.ComboBox cbo入院类别 
      Height          =   300
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   4275
      Width           =   3090
   End
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   285
      Left            =   7995
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4665
      Width           =   255
   End
   Begin VB.CommandButton cmd验卡 
      Caption         =   "重新读卡(&R)"
      Height          =   350
      Left            =   105
      TabIndex        =   18
      Top             =   5310
      Width           =   1305
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7110
      TabIndex        =   61
      Top             =   5310
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   19
      Top             =   615
      Width           =   8475
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -540
      TabIndex        =   17
      Top             =   5055
      Width           =   9030
   End
   Begin VB.TextBox txt病种 
      Height          =   300
      Left            =   840
      TabIndex        =   56
      Top             =   4650
      Width           =   7425
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5955
      TabIndex        =   60
      Top             =   5310
      Width           =   1100
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "出院类别"
      Height          =   180
      Index           =   1
      Left            =   6360
      TabIndex        =   53
      Top             =   4335
      Width           =   720
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "住院类别"
      Height          =   180
      Index           =   0
      Left            =   4380
      TabIndex        =   51
      Top             =   4335
      Width           =   720
   End
   Begin VB.Label lblInfor 
      Caption         =   "入院类别"
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   49
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   5130
      TabIndex        =   58
      Top             =   1005
      Width           =   3135
   End
   Begin VB.Label lblSel 
      AutoSize        =   -1  'True
      Caption         =   "病种"
      Height          =   180
      Left            =   435
      TabIndex        =   55
      Top             =   4710
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "慢病截止日期"
      Height          =   180
      Index           =   22
      Left            =   4020
      TabIndex        =   48
      Top             =   3945
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   22
      Left            =   5130
      TabIndex        =   47
      ToolTipText     =   "慢性病有效期截止日期"
      Top             =   3885
      Width           =   3135
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "慢病病种"
      Height          =   180
      Index           =   21
      Left            =   75
      TabIndex        =   46
      Top             =   3945
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   21
      Left            =   840
      TabIndex        =   45
      Top             =   3885
      Width           =   3075
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "慢性标识"
      Height          =   180
      Index           =   20
      Left            =   6360
      TabIndex        =   44
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   20
      Left            =   7125
      TabIndex        =   43
      ToolTipText     =   "慢性病患者标识"
      Top             =   3495
      Width           =   1140
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "住院日期"
      Height          =   180
      Index           =   19
      Left            =   4380
      TabIndex        =   42
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   19
      Left            =   5130
      TabIndex        =   41
      ToolTipText     =   "住院信息更新日期"
      Top             =   3495
      Width           =   1110
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "住院次数"
      Height          =   180
      Index           =   18
      Left            =   2160
      TabIndex        =   40
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   18
      Left            =   2880
      TabIndex        =   39
      ToolTipText     =   "本年住院次数"
      Top             =   3495
      Width           =   1020
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "门诊日期"
      Height          =   180
      Index           =   17
      Left            =   75
      TabIndex        =   38
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   17
      Left            =   840
      TabIndex        =   37
      ToolTipText     =   "上次门诊就医日期"
      Top             =   3495
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "慢病累计"
      Height          =   180
      Index           =   15
      Left            =   4380
      TabIndex        =   36
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   16
      Left            =   7125
      TabIndex        =   35
      ToolTipText     =   "本年个人帐户支出累计"
      Top             =   3105
      Width           =   1140
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户累计"
      Height          =   180
      Index           =   16
      Left            =   6360
      TabIndex        =   34
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   5130
      TabIndex        =   33
      ToolTipText     =   "本年慢性病费用累计"
      Top             =   3105
      Width           =   1110
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "共付段累计"
      Height          =   180
      Index           =   14
      Left            =   1980
      TabIndex        =   32
      Top             =   3150
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   2880
      TabIndex        =   31
      ToolTipText     =   "共付段金额累计"
      Top             =   3105
      Width           =   1020
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "基金累计"
      Height          =   180
      Index           =   13
      Left            =   75
      TabIndex        =   30
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   840
      TabIndex        =   29
      ToolTipText     =   "本年统筹基金支付累计"
      Top             =   3105
      Width           =   990
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   12
      Left            =   7110
      TabIndex        =   28
      ToolTipText     =   "上次充卡日期"
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   5130
      TabIndex        =   27
      ToolTipText     =   "累计划入个人帐户余额"
      Top             =   2693
      Width           =   1110
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "充卡日期"
      Height          =   180
      Index           =   12
      Left            =   6360
      TabIndex        =   26
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "累计划入帐户"
      Height          =   180
      Index           =   11
      Left            =   4020
      TabIndex        =   25
      Top             =   2745
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   2880
      TabIndex        =   24
      Top             =   2693
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   840
      TabIndex        =   23
      Top             =   2693
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "参保日期"
      Height          =   180
      Index           =   10
      Left            =   2160
      TabIndex        =   22
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "保险编号"
      Height          =   180
      Index           =   9
      Left            =   75
      TabIndex        =   21
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "人员类别"
      Height          =   180
      Index           =   3
      Left            =   4380
      TabIndex        =   3
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
      Height          =   180
      Index           =   8
      Left            =   4380
      TabIndex        =   15
      Top             =   2325
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   8
      Left            =   5130
      TabIndex        =   16
      Top             =   2265
      Width           =   3135
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1005
      Width           =   3075
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "医保病人基本信息显示，可以通过[重新读卡]按钮重新进行读取病人基本信息。"
      Height          =   180
      Index           =   0
      Left            =   630
      TabIndex        =   20
      Top             =   360
      Width           =   6300
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify兴成.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡号"
      Height          =   180
      Index           =   0
      Left            =   435
      TabIndex        =   0
      Top             =   1065
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "社会保障号"
      Height          =   180
      Index           =   1
      Left            =   4230
      TabIndex        =   5
      Top             =   1065
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   2
      Left            =   435
      TabIndex        =   2
      Top             =   1515
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   4
      Left            =   435
      TabIndex        =   7
      Top             =   1890
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   7
      Left            =   75
      TabIndex        =   13
      Top             =   2325
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   6
      Left            =   4380
      TabIndex        =   11
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   5
      Left            =   2520
      TabIndex        =   9
      Top             =   1890
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   840
      TabIndex        =   6
      Top             =   1425
      Width           =   3075
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   5130
      TabIndex        =   4
      Top             =   1425
      Width           =   3135
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   8
      Top             =   1838
      Width           =   990
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   10
      Top             =   1838
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   840
      TabIndex        =   14
      Top             =   2265
      Width           =   3075
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   5130
      TabIndex        =   12
      Top             =   1830
      Width           =   3135
   End
End
Attribute VB_Name = "frmIdentify兴成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐,88-按住院信息进行查询

Private mlng病人ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '第一次起动系统时调用
Private mblnChange As Boolean

Private Sub cbo出院类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub


Private Sub cbo入院类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cbo住院类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
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
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    cmd确定.Enabled = False
    Call LoadBase
    If 获取参保人员信息 = False Then
         Call ClearData
         Exit Sub
     End If
     Call LoadCtrlData
     Call InitCtlData
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
    
    IsValid = False
    If Trim(g病人身份_兴成.姓名) = "" Then
        MsgBox "还没进行身份验证！", vbInformation, gstrSysName
        If cmd验卡.Enabled Then cmd验卡.SetFocus
        Exit Function
    End If
    
      
    If mbytType <> 2 And mbytType <> 88 Then
        If mbytType = 4 Then
            '不检查当前状态
        Else
           '陈宏悦于20051231修改增加，由于医保中心增加异地就医:要求异地卡不能在本地办理住院登记
            If mbytType = 1 Then
               If g病人身份_兴成.异地卡标志 = "1" Then
                  MsgBox "该卡是异地卡，不能在本市住院！", vbOKOnly + vbExclamation, gstrSysName
                  Exit Function
               End If
            End If
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_兴成核工业, g病人身份_兴成.社会保障号)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
         '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_兴成核工业 & " and  医保号='" & g病人身份_兴成.社会保障号 & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
        Else
            mlng病人ID = 0
        End If
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
    
    Dim int当前状态 As Integer
    Dim lng状态 As Long
    Dim str病种名称 As String
    
    
    If IsValid = False Then Exit Sub
    
    int当前状态 = 0
    
    
    '陈宏悦于20050310修改：由于接口文档错误，需要将“YB_HHMD”结构表的字段由“yyjb”修改为“kh”
    
    gstrSQL = "Select * From YB_HHMD where kh='" & g病人身份_兴成.IC卡号 & "'"
    
    rsTemp.Open gstrSQL, gcnSQLSEVER_兴成
    If rsTemp.EOF Then
        g病人身份_兴成.卡状态 = "a"     '正常
    Else
        Select Case Val(Nvl(rsTemp!Kzt))
        Case 0 '挂失
            ShowMsgbox "注意：" & vbCrLf & "该卡已经被挂失!"
        Case 1 '欠费
            ShowMsgbox "注意：" & vbCrLf & "该卡已经欠费!"
        Case 2 '停保
            ShowMsgbox "注意：" & vbCrLf & "该卡已经停保!"
        Case 3 '报损
            ShowMsgbox "注意：" & vbCrLf & "该卡已经报损!"
        End Select
        g病人身份_兴成.卡状态 = Val(Nvl(rsTemp!Kzt))
    End If
    If rsTemp.State = 1 Then rsTemp.Close
        
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_兴成核工业 & " and  医保号='" & g病人身份_兴成.社会保障号 & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
        End If
        rsTemp.Close
    End If
    
    If txt病种.Tag <> "" Then
        g病人身份_兴成.病种编码 = txt病种.Tag
        str病种名称 = Split(txt病种.Text & "]", "]")(1)
    Else
        g病人身份_兴成.病种编码 = ""
        str病种名称 = ""
    End If
    
    If mbytType <> 1 And mbytType <> 4 Then
        g病人身份_兴成.入院类别 = ""
        g病人身份_兴成.住院类别 = ""
        g病人身份_兴成.出院类别 = ""
    Else
        g病人身份_兴成.入院类别 = cbo入院类别.ItemData(cbo入院类别.ListIndex)
        g病人身份_兴成.住院类别 = cbo住院类别.ItemData(cbo住院类别.ListIndex)
        g病人身份_兴成.出院类别 = cbo出院类别.ItemData(cbo出院类别.ListIndex)
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_兴成
        strIdentify = .IC卡号                                '0卡号
        strIdentify = strIdentify & ";" & .社会保障号             '1医保号
        strIdentify = strIdentify & ";"                    '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & .性别             '4性别
        strIdentify = strIdentify & ";" & .出生日期                 '5出生日期
        strIdentify = strIdentify & ";" & .身份证号              '6身份证
        strIdentify = strIdentify & ";"          '7.单位名称(编码)
        strAddition = ";0"                      '8.中心代码
        strAddition = strAddition & ";" & .保险编号                                '9.顺序号
        strAddition = strAddition & ";" & .入院类别                                    '10人员身份
        strAddition = strAddition & ";" & .个人帐户余额                   '11帐户余额

        strAddition = strAddition & ";" & int当前状态                            '12当前状态
        strAddition = strAddition & ";"             '13病种ID
        strAddition = strAddition & ";1"                        '14在职(1,2,3)
        strAddition = strAddition & ";" & .人员类别           '15退休证号
        strAddition = strAddition & ";" & .年龄                     '16年龄段
        strAddition = strAddition & ";"                         '17灰度级
        strAddition = strAddition & ";" & .个人帐户余额                            '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_兴成核工业)
    
    If mlng病人ID = 0 Then Exit Sub

    If mbytType = 1 Or mbytType = 4 Or mbytType = 0 Then
        '更新附加信息(过程参数如下:)
        '   病人ID_IN,本年统筹基金支付累计,共付段金额累计,参保日期,统筹基金共付段金额累计,本年慢性病费用累计,统筹支付累计,
        '   累计划入个人帐户余额,上次充卡日期,本年个人帐户支出累计,当前余额,与帐户余额相同,
        '   上次门诊就医日期,本年住院次数,住院信息更新日期,慢性病患者标识,慢性病病种,慢性病有效期截止日期,病种代码,病种名称
        gstrSQL = "ZL_医保病人附加信息_UPDATE("
        gstrSQL = gstrSQL & "" & mlng病人ID & ","
        gstrSQL = gstrSQL & "'" & g病人身份_兴成.本年统筹支付 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_兴成.共付段累计 & "',"
        gstrSQL = gstrSQL & "" & IIf(IsDate(g病人身份_兴成.参保日期), "to_Date('" & g病人身份_兴成.参保日期 & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "" & g病人身份_兴成.统筹共付段累计 & ","
        gstrSQL = gstrSQL & "" & g病人身份_兴成.慢性病费用累计 & ","
        gstrSQL = gstrSQL & "" & g病人身份_兴成.统筹支付累计 & ","
        gstrSQL = gstrSQL & "" & g病人身份_兴成.累计划入个人帐户 & ","
        gstrSQL = gstrSQL & "" & IIf(IsDate(g病人身份_兴成.上次充卡日期), "to_Date('" & g病人身份_兴成.上次充卡日期 & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "" & g病人身份_兴成.帐户支出累计 & ","
        gstrSQL = gstrSQL & "" & g病人身份_兴成.当前余额 & ","
        gstrSQL = gstrSQL & "" & IIf(IsDate(g病人身份_兴成.上次门诊日期), "to_Date('" & g病人身份_兴成.上次门诊日期 & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "" & g病人身份_兴成.本年住院次数 + 1 & ","
        gstrSQL = gstrSQL & "" & IIf(IsDate(g病人身份_兴成.住院更新日期), "to_Date('" & g病人身份_兴成.住院更新日期 & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "'" & g病人身份_兴成.慢病标识 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_兴成.慢性病病种 & "',"
        gstrSQL = gstrSQL & "" & IIf(IsDate(g病人身份_兴成.慢性有效日期), "to_Date('" & g病人身份_兴成.慢性有效日期 & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "'" & g病人身份_兴成.病种编码 & "',"
        gstrSQL = gstrSQL & "'" & str病种名称 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_兴成.住院类别 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_兴成.出院类别 & "')"
        ExecuteProcedure_兴成 "医保病人附加信息"
    Else
    End If
    'g病人身份_兴成.病人ID = mlng病人ID
    If mbytType = 4 Then
    
    '陈宏悦于20050311修改，语句错误
    'gstrSQL = "Select AF21,AF22 From 医保病人附加信息 where 病人id=mlng病人ID "
        
      gstrSQL = "Select AF21,AF22 From 医保病人附加信息 where 病人id=" & mlng病人ID
      
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取相关的出院信息"
        If rsTemp.EOF Then
            g病人身份_兴成.异地医院 = ""
            g病人身份_兴成.异地医院级别 = ""
        Else
            g病人身份_兴成.异地医院 = Nvl(rsTemp!AF21)
            g病人身份_兴成.异地医院级别 = Nvl(rsTemp!AF22)
        End If
    Else
        g病人身份_兴成.异地医院 = ""
        g病人身份_兴成.异地医院级别 = ""
    End If
    
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    g病人身份_兴成.病人ID = mlng病人ID
    
    '陈宏悦于20050402添加，因为慢性病患者允许以非慢性病进行就医结算
    If mbytType = 0 And g病人身份_兴成.慢病标识 = "1" Then
     If MsgBox("该患者是否以慢性病方式进行医保结算？", vbOKCancel, "中联软件") = vbOK Then
        blnmxb = True
     Else
        blnmxb = False
     End If
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
    With g病人身份_兴成
        lblEdit(0).Caption = .IC卡号
        lblEdit(1).Caption = .社会保障号
        lblEdit(2).Caption = .姓名
        lblEdit(3).Caption = Decode(.人员类别, "10", "公务员", "11", "事业人员", "12", "企业人员", "13", "下岗人员", "14", "内退人员", "15", "停薪留职", "16", "放假人员", "20", "退休人员", "未知")

        lblEdit(4).Caption = .性别
        lblEdit(5).Caption = .年龄     '年龄
        lblEdit(6).Caption = Format(.个人帐户余额, "####0.00;-####0.00;;")
        lblEdit(7).Caption = .身份证号
        lblEdit(8).Caption = .出生日期      '出生日期
        lblEdit(9).Caption = .保险编号
        lblEdit(10).Caption = .参保日期
        lblEdit(11).Caption = Format(.累计划入个人帐户, "####0.00;-####0.00;;")
        lblEdit(12).Caption = .上次充卡日期

        lblEdit(13).Caption = Format(.本年统筹支付, "####0.00;-####0.00;;")
        lblEdit(14).Caption = Format(.共付段累计, "####0.00;-####0.00;;")
        lblEdit(15).Caption = Format(.慢性病费用累计, "####0.00;-####0.00;;")
        lblEdit(16).Caption = Format(.帐户支出累计, "####0.00;-####0.00;;")
        lblEdit(17).Caption = .上次门诊日期
        lblEdit(18).Caption = .本年住院次数
        lblEdit(19).Caption = .住院更新日期
        lblEdit(20).Caption = .慢病标识
        lblEdit(21).Caption = .慢性病病种
        lblEdit(22).Caption = .慢性有效日期
    End With
End Sub

Private Sub Form_Load()
        mblnFirst = True
End Sub


Private Function 获取参保人员信息() As Boolean
    '获取参保人员信息
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant, strArr1 As Variant
    
    获取参保人员信息 = False
    Err = 0:    On Error GoTo errHand:
        
    If 业务请求_兴成(兴成_获取持卡人信息, "", strOutput) = False Then
        Call ClearData
        Exit Function
    End If
    '       IC卡号|公民身份号码|姓名|性别|医疗参保人员类别|个人帐户余额|保险编号|本年统筹基金支付累计|共付段金额累计|异地卡标志
    strArr = Split(strOutput, "|")
    If 业务请求_兴成(兴成_JbylReadIC, "", strOutput) = False Then
        Call ClearData
    End If
    '       社会保障号|卡号|人员类别|参保日期|统筹基金共付段金额累计|本年慢性病费用累计|统筹支付累计|累计划入个人帐户余额|上次充卡日期|本年个人帐户支出累计|当前余额|上次门诊就医日期|本年住院次数|住院信息更新日期|慢性病患者标识|慢性病病种|慢性病有效期截止日期
    strArr1 = Split(strOutput, "|")
    
    With g病人身份_兴成
            .IC卡号 = strArr(0)
            .社会保障号 = strArr1(0)
            .身份证号 = strArr(1)
            .姓名 = strArr(2)
            .性别 = Decode(strArr(3), "1", "男", "0", "女", "9", "未知", strArr(3))
            .人员类别 = strArr(4)
            .个人帐户余额 = Val(strArr(5)) / 100
            .保险编号 = strArr(6)
            .本年统筹支付 = Val(strArr(7)) / 100       '本年统筹基本支付累计
            .共付段累计 = Val(strArr(8)) / 100         '共付段金额累计
            
            '陈宏悦于20051231修改，由于医保中心调整接口文档，增加异地就医卡
            
            .异地卡标志 = strArr(9)
            
            .参保日期 = zlCommFun.AddDate(strArr1(3))
            .统筹共付段累计 = Val(strArr1(4)) / 100   '统筹基金共付段金额累计
            .慢性病费用累计 = Val(strArr1(5)) / 100   '本年慢性病费用累计
            .统筹支付累计 = Val(strArr1(6)) / 100 '统筹支付累计
            .累计划入个人帐户 = Val(strArr1(7)) / 100    '累计划入个人帐户余额
            .上次充卡日期 = zlCommFun.AddDate(strArr1(8))
            .帐户支出累计 = Val(strArr1(9)) / 100     '本年个人帐户支出累计
            .当前余额 = Val(strArr1(10)) / 100
            .上次门诊日期 = zlCommFun.AddDate(strArr1(11))     '上次门诊就医日期
            .本年住院次数 = Val(strArr1(12))
            .住院更新日期 = zlCommFun.AddDate(strArr1(13))    '住院信息更新日期
            .慢病标识 = strArr1(14)          '慢性病患者标识
            If InStr(1, strArr1(15), "1") = 0 Then
               .慢性病病种 = strArr1(15)
            Else
             .慢性病病种 = Lpad(InStr(1, strArr1(15), "1"), 9, "0") '慢性病病种
            End If
            .慢性有效日期 = zlCommFun.AddDate(strArr1(16))      '慢性病有效期截止日期
            If Trim(.身份证号) = "" Then
                .出生日期 = ""
                .年龄 = 0
            Else
                .出生日期 = zlCommFun.GetIDCardDate(.身份证号)
                .年龄 = Get年龄(.出生日期)
            End If
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
    With g病人身份_兴成
        .IC卡号 = ""
        .社会保障号 = ""
        .身份证号 = ""
        .姓名 = ""
        .性别 = ""
        .人员类别 = ""
        .个人帐户余额 = 0
        .保险编号 = ""
        .本年统筹支付 = 0        '本年统筹基本支付累计
        .共付段累计 = 0          '共付段金额累计
        .参保日期 = ""
        .统筹共付段累计 = 0    '统筹基金共付段金额累计
        .慢性病费用累计 = 0    '本年慢性病费用累计
        .统筹支付累计 = 0 '统筹支付累计
        .累计划入个人帐户 = 0     '累计划入个人帐户余额
        .上次充卡日期 = ""
        .帐户支出累计 = 0      '本年个人帐户支出累计
        .当前余额 = 0
        .上次门诊日期 = ""     '上次门诊就医日期
        .本年住院次数 = 0
        .住院更新日期 = ""    '住院信息更新日期
        .慢病标识 = ""          '慢性病患者标识
        .慢性病病种 = ""       '慢性病病种
        .慢性有效日期 = ""     '慢性病有效期截止日期
        .年龄 = 0
        .出生日期 = ""
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub
Private Sub LoadBase()
    '加载数据
    Me.cbo入院类别.Clear
    Me.cbo住院类别.Clear
    Me.cbo出院类别.Clear
    With Me.cbo入院类别
        .AddItem "1-正常入院"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "2-市内转入"
        .ItemData(.NewIndex) = 2
        .AddItem "3-市外转入"
        .ItemData(.NewIndex) = 3
        .AddItem "4-因慢性病加重第一次住院"
        .ItemData(.NewIndex) = 4
    End With
    With Me.cbo住院类别
        .AddItem "0-正常住院"
        .ItemData(.NewIndex) = 0
        .ListIndex = .NewIndex
        .AddItem "1-紧急抢救"
        .ItemData(.NewIndex) = 1
    End With
    With Me.cbo出院类别
        .AddItem "1-正常出院"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "2-转往市内"
        .ItemData(.NewIndex) = 2
        .AddItem "3-转往市外"
        .ItemData(.NewIndex) = 3
    End With
    If mbytType = 1 Or mbytType = 4 Then
        Me.cbo入院类别.Enabled = mbytType = 1
        Me.cbo住院类别.Enabled = True
        Me.cbo出院类别.Enabled = mbytType <> 1
        If Me.cbo出院类别.Enabled Then
            Me.cbo出院类别.ListIndex = -1
        End If
    Else
        Me.cbo入院类别.ListIndex = -1: Me.cbo入院类别.Enabled = False
        Me.cbo住院类别.ListIndex = -1: Me.cbo住院类别.Enabled = False
        Me.cbo出院类别.ListIndex = -1: Me.cbo出院类别.Enabled = False
    End If
   ' Me.cbo出院类别.Enabled = False
End Sub

Private Sub cmd病种_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select jbdm 疾病代码,jbmc 疾病名称,case isnull(jblx,'0') when '0' then '普通病' else '慢性病' end  疾病类型 " & _
            "   From YB_BZML "
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnSQLSEVER_兴成
     
     With rsTemp
         If .EOF Then
             MsgBox "不存在任何病种,请下载！", vbInformation, gstrSysName
             Exit Sub
         End If
         
         If .RecordCount > 1 Then
             Set mshSelect.Recordset = rsTemp
             With mshSelect
                 .Cols = 3
                 .Top = txt病种.Top - .Height
                 .Left = txt病种.Left + txt病种.Width - .Width
                 .Visible = True
                 .SetFocus
                 .ColWidth(0) = 2000
                 .ColWidth(1) = 3000
                 .ColWidth(2) = .Width - .ColWidth(1)
                 .Row = 1
                 .COL = 0
                 .ColSel = .Cols - 1
                 Exit Sub
             End With
         Else
             txt病种 = "[" & Nvl(!疾病代码) & "]" & IIf(IsNull(!疾病名称), "", !疾病名称)
             txt病种.Tag = Nvl(!疾病代码)
             zlCommFun.PressKey vbKeyTab
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
                " Select jbdm 疾病代码,jbmc 疾病名称,case isnull(jblx,'0') when '0' then '普通病' else '慢性病' end  疾病类型 " & _
                " From YB_BZML " & _
                " Where " & zlCommFun.GetLike("", "jbdm", Me.txt病种) & " Or " & _
                        zlCommFun.GetLike("", "jbmc", Me.txt病种)

        With rsTemp
            .CursorLocation = adUseClient
            .Open gstrSQL, gcnSQLSEVER_兴成
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Cols = 3
                    .Top = txt病种.Top - .Height
                    .Left = txt病种.Left + txt病种.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 2000
                    .ColWidth(1) = 3000
                    .ColWidth(2) = .Width - .ColWidth(1) - 30
                    .Row = 1
                    .COL = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt病种 = "[" & Nvl(!疾病代码) & "]" & IIf(IsNull(!疾病名称), "", !疾病名称)
                txt病种.Tag = Nvl(!疾病代码)
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
            txt病种.Text = "[" & .TextMatrix(.Row, 0) & "]" & .TextMatrix(.Row, 1)
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

Private Sub InitCtlData()
    '初始控件数据
    Dim i As Integer
    Dim str入院类别 As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long
    gstrSQL = "Select 病人id,人员身份 From 保险帐户 where 医保号='" & g病人身份_兴成.社会保障号 & "'"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    If rsTemp.EOF Then Exit Sub
    
    str入院类别 = Nvl(rsTemp!人员身份, "0")
    lng病人ID = Nvl(rsTemp!病人ID, 0)
    If lng病人ID = 0 Then Exit Sub
    gstrSQL = "Select AF17,AF18,AF19,AF20 From 医保病人附加信息 where 病人id=" & lng病人ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.EOF Then Exit Sub
    
    '确定入院类别:
    If mbytType = 1 Or mbytType = 4 Then
        For i = 0 To cbo入院类别.ListCount - 1
            If cbo入院类别.ItemData(i) = str入院类别 Then
               cbo入院类别.ListIndex = i: Exit For
            End If
        Next
        For i = 0 To cbo住院类别.ListCount - 1
            If cbo住院类别.ItemData(i) = Nvl(rsTemp!AF19, "0") And cbo住院类别.Enabled Then
               cbo住院类别.ListIndex = i: Exit For
            End If
        Next
        For i = 0 To cbo出院类别.ListCount - 1
            If cbo出院类别.ItemData(i) = Nvl(rsTemp!AF20, "0") And cbo出院类别.Enabled Then
               cbo出院类别.ListIndex = i: Exit For
            End If
        Next
    End If
    
    '确定相关病种
    Me.txt病种.Text = IIf(Nvl(rsTemp!AF17) = "", "", "[" & Nvl(rsTemp!AF17) & "]" & Nvl(rsTemp!AF18))
    Me.txt病种.Tag = Nvl(rsTemp!AF17)
End Sub
