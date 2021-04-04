VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain_北京尚洋病案接口 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病案数据上传"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmMain_北京尚洋病案接口.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7125
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd清除上传标志 
      Caption         =   "清除(&A)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5880
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "清除指定病人的病案数据上传标志"
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmd参数 
      Caption         =   "参数(&M)"
      Height          =   350
      Left            =   5880
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1100
   End
   Begin VB.CheckBox chk全选 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "全选"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5010
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1590
      Width           =   675
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3270
      Top             =   1950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_北京尚洋病案接口.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd上传 
      Caption         =   "上传(&U)"
      Height          =   350
      Left            =   5880
      TabIndex        =   13
      Top             =   780
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw病人清单 
      Height          =   2265
      Left            =   150
      TabIndex        =   12
      Top             =   1800
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "个人编号"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "住院号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "姓名"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "出院日期"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "上传"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.CommandButton CDM查找 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   5880
      TabIndex        =   9
      Top             =   330
      Width           =   1100
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5880
      TabIndex        =   14
      Top             =   3630
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "过滤条件设置(&S)"
      Height          =   1425
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   5565
      Begin VB.CheckBox chk未上传 
         Caption         =   "仅显示未上传数据"
         Height          =   255
         Left            =   1020
         TabIndex        =   17
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2205
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         Left            =   4080
         TabIndex        =   8
         Top             =   690
         Width           =   1275
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   1470
         TabIndex        =   6
         Top             =   690
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker dtp开始日期 
         Height          =   300
         Left            =   1470
         TabIndex        =   2
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   113311747
         CurrentDate     =   39071
      End
      Begin MSComCtl2.DTPicker dtp结束日期 
         Height          =   300
         Left            =   4080
         TabIndex        =   4
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   113311747
         CurrentDate     =   39071
      End
      Begin VB.Label lbl住院号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院号(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3180
         TabIndex        =   7
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   750
         TabIndex        =   5
         Top             =   750
         Width           =   630
      End
      Begin VB.Label lbl结束日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lbl开始日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Label lbl病人清单 
      BackColor       =   &H00C0C0C0&
      Caption         =   "病人清单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   150
      TabIndex        =   10
      Top             =   1590
      Width           =   5565
   End
End
Attribute VB_Name = "frmMain_北京尚洋病案接口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _

Private strSQL As String

Private RSPATIENT As New ADODB.Recordset        '提取符合条件的病人
Private RSREC As New ADODB.Recordset

Private Type TRECORD_INFO
    C1统筹区号 As String
    C2医疗机构编号 As String
    C3住院号 As String
    C4收费操作员 As String
    C5付款方式 As String
    C6本次住院次数 As Integer
    C7病案编号 As String
    C8个人编号 As String
    C9姓名 As String
    C10性别 As String
    C11出生日期 As String
    C12婚姻 As String
    C13职业 As String
    C14出生地 As String
    C15民族 As String
    C16国籍 As String
    C17身份证号 As String
    C18工作单位 As String
    C19单位地址 As String
    C20单位电话 As String
    C21单位邮政编码 As String
    C22户口地址 As String
    C23邮政编码 As String
    C24联系人 As String
    C25与病人关系 As String
    C26联系地址 As String
    C27联系电话 As String
    C28入院日期 As String
    C29入院科室 As String
    C30入院病室 As String
    C31转科科别 As String
    C32出院日期 As String
    C33出院科室 As String
    C34出院病室 As String
    C35入院病情 As String
    C36入院后确认日期 As String
    C37过敏药物 As String
    C38HBSAG As String
    C39HCV_AB As String
    C40HIV_AB As String
    C41门诊与出院 As Integer
    C42入院与出院 As Integer
    C43术前与术后 As Integer
    C44临床与病理 As Integer
    C45放射与病理 As Integer
    C46抢救次数 As Integer
    C47抢救成功次数 As Integer
    C48科主任 As String
    C49主任医师 As String
    C50主治医师 As String
    C51住院医师 As String
    C52进修医师 As String
    C53研究生实习医师 As String
    C54实习医师 As String
    C55编码员 As String
    C56病案质量 As String
    C57质控医师 As String
    C58质控护师 As String
    C59结算日期 As String
    C60尸检标志 As String
    C61手术治疗检查诊断为本院第一例 As Integer
    C62随诊标志 As Integer
    C63随诊期限 As Integer
    C64示教病例 As Integer
    C65血型 As Integer
    C66RH As Integer
    C67输入血反应标志 As Integer
    C68输入红细胞 As Currency
    C69输入血小板 As Currency
    C70输入血浆 As Currency
    C71全血 As Currency
    C72其他 As Currency
    C73经办人 As String
    C74经办时间 As String
End Type

Private Sub READPATIENTS()
    Dim lvwItem As ListItem
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHand
    lvw病人清单.ListItems.Clear
    Me.chk全选.Value = 0
    
    '组合条件
    strSQL = " AND B.出院日期 BETWEEN TO_DATE('" & Format(dtp开始日期.Value, "YYYY-MM-DD") & " 00:00:00','YYYY-MM-DD HH24:MI:SS')" & _
    " AND TO_DATE('" & Format(dtp结束日期.Value, "YYYY-MM-DD") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    If Trim(Me.txt姓名.Text) <> "" Then
        strSQL = strSQL & " AND A.姓名 LIKE '" & Trim(Me.txt姓名.Text) & "%'"
    End If
    If Trim(Me.txt住院号.Text) <> "" Then
        strSQL = strSQL & " AND A.住院号 LIKE '" & Trim(Me.txt住院号.Text) & "%'"
    End If
    
    '提取符合条件的出院病人
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then
        strSQL = " SELECT A.病人ID,B.主页ID,C.医保号,A.姓名,A.住院号,B.出院日期,B.是否上传 " & _
             " FROM 病人信息 A,病案主页 B,保险帐户 C" & _
             " WHERE A.病人ID=B.病人ID AND A.主页ID=B.主页ID " & _
             " AND A.病人ID=C.病人ID AND C.险类=" & TYPE_北京尚洋 & IIf(chk未上传.Value = 1, " AND NVL(B.是否上传,0)=0 ", "") & strSQL
    Else
        strSQL = " SELECT A.病人ID,B.主页ID,C.医保号,A.姓名,A.住院号,B.出院日期,B.是否上传 " & _
             " FROM 病人信息 A,病案主页 B,保险帐户 C" & _
             " WHERE A.病人ID=B.病人ID AND A.住院次数=B.主页ID " & _
             " AND A.病人ID=C.病人ID AND C.险类=" & TYPE_北京尚洋 & IIf(chk未上传.Value = 1, " AND NVL(B.是否上传,0)=0 ", "") & strSQL
        
    End If
    Call OpenRecordset(RSPATIENT, "提取符合条件的病人", strSQL)
    With RSPATIENT
        Do While Not .EOF
            Set lvwItem = lvw病人清单.ListItems.Add(, "K" & .AbsolutePosition, Nvl(!医保号), 1)
            lvwItem.SubItems(1) = Nvl(!住院号) & "_" & !主页ID
            lvwItem.SubItems(2) = Nvl(!姓名)
            lvwItem.SubItems(3) = Format(!出院日期, "YYYY-MM-DD")
            lvwItem.SubItems(4) = IIf(Nvl(!是否上传, 0) = 0, "否", "是")
            lvwItem.Tag = !病人ID & "_" & !主页ID
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHand:
    MsgBox Err.Description, vbInformation, gstrSysName
    Resume
End Sub

Private Function UPLOADREC(ByVal lng病人ID As Long, ByVal lng主页ID As Long, STRERR As String) As Boolean
    '----------------------------------------------------------------
    '分隔段落,分段读取相应的数据并插入
    '----------------------------------------------------------------
    '根据传入的病人标识上传病案信息
    Dim arr病案费目
    Dim STR病案费目 As String
    Dim STR出院情况 As String
    Dim RECORD_INFO As TRECORD_INFO
    Dim bln34 As Boolean
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHand
    
    STR病案费目 = GetSetting("ZLSOFT", "私有模块\FRMSET", "病案费目", "")
    arr病案费目 = Split(STR病案费目, "|")
    
    gcn尚洋.BeginTrans
    
    '----------------------------------------------------------------
    '1、RECORD_INFO
    RECORD_INFO.C1统筹区号 = gstr医保机构编码
    RECORD_INFO.C2医疗机构编号 = Trim(gstr医院编码)
    '取医保档案
    strSQL = " SELECT 医保号" & _
             " FROM 保险帐户" & _
             " WHERE 险类=" & TYPE_北京尚洋 & " AND 病人ID=" & lng病人ID
    Call OpenRecordset(RSREC, "取医保档案", strSQL)
    RECORD_INFO.C8个人编号 = RSREC!医保号
    
    '取病人基本信息
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    bln34 = Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34
    
    #If gverControl < 6 Then
        If bln34 Then
            strSQL = " SELECT A.住院号,B.医疗付款方式,B.主页ID,B.病案号,A.姓名,A.性别,A.出生日期,B.婚姻状况,B.职业,A.出生地点," & _
                "        H.编码 AS 民族,B.国籍,A.身份证号,A.工作单位,B.单位地址,B.单位电话,B.单位邮编,B.家庭地址,B.户口邮编," & _
                "        B.联系人姓名,B.联系人关系,B.联系人地址,B.联系人电话,B.入院日期,D.名称 AS 入院科室,E.名称 AS 入院病区," & _
                "        B.出院日期,F.名称 AS 出院科室,B.入院病况,B.确诊日期,B.抢救次数,B.成功次数,B.出院方式," & _
                "        B.编目员姓名,NVL(B.编目日期,SYSDATE) AS 编目日期,B.尸检标志,B.随诊标志,B.随诊期限,B.血型,B.住院医师" & _
                " FROM 病人信息 A,病案主页 B,合约单位 C,部门表 D,部门表 E,部门表 F,民族 H" & _
                " WHERE A.病人ID=B.病人ID AND A.主页ID=B.主页ID AND A.合同单位ID=C.ID(+)" & _
                " AND B.入院科室ID=D.ID(+) AND B.入院病区ID=E.ID(+) AND B.出院科室ID=F.ID(+) " & _
                " AND A.民族=H.名称 AND B.病人ID=[1] AND B.主页ID=[2]"
        Else
            strSQL = " SELECT A.住院号,B.医疗付款方式,B.主页ID,B.病案号,A.姓名,A.性别,A.出生日期,B.婚姻状况,B.职业,A.出生地点," & _
                "        H.编码 AS 民族,B.国籍,A.身份证号,A.工作单位,B.单位地址,B.单位电话,B.单位邮编,B.家庭地址,B.户口邮编," & _
                "        B.联系人姓名,B.联系人关系,B.联系人地址,B.联系人电话,B.入院日期,D.名称 AS 入院科室,E.名称 AS 入院病区," & _
                "        B.出院日期,F.名称 AS 出院科室,B.入院病况,B.确诊日期,B.抢救次数,B.成功次数,B.出院方式," & _
                "        B.编目员姓名,NVL(B.编目日期,SYSDATE) AS 编目日期,B.尸检标志,B.随诊标志,B.随诊期限,B.血型,B.住院医师" & _
                " FROM 病人信息 A,病案主页 B,合约单位 C,部门表 D,部门表 E,部门表 F,民族 H" & _
                " WHERE A.病人ID=B.病人ID AND A.住院次数=B.主页ID AND A.合同单位ID=C.ID(+)" & _
                " AND B.入院科室ID=D.ID(+) AND B.入院病区ID=E.ID(+) AND B.出院科室ID=F.ID(+) " & _
                " AND A.民族=H.名称 AND B.病人ID=[1] AND B.主页ID=[2]"
        End If
    #Else
        If bln34 Then
            strSQL = " SELECT A.住院号,B.医疗付款方式,B.主页ID,B.病案号,A.姓名,A.性别,A.出生日期,B.婚姻状况,B.职业,A.出生地点," & _
                "        H.编码 AS 民族,B.国籍,A.身份证号,A.工作单位,B.单位地址,B.单位电话,B.单位邮编,B.家庭地址,B.家庭地址邮编 As 户口邮编," & _
                "        B.联系人姓名,B.联系人关系,B.联系人地址,B.联系人电话,B.入院日期,D.名称 AS 入院科室,E.名称 AS 入院病区," & _
                "        B.出院日期,F.名称 AS 出院科室,B.入院病况,B.确诊日期,B.抢救次数,B.成功次数,B.出院方式," & _
                "        B.编目员姓名,NVL(B.编目日期,SYSDATE) AS 编目日期,B.尸检标志,B.随诊标志,B.随诊期限,B.血型,B.住院医师" & _
                " FROM 病人信息 A,病案主页 B,合约单位 C,部门表 D,部门表 E,部门表 F,民族 H" & _
                " WHERE A.病人ID=B.病人ID AND A.主页ID=B.主页ID AND A.合同单位ID=C.ID(+)" & _
                " AND B.入院科室ID=D.ID(+) AND B.入院病区ID=E.ID(+) AND B.出院科室ID=F.ID(+) " & _
                " AND A.民族=H.名称 AND B.病人ID=[1] AND B.主页ID=[2]"
        Else
            strSQL = " SELECT A.住院号,B.医疗付款方式,B.主页ID,B.病案号,A.姓名,A.性别,A.出生日期,B.婚姻状况,B.职业,A.出生地点," & _
                "        H.编码 AS 民族,B.国籍,A.身份证号,A.工作单位,B.单位地址,B.单位电话,B.单位邮编,B.家庭地址,B.家庭地址邮编 As 户口邮编," & _
                "        B.联系人姓名,B.联系人关系,B.联系人地址,B.联系人电话,B.入院日期,D.名称 AS 入院科室,E.名称 AS 入院病区," & _
                "        B.出院日期,F.名称 AS 出院科室,B.入院病况,B.确诊日期,B.抢救次数,B.成功次数,B.出院方式," & _
                "        B.编目员姓名,NVL(B.编目日期,SYSDATE) AS 编目日期,B.尸检标志,B.随诊标志,B.随诊期限,B.血型,B.住院医师" & _
                " FROM 病人信息 A,病案主页 B,合约单位 C,部门表 D,部门表 E,部门表 F,民族 H" & _
                " WHERE A.病人ID=B.病人ID AND A.住院次数=B.主页ID AND A.合同单位ID=C.ID(+)" & _
                " AND B.入院科室ID=D.ID(+) AND B.入院病区ID=E.ID(+) AND B.出院科室ID=F.ID(+) " & _
                " AND A.民族=H.名称 AND B.病人ID=[1] AND B.主页ID=[2]"
        End If
    #End If
    Call OpenRecordset(RSREC, "取病人基本信息", strSQL)
    STR出院情况 = TRANDATA("出院情况", Nvl(RSREC!出院方式))
    With RSREC
        RECORD_INFO.C31转科科别 = "无"
        
        RECORD_INFO.C3住院号 = Nvl(!住院号) & "_" & !主页ID
        RECORD_INFO.C5付款方式 = TRANDATA("医疗付款方式", !医疗付款方式)
        RECORD_INFO.C6本次住院次数 = !主页ID
        RECORD_INFO.C7病案编号 = Nvl(!住院号)
        RECORD_INFO.C9姓名 = !姓名
        RECORD_INFO.C10性别 = TRANDATA("性别", !性别)
        RECORD_INFO.C11出生日期 = Format(!出生日期, "YYYY-MM-DD HH:MM:SS")
        RECORD_INFO.C12婚姻 = TRANDATA("婚姻", Nvl(!婚姻状况, "未婚"))
        RECORD_INFO.C13职业 = ToVarchar(Nvl(!职业), 20)
        RECORD_INFO.C14出生地 = Nvl(!出生地点)
        RECORD_INFO.C15民族 = !民族
        RECORD_INFO.C16国籍 = 86 '!国籍
        RECORD_INFO.C17身份证号 = Nvl(!身份证号)
        RECORD_INFO.C18工作单位 = Nvl(!工作单位)
        RECORD_INFO.C19单位地址 = Nvl(!单位地址)
        RECORD_INFO.C20单位电话 = Nvl(!单位电话)
        RECORD_INFO.C21单位邮政编码 = Nvl(!单位邮编)
        RECORD_INFO.C22户口地址 = Nvl(!家庭地址)
        RECORD_INFO.C23邮政编码 = Nvl(!户口邮编)
        RECORD_INFO.C24联系人 = Nvl(!联系人姓名)
        RECORD_INFO.C25与病人关系 = TRANDATA("与病人关系", Nvl(!联系人关系))
        RECORD_INFO.C26联系地址 = Nvl(!联系人地址)
        RECORD_INFO.C27联系电话 = Nvl(!联系人电话)
        RECORD_INFO.C28入院日期 = Format(!入院日期, "YYYY-MM-DD HH:MM:SS")
        RECORD_INFO.C29入院科室 = !入院科室
        RECORD_INFO.C30入院病室 = Nvl(!入院病区)
        RECORD_INFO.C32出院日期 = Format(!出院日期, "YYYY-MM-DD HH:MM:SS")
        RECORD_INFO.C33出院科室 = !出院科室
        RECORD_INFO.C34出院病室 = !出院科室
        RECORD_INFO.C35入院病情 = TRANDATA("入院病情", Nvl(!入院病况))
        RECORD_INFO.C36入院后确认日期 = Format(!确诊日期, "YYYY-MM-DD HH:MM:SS")
        RECORD_INFO.C46抢救次数 = Nvl(!抢救次数, 0)
        RECORD_INFO.C47抢救成功次数 = Nvl(!成功次数, 0)
        RECORD_INFO.C51住院医师 = Nvl(!住院医师)
        RECORD_INFO.C55编码员 = Nvl(!编目员姓名)
        RECORD_INFO.C60尸检标志 = TRANDATA("尸检标志", Nvl(!尸检标志, "否"))
        RECORD_INFO.C62随诊标志 = TRANDATA("随诊标志", Nvl(!随诊标志))
        RECORD_INFO.C63随诊期限 = Nvl(!随诊期限, 0)
        RECORD_INFO.C65血型 = TRANDATA("血型", Nvl(!血型))
        RECORD_INFO.C73经办人 = ToVarchar(Nvl(!编目员姓名), 20)
        RECORD_INFO.C74经办时间 = Format(Now, "YYYY-MM-DD HH:MM:SS")
    End With
    
    '取病案主页从表
    Dim STR信息值 As String
    strSQL = "SELECT UPPER(信息名) AS 信息名,信息值 FROM 病案主页从表 WHERE 病人ID=" & lng病人ID & " AND 主页ID=" & lng主页ID
    Call OpenRecordset(RSREC, "取病案主页从表", strSQL)
    With RSREC
        Do While Not .EOF
            STR信息值 = Nvl(!信息值)
            Select Case !信息名
            Case "HBSAG"
                RECORD_INFO.C38HBSAG = TRANDATA("HBSAG", STR信息值)
            Case "HCV-AB"
                RECORD_INFO.C39HCV_AB = TRANDATA("HCV_AB", STR信息值)
            Case "HIV-AB"
                RECORD_INFO.C40HIV_AB = TRANDATA("HIV_AB", STR信息值)
            Case "科主任"
                RECORD_INFO.C48科主任 = STR信息值
            Case "主任医师"
                RECORD_INFO.C49主任医师 = STR信息值
            Case "主治医师"
                RECORD_INFO.C50主治医师 = STR信息值
            Case "进修医师"
                RECORD_INFO.C52进修医师 = STR信息值
            Case "研究生实习医师"
                RECORD_INFO.C53研究生实习医师 = STR信息值
            Case "实习医师"
                RECORD_INFO.C54实习医师 = STR信息值
            Case "病案质量"
                RECORD_INFO.C56病案质量 = TRANDATA("病案质量", STR信息值)
            Case "质控医师"
                RECORD_INFO.C57质控医师 = STR信息值
            Case "质控护士"
                RECORD_INFO.C58质控护师 = STR信息值
            Case "首例"
                RECORD_INFO.C61手术治疗检查诊断为本院第一例 = TRANDATA("首例", STR信息值)
            Case "示教病案"
                RECORD_INFO.C64示教病例 = TRANDATA("示教病例", STR信息值)
            Case "RH"
                RECORD_INFO.C66RH = TRANDATA("RH", STR信息值)
            Case "输血反应"
                RECORD_INFO.C67输入血反应标志 = TRANDATA("输血反应", STR信息值)
            Case "输红细胞"
                RECORD_INFO.C68输入红细胞 = Val(STR信息值)
            Case "输血小板"
                RECORD_INFO.C69输入血小板 = Val(STR信息值)
            Case "输血浆"
                RECORD_INFO.C70输入血浆 = Val(STR信息值)
            Case "输全血"
                RECORD_INFO.C71全血 = Val(STR信息值)
            Case "输其他"
                RECORD_INFO.C72其他 = Val(STR信息值)
            End Select
            .MoveNext
        Loop
    End With
    
    '任选一种过敏药物
    strSQL = " SELECT 过敏药物 FROM 病人过敏药物 WHERE 病人ID=" & lng病人ID
    Call OpenRecordset(RSREC, "任选一种过敏药物", strSQL)
    Do While Not RSREC.EOF
        RECORD_INFO.C37过敏药物 = RECORD_INFO.C37过敏药物 & "," & Nvl(RSREC!过敏药物)
        RSREC.MoveNext
    Loop
    RECORD_INFO.C37过敏药物 = Mid(RECORD_INFO.C37过敏药物, 2)
    RECORD_INFO.C37过敏药物 = ToVarchar(RECORD_INFO.C37过敏药物, 50)
'    '取病案评分结果（从病案主页从表中读取，只要填写了病案的都会有数据）
'    STRSQL = "SELECT 等级 FROM 病案评分结果 WHERE 病人ID=" & LNG病人ID & " AND 主页ID=" & LNG主页ID
'    CALL OPENRECORD(RSREC, STRSQL, "取病案评分结果")
'    IF RSREC.RECORDCOUNT <> 0 THEN
'        RECORD_INFO.C56病案质量 = RSREC!等级
'    END IF
    '取结算数据
    strSQL = "SELECT 操作员姓名,收费时间 FROM 病人结帐记录 WHERE ID = (SELECT MAX(ID) FROM 病人结帐记录 WHERE 病人ID=" & lng病人ID & " AND 记录状态=1)"
    Call OpenRecordset(RSREC, "取结算数据", strSQL)
    If RSREC.RecordCount <> 0 Then
'        RECORD_INFO.C4收费操作员 = RSREC!操作员姓名
        RECORD_INFO.C59结算日期 = Format(RSREC!收费时间, "YYYY-MM-DD HH:MM:SS")
    End If
    '取诊断情况
    strSQL = "SELECT 符合类型,NVL(符合情况,0) AS 符合情况 FROM 诊断符合情况 WHERE 病人ID=" & lng病人ID & " AND 主页ID=" & lng主页ID
    Call OpenRecordset(RSREC, "取诊断情况", strSQL)
    With RSREC
        Do While Not .EOF
            Select Case !符合类型
            Case 1  '门诊与出院
                RECORD_INFO.C41门诊与出院 = !符合情况
            Case 2  '入院与出院
                RECORD_INFO.C42入院与出院 = !符合情况
            Case 3  '放射与病理
                RECORD_INFO.C45放射与病理 = !符合情况
            Case 4  '临床与病理
                RECORD_INFO.C44临床与病理 = !符合情况
            Case 6  '术前与术后
                RECORD_INFO.C43术前与术后 = !符合情况
            End Select
            .MoveNext
        Loop
    End With
    
    '插入数据表:MEDICAL_RECORD_INFO
'    gcn尚洋.Execute " DELETE MEDICAL_RECORD_INFO " & _
'                  " WHERE AREAID='" & RECORD_INFO.C1统筹区号 & "'" & _
'                  " AND PERSONAL_NUMBER='" & RECORD_INFO.C8个人编号 & "'" & _
'                  " AND RESIDENCE_NO='" & RECORD_INFO.C3住院号 & "'"
    strSQL = " INSERT INTO MEDICAL_RECORD_INFO" & _
         " (AREAID,HOSPITAL_NUMBER,RESIDENCE_NO,CHARGE_NUMBER,PAY_MODE,IN_COUNT,MEDICAL_RECORD_NO,PERSONAL_NUMBER, " & _
         " NAME,SEX,BIRTH_DATE,MARITAL_STATUS,STATUS,BIRTH_ADDRESS,NATIONALITY,CITIZENSHIP,IDENTITY_NUMBER, " & _
         " UNIT_NAME,UNIT_ADDRESS,UNIT_PHONE,UNIT_ZIPCODE,REGISTER_ADDRESS,REGISTER_ZIPCODE,CONTACT_PERSON, " & _
         " RELATIONSHIP,CONTACT_ADDRESS,CONTACT_PHONE,ADMISSION_DATE,ADMISSION_DEPT,IN_DEPT_ZONE,DEPT_TRANSFERED_TO, " & _
         " DISCHARGE_DATE,DISCHARGE_DEPT,OUT_DEPT_ZONE,PAT_ADM_CONDITION,DIAGNOSIS_DATE,ALERGY_DRUGS,HBsAg,HCV_Ab, " & _
         " HIV_Ab,CLINIC_INHOSPITAL,IN_OUT,BEFORE_AFTER_TREATMENT,CLINIC_PATHOLOGY,EMIT_PATHOLOGY,EMER_TREAT_TIMES,ESC_EMER_TIMES, " & _
         " DIRECTOR,DIRECTOR_DOCTOR,ATTENDING_DOCTOR,INHOSPITAL_DOCTOR,REFRESH_DOCTOR,GRADUATE_DOCTOR,INTERM,CODE_NAME, " & _
         " MEDICAL_RECORD_MASS,CONTROL_DOCTOR,CONTROL_NURSE,BAL_DATE,BODY_EXAMINE_FLAG,FIRST_FLAG,FOLLOW_FLAG,FOLLOW_TERM, " & _
         " TEACH_MR_FLAG,BLOOD_TYPE,Rh,BLOOD_TRAN_REACT_FLAG,ERYTHROCYTE,HEMOBLAST,PLASM,BLOOD,OTHER_BLOOD,HANDLE,HANDLE_DATE)" & _
         " VALUES ("
    strSQL = strSQL & _
         "'" & RECORD_INFO.C1统筹区号 & "','" & RECORD_INFO.C2医疗机构编号 & "','" & RECORD_INFO.C3住院号 & "','" & RECORD_INFO.C4收费操作员 & "'," & _
         "'" & RECORD_INFO.C5付款方式 & "'," & RECORD_INFO.C6本次住院次数 & ",'" & RECORD_INFO.C7病案编号 & "','" & RECORD_INFO.C8个人编号 & "'," & _
         "'" & RECORD_INFO.C9姓名 & "','" & RECORD_INFO.C10性别 & "','" & RECORD_INFO.C11出生日期 & "','" & RECORD_INFO.C12婚姻 & "'," & _
         "'" & RECORD_INFO.C13职业 & "','" & RECORD_INFO.C14出生地 & "','" & RECORD_INFO.C15民族 & "','" & RECORD_INFO.C16国籍 & "'," & _
         "'" & RECORD_INFO.C17身份证号 & "','" & RECORD_INFO.C18工作单位 & "','" & RECORD_INFO.C19单位地址 & "','" & RECORD_INFO.C20单位电话 & "'," & _
         "'" & RECORD_INFO.C21单位邮政编码 & "','" & RECORD_INFO.C22户口地址 & "','" & RECORD_INFO.C23邮政编码 & "','" & RECORD_INFO.C24联系人 & "'," & _
         "'" & RECORD_INFO.C25与病人关系 & "','" & RECORD_INFO.C26联系地址 & "','" & RECORD_INFO.C27联系电话 & "','" & RECORD_INFO.C28入院日期 & "'," & _
         "'" & RECORD_INFO.C29入院科室 & "','" & RECORD_INFO.C30入院病室 & "','" & RECORD_INFO.C31转科科别 & "','" & RECORD_INFO.C32出院日期 & "'," & _
         "'" & RECORD_INFO.C33出院科室 & "','" & RECORD_INFO.C34出院病室 & "','" & RECORD_INFO.C35入院病情 & "','" & RECORD_INFO.C36入院后确认日期 & "'," & _
         "'" & RECORD_INFO.C37过敏药物 & "','" & RECORD_INFO.C38HBSAG & "','" & RECORD_INFO.C39HCV_AB & "','" & RECORD_INFO.C40HIV_AB & "'," & _
         "'" & RECORD_INFO.C41门诊与出院 & "','" & RECORD_INFO.C42入院与出院 & "','" & RECORD_INFO.C43术前与术后 & "','" & RECORD_INFO.C44临床与病理 & "'," & _
         "'" & RECORD_INFO.C45放射与病理 & "'," & RECORD_INFO.C46抢救次数 & "," & RECORD_INFO.C47抢救成功次数 & ",'" & RECORD_INFO.C48科主任 & "'," & _
         "'" & RECORD_INFO.C49主任医师 & "','" & RECORD_INFO.C50主治医师 & "','" & RECORD_INFO.C51住院医师 & "','" & RECORD_INFO.C52进修医师 & "',"
    strSQL = strSQL & _
         "'" & RECORD_INFO.C53研究生实习医师 & "','" & RECORD_INFO.C54实习医师 & "','" & RECORD_INFO.C55编码员 & "','" & RECORD_INFO.C56病案质量 & "'," & _
         "'" & RECORD_INFO.C57质控医师 & "','" & RECORD_INFO.C58质控护师 & "','" & RECORD_INFO.C59结算日期 & "','" & RECORD_INFO.C60尸检标志 & "'," & _
         "'" & RECORD_INFO.C61手术治疗检查诊断为本院第一例 & "','" & RECORD_INFO.C62随诊标志 & "'," & RECORD_INFO.C63随诊期限 & ",'" & RECORD_INFO.C64示教病例 & "'," & _
         "'" & RECORD_INFO.C65血型 & "','" & RECORD_INFO.C66RH & "','" & RECORD_INFO.C67输入血反应标志 & "'," & RECORD_INFO.C68输入红细胞 & "," & _
         "" & RECORD_INFO.C69输入血小板 & "," & RECORD_INFO.C70输入血浆 & "," & RECORD_INFO.C71全血 & "," & RECORD_INFO.C72其他 & ",'" & RECORD_INFO.C73经办人 & "','" & RECORD_INFO.C74经办时间 & "')"
    gcn尚洋.Execute strSQL
    
    '----------------------------------------------------------------
    '2、DIAGNOSIS
    'TODO:诊疗结果如何填写:每条诊断记录都要填写诊疗结果，就按出院情况的代码表填写
    strSQL = " SELECT A.诊断类型,A.诊断次序,A.编码序号,B.编码 AS 疾病编码,A.诊断描述,A.记录人,NVL(A.记录日期,SYSDATE) AS 记录日期" & _
             " FROM 病人诊断记录 A,疾病编码目录 B" & _
             " WHERE A.疾病ID=B.ID AND A.记录来源=3 AND A.诊断类型<8 AND A.病人ID=" & lng病人ID & " AND A.主页ID=" & lng主页ID
    Call OpenRecordset(RSREC, "读诊断记录", strSQL)
    With RSREC
        Do While Not .EOF
            strSQL = " INSERT INTO MR_DIAGNOSIS" & _
                     " (HOSPITAL_NUMBER,MEDICAL_RECORD_NO,IN_COUNT,DIAGNOSIS_TYPE,DIAGNOSIS_NO,ILLNESS_CODE,DIAGNOSIS_DESC,DIAGNOSIS_DATE,TREAT_RESULT,HANDLE,HANDLE_DATE)" & _
                     " VALUES (" & _
                     "'" & RECORD_INFO.C2医疗机构编号 & "','" & RECORD_INFO.C7病案编号 & "'," & RECORD_INFO.C6本次住院次数 & "," & _
                     "'" & TRANDATA("诊断类型", !诊断类型) & "'," & Nvl(!诊断次序, 0) & ",'" & Nvl(!疾病编码) & "','" & Nvl(!诊断描述) & "'," & _
                     "'" & Format(!记录日期, "YYYY-MM-DD HH:MM:SS") & "','" & STR出院情况 & "','" & ToVarchar(Nvl(!记录人), 20) & "','" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
            gcn尚洋.Execute strSQL
            .MoveNext
        Loop
    End With
    
    '----------------------------------------------------------------
    '3、OPERATION
    strSQL = " SELECT B.编码,B.名称,A.切口,A.愈合,A.手术日期,A.麻醉类型,A.主刀医师,A.第一助手,A.第二助手,A.麻醉医师,A.记录人,NVL(A.记录日期,SYSDATE) AS 记录日期 " & _
             " FROM 病人手麻记录 A ,疾病编码目录 B " & _
             " WHERE A.病人ID=" & lng病人ID & " AND A.主页ID=" & lng主页ID & " AND A.手术操作ID=B.ID"
    Call OpenRecordset(RSREC, "取病人手麻记录", strSQL)
    With RSREC
        Do While Not .EOF
            strSQL = " INSERT INTO MR_OPERATION" & _
                     " (HOSPITAL_NUMBER,MEDICAL_RECORD_NO,IN_COUNT,OPERATION_NO,OPERATION_CODE,OPERATION_DESC,WOUND_GRADE," & _
                     " HEAL,OPERATING_DATE,ANAESTHESIA_METHOD,OPERATOR,IASIST1,IASIST2,ANAESTHESIA_OPERATOR,HANDLE,HANDLE_DATE)" & _
                     " VALUES (" & _
                     "'" & RECORD_INFO.C2医疗机构编号 & "','" & RECORD_INFO.C7病案编号 & "'," & RECORD_INFO.C6本次住院次数 & "," & _
                     "" & .AbsolutePosition & ",'" & !编码 & "','" & !名称 & "','" & TRANDATA(Nvl(!切口), Nvl(!愈合)) & "','" & TRANDATA(Nvl(!切口), Nvl(!愈合)) & "'," & _
                     "'" & Format(!手术日期, "YYYY-MM-DD HH:MM:SS") & "','" & TRANDATA("麻醉类型", !麻醉类型) & "','" & !主刀医师 & "'," & _
                     "'" & Nvl(!第一助手) & "','" & Nvl(!第二助手) & "','" & ToVarchar(Nvl(!麻醉医师), 20) & "','" & ToVarchar(Nvl(!记录人), 20) & "','" & Format(Now, "YYYY-MM-DD") & "')"
            gcn尚洋.Execute strSQL
            .MoveNext
        Loop
    End With
    
    '----------------------------------------------------------------
    '4、RECEIPT_DETAIL
    strSQL = " SELECT 费用名,金额 FROM 病人费用 WHERE 病人ID=" & lng病人ID & " AND 主页ID=" & lng主页ID
    Call OpenRecordset(RSREC, "取病人费用", strSQL)
    With RSREC
        Do While Not .EOF
            strSQL = " INSERT INTO MR_RECEIPT_DETAIL" & _
                     " (HOSPITAL_NUMBER,MEDICAL_RECORD_NO,IN_COUNT,RECEIPT_NAME,ITEM_COST,SEND_FLAG,HANDLE,HANDLE_DATE)" & _
                     " VALUES (" & _
                     "'" & RECORD_INFO.C2医疗机构编号 & "','" & RECORD_INFO.C7病案编号 & "'," & RECORD_INFO.C6本次住院次数 & "," & _
                     "'" & GET费用项目编码(!费用名, arr病案费目) & "'," & !金额 & ",0,'" & gstrUserName & "','" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
            gcn尚洋.Execute strSQL
            .MoveNext
        Loop
    End With
    
    gcn尚洋.CommitTrans
    '更新上传标记
    strSQL = "ZL_病案主页_上传(" & lng病人ID & "," & lng主页ID & ",1)"
    gcnOracle.Execute strSQL, , adCmdStoredProc
    
    UPLOADREC = True
    Exit Function
errHand:
    STRERR = Err.Description
    Debug.Print "Error SQL:" & strSQL
    gcn尚洋.RollbackTrans
End Function

Private Sub CDM查找_CLICK()
    Call READPATIENTS
End Sub

Private Sub CHK全选_CLICK()
    Dim BLNSEL As Boolean
    Dim LNGDO As Long, LNGMAX As Long
    
    BLNSEL = (chk全选.Value = 1)
    LNGMAX = lvw病人清单.ListItems.Count
    
    For LNGDO = 1 To LNGMAX
        lvw病人清单.ListItems(LNGDO).Checked = BLNSEL
    Next
    chk全选.Caption = IIf(BLNSEL, "全清", "全选")
End Sub

Private Sub CMD参数_CLICK()
'    frmSet.Show 1, Me
End Sub

Private Sub CMD放弃_Click()
    Unload Me
End Sub

Private Sub cmd清除上传标志_Click()
    Dim lng病人ID As Long, lng主页ID As Long
    On Error GoTo errHand
    
    If lvw病人清单.ListItems.Count = 0 Then Exit Sub
    If lvw病人清单.SelectedItem Is Nothing Then Exit Sub
    lng病人ID = Val(Split(lvw病人清单.SelectedItem.Tag, "_")(0))
    lng主页ID = Val(Split(lvw病人清单.SelectedItem.Tag, "_")(1))
    
    If MsgBox("你确定要清除该病人的病案数据上传标志吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '更新上传标记
    strSQL = "ZL_病案主页_上传(" & lng病人ID & "," & lng主页ID & ",0)"
    gcnOracle.Execute strSQL, , adCmdStoredProc
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CMD上传_CLICK()
    Dim STR病案费目 As String
    Dim LNGDO As Long, LNGMAX As Long
    Dim lng病人ID As Long, lng主页ID As Long
    Dim STR姓名 As String, STRERR As String
    
    STR病案费目 = GetSetting("ZLSOFT", "私有模块\FRMSET", "病案费目", "")
    If STR病案费目 = "" Then
        MsgBox "请先在参数设置中进行病案费目对照！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    LNGMAX = lvw病人清单.ListItems.Count
    For LNGDO = 1 To LNGMAX
        If lvw病人清单.ListItems(LNGDO).Checked Then
            STR姓名 = lvw病人清单.ListItems(LNGDO).SubItems(2)
            lng病人ID = Split(lvw病人清单.ListItems(LNGDO).Tag, "_")(0)
            lng主页ID = Split(lvw病人清单.ListItems(LNGDO).Tag, "_")(1)
            
            If lvw病人清单.SelectedItem.SubItems(4) = "否" Then
                Me.Caption = "病案数据上传 正在上传:" & STR姓名 & "的数据,请稍候..."
                If Not UPLOADREC(lng病人ID, lng主页ID, STRERR) Then
                    If MsgBox("上传病人[" & STR姓名 & "]时发生错误,继续上传其他病人吗？" & vbCrLf & _
                        STRERR, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Me.Caption = "病案数据上传"
                            Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    Me.Caption = "病案数据上传"
    Call READPATIENTS
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If 医保初始化_北京尚洋 = False Then
        Unload Me
        Exit Sub
    End If
    
    Me.dtp结束日期.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    Me.dtp开始日期.Value = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd")
End Sub

Private Function TRANDATA(ByVal STR信息名 As String, ByVal STR信息值 As String) As String
    '根据接口文档转换HIS中的值
    Select Case STR信息名
    Case "医疗付款方式"
        Select Case STR信息值
        Case "社会基本医疗保险"
            TRANDATA = 1
        Case "商业保险"
            TRANDATA = 2
        Case "自费医疗"
            TRANDATA = 3
        Case "公费医疗"
            TRANDATA = 4
        Case "大病统筹"
            TRANDATA = 5
        Case Else   '其他
            TRANDATA = 6
        End Select
    Case "性别"
        Select Case STR信息值
        Case "男"
            TRANDATA = 1
        Case Else   '女
            TRANDATA = 2
        End Select
    Case "婚姻"
        Select Case STR信息值
        Case "未婚"
            TRANDATA = 1
        Case "已婚"
            TRANDATA = 2
        Case "离婚"
            TRANDATA = 3
        Case Else   '丧
            TRANDATA = 4
        End Select
    Case "与病人关系"
        Select Case STR信息值
        Case "配偶"
            TRANDATA = 1
        Case "子", "女"
            TRANDATA = 2
        Case "父母"
            TRANDATA = 3
        Case Else   '孙子\孙女\祖父\祖母\本人等等,都归入其他
            TRANDATA = 9
        End Select
    Case "尸检标志", "首例", "随诊标志", "示教病例", "RH", "输血反应"
        Select Case STR信息值
        Case "是"
            TRANDATA = 1
        Case Else
            TRANDATA = 2
        End Select
    Case "血型"
        Select Case STR信息值
        Case "A"
            TRANDATA = 1
        Case "B"
            TRANDATA = 2
        Case "AB"
            TRANDATA = 3
        Case "O"
            TRANDATA = 4
        Case Else
            TRANDATA = 5
        End Select
    Case "麻醉类型"
        Select Case STR信息值
        Case "全麻"
            TRANDATA = 1
        Case "局麻"
            TRANDATA = 3
        Case Else
            TRANDATA = 2
        End Select
    Case "病案质量"
        Select Case STR信息值
        Case "甲"
            TRANDATA = 1
        Case "乙"
            TRANDATA = 2
        Case Else
            TRANDATA = 3
        End Select
    Case "诊疗结果", "出院情况"
        Select Case STR信息值
        Case "治愈", "正常"
            TRANDATA = 1
        Case "好转"
            TRANDATA = 2
        Case "未愈"
            TRANDATA = 3
        Case "死亡"
            TRANDATA = 4
        Case Else
            TRANDATA = 5
        End Select
    Case "HBSAG", "HCV_AB", "HIV_AB"
        Select Case STR信息值
        Case "阴性"
            TRANDATA = 1
        Case "阳性"
            TRANDATA = 2
        Case Else
            TRANDATA = 0
        End Select
    Case "入院病情"
        Select Case STR信息值
        Case "危"
            TRANDATA = 1
        Case "急"
            TRANDATA = 2
        Case Else
            TRANDATA = 3
        End Select
    Case "切口", "愈合"     '接口内是统一判断的
        Select Case STR信息值
        Case "Ⅰ/甲"
            TRANDATA = "01"
        Case "Ⅱ/甲"
            TRANDATA = "02"
        Case "Ⅲ/甲"
            TRANDATA = "03"
        Case "Ⅰ/乙"
            TRANDATA = "04"
        Case "Ⅱ/乙"
            TRANDATA = "05"
        Case "Ⅲ/乙"
            TRANDATA = "06"
        Case "Ⅰ/丙"
            TRANDATA = "07"
        Case "Ⅱ/丙"
            TRANDATA = "08"
        Case Else
            TRANDATA = "09"
        End Select
    Case "诊断类型"
        Select Case STR信息值
        Case 5, 6, 7
            TRANDATA = Val(STR信息值) - 1
        Case 1, 2, 3
            TRANDATA = Val(STR信息值)
        End Select
    End Select
End Function

Private Function GET费用项目编码(ByVal STRNAME As String, ByVal ARRNAME As Variant) As String
    Dim intDO As Integer, intCOUNT As Integer
    intCOUNT = UBound(ARRNAME)
    For intDO = 0 To intCOUNT
        If STRNAME = Split(ARRNAME(intDO), ",")(0) Then
            GET费用项目编码 = Split(Split(ARRNAME(intDO), ",")(1), "-")(0)
            Exit Function
        End If
    Next
End Function


Private Sub lvw病人清单_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvw病人清单.ListItems.Count = 0 Then Exit Sub
    If lvw病人清单.SelectedItem Is Nothing Then Exit Sub
    
    If lvw病人清单.SelectedItem.SubItems(4) = "是" Then
        cmd清除上传标志.Enabled = True
    Else
        cmd清除上传标志.Enabled = False
    End If
End Sub
