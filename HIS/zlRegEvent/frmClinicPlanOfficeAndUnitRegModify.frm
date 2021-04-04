VERSION 5.00
Begin VB.Form frmClinicPlanOfficeAndUnitRegModify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊室调整"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9375
   Icon            =   "frmClinicPlanOfficeAndUnitRegModify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8130
      TabIndex        =   32
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8130
      TabIndex        =   33
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   8130
      TabIndex        =   34
      Top             =   6210
      Width           =   1100
   End
   Begin VB.Frame fra出诊信息 
      Caption         =   "出诊信息"
      Height          =   1065
      Left            =   30
      TabIndex        =   14
      Top             =   1560
      Width           =   7965
      Begin VB.TextBox txt限约数 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6720
         TabIndex        =   21
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txt限号数 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4800
         TabIndex        =   20
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txt上班时段 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   18
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txt出诊日期 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   16
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txt替诊医生 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   23
         Top             =   675
         Width           =   1065
      End
      Begin VB.CheckBox chk序号控制 
         Caption         =   "启用序号控制"
         Enabled         =   0   'False
         Height          =   225
         Left            =   4800
         TabIndex        =   26
         Top             =   713
         Width           =   1395
      End
      Begin VB.CheckBox chk时段 
         Caption         =   "启用时段"
         Enabled         =   0   'False
         Height          =   225
         Left            =   6690
         TabIndex        =   27
         Top             =   713
         Width           =   1035
      End
      Begin VB.TextBox txt预约控制 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   25
         Top             =   675
         Width           =   1605
      End
      Begin VB.Label lbl限约数 
         AutoSize        =   -1  'True
         Caption         =   "限约数"
         Height          =   180
         Left            =   6120
         TabIndex        =   28
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lbl限号数 
         AutoSize        =   -1  'True
         Caption         =   "限号数"
         Height          =   180
         Left            =   4230
         TabIndex        =   19
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lbl出诊日期 
         AutoSize        =   -1  'True
         Caption         =   "出诊日期"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl替诊医生 
         AutoSize        =   -1  'True
         Caption         =   "替诊医生"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   735
         Width           =   720
      End
      Begin VB.Label lbl上班时段 
         AutoSize        =   -1  'True
         Caption         =   "上班时段"
         Height          =   180
         Left            =   2130
         TabIndex        =   17
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl预约控制 
         AutoSize        =   -1  'True
         Caption         =   "预约控制"
         Height          =   180
         Left            =   2130
         TabIndex        =   24
         Top             =   735
         Width           =   720
      End
   End
   Begin VB.Frame fra应诊诊室 
      Caption         =   "应诊诊室"
      Height          =   4125
      Left            =   30
      TabIndex        =   29
      Top             =   2730
      Width           =   7965
      Begin zl9RegEvent.ClinicPlanOffice cpoRoom 
         Height          =   3855
         Left            =   60
         TabIndex        =   30
         Top             =   210
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   6800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin zl9RegEvent.ClinicPlanUnit cpuUnit 
         Height          =   3855
         Left            =   60
         TabIndex        =   31
         Top             =   210
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   6800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fra号源信息 
      Caption         =   "号源基本信息"
      Height          =   1395
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7965
      Begin VB.TextBox txt号类 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3210
         TabIndex        =   4
         Top             =   285
         Width           =   1275
      End
      Begin VB.TextBox txtSignalNO 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   2
         Top             =   285
         Width           =   1035
      End
      Begin VB.CheckBox chk建档 
         Caption         =   "挂号时必须建档"
         Enabled         =   0   'False
         Height          =   180
         Left            =   5160
         TabIndex        =   13
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txt假日控制 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   12
         Top             =   1020
         Width           =   1935
      End
      Begin VB.TextBox txtDoctor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5160
         TabIndex        =   10
         Top             =   652
         Width           =   2625
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5160
         TabIndex        =   6
         Top             =   285
         Width           =   2625
      End
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   8
         Top             =   652
         Width           =   3615
      End
      Begin VB.Label lblSignalNO 
         AutoSize        =   -1  'True
         Caption         =   "号码"
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lbl号类 
         AutoSize        =   -1  'True
         Caption         =   "号类"
         Height          =   180
         Left            =   2820
         TabIndex        =   3
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "项目"
         Height          =   180
         Left            =   480
         TabIndex        =   7
         Top             =   712
         Width           =   360
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   180
         Left            =   4770
         TabIndex        =   5
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         Height          =   180
         Left            =   4770
         TabIndex        =   9
         Top             =   705
         Width           =   360
      End
      Begin VB.Label lbl假日控制 
         AutoSize        =   -1  'True
         Caption         =   "假日控制"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmClinicPlanOfficeAndUnitRegModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As Byte '1-门诊诊室设置,2-合作单位挂号控制
Private mobj号源 As 出诊号源, mobj出诊记录 As 出诊记录
Private mblnRecord As Boolean '是否是门诊记录
Private mblnFirst As Boolean

Private mblnOk As Boolean

Public Function ShowMe(frmParent As Form, ByVal bytFun As Byte, _
    ByVal obj号源 As 出诊号源, ByVal obj出诊记录 As 出诊记录, Optional ByVal blnRecord As Boolean) As Boolean
    '程序入口
    '参数：
    '   bytFun 1-门诊诊室设置,2-合作单位挂号控制
    If obj号源 Is Nothing Then Exit Function
    If obj出诊记录 Is Nothing Then Exit Function
    
    mbytFun = bytFun
    Set mobj号源 = obj号源: Set mobj出诊记录 = obj出诊记录
    mblnRecord = blnRecord
    
    On Error Resume Next
    If CheckDepend() = False Then Exit Function
    mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Function CheckDepend() As Boolean
    '功能:检查数据
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    '不能对历史的安排进行操作
    If DateDiff("s", mobj出诊记录.终止时间, zlDatabase.Currentdate) >= 0 Then
        MsgBox "当前系统时间已大于了安排时段的终止时间，不能进行" & IIf(mbytFun = 1, "门诊诊室", "合作单位挂号控制") & "调整操作！", vbInformation, gstrSysName
        Exit Function
    End If
    '已经停诊或未出诊安排的，不允许调整
    strSQL = "Select 1 from 临床出诊记录 Where ID=[1] and 上班时段=[2] And 停诊开始时间 Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查出诊记录", mobj出诊记录.记录ID, mobj出诊记录.时间段)
    If rsTemp.EOF Then
        MsgBox "当前安排时段不存在或已停诊，不能进行" & IIf(mbytFun = 1, "门诊诊室", "合作单位挂号控制") & "调整操作！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckDepend = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Err = 0: On Error GoTo errHandler
    If mbytFun = 1 Then '门诊诊室
        If cpoRoom.IsValied() = False Then Exit Sub
    ElseIf mbytFun = 2 Then '合作单位控制
        If cpuUnit.IsValied() = False Then Exit Sub
    End If
    
    If SaveData() = False Then Exit Sub
    mblnOk = True
    Unload Me
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function InitData() As Boolean
    Dim i As Integer
'    Dim obj所有分诊诊室 As 分诊诊室集
    Dim obj所有号序信息集 As 号序信息集, obj所有合作单位 As 合作单位控制集
    
    Err = 0: On Error GoTo errHandler
    cpoRoom.Visible = False
    cpuUnit.Visible = False
    If mbytFun = 1 Then '诊室设置
        cpoRoom.Visible = True
        Me.Caption = "出诊诊室调整"
        fra应诊诊室.Caption = "出诊诊室"
    Else '合作单位控制
        cpuUnit.Visible = True
        Me.Caption = "合作单位挂号控制调整"
        fra应诊诊室.Caption = "合作单位挂号控制"
    End If
    
    '号源信息
    txtSignalNO.Text = mobj号源.号码
    txt号类.Text = mobj号源.号类
    txtDept.Text = mobj号源.科室名称
    txtItem.Text = mobj号源.项目名称
    txtDoctor.Text = mobj号源.医生姓名
    txt假日控制.Text = Decode(mobj号源.假日控制状态, 1, "开放预约", 2, "禁止预约", 3, "受节假日设置控制", "不上班")
    chk建档.Value = IIf(mobj号源.是否建病案, vbChecked, vbUnchecked)
    If IsDate(mobj出诊记录.出诊日期) Then
        txt出诊日期.Text = Format(mobj出诊记录.出诊日期, "yyyy-mm-dd")
    Else
        txt出诊日期.Text = mobj出诊记录.出诊日期
    End If
    
    txt上班时段.Text = mobj出诊记录.时间段
    txt替诊医生.Text = mobj出诊记录.替诊医生
    txt预约控制.Text = Choose(mobj出诊记录.预约控制 + 1, "允许预约", "禁止预约", "仅禁止三方机构预约")
    chk序号控制.Value = IIf(mobj出诊记录.是否序号控制, vbChecked, vbUnchecked)
    chk时段.Value = IIf(mobj出诊记录.是否分时段, vbChecked, vbUnchecked)
    txt限号数.Text = IIf(mobj出诊记录.限号数 = 0, "", mobj出诊记录.限号数)
    txt限约数.Text = IIf(mobj出诊记录.限约数 = 0, "", mobj出诊记录.限约数)
    
    If mbytFun = 1 Then '诊室设置
'        Set obj所有分诊诊室 = GetVisitRoomsObjects(GetDoctorRooms(mobj出诊记录.科室ID))
'        obj所有分诊诊室.分诊方式 = mobj出诊记录.分诊方式
'        cpoRoom.LoadData mobj出诊记录.安排门诊诊室集, obj所有分诊诊室
    Else '合作单位控制
        Set obj所有号序信息集 = GetTimeIntervalObjects(GetTimeInterval(mobj出诊记录.记录ID, True))
        With obj所有号序信息集
            .是否分时段 = mobj出诊记录.是否分时段
            .是否序号控制 = mobj出诊记录.是否序号控制
            .限号数 = mobj出诊记录.限号数
            .限约数 = mobj出诊记录.限约数
            .预约控制 = mobj出诊记录.预约控制
        End With
        Set obj所有合作单位 = GetUnitsObjects(GetUnitAll())
        cpuUnit.LoadData mobj出诊记录.合作单位控制集, obj所有号序信息集, obj所有合作单位
    End If
    InitData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Activate()
    Dim obj所有分诊诊室 As 分诊诊室集
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = True
    
    If mbytFun = 1 Then '诊室设置
        '在这里加载是因为ListView控件的原因，如果在Load事件中加载数据，会导致数据显示不全，会出现省略号
        Set obj所有分诊诊室 = GetVisitRoomsObjects(GetDoctorRooms(mobj出诊记录.科室ID))
        obj所有分诊诊室.分诊方式 = mobj出诊记录.分诊方式
        cpoRoom.LoadData mobj出诊记录.安排门诊诊室集, obj所有分诊诊室
    End If
    If cpoRoom.Visible And cpoRoom.EditMode = ED_RegistPlan_Edit Then cpoRoom.SetFocus
    If cpuUnit.Visible And cpuUnit.EditMode = ED_RegistPlan_Edit Then cpuUnit.SetFocus
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    
    mblnFirst = True
    If mbytFun = 2 Then
        If mobj出诊记录.预约控制 = 1 Then
            '禁止预约
            MsgBox "当前安排为禁止预约，不能调整合作单位控制！", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    Call InitData
    Call SetEnabledBackColor(Me.Controls)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, cllPro As New Collection, i As Integer
    Dim byt分诊方式 As Byte, str诊室 As String, obj诊室 As 分诊诊室
    Dim obj合作单位 As 合作单位控制, obj号序 As 号序信息
    Dim cll号序 As Collection, str号序 As String, strTemp As String
    Dim blnTrans As Boolean
    Dim lng变动ID As Long
    
    Err = 0: On Error GoTo errHandler
    If mbytFun = 1 Then '门诊诊室
        Set mobj出诊记录.安排门诊诊室集 = cpoRoom.Get安排门诊诊室集
        '门诊诊室
        byt分诊方式 = mobj出诊记录.安排门诊诊室集.分诊方式
        str诊室 = ""
        For Each obj诊室 In mobj出诊记录.安排门诊诊室集
            '诊室_In:诊室1,诊室2,...
            str诊室 = str诊室 & "," & obj诊室.诊室ID
        Next
        If str诊室 <> "" Then str诊室 = Mid(str诊室, 2)
        
        'Zl_临床出诊诊室_Update(
        strSQL = "Zl_临床出诊诊室_Update("
        'Id_In       临床出诊限制.Id%Type,
        strSQL = strSQL & "" & mobj出诊记录.记录ID & ","
        '分诊方式_In 临床出诊限制.分诊方式%Type := Null,
        strSQL = strSQL & "" & byt分诊方式 & ","
        '诊室_In     Varchar2 := Null,
        strSQL = strSQL & "'" & str诊室 & "',"
        '出诊记录_In Number:=0--是否是对出诊记录进行删除
        strSQL = strSQL & "" & IIf(mblnRecord, 1, 0) & ")"
        cllPro.Add strSQL
    Else '合作单位控制
        Set mobj出诊记录.合作单位控制集 = cpuUnit.Get合作单位控制信息集
        If mblnRecord Then
            lng变动ID = zlDatabase.GetNextId("临床出诊变动记录")
            'Zl_临床出诊预约控制变动(
            strSQL = "Zl_临床出诊预约控制变动("
            '变动性质_In   临床出诊变动明细.变动性质%Type,
            strSQL = strSQL & "" & 1 & ","
            'Id_In         临床出诊变动记录.Id%Type,
            strSQL = strSQL & "" & lng变动ID & ","
            '记录id_In     临床出诊变动记录.记录id%Type := Null,
            strSQL = strSQL & "" & mobj出诊记录.记录ID & ","
            '现预约控制_In 临床出诊变动记录.现预约控制%Type := Null
            strSQL = strSQL & "" & "NULL" & ")"
            cllPro.Add strSQL
        End If
        For Each obj合作单位 In mobj出诊记录.合作单位控制集
            '预约控制:0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
            '类型:1-三方机构;2-预约方式
            Set cll号序 = New Collection
            str号序 = ""
            For Each obj号序 In obj合作单位.号序信息集
                strTemp = obj号序.序号 & "," & obj号序.数量
                If zlCommFun.ActualLen(str号序 & "|" & strTemp) > 2000 Then
                    '安排控制_in:序号1,数量|序号2,数量|...
                    str号序 = Mid(str号序, 2)
                    cll号序.Add str号序
                    str号序 = ""
                End If
                str号序 = str号序 & "|" & strTemp
            Next
            If str号序 <> "" Then
                str号序 = Mid(str号序, 2)
                cll号序.Add str号序
            End If
            For i = 1 To IIf(cll号序.Count = 0, 1, cll号序.Count)
                If mblnRecord Then
                    'Zl_临床出诊挂号控制记录_Insert(
                    strSQL = "Zl_临床出诊挂号控制记录_Insert("
                    '记录id_In   临床出诊挂号控制记录.记录id%Type,
                    strSQL = strSQL & "" & mobj出诊记录.记录ID & ","
                    '类型_In     临床出诊挂号控制记录.类型%Type,
                    strSQL = strSQL & "" & obj合作单位.类型 & ","
                    '性质_In     临床出诊挂号控制记录.性质%Type,
                    strSQL = strSQL & "" & 1 & ","
                    '名称_In     临床出诊挂号控制记录.名称%Type,
                    strSQL = strSQL & "'" & obj合作单位.合作单位名称 & "',"
                    '控制方式_In 临床出诊挂号控制记录.控制方式%Type,
                    strSQL = strSQL & "" & obj合作单位.预约控制方式 & ","
                    '是否独占_In 临床出诊记录.是否独占%Type,
                    strSQL = strSQL & "" & IIf(mobj出诊记录.合作单位控制集.是否独占, 1, 0) & ","
                    '安排控制_In Varchar2,
                    str号序 = ""
                    If cll号序.Count > 0 Then str号序 = cll号序(i)
                    strSQL = strSQL & "'" & str号序 & "',"
                    '删除_In Number:=0
                    strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                    cllPro.Add strSQL
                Else
                    'Zl_临床出诊挂号控制_Insert(
                    strSQL = "Zl_临床出诊挂号控制_Insert("
                    '限制id_In   临床出诊挂号控制.限制id%Type,
                    strSQL = strSQL & "" & mobj出诊记录.记录ID & ","
                    '类型_In     临床出诊挂号控制.类型%Type,
                    strSQL = strSQL & "" & obj合作单位.类型 & ","
                    '性质_In     临床出诊挂号控制.性质%Type,
                    strSQL = strSQL & "" & 1 & ","
                    '名称_In     临床出诊挂号控制.名称%Type,
                    strSQL = strSQL & "'" & obj合作单位.合作单位名称 & "',"
                    '控制方式_In 临床出诊挂号控制.控制方式%Type,
                    strSQL = strSQL & "" & obj合作单位.预约控制方式 & ","
                    '是否独占_In 临床出诊限制.是否独占%Type,
                    strSQL = strSQL & "" & IIf(mobj出诊记录.合作单位控制集.是否独占, 1, 0) & ","
                    '安排控制_In Varchar2,
                    str号序 = ""
                    If cll号序.Count > 0 Then str号序 = cll号序(i)
                    strSQL = strSQL & "'" & str号序 & "',"
                    '删除_In Number:=0
                    strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                    cllPro.Add strSQL
                End If
            Next
        Next
        
        'Zl_临床出诊预约控制变动(
        strSQL = "Zl_临床出诊预约控制变动("
        '变动性质_In   临床出诊变动明细.变动性质%Type,
        strSQL = strSQL & "" & 2 & ","
        'Id_In         临床出诊变动记录.Id%Type,
        strSQL = strSQL & "" & lng变动ID & ","
        '记录id_In     临床出诊变动记录.记录id%Type := Null,
        strSQL = strSQL & "" & mobj出诊记录.记录ID & ")"
        '现预约控制_In 临床出诊变动记录.现预约控制%Type := Null
        cllPro.Add strSQL
    End If
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    blnTrans = False
    SaveData = True
    Exit Function
errHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
