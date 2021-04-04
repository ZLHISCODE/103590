VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPacsNewFilter 
   Caption         =   "数据查询"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPacsNewFilter.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   7710
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.Slider sldDays 
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   529
      _Version        =   393216
      Max             =   180
      TickFrequency   =   7
   End
   Begin VB.ComboBox cboAgeType 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":000C
      Left            =   6720
      List            =   "frmPacsNewFilter.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox cboAgeWhere 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":0030
      Left            =   5280
      List            =   "frmPacsNewFilter.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtEndAge 
      Height          =   300
      Left            =   6120
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.ComboBox cboQueryTime 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":0074
      Left            =   1320
      List            =   "frmPacsNewFilter.frx":0081
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox cboPatientFrom 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":00A3
      Left            =   5400
      List            =   "frmPacsNewFilter.frx":00B6
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame fraControl 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   52
      Top             =   6240
      Width           =   7215
      Begin VB.CommandButton cmdCustomQuery 
         Caption         =   "自定义查询(&C)"
         Height          =   375
         Left            =   2760
         TabIndex        =   36
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "确  定(&O)"
         Height          =   375
         Left            =   4440
         TabIndex        =   34
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "退 出(&Q)"
         Height          =   375
         Left            =   5760
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkInputReportInf 
         Caption         =   "更多条件"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveSchema 
         Caption         =   "保存方案(&S)"
         Height          =   375
         Left            =   5400
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelSchema 
         Caption         =   "删除方案(&D)"
         Height          =   375
         Left            =   3720
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboSchemaName 
         Height          =   330
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label labSchema 
         Caption         =   "查询方案："
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   42
      Top             =   2880
      Width           =   7215
      Begin VB.ComboBox cboYangXingLv 
         Height          =   330
         ItemData        =   "frmPacsNewFilter.frx":00D4
         Left            =   4800
         List            =   "frmPacsNewFilter.frx":00E2
         TabIndex        =   21
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cboStudyDoctor 
         Height          =   330
         Left            =   4800
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboBodyPart 
         Height          =   330
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboImageType 
         Height          =   330
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboDevice 
         Height          =   330
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboPatientRoom 
         Height          =   330
         Left            =   1320
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboProcedure 
         Height          =   330
         ItemData        =   "frmPacsNewFilter.frx":00FC
         Left            =   4800
         List            =   "frmPacsNewFilter.frx":011E
         TabIndex        =   19
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txt建议 
         Height          =   300
         Left            =   4800
         TabIndex        =   29
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txt诊断意见 
         Height          =   300
         Left            =   1320
         TabIndex        =   28
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txt检查所见 
         Height          =   300
         Left            =   4800
         TabIndex        =   27
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtReportContext 
         Height          =   300
         Left            =   1320
         TabIndex        =   26
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txt随访 
         Height          =   300
         Left            =   4800
         TabIndex        =   25
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtIllnessRes 
         Height          =   300
         Left            =   1320
         TabIndex        =   24
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox cboQuality 
         Height          =   330
         ItemData        =   "frmPacsNewFilter.frx":016C
         Left            =   1320
         List            =   "frmPacsNewFilter.frx":0179
         TabIndex        =   20
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cboDiagnoseDoctor 
         Height          =   330
         Left            =   4800
         TabIndex        =   23
         Top             =   1680
         Width           =   2175
      End
      Begin VB.ComboBox cboAuditingDoctor 
         Height          =   330
         Left            =   1320
         TabIndex        =   22
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label labFilter 
         Caption         =   "阴 阳 性："
         Height          =   255
         Index           =   22
         Left            =   3720
         TabIndex        =   62
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "检查部位："
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "检查技师："
         Height          =   255
         Index           =   7
         Left            =   3720
         TabIndex        =   60
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "影像类别："
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   59
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "检查设备："
         Height          =   255
         Index           =   9
         Left            =   3720
         TabIndex        =   58
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "病人科室："
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   57
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "检查过程："
         Height          =   255
         Index           =   11
         Left            =   3720
         TabIndex        =   56
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "建    议："
         Height          =   255
         Index           =   20
         Left            =   3720
         TabIndex        =   51
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "诊断意见："
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   50
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "影像质量："
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   49
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "审核医生："
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   48
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "诊断医生："
         Height          =   255
         Index           =   16
         Left            =   3720
         TabIndex        =   47
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "疾病诊断："
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   46
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "随    访："
         Height          =   255
         Index           =   14
         Left            =   3720
         TabIndex        =   45
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "报告内容："
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   44
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "检查所见："
         Height          =   255
         Index           =   12
         Left            =   3720
         TabIndex        =   43
         Top             =   2520
         Width           =   1095
      End
   End
   Begin VB.ComboBox cboNumType 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":0187
      Left            =   1320
      List            =   "frmPacsNewFilter.frx":01A3
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpBirthDay 
      Height          =   420
      Left            =   4680
      TabIndex        =   9
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   63438849
      CurrentDate     =   40372
   End
   Begin VB.TextBox txtStartAge 
      Height          =   300
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.ComboBox cboSex 
      Height          =   330
      Left            =   1320
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtNum 
      Height          =   300
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   330
      Left            =   5400
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   63438851
      CurrentDate     =   38082
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   330
      Left            =   3000
      TabIndex        =   11
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   63438851
      CurrentDate     =   40372
   End
   Begin VB.Label labDays 
      Alignment       =   2  'Center
      Caption         =   "今天"
      Height          =   255
      Left            =   840
      TabIndex        =   66
      Top             =   2640
      Width           =   5895
   End
   Begin VB.Label Label 
      Caption         =   "近半年"
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   65
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "今天"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   64
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   63
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label labFilter 
      Caption         =   "查询时间："
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   55
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "病人来源："
      Height          =   255
      Index           =   21
      Left            =   4320
      TabIndex        =   54
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "出生日期："
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   41
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "年    龄："
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   40
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "性    别："
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   39
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "姓    名："
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   38
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "查询号码："
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   37
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmPacsNewFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDepartmentId As Long   '保存当前所属的部门ID
Private marrParValue(25) As String '保存界面所设置的参数值
Private mblnOk As Boolean


Public index_查询号码 As Integer
Public index_病人来源 As Integer
Public index_病人姓名 As Integer
Public index_开始年龄 As Integer
Public index_结束年龄 As Integer
Public index_病人性别 As Integer
Public index_出生日期 As Integer
Public index_查询开始时间 As Integer
Public index_查询结束时间 As Integer
Public index_检查部位 As Integer
Public index_检查技师 As Integer
Public index_影像类别 As Integer
Public index_检查设备 As Integer
Public index_病人科室 As Integer
Public index_检查过程 As Integer
Public index_影像质量 As Integer
Public index_阴阳性 As Integer
Public index_审核医生 As Integer
Public index_诊断医生 As Integer
Public index_疾病诊断 As Integer
Public index_随访 As Integer
Public index_报告内容 As Integer
Public index_检查所见 As Integer
Public index_诊断意见 As Integer
Public index_建议 As Integer



Public Function ShowFilter(ByVal lngDepartmentId, arrParameter() As String, owner As Form) As String
    mblnOk = False
    mDepartmentId = lngDepartmentId
    
    Call SetParIndex
    
    Me.Show 1, owner
    
    If mblnOk Then
        Call SetParValue(arrParameter)
        ShowFilter = GetQueryFilter()
    End If
End Function

Private Sub SetParIndex()
'********************************************
'
'初始化参数索引取值
'
'********************************************
    On Error GoTo errHandle
        index_查询号码 = 1
        index_病人来源 = 2
        index_病人姓名 = 3
        index_开始年龄 = 4
        index_结束年龄 = 5
        index_病人性别 = 6
        index_出生日期 = 7
        index_查询开始时间 = 8
        index_查询结束时间 = 9
        index_检查部位 = 10
        index_检查技师 = 11
        index_影像类别 = 12
        index_检查设备 = 13
        index_病人科室 = 14
        index_检查过程 = 15
        index_影像质量 = 16
        index_阴阳性 = 17
        index_审核医生 = 18
        index_诊断医生 = 19
        index_疾病诊断 = 20
        index_随访 = 21
        index_报告内容 = 22
        index_检查所见 = 23
        index_诊断意见 = 24
        index_建议 = 25
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub SetParValue(arrParValue() As String)
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrParValue(index_查询号码) = txtNum.Text
    
    '取得病人来源
    'arrParValue(index_病人来源) = cboPatientFrom.Text
    Select Case cboPatientFrom.Text
        Case "门诊"
            arrParValue(index_病人来源) = 1
        Case "住院"
            arrParValue(index_病人来源) = 2
        Case "外诊"
            arrParValue(index_病人来源) = 3
        Case "体检"
            arrParValue(index_病人来源) = 4
    End Select
    
    arrParValue(index_病人姓名) = txtName.Text
    
    '需要将病人年龄转换成指定的天数
    'arrParValue(index_开始年龄) = txtStartAge.Text & cboAgeType.Text
    'arrParValue(index_结束年龄) = txtEndAge.Text & cboAgeType.Text
    Select Case cboAgeType.Text
        Case "岁"
            arrParValue(index_开始年龄) = Val(txtStartAge.Text) * 365
            arrParValue(index_结束年龄) = Val(txtEndAge.Text) * 365
        Case "月"
            arrParValue(index_开始年龄) = Val(txtStartAge.Text) * 30
            arrParValue(index_结束年龄) = Val(txtEndAge.Text) * 30
        Case "周"
            arrParValue(index_开始年龄) = Val(txtStartAge.Text) * 7
            arrParValue(index_结束年龄) = Val(txtEndAge.Text) * 7
        Case "天"
            arrParValue(index_开始年龄) = Val(txtStartAge.Text) * 1
            arrParValue(index_结束年龄) = Val(txtEndAge.Text) * 1
    End Select
    
    
    arrParValue(index_病人性别) = cboSex.Text
    
    If Not dtpBirthDay Is Nothing Then
        arrParValue(index_出生日期) = dtpBirthDay.value
    End If
    
    arrParValue(index_查询开始时间) = dtpBegin.value
    arrParValue(index_查询结束时间) = dtpEnd.value
    arrParValue(index_检查部位) = cboBodyPart.Text
    arrParValue(index_检查技师) = cboStudyDoctor.Text
    arrParValue(index_影像类别) = cboImageType.Text
    
    If Trim(cboDevice.Text <> "") Then arrParValue(index_检查设备) = cboDevice.ItemData(cboDevice.ListIndex) '保存检查设备号
    If Trim(cboPatientRoom.Text <> "") Then arrParValue(index_病人科室) = cboPatientRoom.ItemData(cboPatientRoom.ListIndex) '保存病人的科室ID
    
    arrParValue(index_检查过程) = cboProcedure.Text
    arrParValue(index_影像质量) = cboQuality.Text
    
    '在处理阴阳性时，0表示阴性，1表示阳性
    If cboYangXingLv.Text = "结果阳性" Then
        arrParValue(index_阴阳性) = 1
    Else
        arrParValue(index_阴阳性) = 0
    End If
    
    arrParValue(index_审核医生) = cboAuditingDoctor.Text
    arrParValue(index_诊断医生) = cboDiagnoseDoctor.Text
    arrParValue(index_疾病诊断) = txtIllnessRes.Text
    arrParValue(index_随访) = txt随访.Text
    arrParValue(index_报告内容) = txtReportContext.Text
    arrParValue(index_检查所见) = txt检查所见.Text
    arrParValue(index_诊断意见) = txt诊断意见.Text
    arrParValue(index_建议) = txt建议.Text
           
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetQueryFilter() As String
    Dim strFilter As String
    Dim strSubFilter As String
    Dim strQueryField As String
    
    On Error GoTo errHandle
    
    '查询号码
    If Trim(txtNum.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strQueryField = GetQueryNumField(cboNumType.Text)
        
        If strQueryField <> "影像标本核收取材.病理号" Then
            strFilter = strFilter & GetQueryNumField(cboNumType.Text) & "=[" & index_查询号码 & "]"
        Else
            strFilter = strFilter & "病人医嘱记录.诊疗项目ID IN (select a.诊疗项目 from 影像标本核收取材 a where a.病理号=[" & index_查询号码 & "])"
        End If
    End If
    
    '病人来源
    If Trim(cboPatientFrom.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "病人医嘱记录.病人来源=[" & index_病人来源 & "]"
    End If
    
    '病人姓名
    If Trim(txtName.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "病人信息.姓名=[" & index_病人姓名 & "]"
    End If
    
    '病人年龄-开始年龄(只有当条件使用“到”，即在多少年龄之间时，才使用开始年龄)
    If Trim(txtStartAge.Text) <> "" Then
        If cboAgeWhere.Text = "到" Then
            If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "AgeToDays(病人信息.年龄)>=[" & index_开始年龄 & "]"
        End If
    End If
    
    '病人年龄-结束年龄
    If Trim(txtEndAge.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        If cboAgeWhere.Text = "到" Then
            strFilter = strFilter & "AgeToDays(病人信息.年龄)<=[" & index_结束年龄 & "]"
        Else
            strFilter = strFilter & "AgeToDays(病人信息.年龄)" & GetQueryAgeWhere(cboAgeWhere.Text) & "[" & index_结束年龄 & "]"
        End If
    End If
    
    '病人性别
    If Trim(cboSex.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "病人信息.性别=[" & index_病人性别 & "]"
    End If
    
    '出生日期
    If Trim(dtpBirthDay.value) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "病人信息.出生日期=[" & index_出生日期 & "]"
    End If
    
    '检查日期-开始日期(该条件是必选的条件)
    If Trim(dtpBegin.value) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & GetQueryTimeField(cboQueryTime.Text) & ">=[" & index_查询开始时间 & "]"
    End If
    
    '检查日期-结束日期(该条件是必选的条件)
    If Trim(dtpEnd.value) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & GetQueryTimeField(cboQueryTime.Text) & "<=[" & index_查询结束时间 & "]"
    End If
    
    '检查部位
    If Trim(cboBodyPart.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "Instr(病人医嘱记录.医嘱内容, [" & index_检查部位 & "]) > 0"
    End If
    
    '检查技师
    If Trim(cboStudyDoctor.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "影像检查记录.检查技师=[" & index_检查技师 & "]"
    End If
    
    '影像类别
    If Trim(cboImageType.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "影像检查记录.影像类别=[" & index_影像类别 & "]"
    End If
    
    '检查设备
    If Trim(cboDevice.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "影像检查记录.检查设备=[" & index_检查设备 & "]"
    End If
    
    '病人科室 "+0"表示不走索引，有些地方使用索引查询在效率上相对比较低效
    If Trim(cboPatientRoom.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "病人医嘱记录.病人科室ID+0=[" & index_病人科室 & "]"
    End If
    
    '检查过程
    If Trim(cboProcedure.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "

        Select Case cboProcedure.Text
            Case "已登记"
                strFilter = strFilter & " ( 病人医嘱发送.执行过程=0 or 病人医嘱发送.执行过程 = 1 or 病人医嘱发送.执行过程 IS NULL) "
            Case "已报到"
                strFilter = strFilter & " ( 病人医嘱发送.执行过程=2 and 影像检查记录.报告人 IS NULL)"
            Case "已检查"
                strFilter = strFilter & " ( 病人医嘱发送.执行过程=3 and 影像检查记录.报告人 IS NULL)"
            Case "处理中"
                strFilter = strFilter & " ( not 影像检查记录.报告操作 IS NULL)"
            Case "报告中"
                strFilter = strFilter & " (( 病人医嘱发送.执行过程 =2 or 病人医嘱发送.执行过程=3) and not 影像检查记录.报告人 is null and 影像检查记录.报告操作 is null) "
            Case "已报告"
                strFilter = strFilter & " (病人医嘱发送.执行过程=4 and 影像检查记录.复核人 is null) "
            Case "审核中"
                strFilter = strFilter & " (病人医嘱发送.执行过程=4 and not 影像检查记录.复核人 is null) "
            Case "已审核"
                strFilter = strFilter & " 病人医嘱发送.执行过程=5 "
            Case "已完成"
                strFilter = strFilter & " 病人医嘱发送.执行过程=6 "
        End Select
    End If
    
    '影像质量
    If Trim(cboQuality.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "影像检查记录.影像质量=[" & index_影像质量 & "]"
    End If
    
    '阴阳性
    If Trim(cboYangXingLv.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " Nvl(病人医嘱发送.结果阳性, 0)=[" & index_阴阳性 & "]"
    End If
    
    '审核医生
    If Trim(cboAuditingDoctor.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "影像检查记录.复核人=[" & index_审核医生 & "]"
    End If
    
    '诊断医生
    If Trim(cboDiagnoseDoctor.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "影像检查记录.报告人=[" & index_诊断医生 & "]"
    End If
    

    '随访
    If Trim(txt随访.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "影像检查记录.随访描述=[" & index_随访 & "]"
    End If
    
    '疾病诊断 - 需要与其他表进行关联查询
    If Trim(txtIllnessRes.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "病人医嘱记录.ID IN (Select t.医嘱id From 病人医嘱报告 t Where t.病历id IN " & _
                                                                        " (Select Distinct a.ID  " & _
                                                                        " From 电子病历记录 a,电子病历内容 b " & _
                                                                        " Where a.创建时间>[" & index_查询开始时间 & "] AND a.Id=b.文件ID  " & _
                                                                        " And b.对象类型=7 And instr(b.对象属性,'52;')>0 And instr(b.内容文本,[" & index_疾病诊断 & "])>0))"
    End If
    
    '报告内容
    If Trim(txtReportContext.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " And 病人医嘱记录.ID IN (Select t.医嘱id From 病人医嘱报告 t Where t.病历id IN " & _
                                                                " (Select Distinct a.ID " & _
                                                                " From 电子病历记录 a,电子病历内容 b " & _
                                                                " Where a.创建时间>[" & index_查询开始时间 & "] AND A.Id=b.文件ID " & _
                                                                " And b.对象类型=2 And instr(b.内容文本,[" & index_报告内容 & "])>0 And b.终止版 = 0)) "
    End If
    
    '检查所见
    If Trim(txt检查所见.Text) <> "" Then
        If Trim(strSubFilter) <> "" Then strSubFilter = strSubFilter & " or "
        strSubFilter = strSubFilter & " (b.内容文本 ='检查所见' And Instr(c.内容文本, [" & index_检查所见 & "]) > 0)"
    End If
    
    '诊断意见
    If Trim(txt诊断意见.Text) <> "" Then
        If Trim(strSubFilter) <> "" Then strSubFilter = strSubFilter & " or "
        strSubFilter = strSubFilter & " (b.内容文本 ='诊断意见' And Instr(c.内容文本, [" & index_诊断意见 & "]) > 0)"
    End If
    
    '建议
    If Trim(txt建议.Text) <> "" Then
        If Trim(strSubFilter) <> "" Then strSubFilter = strSubFilter & " or "
        strSubFilter = strSubFilter & " (b.内容文本 ='建议' And Instr(c.内容文本, [" & index_建议 & "]) > 0)"
    End If
    
    If strSubFilter <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        
        strSubFilter = " (" & strSubFilter & ")"
        
        strFilter = strFilter & " 病人医嘱记录.ID IN ( Select t.医嘱id From 病人医嘱报告 t Where t.病历id IN " & _
            " (Select Distinct a.ID From 电子病历记录 a, 电子病历内容 b,电子病历内容 c " _
            & " Where a.创建时间 > [" & index_查询开始时间 & "] And a.Id = b.文件id And b.Id = C.父ID And b.对象类型 = 3 And c.对象类型 = 2 And c.终止版 = 0 and " _
            & strSubFilter & "))"
    End If
    
    GetQueryFilter = strFilter
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetQueryTimeField(ByVal strTimeType As String) As String
'********************************************
'
'取得日期查询的相关字段
'
'********************************************
    Dim strField As String
    
    On Error GoTo errHandle
    
    Select Case strTimeType
        Case "申请时间"
            strField = "病人医嘱发送.发送时间"
        Case "报到时间"
            strField = "病人医嘱发送.首次时间"
        Case "采图时间"
            strField = "影像检查记录.接收日期"
    End Select
    
    GetQueryTimeField = strField
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog

End Function


Private Function GetQueryAgeWhere(ByVal strAgeWhere) As String
'********************************************
'
'获取当使用年龄进行查询时，所用的查询条件符号
'
'********************************************
    Dim strWhere As String
    
    On Error GoTo errHandle
    
    Select Case strAgeWhere
        Case "大于"
            strWhere = ">"
        Case "大于等于"
            strWhere = ">="
        Case "小于"
            strWhere = "<"
        Case "小于等于"
            strWhere = "<="
        Case "等于"
            strWhere = "="
        Case "到"
            strWhere = ""
    End Select
    
    GetQueryAgeWhere = strWhere
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    
End Function


Private Function GetQueryNumField(ByVal strNumType As String) As String
'********************************************
'
'取得号码查询所需要的字段
'
'********************************************
    Dim strField As String
    
    On Error GoTo errHandle
    
    Select Case strNumType
        Case "门诊号"
            strField = "病人信息.门诊号"
        Case "住院号"
            strField = "病人信息.住院号"
        Case "就诊卡号"
            strField = "病人信息.就诊卡号"
        Case "单据号"
            strField = "病人医嘱发送.No"
        Case "IC卡卡号"
            strField = "病人信息.IC卡号"
        Case "检查号"
            strField = "影像检查记录.检查号"
        Case "病理号"
            strField = "影像标本核收取材.病理号"
        Case "身份证"
            strField = "病人信息.身份证号"
    End Select
    
    GetQueryNumField = strField
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub InitFaceData()
'********************************************
'
'初始化查询界面数据
'
'********************************************

    '将日期控件的日期修改为当天日期
    dtpBirthDay.value = Now
    dtpBegin.value = CDate(Now - time)
    dtpEnd.value = Now
    
    dtpBirthDay.value = ""
            
            
    '载入检查部位
    Call LoadStudyPart
    '载入图象类别
    Call LoadImageType
    '载入相关医生
    Call LoadDoctor
    '载入病人性别
    Call LoadSex
    '载入病人科室
    Call LoadPatientRoom
    '读取检查设备
    Call LoadStudyDevice
    
    
    cboNumType.ListIndex = 2 '默认按照就诊卡号查询
    cboAgeWhere.ListIndex = 4 '默认年龄的查询条件为等于
    cboAgeType.ListIndex = 0 '默认的年龄单位为岁
    cboQueryTime.ListIndex = 0 '默认的查询时间为申请时间
    cboPatientFrom.ListIndex = 0 '默认不对病人来源进行判断


    'Call txtName.SetFocus
End Sub


Private Sub LoadStudyDevice()
'********************************************
'
'读取检查设备
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    strSql = "Select 设备号, 设备名 From 影像设备目录 where 类型=4 and 状态=1"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取检查设备")
    cboDevice.Clear
    cboDevice.AddItem ""
    cboDevice.ItemData(0) = -1
        
    With Me.cboDevice
        Do While Not rsTmp.EOF
            .AddItem rsTmp!设备号 & "-" & Nvl(rsTmp!设备名)
            .ItemData(cboDevice.NewIndex) = rsTmp!设备号
            
            rsTmp.MoveNext
        Loop
    End With
    
    cboDevice.ListIndex = 0
 
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadPatientRoom()
'********************************************
'
'读取病人科室
'
'********************************************
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Long
    
    strSql = "Select Distinct A.ID,A.编码,A.名称,B.服务对象" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.工作性质 IN('临床','手术')" & _
        " And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.编码"
        
    On Error GoTo errHandle
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    cboPatientRoom.Clear
    cboPatientRoom.AddItem ""
    cboPatientRoom.ItemData(0) = -1
    
    For i = 1 To rsTmp.RecordCount
        cboPatientRoom.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboPatientRoom.ItemData(cboPatientRoom.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next
    
    cboPatientRoom.ListIndex = 0

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadSex()
'********************************************
'
'读取病人性别
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    strSql = "Select 名称 From 性别"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病人性别")
    cboSex.Clear
    cboSex.AddItem ""
        
    With Me.cboSex
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("名称"))) & "-" & Nvl(rsTmp("名称"))
            rsTmp.MoveNext
        Loop
    End With
    
    cboSex.ListIndex = 0
 
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadDoctor()
'********************************************
'
'读取诊断医生、审核医生、检查医生
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    
    cboDiagnoseDoctor.Clear
    cboAuditingDoctor.Clear
    cboStudyDoctor.Clear
    
    cboDiagnoseDoctor.AddItem ""
    cboAuditingDoctor.AddItem ""
    cboStudyDoctor.AddItem ""
        
    
    strSql = "select distinct A.简码,A.姓名 from 人员表 A,部门人员 B where B.部门ID=[1] AND A.ID=B.人员ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取本科室医生", mDepartmentId)
    
    If rsTmp Is Nothing Then Exit Sub
    
    Do While Not rsTmp.EOF
        cboDiagnoseDoctor.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cboAuditingDoctor.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cboStudyDoctor.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        rsTmp.MoveNext
    Loop
    
    cboDiagnoseDoctor.ListIndex = 0
    cboAuditingDoctor.ListIndex = 0
    cboStudyDoctor.ListIndex = 0
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub LoadStudyPart()
'********************************************
'
'读取检查部位
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    strSql = "Select Distinct 名称 From 诊疗检查部位"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "提取部位")
    
    cboBodyPart.Clear
    cboBodyPart.AddItem ""
        
    With Me.cboBodyPart
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("名称"))) & "-" & Nvl(rsTmp("名称"))
            rsTmp.MoveNext
        Loop
    End With
    
    cboBodyPart.ListIndex = 0
 
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadImageType()
'********************************************
'
'读取图像类别
'
'********************************************
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "select 编码,名称 from 影像检查类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "影像检查类别")
    
    cboImageType.Clear
    cboImageType.AddItem ""
    
    Do While Not rsTemp.EOF
        cboImageType.AddItem rsTemp!名称 & "-" & rsTemp!编码
        rsTemp.MoveNext
    Loop
    
    cboImageType.ListIndex = 0
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetDaysHint(ByVal intDays As Integer) As String
'********************************************
'
'取得当前天数对应的文字说明
'
'intDays：需要转换成文字说明的具体天数
'
'********************************************
    Dim strReturn As String
    
    On Error GoTo errHandle
    
    If intDays = 0 Then strReturn = "今天"
    If intDays >= 1 And intDays < 7 Then strReturn = intDays & "天(近" & intDays & "天)"
    If intDays >= 7 And intDays < 30 Then strReturn = intDays & "天(近" & Int(intDays / 7) & "周)"
    If intDays >= 30 And intDays < 180 Then strReturn = intDays & "天(近" & Int(intDays / 30) & "月)"
    If intDays >= 180 And intDays < 365 Then strReturn = intDays & "天(近半年)"
    
    GetDaysHint = strReturn
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function



Private Sub cmdQuit_Click()
    mblnOk = False
    Me.Hide
End Sub

Private Sub cmdSure_Click()
    If dtpEnd.value < dtpBegin.value Then
        MsgBox "查询时间的开始时间不能大于截止时间，请检查！", vbInformation, gstrSysName
        dtpEnd.SetFocus
        Exit Sub
    End If
    
    mblnOk = True
    Me.Hide
End Sub

Private Sub Form_Load()
    mblnOk = False
    
    Call InitFaceData
End Sub

Private Sub sldDays_Change()
    '设置查询时间范围
    dtpBegin.value = CDate(dtpEnd.value - sldDays.value)
End Sub


Private Sub sldDays_Scroll()
    labDays.Caption = GetDaysHint(sldDays.value)
End Sub

