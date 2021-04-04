VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLabSampleRegisterFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabSampleRegisterFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   345
      Left            =   2130
      TabIndex        =   20
      Top             =   2850
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   3570
      TabIndex        =   21
      Top             =   2850
      Width           =   1095
   End
   Begin VB.Frame fraFilter 
      Height          =   2685
      Left            =   30
      TabIndex        =   22
      Top             =   30
      Width           =   4875
      Begin VB.Frame Frame4 
         Height          =   45
         Left            =   60
         TabIndex        =   26
         Top             =   2070
         Width           =   4725
      End
      Begin VB.Frame Frame3 
         Height          =   45
         Left            =   60
         TabIndex        =   25
         Top             =   1560
         Width           =   4725
      End
      Begin VB.ComboBox cboCapture 
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1140
         Width           =   1395
      End
      Begin VB.ComboBox cboSample 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1140
         Width           =   1365
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   60
         TabIndex        =   24
         Top             =   1560
         Width           =   4725
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   60
         TabIndex        =   23
         Top             =   990
         Width           =   4725
      End
      Begin VB.CheckBox chkOutPatient 
         Caption         =   "门诊"
         Height          =   255
         Left            =   990
         TabIndex        =   13
         Top             =   1710
         Width           =   795
      End
      Begin VB.CheckBox chkInpatient 
         Caption         =   "住院"
         Height          =   255
         Left            =   2145
         TabIndex        =   14
         Top             =   1710
         Width           =   795
      End
      Begin VB.CheckBox chkPhysical 
         Caption         =   "体检"
         Height          =   255
         Left            =   3300
         TabIndex        =   15
         Top             =   1710
         Width           =   795
      End
      Begin VB.TextBox TxtID 
         Height          =   285
         Left            =   990
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox TxtSickCard 
         Height          =   285
         Left            =   3300
         TabIndex        =   3
         Top             =   240
         Width           =   1395
      End
      Begin VB.TextBox TxtName 
         Height          =   285
         Left            =   990
         TabIndex        =   5
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox TxtNo 
         Height          =   285
         Left            =   3300
         TabIndex        =   7
         Top             =   630
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   285
         Left            =   990
         TabIndex        =   17
         Top             =   2220
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Format          =   110231553
         CurrentDate     =   39034
      End
      Begin MSComCtl2.DTPicker DTPEND 
         Height          =   285
         Left            =   3300
         TabIndex        =   19
         Top             =   2220
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Format          =   110231553
         CurrentDate     =   39034
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "采集方式"
         Height          =   195
         Left            =   2460
         TabIndex        =   10
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "标         本"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人来源"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label lblLabel6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发送时间"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   2265
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ">>>>>>"
         Height          =   195
         Left            =   2490
         TabIndex        =   18
         Top             =   2265
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标识号(&1)"
         Height          =   180
         Left            =   150
         TabIndex        =   0
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡(&2)"
         Height          =   180
         Left            =   2460
         TabIndex        =   2
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓    名(&3)"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   675
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号(&4)"
         Height          =   180
         Left            =   2460
         TabIndex        =   6
         Top             =   675
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmLabSampleRegisterFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDateOldEnd As Date                                            '记录旧的时间
Private mstrFilter As String                                            '过滤字串
Private Enum mFilter
    标识号 = 0
    就诊卡
    姓名
    单据号
    标本
    采集方式
    门诊
    住院
    体检
    病人科室
    间隔时间
    开始时间
    结束时间
End Enum

Private Sub cmdOK_Click()
    Dim dateSpace As Integer
    Dim strFilter As String                             '过滤条件字串
    
    dateSpace = DateDiff("d", Me.DTPBegin.Value, Me.DTPEND.Value)
    
    strFilter = Me.TxtID & ";" & TxtSickCard & ";" & TxtName & ";" & TxtNo & ";" & Mid(cboSample, InStr(1, cboSample, "-") + 1) & _
                ";" & cboCapture.ItemData(cboCapture.ListIndex) & ";" & IIf(chkOutPatient, 1, "") & ";" & _
                IIf(chkInpatient, 2, "") & ";" & IIf(chkPhysical, 4, "") & ";" & "0" & ";" & _
                dateSpace & ";" & IIf(mDateOldEnd <> DTPEND.Value, DTPBegin.Value, "") & ";" & _
                IIf(mDateOldEnd <> DTPEND.Value, DTPEND.Value, "")

    zlDatabase.SetPara "标本登记过滤", strFilter, 100, 1212
    '传出供主窗体调用
    mstrFilter = strFilter
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub DTPBegin_Change()
    If Me.DTPBegin > Me.DTPEND Then
        Me.DTPBegin = Me.DTPEND
    End If
End Sub

Private Sub DTPEND_Change()
    If Me.DTPEND < Me.DTPBegin Then
        Me.DTPEND = Me.DTPBegin
    End If
End Sub

Private Sub Form_Load()
    InitinterFace
End Sub

Private Sub InitinterFace()
    '初使化界面
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String
    Dim intLoop As Integer                          '循环变量
    Dim strTmp As String                            '临时字串变量
    Dim varFilter As Variant                        '过滤字串分解
        
    On Error GoTo errH
    
    strTmp = zlDatabase.GetPara("标本登记过滤", 100, 1212, "")
    
    If strTmp <> "" Then
        varFilter = Split(strTmp, ";")
        Me.chkOutPatient = IIf(Val(varFilter(mFilter.门诊)) = 0, 0, 1)
        Me.chkInpatient = IIf(Val(varFilter(mFilter.住院)) = 0, 0, 1)
        Me.chkPhysical = IIf(Val(varFilter(mFilter.体检)) = 0, 0, 1)
    Else
        Me.chkOutPatient = 1
        Me.chkInpatient = 1
        Me.chkPhysical = 1
    End If
    
    mDateOldEnd = Me.DTPEND.Value
    
    '===病人所在科室
    strsql = "Select Distinct A.ID,A.编码,A.名称,B.服务对象" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.工作性质 IN('临床','手术')" & _
        " And B.服务对象 IN(3,[1],[2])" & _
        " And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.编码"
    
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, IIf(chkOutPatient.Value = 1 Or chkPhysical.Value = 1, 1, -1), IIf(chkInpatient.Value = 1, 2, -1))

    
    
    '===读入检验标本
    strsql = "select 编码,名称 from 诊疗检验标本 order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName)
    cboSample.Clear
    cboSample.AddItem "所有标本"
    cboSample.ItemData(cboSample.NewIndex) = 0
    Do Until rsTmp.EOF
        cboSample.AddItem rsTmp("编码") & "-" & rsTmp("名称")
        cboSample.ItemData(cboSample.NewIndex) = rsTmp("编码")
        If strTmp <> "" Then
            If rsTmp("名称") = varFilter(mFilter.标本) Then
                cboSample.ListIndex = cboSample.NewIndex
            End If
        End If
        rsTmp.MoveNext
    Loop
    If cboSample.Text = "" And cboSample.ListCount > 0 Then cboSample.ListIndex = 0
    
    '===读入采集方式
    strsql = "select ID,名称 from 诊疗项目目录 where 类别='E' and 操作类型 = '6'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName)
    cboCapture.Clear
    cboCapture.AddItem "所有方式"
    cboCapture.ItemData(cboCapture.NewIndex) = 0
    Do Until rsTmp.EOF
        cboCapture.AddItem rsTmp("名称")
        cboCapture.ItemData(cboCapture.NewIndex) = rsTmp("ID")
        If strTmp <> "" Then
            If CLng(varFilter(mFilter.采集方式)) = rsTmp("ID") Then
                cboCapture.ListIndex = cboCapture.NewIndex
            End If
        End If
        rsTmp.MoveNext
    Loop
    If cboCapture.Text = "" And cboCapture.ListCount > 0 Then cboCapture.ListIndex = 0
    
    '读入时间
    If strTmp <> "" Then
        Me.DTPBegin.Value = Now - varFilter(mFilter.间隔时间)
        Me.DTPEND.Value = Now
    Else
        Me.DTPBegin.Value = Now - 3
        Me.DTPEND.Value = Now
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowMe(Objfrm As Object, ByRef strFilter As String)
    Me.Show vbModal, Objfrm
    strFilter = mstrFilter
End Sub
