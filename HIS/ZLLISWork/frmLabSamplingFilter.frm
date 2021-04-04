VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLabSamplingFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabSamplingFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   345
      Left            =   3450
      TabIndex        =   15
      Top             =   4935
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   4890
      TabIndex        =   16
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame fraFilter 
      Height          =   4665
      Left            =   75
      TabIndex        =   23
      Top             =   90
      Width           =   6075
      Begin VB.CheckBox chkCapture 
         Caption         =   "采集方式"
         Height          =   225
         Left            =   4080
         TabIndex        =   13
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox chkSample 
         Caption         =   "标本"
         Height          =   255
         Left            =   2130
         TabIndex        =   11
         Top             =   2145
         Width           =   1125
      End
      Begin VB.CheckBox chkSampeleType 
         Caption         =   "检验类型"
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   2130
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   60
         TabIndex        =   27
         Top             =   2010
         Width           =   5925
      End
      Begin VB.ListBox lstCapture 
         Height          =   2085
         ItemData        =   "frmLabSamplingFilter.frx":000C
         Left            =   4050
         List            =   "frmLabSamplingFilter.frx":0013
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   2430
         Width           =   1935
      End
      Begin VB.ListBox lstSample 
         Height          =   2085
         ItemData        =   "frmLabSamplingFilter.frx":0020
         Left            =   2070
         List            =   "frmLabSamplingFilter.frx":0027
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   2430
         Width           =   1935
      End
      Begin VB.ListBox lst类型 
         Height          =   2085
         ItemData        =   "frmLabSamplingFilter.frx":0034
         Left            =   105
         List            =   "frmLabSamplingFilter.frx":003B
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   2430
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Height          =   45
         Left            =   60
         TabIndex        =   25
         Top             =   1440
         Width           =   5925
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   60
         TabIndex        =   24
         Top             =   990
         Width           =   5925
      End
      Begin VB.CheckBox chkOutPatient 
         Caption         =   "门诊"
         Height          =   255
         Left            =   990
         TabIndex        =   4
         Top             =   1140
         Width           =   795
      End
      Begin VB.CheckBox chkInpatient 
         Caption         =   "住院"
         Height          =   255
         Left            =   2145
         TabIndex        =   5
         Top             =   1140
         Width           =   795
      End
      Begin VB.CheckBox chkPhysical 
         Caption         =   "体检"
         Height          =   255
         Left            =   3300
         TabIndex        =   6
         Top             =   1140
         Width           =   795
      End
      Begin VB.TextBox TxtID 
         Height          =   285
         Left            =   990
         TabIndex        =   0
         Top             =   240
         Width           =   1965
      End
      Begin VB.TextBox TxtSickCard 
         Height          =   285
         Left            =   4020
         TabIndex        =   1
         Top             =   240
         Width           =   1965
      End
      Begin VB.TextBox TxtName 
         Height          =   285
         Left            =   990
         TabIndex        =   2
         Top             =   630
         Width           =   1965
      End
      Begin VB.TextBox TxtNo 
         Height          =   285
         Left            =   4020
         TabIndex        =   3
         Top             =   630
         Width           =   1965
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   285
         Left            =   990
         TabIndex        =   7
         Top             =   1620
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   503
         _Version        =   393216
         Format          =   195952641
         CurrentDate     =   39034
      End
      Begin MSComCtl2.DTPicker DTPEND 
         Height          =   285
         Left            =   4020
         TabIndex        =   8
         Top             =   1620
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   503
         _Version        =   393216
         Format          =   195952641
         CurrentDate     =   39034
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发送时间"
         Height          =   195
         Left            =   105
         TabIndex        =   26
         Top             =   1665
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人来源"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   21
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ">>>>>>"
         Height          =   195
         Left            =   3180
         TabIndex        =   22
         Top             =   1665
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标识号(&1)"
         Height          =   180
         Left            =   105
         TabIndex        =   17
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡(&2)"
         Height          =   180
         Left            =   3180
         TabIndex        =   18
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓    名(&3)"
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   675
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号(&4)"
         Height          =   180
         Left            =   3180
         TabIndex        =   20
         Top             =   675
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmLabSamplingFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDateOldEnd As Date                                             '记录旧的时间
Private mstrFilter As String                                            '过滤字串
Private mintType As Integer                                             '复选框和列表的点击状态，1=点击复选框，2=点击列表

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
    间隔时间
    发送或审核时间          '=0 发送时间 = 1 审核时间
    开始时间
    结束时间
    检验类型
End Enum


Private Sub chkCapture_Click()
    Dim intLoop As Integer
    
    If mintType = 1 Then Exit Sub
    mintType = 2
    With Me.lstCapture
        For intLoop = 0 To .ListCount - 1
            .Selected(intLoop) = Me.chkCapture.Value
        Next
    End With
    mintType = 0
End Sub

Private Sub chkCapture_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Me.chkCapture
            If .Value = 1 Then
                .Value = 0
            Else
                .Value = 1
            End If
        End With
    End If
End Sub

Private Sub chkInpatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Me.chkInpatient
            If .Value = 1 Then
                .Value = 0
            Else
                .Value = 1
            End If
        End With
    End If
End Sub

Private Sub chkOutPatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Me.chkOutPatient
            If .Value = 1 Then
                .Value = 0
            Else
                .Value = 1
            End If
        End With
    End If
End Sub

Private Sub chkPhysical_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Me.chkPhysical
            If .Value = 1 Then
                .Value = 0
            Else
                .Value = 1
            End If
        End With
    End If
End Sub

Private Sub chkSampeleType_Click()
    Dim intLoop As Integer
    
    If mintType = 1 Then Exit Sub
    mintType = 2
    With Me.lst类型
        For intLoop = 0 To .ListCount - 1
            .Selected(intLoop) = Me.chkSampeleType.Value
        Next
    End With
    mintType = 0
End Sub

Private Sub chkSampeleType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Me.chkSampeleType
            If .Value = 1 Then
                .Value = 0
            Else
                .Value = 1
            End If
        End With
    End If
End Sub

Private Sub chkSample_Click()
    Dim intLoop As Integer
    
    If mintType = 1 Then Exit Sub
    mintType = 2
    With Me.lstSample
        For intLoop = 0 To .ListCount - 1
            .Selected(intLoop) = Me.chkSample.Value
        Next
    End With
    mintType = 0
End Sub

Private Sub setChkType(ByVal objChk As CheckBox, ByVal objList As ListBox)
    Dim intLoop As Integer
    Dim intType As Integer
    
'    If Me.Visible = False Then Exit Sub
    With objList
        If .SelCount >= .ListCount Then
            objChk.Value = 1
        ElseIf .SelCount < .ListCount And .SelCount <> 0 Then
            objChk.Value = 2
        ElseIf .SelCount = 0 Then
            objChk.Value = 0
        End If
        
    End With
End Sub
    

Private Sub chkSample_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Me.chkSample
            If .Value = 1 Then
                .Value = 0
            Else
                .Value = 1
            End If
        End With
    End If
End Sub

Private Sub cmdOK_Click()
    Dim dateSpace As Integer
    Dim strFilter As String                             '过滤条件字串
    Dim i As Integer, str类型 As String, strSample As String, strCapture As String
    
    dateSpace = DateDiff("d", Me.DTPBegin.Value, Me.DTPEND.Value)
    
    If dateSpace > 30 Then
        If MsgBox("你选择的时间大于30天，可能导致查很慢！是否继续？", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    
    strFilter = Me.TxtID & ";" & TxtSickCard & ";" & TxtName & ";" & TxtNo
    '标本
    For i = 0 To lstSample.ListCount - 1
        If lstSample.Selected(i) Then
            strSample = strSample & "," & lstSample.List(i)
        End If
    Next
    strFilter = strFilter & ";" & strSample
    '采集方式
    For i = 0 To lstCapture.ListCount - 1
        If lstCapture.Selected(i) Then
            strCapture = strCapture & "," & lstCapture.ItemData(i)
        End If
    Next
    strFilter = strFilter & ";" & strCapture
    
     strFilter = strFilter & ";" & IIf(chkOutPatient, 1, "") & ";" & _
                IIf(chkInpatient, 2, "") & ";" & IIf(chkPhysical, 4, "") & ";" & _
                dateSpace & ";0;" & _
                IIf(mDateOldEnd <> DTPEND.Value, DTPBegin.Value, "") & ";" & _
                IIf(mDateOldEnd <> DTPEND.Value, DTPEND.Value, "")
                
    '检验类型
    For i = 0 To lst类型.ListCount - 1
        If lst类型.Selected(i) Then
            str类型 = str类型 & "," & lst类型.List(i)
        End If
    Next
    strFilter = strFilter & ";" & str类型
    
    zlDatabase.SetPara "采集工作站过滤", strFilter, 100, 1211
    '传出供主窗体调用
    mstrFilter = strFilter
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mstrFilter = ""
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
    Dim strSQL As String
    Dim intLoop As Integer                          '循环变量
    Dim strTmp As String                            '临时字串变量
    Dim varFilter As Variant                        '过滤字串分解
    Dim NowDate As Date                             '当前时间
    Dim strBloodType As String                      '输血检验类型
    Dim strOldType As String                        '老版LIS标本类型
    Dim objLisInsideComm As Object                  '新版LIS接口部件
    
    On Error GoTo errH
    
    strTmp = zlDatabase.GetPara("采集工作站过滤", 100, 1211, "")
    strBloodType = zlDatabase.GetPara(273, 100)
    
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
'    strSQL = "Select Distinct A.ID,A.编码,A.名称,B.服务对象" & _
'        " From 部门表 A,部门性质说明 B" & _
'        " Where A.ID=B.部门ID And B.工作性质 IN('临床','手术')" & _
'        " And B.服务对象 IN(3,[1],[2])" & _
'        " And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
'        " Order by A.编码"
'
'
'    Set rsTmp =zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(chkOutPatient.Value = 1 Or chkPhysical.Value = 1, 1, -1), IIf(chkInpatient.Value = 1, 2, -1))
'
'    cboPatientDept.Clear
'    cboPatientDept.AddItem "所有科室"
'    cboPatientDept.ItemData(cboPatientDept.NewIndex) = 0
'    cboPatientDept.ListIndex = 0
'    Do Until rsTmp.EOF
'        cboPatientDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
'        cboPatientDept.ItemData(cboPatientDept.NewIndex) = rsTmp!ID
'        If strTmp <> "" Then
'            If rsTmp!ID = CLng(varFilter(mFilter.病人科室)) Then
'                cboPatientDept.ListIndex = cboPatientDept.NewIndex
'            End If
'        End If
'        rsTmp.MoveNext
'    Loop
    
    '===读入检验标本
    strSQL = "select 编码,名称 from 诊疗检验标本 order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    With Me.lstSample
        .Clear
        Do While Not rsTmp.EOF
            .AddItem Trim(rsTmp("名称") & "")
            strOldType = strOldType & "," & Trim(rsTmp("名称") & "")
            If strTmp <> "" Then
            If UBound(varFilter) >= mFilter.标本 Then
                If Trim("" & rsTmp!名称) <> "" And varFilter(mFilter.标本) <> "" Then
                    If InStr("," & varFilter(mFilter.标本) & ",", "," & Trim("" & rsTmp!名称) & ",") > 0 Then
                        lstSample.Selected(lstSample.NewIndex) = True
                    End If
                ElseIf varFilter(mFilter.标本) = "" Then
'                    lstSample.Selected(lstSample.NewIndex) = True
                End If
            Else
                lstSample.Selected(lstSample.NewIndex) = True
            End If
        End If
            rsTmp.MoveNext
        Loop
    End With
    '获取新版LIS中的标本类型。由于使用申请单新开医嘱时，选择的标本类型为新版LIS中的标本类型，
    '而采集工作站过滤使用的标本类型为老版LIS中的标本类型，
    '而两者之间没有直接的关联，所以，在采集工作站过滤功能的标本类型中添加上新版的标本类型
    '创建LIS接口
    If objLisInsideComm Is Nothing Then
        Dim strErr As String
        Set objLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        '初始化LIS接口部件
        If Not objLisInsideComm Is Nothing Then
            If objLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "初始化LIS接口失败！" & vbCrLf & strErr
                End If
                Set objLisInsideComm = Nothing
            End If
        End If
    End If
    If Not objLisInsideComm Is Nothing Then
        Set rsTmp = objLisInsideComm.GetSampleTypeNew()   '获取新版LIS中的标本类型
        With Me.lstSample
            Do While Not rsTmp.EOF
                If InStr(strOldType & ",", "," & Trim(rsTmp("名称") & "") & ",") <= 0 Then
                    .AddItem Trim(rsTmp("名称") & "")
                    If strTmp <> "" Then
                        If UBound(varFilter) >= mFilter.标本 Then
                            If Trim("" & rsTmp!名称) <> "" And varFilter(mFilter.标本) <> "" Then
                                If InStr("," & varFilter(mFilter.标本) & ",", "," & Trim("" & rsTmp!名称) & ",") > 0 Then
                                    lstSample.Selected(lstSample.NewIndex) = True
                                End If
                            ElseIf varFilter(mFilter.标本) = "" Then
            '                    lstSample.Selected(lstSample.NewIndex) = True
                            End If
                        Else
                            lstSample.Selected(lstSample.NewIndex) = True
                        End If
                    End If
                End If
                rsTmp.MoveNext
            Loop
        End With
    End If
    
    
    '===读入采集方式(加入输血采集)
    strSQL = "select ID,名称 from 诊疗项目目录 where 类别='E' and 操作类型 in ('6','9')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
    With Me.lstCapture
        .Clear
        Do While Not rsTmp.EOF
            .AddItem Trim(rsTmp("名称") & "")
            .ItemData(.NewIndex) = Val(rsTmp("ID") & "")
            If strTmp <> "" Then
            If UBound(varFilter) >= mFilter.采集方式 Then
                If Val("" & rsTmp!ID) <> 0 And varFilter(mFilter.采集方式) <> "" Then
                    If InStr("," & varFilter(mFilter.采集方式) & ",", "," & Val("" & rsTmp!ID) & ",") > 0 Then
                        lstCapture.Selected(lstCapture.NewIndex) = True
                    End If
                ElseIf varFilter(mFilter.采集方式) = "" Then
'                    lstCapture.Selected(lstCapture.NewIndex) = True
                End If
            Else
                lstCapture.Selected(lstCapture.NewIndex) = True
            End If
        End If
            rsTmp.MoveNext
        Loop
    End With
       
    
    NowDate = zlDatabase.Currentdate
    
    Me.DTPBegin.Value = NowDate - 3
    Me.DTPEND.Value = NowDate
    
    '读入时间
    If strTmp <> "" Then
        Me.DTPBegin.Value = NowDate - varFilter(mFilter.间隔时间)
        Me.DTPEND.Value = NowDate
    Else
        Me.DTPBegin.Value = NowDate - 3
        Me.DTPEND.Value = NowDate
    End If
    
    '检验类型
    strSQL = "select  distinct  操作类型  from 诊疗项目目录 Where 类别='C' and 操作类型 is not null  "
    If strBloodType <> "" Then
        strSQL = strSQL & " UNION " & vbNewLine & "Select '" & strBloodType & "' 操作类型 from dual"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    lst类型.Clear
    Do Until rsTmp.EOF
        lst类型.AddItem Trim("" & rsTmp!操作类型)
        If strTmp <> "" Then
            If UBound(varFilter) >= mFilter.检验类型 Then
                If Trim("" & rsTmp!操作类型) <> "" And varFilter(mFilter.检验类型) <> "" Then
                    If InStr("," & varFilter(mFilter.检验类型) & ",", "," & Trim("" & rsTmp!操作类型) & ",") > 0 Then
                        lst类型.Selected(lst类型.NewIndex) = True
                    End If
                ElseIf varFilter(mFilter.检验类型) = "" Then
'                    lst类型.Selected(lst类型.NewIndex) = True
                End If
            Else
                lst类型.Selected(lst类型.NewIndex) = True
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    mintType = 1
    Call setChkType(Me.chkCapture, Me.lstCapture)
    Call setChkType(Me.chkSample, Me.lstSample)
    Call setChkType(Me.chkSampeleType, Me.lst类型)
    mintType = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowME(Objfrm As Object, ByRef strFilter As String)
    Me.Show vbModal, Objfrm
    strFilter = mstrFilter
End Sub

Private Sub lstCapture_Click()
    If mintType = 2 Then Exit Sub
    mintType = 1
    Call setChkType(Me.chkCapture, Me.lstCapture)
    mintType = 0
End Sub

Private Sub lstSample_Click()
    If mintType = 2 Then Exit Sub
    mintType = 1
    Call setChkType(Me.chkSample, Me.lstSample)
    mintType = 0
End Sub

Private Sub lst类型_Click()
    If mintType = 2 Then Exit Sub
    mintType = 1
    Call setChkType(Me.chkSampeleType, Me.lst类型)
    mintType = 0
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtSickCard.SetFocus
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtNo.SetFocus
End Sub

Private Sub TxtNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkOutPatient.SetFocus
End Sub

Private Sub TxtSickCard_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtName.SetFocus
End Sub
