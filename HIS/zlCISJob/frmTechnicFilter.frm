VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frmTechnicFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤条件"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ControlBox      =   0   'False
   Icon            =   "frmTechnicFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   23
      Top             =   3330
      Width           =   1100
   End
   Begin VB.ComboBox cboDoctor 
      Height          =   300
      Left            =   4035
      TabIndex        =   11
      Text            =   "cboDoctor"
      Top             =   2715
      Width           =   1710
   End
   Begin VB.CheckBox chk本次住院 
      Caption         =   "只显示本次住院的项目"
      Height          =   195
      Left            =   3660
      TabIndex        =   5
      Top             =   2100
      Value           =   1  'Checked
      Width           =   2100
   End
   Begin VB.CheckBox chk来源 
      Caption         =   "体检"
      Height          =   195
      Index           =   2
      Left            =   2820
      TabIndex        =   8
      Top             =   2400
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.CheckBox chk期效 
      Caption         =   "长期"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   3
      Top             =   2100
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.CheckBox chk期效 
      Caption         =   "临时"
      Height          =   195
      Index           =   1
      Left            =   1965
      TabIndex        =   4
      Top             =   2100
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   2
      Left            =   0
      TabIndex        =   20
      Top             =   1460
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   18
      Top             =   720
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   -105
      TabIndex        =   17
      Top             =   3180
      Width           =   6360
   End
   Begin VB.CommandButton cmdDefault 
      Cancel          =   -1  'True
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   330
      TabIndex        =   16
      Top             =   3330
      Width           =   1100
   End
   Begin VB.CheckBox chk来源 
      Caption         =   "住院"
      Height          =   195
      Index           =   1
      Left            =   1965
      TabIndex        =   7
      Top             =   2400
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.CheckBox chk来源 
      Caption         =   "门诊"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   6
      Top             =   2400
      Value           =   1  'Checked
      Width           =   660
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2715
      Width           =   2115
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   3915
      TabIndex        =   2
      Top             =   1650
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   169607171
      CurrentDate     =   38082
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1140
      TabIndex        =   1
      Top             =   1650
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   169607171
      CurrentDate     =   38082
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3420
      TabIndex        =   12
      Top             =   3330
      Width           =   1100
   End
   Begin zlIDKind.PatiIdentify PatiIdentify 
      Height          =   270
      Left            =   1140
      TabIndex        =   24
      Top             =   975
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmTechnicFilter.frx":000C
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      IDKindAppearance=   0
      ShowPropertySet =   -1  'True
      DefaultCardType =   "就诊卡"
      IDKindWidth     =   555
      FindPatiShowName=   0   'False
      HiddenMoseRightKey=   0   'False
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开单人"
      Height          =   180
      Left            =   3375
      TabIndex        =   10
      Top             =   2775
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      Height          =   180
      Left            =   3330
      TabIndex        =   22
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label lbl期效 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医嘱期效"
      Height          =   180
      Left            =   270
      TabIndex        =   21
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人"
      Height          =   180
      Left            =   630
      TabIndex        =   0
      Top             =   1020
      Width           =   360
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   270
      Picture         =   "frmTechnicFilter.frx":00D3
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    设置过滤条件以便准确查找执行记录；建议时间范围尽量精确，以提高查找速度。"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   915
      TabIndex        =   19
      Top             =   180
      Width           =   3780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人来源"
      Height          =   180
      Left            =   270
      TabIndex        =   15
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人科室"
      Height          =   180
      Left            =   270
      TabIndex        =   14
      Top             =   2775
      Width           =   720
   End
   Begin VB.Label lbl查询时间 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行时间"
      Height          =   180
      Left            =   270
      TabIndex        =   13
      Top             =   1710
      Width           =   720
   End
End
Attribute VB_Name = "frmTechnicFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Public mblnOK As Boolean
Public mstrDeptNode As String   '当前医技科室所属的站点

Public mobjSquareCard As Object      '卡结算对象
Public mstrCardKind As String        '卡结算对象返回的可用的医疗卡
Public mstrFindType As String        '存储当前查找类型名称
Public mlngPatiID As Long

Private mblnLoad As Boolean
Private mstrDeptNodePre As String

Private Sub cboDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboDoctor.Tag = "keypress"
        If SeekDoctor = False Then
            cboDoctor.Tag = ""
            cboDoctor.SetFocus
        End If
    End If
End Sub

Private Sub cboDoctor_Validate(Cancel As Boolean)
    If cboDoctor.Tag = "keypress" Then
        cboDoctor.Tag = ""
    ElseIf cboDoctor.ListIndex = -1 And cboDoctor.Text <> "" Then
        If SeekDoctor = False Then
            cboDoctor.Text = ""
        End If
    End If
End Sub

Private Function SeekDoctor() As Boolean
'功能：根据当前输入内容查找医生列表

    Dim strTxt As String, blnYes As Boolean
    Dim i As Long, bytKind As Byte
    
    strTxt = UCase(Trim(cboDoctor.Text))
    If strTxt = "所有医生" Then
        cboDoctor.ListIndex = 0
        SeekDoctor = True
        Exit Function
    End If
    
    If zlCommFun.IsCharAlpha(strTxt) Then
        bytKind = 0
    ElseIf InStr(strTxt, "-") > 0 Then
        bytKind = 1
    Else
        bytKind = 2
    End If
    
    'i=0是“所有医生”
    For i = 1 To cboDoctor.ListCount - 1
            If bytKind = 0 Then
            If cboDoctor.List(i) Like "*/" & strTxt & "-*" Or cboDoctor.List(i) Like strTxt & "/*" Then
                blnYes = True
            End If
        ElseIf bytKind = 2 Then
            If cboDoctor.List(i) Like "*-" & strTxt Then
                blnYes = True
            End If
        Else
            If cboDoctor.List(i) = strTxt Then
                blnYes = True
            End If
        End If
        If blnYes Then
            cboDoctor.ListIndex = i
            SeekDoctor = True
            Exit Function
        End If
    Next
    If cboDoctor.ListCount > 0 Then
        cboDoctor.ListIndex = 0
        SeekDoctor = True
    End If
End Function

Private Sub chk来源_Click(Index As Integer)
    If chk来源(0).Value = 0 And chk来源(1).Value = 0 And chk来源(2).Value = 0 Then
        chk来源((Index + 1) Mod 3).Value = 1
    End If
    
    chk本次住院.Enabled = chk来源(1).Value = 1
    
    If Me.Visible Then
        Call LoadDeptList
        Call LoadDoctorList
    End If
End Sub

Private Sub chk期效_Click(Index As Integer)
    If chk期效(0).Value = 0 And chk期效(1).Value = 0 Then
        chk期效((Index + 1) Mod 2).Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Me.Hide
End Sub

Private Sub cmdDefault_Click()
    Call Form_Load
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    
    '保存参数
    Call zlDatabase.SetPara("病人来源", chk来源(0).Value & chk来源(1).Value & chk来源(2).Value, glngSys, p医技工作站, InStr(mstrPrivs, "参数设置") > 0)
    Call zlDatabase.SetPara("医嘱期效", chk期效(0).Value & chk期效(1).Value, glngSys, p医技工作站, InStr(mstrPrivs, "参数设置") > 0)
    Call zlDatabase.SetPara("只显示本次住院项目", chk本次住院.Value, glngSys, p医技工作站, InStr(mstrPrivs, "参数设置") > 0)
    With cboDoctor
        If .ListIndex = 0 Or .ListIndex = -1 Then
            strTmp = ""
        Else
            strTmp = Split(.Text, "-")(1)
        End If
        Call zlDatabase.SetPara("开单人", strTmp, glngSys, p医技工作站, InStr(mstrPrivs, "参数设置") > 0)
    End With
        
    mblnOK = True
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim curDate As Date
    
    '如果上一次是取的当前时间,则重新设置时刷新结果时间为当前时间
    If Not mblnLoad Then
        If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
            curDate = zlDatabase.Currentdate
            dtpBegin.MaxDate = curDate + 7
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59")
            dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
        End If
    End If
    If mblnLoad Then mblnLoad = False
    
    If mstrDeptNodePre <> mstrDeptNode Then
        mstrDeptNodePre = mstrDeptNode
        
        Call LoadDeptList
        Call LoadDoctorList
    End If
    
    '自动定位
    dtpBegin.SetFocus
    If PatiIdentify.Text <> "" Then
        mlngPatiID = 0
        PatiIdentify.Text = "": PatiIdentify.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim strKey As String

    mblnLoad = True
    
    mstrDeptNodePre = ""
    If Not mobjSquareCard Is Nothing Then
        Call PatiIdentify.zlInit(Me, glngSys, p医技工作站, gcnOracle, gstrDBUser, mobjSquareCard, mstrCardKind, "zl9CISJob")
        PatiIdentify.objIDKind.AllowAutoICCard = True
        PatiIdentify.objIDKind.AllowAutoIDCard = True
        PatiIdentify.Text = ""
    End If
    '病人过滤方式设置
    lbl查询时间.Caption = IIf(Val(zlDatabase.GetPara("病人过滤方式", glngSys, p医技工作站)) = 1, "发送时间", "执行时间")

    '本次住院
    chk本次住院.Value = Val(zlDatabase.GetPara("只显示本次住院项目", glngSys, p医技工作站, "1", Array(chk本次住院), InStr(mstrPrivs, "参数设置") > 0))
    
    '来源
    strKey = zlDatabase.GetPara("病人来源", glngSys, p医技工作站, "111", Array(chk来源(0), chk来源(1), chk来源(2)), InStr(mstrPrivs, "参数设置") > 0)
    chk来源(0).Value = Val(Mid(strKey, 1, 1))
    chk来源(1).Value = Val(Mid(strKey, 2, 1))
    chk来源(2).Value = Val(Mid(strKey, 3, 1))
    
    '期效
    strKey = zlDatabase.GetPara("医嘱期效", glngSys, p医技工作站, "11", Array(chk期效(0), chk期效(1)), InStr(mstrPrivs, "参数设置") > 0)
    chk期效(0).Value = Val(Mid(strKey, 1, 1))
    chk期效(1).Value = Val(Mid(strKey, 2, 1))
    
    '发送时间
    curDate = zlDatabase.Currentdate
    dtpBegin.MaxDate = curDate + 7
    dtpBegin.Value = Format(curDate - 1, "yyyy-MM-dd 00:00")
    dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59")
    dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
            
    Call LoadDeptList
    Call LoadDoctorList
    
    strKey = zlDatabase.GetPara("开单人", glngSys, p医技工作站, "", , InStr(mstrPrivs, "参数设置") > 0)
    Call Cbo.Locate(cboDoctor, IIf(strKey = "ALL", "所有医生", strKey))
    mblnOK = False
End Sub

Private Sub LoadDeptList()
'功能：根据病人来源读取病人科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngPre As Long
    
    If cboDept.ListIndex <> -1 Then
        lngPre = cboDept.ItemData(cboDept.ListIndex)
    End If
    strSQL = "Select Distinct A.ID,A.编码,A.名称,B.服务对象" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.工作性质 IN('临床','手术')" & _
        " And B.服务对象 IN(3,[1],[2])" & _
        IIf(mstrDeptNode <> "", " And (A.站点 = [3] Or A.站点 is Null)", "") & _
        " And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(chk来源(0).Value = 1 Or chk来源(2).Value = 1, 1, -1), IIf(chk来源(1).Value = 1, 2, -1), mstrDeptNode)
    On Error GoTo 0
    cboDept.Clear
    cboDept.AddItem "所有科室"
    cboDept.ListIndex = 0
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPre Then cboDept.ListIndex = cboDept.NewIndex
        rsTmp.MoveNext
    Next
            
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDoctorList()
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim lngPre As Long
    
    If cboDoctor.ListIndex <> -1 Then
        lngPre = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    
    cboDoctor.Clear
    cboDoctor.AddItem "所有医生"
    cboDoctor.ListIndex = 0
    
    Set rsTmp = GetDoctorRs
    For i = 1 To rsTmp.RecordCount
        cboDoctor.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cboDoctor.ItemData(cboDoctor.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPre Then cboDoctor.ListIndex = cboDoctor.NewIndex
        rsTmp.MoveNext
    Next
    
End Sub

Private Function GetDoctorRs() As ADODB.Recordset
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(3,[1],[2])"
    strSQL = "Select Distinct A.ID,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        IIf(mstrDeptNode <> "", " And (A.站点 = [3] Or A.站点 is Null)", "") & _
        " And B.部门ID IN(" & strSQL & ")" & _
        " Order by A.简码"
        
    Set GetDoctorRs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(chk来源(0).Value = 1 Or chk来源(2).Value = 1, 1, -1), IIf(chk来源(1).Value = 1, 2, -1), mstrDeptNode)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub PatiIdentify_Change()
    mlngPatiID = 0
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    mlngPatiID = 0
    If objHisPati Is Nothing Then Exit Sub
    mlngPatiID = objHisPati.病人ID
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    If mstrFindType = objCard.名称 And InStr(";就诊卡;标识号;单据号;姓名;二代身份证;IC卡;医保号;", ";" & mstrFindType & ";") > 0 Then
        mlngPatiID = 0
        blnCancel = True
    End If
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    mstrFindType = objCard.名称
    PatiIdentify.Text = ""
    mlngPatiID = 0
    PatiIdentify.PasswordChar = IIf(gblnCardHide And mstrFindType = "就诊卡", "*", "")
End Sub

Private Sub PatiIdentify_Validate(Cancel As Boolean)
    If mstrFindType = "单据号" And IsNumeric(Trim(PatiIdentify.Text)) And mlngPatiID = 0 Then
        PatiIdentify.Text = GetFullNO(Trim(PatiIdentify.Text), 14): PatiIdentify.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjSquareCard = Nothing
End Sub