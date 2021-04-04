VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanInfoEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "编辑"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5580
   Icon            =   "frmClinicPlanInfoEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picTemplet 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   90
      ScaleHeight     =   1755
      ScaleWidth      =   5295
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   5295
      Begin VB.Frame fraTempletType 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   750
         TabIndex        =   21
         Top             =   60
         Width           =   1875
         Begin VB.OptionButton optTempletType 
            Caption         =   "周排班"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optTempletType 
            Caption         =   "月排班"
            Height          =   180
            Index           =   1
            Left            =   930
            TabIndex        =   22
            Top             =   0
            Width           =   885
         End
      End
      Begin VB.CheckBox chkTempletByDay 
         Caption         =   "按天安排出诊"
         Height          =   180
         Left            =   2640
         TabIndex        =   4
         Top             =   60
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txt备注 
         Height          =   1050
         Left            =   750
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   4485
      End
      Begin VB.ComboBox cbo科室 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3660
         TabIndex        =   9
         Text            =   "cbo科室"
         Top             =   330
         Width           =   1665
      End
      Begin VB.OptionButton opt应用范围 
         Caption         =   "所属科室"
         Height          =   180
         Index           =   2
         Left            =   2610
         TabIndex        =   8
         Top             =   390
         Width           =   1065
      End
      Begin VB.OptionButton opt应用范围 
         Caption         =   "仅本人"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   390
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton opt应用范围 
         Caption         =   "全院"
         Height          =   180
         Index           =   0
         Left            =   750
         TabIndex        =   6
         Top             =   390
         Width           =   705
      End
      Begin VB.Label lblTempletType 
         AutoSize        =   -1  'True
         Caption         =   "模板类型"
         Height          =   180
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lbl备注 
         AutoSize        =   -1  'True
         Caption         =   "备注"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   690
         Width           =   360
      End
      Begin VB.Label lbl应用范围 
         AutoSize        =   -1  'True
         Caption         =   "应用范围"
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.Frame fraSplitY 
      Height          =   25
      Left            =   -30
      TabIndex        =   20
      Top             =   2340
      Width           =   5730
   End
   Begin VB.Frame fraSplitX 
      Height          =   1875
      Left            =   3090
      TabIndex        =   19
      Top             =   -120
      Width           =   25
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   870
      MaxLength       =   50
      TabIndex        =   1
      Top             =   180
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   330
      Left            =   4200
      TabIndex        =   18
      Top             =   2580
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   330
      Left            =   2970
      TabIndex        =   17
      Top             =   2580
      Width           =   915
   End
   Begin VB.PictureBox picFixedRule 
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   90
      ScaleHeight     =   1065
      ScaleWidth      =   2985
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   540
      Width           =   2985
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   300
         Left            =   780
         TabIndex        =   16
         Top             =   630
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483630
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy/MM/dd HH:mm:ss"
         Format          =   171180035
         CurrentDate     =   42340
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   300
         Left            =   780
         TabIndex        =   14
         Top             =   150
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483630
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy/MM/dd HH:mm:ss"
         Format          =   171180035
         CurrentDate     =   42340
      End
      Begin VB.Label lblEndTime 
         AutoSize        =   -1  'True
         Caption         =   "终止时间"
         Height          =   180
         Left            =   30
         TabIndex        =   15
         Top             =   690
         Width           =   720
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         Caption         =   "开始时间"
         Height          =   180
         Left            =   30
         TabIndex        =   13
         Top             =   210
         Width           =   720
      End
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "模板名称"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmClinicPlanInfoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As Byte '1-模板，2-固定安排
Private mlngModule As Long
Private mblnOK As Boolean
Private mobj出诊安排 As 出诊安排
Private mrsDepts As ADODB.Recordset '人员所属部门
Private mblnSaveAsTemplet As Boolean '是否出诊表保存为模板，若为True，则不允许修改模板类型
Private mblnUpdate As Boolean

Public Function ShowMe(frmParent As Form, ByVal lngModule As Long, ByVal bytFun As Byte, _
    Optional ByRef obj出诊安排 As 出诊安排, Optional ByVal blnUpdate As Boolean, _
    Optional ByVal blnSaveAsTemplet As Boolean)
    '程序入口
    mbytFun = bytFun: Set mobj出诊安排 = obj出诊安排
    mlngModule = lngModule
    If mobj出诊安排 Is Nothing Then Set mobj出诊安排 = New 出诊安排
    mblnSaveAsTemplet = blnSaveAsTemplet: mblnUpdate = blnUpdate
    
    On Error Resume Next
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Err = 0: On Error GoTo errHandle
    If KeyAscii = 13 Then
        If cbo科室.Text = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If cbo科室.ListIndex >= 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If Select科室(Me, mlngModule, mrsDepts, cbo科室, cbo科室.Text) = True Then
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        If cbo科室.Enabled Then cbo科室.SetFocus
        zlControl.TxtSelAll cbo科室
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, blnDoCheck As Boolean
    Dim rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Sub
    If zlControl.TxtCheckInput(txtName, lblName.Caption, 50) = False Then Exit Sub
    If mbytFun = 1 Then
        If zlControl.TxtCheckInput(txt备注, lbl备注.Caption, 100, True) = False Then Exit Sub
        If opt应用范围(2).Value And cbo科室.ListIndex = -1 Then
            MsgBox "所属科室不能为空！", vbInformation, gstrSysName
            If cbo科室.Visible And cbo科室.Enabled Then cbo科室.SetFocus
            Exit Sub
        End If
    ElseIf dtpStartTime.Enabled Then
        If dtpEndTime.Value <= dtpStartTime.Value Then
            MsgBox "终止时间必须大于开始时间！", vbInformation, gstrSysName
            If dtpEndTime.Visible And dtpEndTime.Enabled Then dtpEndTime.SetFocus
            Exit Sub
        End If
        If dtpStartTime.Value < Now Then
            MsgBox "开始时间不能小于当前时间！", vbInformation, gstrSysName
            If dtpStartTime.Visible And dtpStartTime.Enabled Then dtpStartTime.SetFocus
            Exit Sub
        End If
        
        If mobj出诊安排.开始时间 = "" Then
            blnDoCheck = True
        Else
            If DateDiff("s", mobj出诊安排.开始时间, dtpStartTime.Value) <> 0 Then blnDoCheck = True
        End If
        If blnDoCheck Then
'            strSQL = "Select Max(a.开始时间) As 开始时间" & vbNewLine & _
'                    " From 临床出诊安排 A, 临床出诊表 B" & vbNewLine & _
'                    " Where a.出诊id = b.Id And b.排班方式 = 0 And a.开始时间 > [1] And b.发布时间 Is Not Null"
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊表信息", dtpStartTime.Value)
'            If Not rsTemp.EOF Then
'                If Nvl(rsTemp!开始时间) <> "" Then
'                    MsgBox "当前开始时间不能小于上一个已发布的固定安排的开始时间(" & Nvl(rsTemp!开始时间, "yyyy-mm-dd hh:mm:ss") & ")！", vbInformation, gstrSysName
'                    Exit Sub
'                End If
'            End If
        End If
    End If
    
    If mobj出诊安排.出诊表名 <> Trim(txtName.Text) Then
        strSQL = "Select 1 From 临床出诊表 Where 出诊表名 = [1] And 排班方式 = [2] And Nvl(站点,'-') = Nvl([3],'-') And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊表信息", Trim(txtName.Text), IIf(mbytFun = 1, 3, 0), gstrNodeNo)
        If Not rsTemp.EOF Then
            MsgBox "当前已存在名为“" & Trim(txtName.Text) & "”的" & IIf(mbytFun = 1, "模板！", "出诊表！"), vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    With mobj出诊安排
        .出诊表名 = Trim(txtName.Text)
        If mbytFun = 1 Then
            '模板类型：0-周排班模板，1-不是按天排班的月排班模板，2-按天排班的月排班模板
            .模板类型 = IIf(optTempletType(1).Value, IIf(chkTempletByDay.Value = vbChecked, 2, 1), 0)
            '应用范围：0-本人;1-人员所属科室(指定科室);2-全院通用
            .应用范围 = Choose(GetSelectedIndex(opt应用范围) + 1, 2, 0, 1)
            If .应用范围 = 1 Then  '所属科室
                .科室ID = cbo科室.ItemData(cbo科室.ListIndex)
                .科室名称 = cbo科室.Text
            End If
            .备注 = Trim(txt备注.Text)
        Else
            mobj出诊安排.开始时间 = Format(dtpStartTime.Value, "yyyy-mm-dd hh:mm:ss")
            mobj出诊安排.终止时间 = Format(dtpEndTime.Value, "yyyy-mm-dd hh:mm:ss")
        End If
    End With
    mblnOK = True: Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dtpEndTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStartTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStartTime_Validate(Cancel As Boolean)
    If dtpEndTime.Value < dtpStartTime.Value Then
        dtpEndTime.Value = dtpStartTime.Value
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim int应用范围 As Integer, bln系统导入 As Boolean
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = 1 Then
        strSQL = "Select b.ID,b.编码,b.简码,b.名称" & vbNewLine & _
                " From 部门人员 A, 部门表 B" & vbNewLine & _
                " Where a.部门ID=b.ID And a.人员ID=[1]" & vbNewLine & _
                "       And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null)" & vbNewLine & _
                "       And (b.站点='" & gstrNodeNo & "' Or b.站点 is Null)"
        Set mrsDepts = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.id)
        cbo科室.Clear
        If Not mrsDepts Is Nothing Then
            Do While Not mrsDepts.EOF
                cbo科室.AddItem Nvl(mrsDepts!名称)
                cbo科室.ItemData(cbo科室.NewIndex) = Nvl(mrsDepts!id)
                mrsDepts.MoveNext
            Loop
        End If
    End If
    
    With mobj出诊安排
        txtName.Text = .出诊表名
        If mbytFun = 1 Then
            '模板类型：0-周排班模板，1-不是按天排班的月排班模板，2-按天排班的月排班模板
            If .模板类型 = 0 Then '周排班
                optTempletType(0).Value = True
            Else '月排班
                optTempletType(1).Value = True
                chkTempletByDay.Value = IIf(.模板类型 = 2, vbChecked, vbUnchecked)
            End If
            int应用范围 = .应用范围 '0-本人;1-所属科室;2-全院通用
            If int应用范围 = 1 Then
                opt应用范围(2).Value = True
                zlControl.CboLocate cbo科室, .科室ID, True
            Else
                opt应用范围(IIf(int应用范围 = 0, 1, 0)).Value = True
            End If
            txt备注.Text = .备注
            
            If mblnSaveAsTemplet Or mblnUpdate Then
                optTempletType(0).Enabled = False
                optTempletType(1).Enabled = False
                chkTempletByDay.Enabled = False
            Else
                optTempletType(0).Enabled = True
                optTempletType(1).Enabled = True
                chkTempletByDay.Enabled = True
            End If
        Else
            bln系统导入 = (.排班方式 = 0 And .备注 = "系统导入")
            
            dtpStartTime.Value = Format(IIf(mobj出诊安排.开始时间 = "", Format(Now + 1, "yyyy-mm-dd 00:00:00"), mobj出诊安排.开始时间), "yyyy-mm-dd hh:mm:ss")
            dtpEndTime.Value = Format(IIf(mobj出诊安排.终止时间 = "", "3000-01-01 00:00:00", mobj出诊安排.终止时间), "yyyy-mm-dd hh:mm:ss")
        End If
    End With
    
    picTemplet.Visible = False
    picFixedRule.Visible = False
    fraSplitX.Visible = False
    fraSplitY.Visible = False
    If mbytFun = 1 Then '模板
        picTemplet.Visible = True
        picTemplet.Left = 90
        picTemplet.Top = 540
        fraSplitY.Visible = True
        cmdOk.Left = 2970
        cmdOk.Top = 2450
        cmdCancel.Left = cmdOk.Left + cmdOk.Width + 300
        cmdCancel.Top = cmdOk.Top
        Me.Width = 5560
        Me.Height = 3370
        Me.Caption = IIf(mobj出诊安排.出诊ID = 0, "增加模板", "调整模板")
        lblName.Caption = "模板名称"
    Else '固定安排
        picFixedRule.Visible = True
        picFixedRule.Left = 90
        picFixedRule.Top = 540
        fraSplitX.Visible = True
        cmdOk.Left = 3240
        cmdOk.Top = txtName.Top
        cmdCancel.Left = cmdOk.Left
        cmdCancel.Top = cmdOk.Top + cmdOk.Height + 100
        Me.Width = 4360
        Me.Height = 2110
        Me.Caption = IIf(mobj出诊安排.出诊ID = 0, "增加固定安排", "调整安排")
        lblName.Caption = "安排名称"
        If bln系统导入 Then
            dtpStartTime.Enabled = False
            dtpEndTime.Enabled = False
        End If
    End If
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsDepts = Nothing
End Sub

Private Sub optTempletType_Click(index As Integer)
    chkTempletByDay.Visible = index = 1
End Sub

Private Sub opt应用范围_Click(index As Integer)
    cbo科室.Enabled = index = 2
    If index <> 2 Then
        cbo科室.ListIndex = -1
    End If
End Sub

Private Sub opt应用范围_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt备注_GotFocus()
    zlControl.TxtSelAll txt备注
End Sub

Private Sub txt备注_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
