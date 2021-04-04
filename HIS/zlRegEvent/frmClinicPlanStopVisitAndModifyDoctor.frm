VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "ZLIDKIND.OCX"
Begin VB.Form frmClinicPlanStopVisitAndModifyDoctor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "停诊"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   Icon            =   "frmClinicPlanStopVisitAndModifyDoctor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6120
      TabIndex        =   28
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6120
      TabIndex        =   27
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6120
      TabIndex        =   29
      Top             =   3330
      Width           =   1100
   End
   Begin VB.Frame fra替诊医生 
      Caption         =   "替诊医生"
      Height          =   765
      Left            =   60
      TabIndex        =   23
      Top             =   3960
      Width           =   5895
      Begin VB.ComboBox cbo替诊医生 
         Height          =   300
         Left            =   1110
         TabIndex        =   25
         Text            =   "张三"
         Top             =   300
         Width           =   4575
      End
      Begin zlIDKind.IDKindNew idkDoctor 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         IDKindStr       =   "内|院内医生|0|0|0|0|0||0|0|0;外|院外医生|0|0|0|0|0||0|0|0"
         CaptionAlignment=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         DefaultCardType =   "0"
         NotAutoAppendKind=   -1  'True
         BackColor       =   -2147483633
      End
   End
   Begin VB.Frame fra出诊信息 
      Caption         =   "出诊信息"
      Height          =   1185
      Left            =   60
      TabIndex        =   14
      Top             =   2610
      Width           =   5895
      Begin VB.TextBox txt上班时段 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   300
         Left            =   3150
         TabIndex        =   21
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   184680451
         UpDown          =   -1  'True
         CurrentDate     =   42360.3333333333
      End
      Begin VB.ComboBox cbo停诊原因 
         Height          =   300
         Left            =   3150
         TabIndex        =   18
         Text            =   "手术"
         Top             =   330
         Width           =   2535
      End
      Begin VB.TextBox txt出诊日期 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   16
         Text            =   "2016-04-05"
         Top             =   330
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   4620
         TabIndex        =   22
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   184680451
         UpDown          =   -1  'True
         CurrentDate     =   42360.5
      End
      Begin VB.Label lblTimeRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   180
         Left            =   4350
         TabIndex        =   30
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lbl停诊时间 
         AutoSize        =   -1  'True
         Caption         =   "停诊时间"
         Height          =   180
         Left            =   2400
         TabIndex        =   20
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lbl停诊原因 
         AutoSize        =   -1  'True
         Caption         =   "停诊原因"
         Height          =   180
         Left            =   2400
         TabIndex        =   17
         Top             =   390
         Width           =   720
      End
      Begin VB.Label lbl上班时段 
         AutoSize        =   -1  'True
         Caption         =   "上班时段"
         Height          =   180
         Left            =   90
         TabIndex        =   19
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lbl出诊日期 
         AutoSize        =   -1  'True
         Caption         =   "出诊日期"
         Height          =   180
         Left            =   90
         TabIndex        =   15
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.Frame fra号源信息 
      Caption         =   "号源基本信息"
      Height          =   2295
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtSignalNO 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   2
         Text            =   "4"
         Top             =   330
         Width           =   1335
      End
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   8
         Text            =   "主任医师号"
         Top             =   1110
         Width           =   4875
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   6
         Text            =   "门诊内科"
         Top             =   720
         Width           =   4875
      End
      Begin VB.TextBox txtDoctor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   10
         Text            =   "王二"
         Top             =   1500
         Width           =   4875
      End
      Begin VB.TextBox txt假日控制 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   12
         Text            =   "不上班"
         Top             =   1890
         Width           =   1965
      End
      Begin VB.CheckBox chk建档 
         Caption         =   "挂号时必须建档"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3420
         TabIndex        =   13
         Top             =   1935
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txt号类 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3420
         TabIndex        =   4
         Text            =   "普通"
         Top             =   330
         Width           =   2265
      End
      Begin VB.Label lbl假日控制 
         AutoSize        =   -1  'True
         Caption         =   "假日控制"
         Height          =   180
         Left            =   60
         TabIndex        =   11
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         Height          =   180
         Left            =   420
         TabIndex        =   9
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   180
         Left            =   420
         TabIndex        =   5
         Top             =   780
         Width           =   360
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "项目"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label lbl号类 
         AutoSize        =   -1  'True
         Caption         =   "号类"
         Height          =   180
         Left            =   3030
         TabIndex        =   3
         Top             =   390
         Width           =   360
      End
      Begin VB.Label lblSignalNO 
         AutoSize        =   -1  'True
         Caption         =   "号码"
         Height          =   180
         Left            =   420
         TabIndex        =   1
         Top             =   390
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmClinicPlanStopVisitAndModifyDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mlngModule As Long
Private Enum m_FunType
    F_停诊 = 1
    F_取消停诊 = 2
    F_替诊 = 3
    F_取消替诊 = 4
End Enum
Private mbytFun As m_FunType
Private mlng记录ID As Long '记录ID
Private mrsStopReason As ADODB.Recordset

'参数
Private mblnOnly院内医生 As Boolean '仅只能输院内医生
Private mbln替诊医生级别检查 As Boolean
Private mbyt预约清单控制方式 As Byte
Private mbyt预约清单打印方式 As Byte

Private mblnCboClick As Boolean     '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件

Public Function ShowMe(frmParent As Form, ByVal lngModule As Long, ByVal bytFun As Byte, _
    ByVal lng记录ID As Long) As Boolean
    '程序入口
    '入参：
    '   frmParent 父窗口
    '   lngModule 模块号
    '   bytFun 功能：1-停诊,2-取消停诊,3-替诊,4-取消替诊
    '   lng记录ID 出诊记录ID
    mbytFun = bytFun: mlngModule = lngModule
    mlng记录ID = lng记录ID
    
    On Error Resume Next
    If EditBeforCheck(bytFun, lng记录ID) = False Then Exit Function
    mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Function EditBeforCheck(ByVal bytFun As m_FunType, ByVal lng记录ID As Long) As Boolean
    '对出诊安排进行限制检查
    Dim strSQL, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    '不能对历史的安排进行操作
    strSQL = "Select 1 From 临床出诊记录 A Where ID = [1] And a.终止时间 < Sysdate"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "不能对历史的安排进行操作", lng记录ID)
    If Not rsTemp.EOF Then
        MsgBox "不能对历史的安排进行操作！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If bytFun = F_替诊 Then
        '针对未设置医生的号源，不允许替诊操作
        strSQL = "Select 1 From 临床出诊记录 A Where ID = [1] And a.医生姓名 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "未设置医生的号源不允许替诊", lng记录ID)
        If rsTemp.EOF Then
            MsgBox "该号源未设置医生，不允许替诊操作！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    EditBeforCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo替诊医生_Click()
    Dim strSQL As String, blnCancel As Boolean
    Dim rsReturn As ADODB.Recordset
    Dim vRect  As RECT
    
    Err = 0: On Error GoTo errHandler
    If cbo替诊医生.Text = "其他科室医生..." Then
        '选择"其他科室医生..."时，弹出选择器
        cbo替诊医生.ListIndex = -1
        Call GetDoctor(Val(txtDept.Tag), "", True, True, strSQL)  '获取SQL语句
        vRect = zlControl.GetControlRect(cbo替诊医生.Hwnd)
        
        Set rsReturn = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "替诊医生", False, _
                       "", "", False, False, False, vRect.Left, vRect.Top, cbo替诊医生.Height, blnCancel, True, False)
        If blnCancel Then Exit Sub
        If rsReturn Is Nothing Then Exit Sub
        If rsReturn.EOF Then Exit Sub
        
        With rsReturn
            zlControl.CboLocate cbo替诊医生, Nvl(!ID), True
            If cbo替诊医生.ListIndex = -1 Then
                cbo替诊医生.AddItem Nvl(!姓名) & IIf(Nvl(!专业技术职务) = "", "", "(" & Nvl(!专业技术职务) & ")"), cbo替诊医生.ListCount - 1
                cbo替诊医生.ItemData(cbo替诊医生.NewIndex) = Val(Nvl(!ID))
                cbo替诊医生.ListIndex = cbo替诊医生.NewIndex
            End If
        End With
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo替诊医生_GotFocus()
    zlControl.TxtSelAll cbo替诊医生
End Sub

Private Sub cbo替诊医生_KeyPress(KeyAscii As Integer)
    Dim strSQL As String, blnCancel As Boolean
    Dim rsReturn As ADODB.Recordset
    Dim vRect  As RECT
    Dim strKey As String, strWhere As String
    
    Err = 0: On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    If cbo替诊医生.ListIndex <> -1 Or mblnOnly院内医生 = False Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If Trim(cbo替诊医生.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    '模糊匹配,选择医生
    strKey = gstrLike & Trim(cbo替诊医生.Text) & "%"
    If zlCommFun.IsCharChinese(Trim(cbo替诊医生.Text)) Then
         strWhere = " And 姓名 like [1] "
    ElseIf zlCommFun.IsNumOrChar(Trim(cbo替诊医生.Text)) Then
         strWhere = " And (简码 like upper([1]) or 编号 like upper([1]))"
    End If
        
    Call GetDoctor(0, strWhere, False, True, strSQL) '获取SQL语句
    vRect = zlControl.GetControlRect(cbo替诊医生.Hwnd)
    Set rsReturn = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "替诊医生", False, _
                   "", "", False, False, True, vRect.Left, vRect.Top, cbo替诊医生.Height, blnCancel, True, False, strKey)
    If blnCancel Then Exit Sub
    If rsReturn Is Nothing Then Exit Sub
    If rsReturn.EOF Then Exit Sub
    
    zlControl.CboLocate cbo替诊医生, Nvl(rsReturn!ID), True
    If cbo替诊医生.ListIndex = -1 And Nvl(rsReturn!姓名) <> "其他科室医生..." Then
        cbo替诊医生.AddItem Nvl(rsReturn!姓名) & IIf(Nvl(rsReturn!专业技术职务) = "", "", "(" & Nvl(rsReturn!专业技术职务) & ")"), cbo替诊医生.ListCount - 1
        cbo替诊医生.ItemData(cbo替诊医生.NewIndex) = Val(Nvl(rsReturn!ID))
        cbo替诊医生.ListIndex = cbo替诊医生.NewIndex
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo替诊医生_Validate(Cancel As Boolean)
    If mblnOnly院内医生 Then
        If cbo替诊医生.ListIndex < 0 Then cbo替诊医生.Text = ""
    End If
End Sub

Private Sub cbo停诊原因_GotFocus()
    zlControl.TxtSelAll cbo停诊原因
End Sub

Private Sub cbo停诊原因_KeyPress(KeyAscii As Integer)
    Dim strReason As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(cbo停诊原因.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    strReason = SearchStopVisitReason(Me, cbo停诊原因, Trim(cbo停诊原因.Text))
    If strReason = "" Then Exit Sub
    
    zlControl.CboLocate cbo停诊原因, strReason
    If cbo停诊原因.ListIndex = -1 Then cbo停诊原因.Text = strReason
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function IsValied() As Boolean
    Dim arrTime As Variant
    Dim dtStartTime As Date, dtEndTime As Date
    Dim dtStartTimeNew As Date, dtEndTimeNew As Date
    Dim strSQL As String, strWhere As String, rsTemp As ADODB.Recordset
    Dim lngDoctor As Long
    
    Err = 0: On Error GoTo errHandle
    If mbytFun = F_取消停诊 Or mbytFun = F_取消替诊 Then
        '并发检查
        If mbytFun = F_取消停诊 Then
            strSQL = "Select 1 From 临床出诊记录 A Where a.ID = [1] And  a.停诊开始时间 Is Null"
        Else
            strSQL = "Select 1 From 临床出诊记录 A Where a.ID = [1] And  a.替诊开始时间 Is Null"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记录ID)
        If Not rsTemp.EOF Then
            MsgBox "当前安排已被他人取消" & IIf(mbytFun = F_取消替诊, "替诊", "停诊") & "，请刷新数据后查看！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If mbytFun = F_取消停诊 Then
            strSQL = "Select 1 From 临床出诊记录 A Where a.ID = [1] And  a.停诊终止时间< Sysdate"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记录ID)
            If Not rsTemp.EOF Then
                MsgBox "停诊时间的终止时间小于了当前时间，不能进行取消停诊操作！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            strSQL = "Select 1 From 临床出诊记录 A Where a.ID = [1] And  a.替诊开始时间< Sysdate"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记录ID)
            If Not rsTemp.EOF Then
                MsgBox "替诊时间的开始时间小于了当前时间，不能进行取消替诊操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If mbytFun = F_取消停诊 Then '启用后当前有效时间不能有交叉
            '除去当前启用上班时段
            strSQL = "Select a.开始时间, a.终止时间, a.停诊开始时间, a.停诊终止时间" & vbNewLine & _
                    " From 临床出诊记录 A, 临床出诊记录 B" & vbNewLine & _
                    " Where a.号源id = b.号源id And a.出诊日期 = b.出诊日期 And b.Id = [1] And a.Id <> b.Id" & vbNewLine

            '检查启用后是否有交叉
            strSQL = "Select 1 From " & _
                    "  (Select 开始时间, 停诊开始时间 As 终止时间 From (" & strSQL & ") Where 开始时间 < 停诊开始时间 And 终止时间 = 停诊终止时间" & vbNewLine & _
                    "   Union All" & vbNewLine & _
                    "   Select 停诊终止时间 As 开始时间, 终止时间 From (" & strSQL & ") Where 开始时间 = 停诊开始时间 And 终止时间 > 停诊终止时间" & vbNewLine & _
                    "   Union All" & vbNewLine & _
                    "   Select 开始时间, 停诊开始时间 As 终止时间 From (" & strSQL & ") Where 开始时间 < 停诊开始时间 And 终止时间 > 停诊终止时间" & vbNewLine & _
                    "   Union All" & vbNewLine & _
                    "   Select 停诊终止时间 As 开始时间, 终止时间 From (" & strSQL & ") Where 开始时间 < 停诊开始时间 And 终止时间 > 停诊终止时间) M, 临床出诊记录 N" & vbNewLine & _
                    " Where m.开始时间 < n.终止时间 And m.终止时间 > n.开始时间 And n.Id = [1] And Rownum < 2"
            '不能使用With语句，要报错
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记录ID)
            If Not rsTemp.EOF Then
                MsgBox "当前上班时段的时间范围与该号源今日目前有效的上班时段的时间范围有交叉，你不能取消停诊！", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
        IsValied = True: Exit Function
    End If
    
    '并发检查
    If mbytFun = F_停诊 Then
        strSQL = "Select 1 From 临床出诊记录 A Where a.ID = [1] And  a.停诊开始时间 Is Not Null"
    Else
        strSQL = "Select 1 From 临床出诊记录 A Where a.ID = [1] And  a.替诊开始时间 is Not Null"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记录ID)
    If Not rsTemp.EOF Then
        MsgBox "当前安排已被他人进行了" & IIf(mbytFun = F_替诊, "替诊", "停诊") & "，请刷新数据后查看！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If zlControl.TxtCheckInput(cbo停诊原因, "停诊原因", 50, False) = False Then Exit Function
    If Trim(cbo替诊医生.Text) = "" And cbo替诊医生.Visible Then
        MsgBox "替诊医生不能为空！", vbInformation, gstrSysName
        If cbo替诊医生.Visible And cbo替诊医生.Enabled Then cbo替诊医生.SetFocus
        Exit Function
    End If
    If mblnOnly院内医生 Then
        If cbo替诊医生.ListIndex < 0 And Trim(cbo替诊医生.Text) <> "" And cbo替诊医生.Visible Then
            MsgBox "你选择的医生不存在，请重新输入医生！", vbInformation + vbOKOnly, gstrSysName
            If cbo替诊医生.Visible And cbo替诊医生.Enabled Then cbo替诊医生.SetFocus
            Exit Function
        End If
    End If
    
    '停诊/替诊时间检查
    dtStartTime = Format(dtpStart.Tag, "yyyy-mm-dd hh:mm:ss")
    dtEndTime = Format(dtpEnd.Tag, "yyyy-mm-dd hh:mm:ss")
    dtStartTimeNew = GetWorkTrueDate(dtStartTime, Format(dtStartTime, "yyyy-mm-dd ") & Format(dtpStart.Value, "hh:mm:ss"), True, False)
    dtEndTimeNew = GetWorkTrueDate(dtStartTime, Format(dtStartTime, "yyyy-mm-dd ") & Format(dtpEnd.Value, "hh:mm:ss"))
    If dtStartTimeNew >= dtEndTimeNew Then
        MsgBox IIf(mbytFun = F_替诊, "替诊", "停诊") & "时间范围的结束时间必须大于开始时间！", vbInformation, gstrSysName
        If dtpEnd.Visible And dtpEnd.Enabled Then dtpEnd.SetFocus
        Exit Function
    End If
    If Not ((DateDiff("n", dtStartTime, dtStartTimeNew) >= 0 And DateDiff("n", dtStartTimeNew, dtEndTime) >= 0) _
            And (DateDiff("n", dtStartTime, dtEndTimeNew) >= 0 And DateDiff("n", dtEndTimeNew, dtEndTime) >= 0)) Then
        MsgBox IIf(mbytFun = F_替诊, "替诊", "停诊") & "时间必须在上班时段时间范围(" & Format(dtStartTime, "hh:mm") & "-" & Format(dtEndTime, "hh:mm") & ")内！", vbInformation, gstrSysName
        If dtpEnd.Visible And dtpEnd.Enabled Then dtpEnd.SetFocus
        Exit Function
    End If
    
    If zlDatabase.Currentdate > dtStartTimeNew Then
        MsgBox IIf(mbytFun = F_替诊, "替诊", "停诊") & "时间的开始时间小于了当前时间，不能进行" & IIf(mbytFun = F_替诊, "替诊", "停诊") & "操作！", vbInformation, gstrSysName
        If dtpStart.Visible And dtpStart.Enabled Then dtpStart.SetFocus
        Exit Function
    End If
    
    If mbytFun = F_替诊 Then
        If mblnOnly院内医生 Then
            strWhere = " And a.医生ID = [4]"
            lngDoctor = cbo替诊医生.ItemData(cbo替诊医生.ListIndex)
        Else
            strWhere = " And a.医生姓名 = [5] And a.医生ID Is Null"
        End If
        
        If lngDoctor <> 0 Then
            strSQL = "Select 1 From 临床出诊记录 A Where ID = [1] And Nvl(医生ID,替诊医生ID)= [2] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记录ID, lngDoctor)
            If Not rsTemp.EOF Then
                MsgBox "替诊医生不能为原安排医生，请选择其它医生！", vbInformation, gstrSysName
                If cbo替诊医生.Visible And cbo替诊医生.Enabled Then cbo替诊医生.SetFocus
                Exit Function
            End If
        End If
        
        '在该时段内，替诊医生不能存在其他的出诊安排
        '若A[A1,A2],B[B1,B2],且B为空或完全包含于A中(A1<=B1,A2>=B2).那么X[X1,X2]与A-B有交集，则
        '(X1>=A1 And X1<=NVL(B1,A2)) Or (X2>=A1 And X2<=NVL(B1,A2)) Or (X1>=NVL(B2,A1) And X1<=A2) Or (X2>=NVL(B2,A1) And X2<=A2)
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊记录 A" & vbNewLine & _
                " Where a.出诊日期 = To_Date([1], 'yyyy-mm-dd')" & strWhere & vbNewLine & _
                "       And (([2] Between a.开始时间 And Nvl(a.停诊开始时间, a.终止时间)) Or ([3] Between a.开始时间 And Nvl(a.停诊开始时间, a.终止时间))" & vbNewLine & _
                "       Or ([2] Between Nvl(a.停诊终止时间, a.开始时间) And a.终止时间) Or ([3] Between Nvl(a.停诊终止时间, a.开始时间) And a.终止时间))"
        '不能有替诊
        strSQL = strSQL & vbNewLine & _
                "       And [2] < Nvl(a.替诊终止时间, To_Date('1900-01-01','yyyy-mm-dd'))" & vbNewLine & _
                "       And [3] > Nvl(a.替诊开始时间, To_Date('3000-01-01','yyyy-mm-dd'))"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt出诊日期.Text, dtStartTimeNew, dtEndTimeNew, _
            lngDoctor, GetReplaceDoctor(Trim(cbo替诊医生.Text)))
        If Not rsTemp.EOF Then
            MsgBox "替诊医生在替诊时间(" & Format(dtStartTimeNew, "hh:mm") & "-" & Format(dtEndTimeNew, "hh:mm") & ")范围内已存在其它出诊安排，请选择其它医生！", vbInformation, gstrSysName
            If cbo替诊医生.Visible And cbo替诊医生.Enabled Then cbo替诊医生.SetFocus
            Exit Function
        End If
        
        '替诊医生级别检查
        If mbln替诊医生级别检查 Then
           strSQL = "Select Zl1_Ex_Isdoctorsamelevel(a.医生id, a.医生姓名, [2], [3]) As 结果 From 临床出诊记录 A Where ID = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记录ID, lngDoctor, GetReplaceDoctor(Trim(cbo替诊医生.Text)))
            If Not rsTemp.EOF Then
                 If Val(Nvl(rsTemp!结果)) = -1 Then
                    MsgBox "替诊医生的职务级别不够，不允许替诊，请选择其它医生！", vbInformation, gstrSysName
                    If cbo替诊医生.Visible And cbo替诊医生.Enabled Then cbo替诊医生.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str记录IDs As String
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandle
    
    If IsValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    '检查是否要输出预约清单
    If mbytFun = F_停诊 Or mbytFun = F_替诊 Then
        If mbytFun = F_停诊 Then
            strSQL = "Select a.ID As 记录ID" & vbNewLine & _
                    " From 临床出诊记录 A, 病人挂号记录 B, 临床出诊号源 C" & vbNewLine & _
                    " Where a.Id = b.出诊记录id And a.号源id = c.Id And b.记录状态 = 1 And Nvl(b.执行状态, 0) = 0 And a.Id = [1] " & vbNewLine & _
                    "       And (b.记录性质 = 1 And b.发生时间 Between a.停诊开始时间 And a.停诊终止时间" & vbNewLine & _
                    "           Or b.记录性质 = 2 And b.预约时间 Between a.停诊开始时间 And a.停诊终止时间) And Rownum < 2"
        Else
            strSQL = "Select a.ID As 记录ID" & vbNewLine & _
                    " From 临床出诊记录 A, 病人挂号记录 B, 临床出诊号源 C" & vbNewLine & _
                    " Where a.Id = b.出诊记录id And a.号源id = c.Id And b.记录状态 = 1 And Nvl(b.执行状态, 0) = 0  And a.Id = [1] " & vbNewLine & _
                    "       And (b.记录性质 = 1 And b.发生时间 Between a.替诊开始时间 And a.替诊终止时间" & vbNewLine & _
                    "           Or b.记录性质 = 2 And b.预约时间 Between a.替诊开始时间 And a.替诊终止时间) And Rownum < 2"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记录ID)
        
        If rsTemp Is Nothing Then GoTo UnloadForm:
        If rsTemp.EOF Then GoTo UnloadForm:
    
        Do While Not rsTemp.EOF
            If InStr(strTemp & ",", "," & Nvl(rsTemp!记录ID) & ",") = 0 Then
                str记录IDs = str记录IDs & "," & Nvl(rsTemp!记录ID)
            End If
            rsTemp.MoveNext
        Loop
        If str记录IDs <> "" Then str记录IDs = Mid(str记录IDs, 2)
        
        If mbyt预约清单控制方式 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "出诊记录IDS=" & str记录IDs, 3)
        ElseIf mbyt预约清单控制方式 = 2 Then
            If MsgBox("当前号源停诊时间内存在预约或挂号病人，是否将预约清单输出到Excel中？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "出诊记录IDS=" & str记录IDs, 3)
            End If
        End If
        
        If mbyt预约清单打印方式 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "出诊记录IDS=" & str记录IDs, 2)
        ElseIf mbyt预约清单打印方式 = 2 Then
            If MsgBox("当前号源停诊时间内存在预约或挂号病人，你确定要打印预约清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me, "出诊记录IDS=" & str记录IDs, 2)
            End If
        End If
    End If
UnloadForm:
    mblnOk = True
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitFace()
    Select Case mbytFun
    Case F_停诊
        Me.Caption = "停诊": Me.Height = 4350
        lbl停诊原因.Caption = "停诊原因"
        lbl停诊时间.Caption = "停诊时间"
        cbo替诊医生.Visible = False
        dtpStart.Enabled = True: dtpEnd.Enabled = True
    Case F_取消停诊
        Me.Caption = "取消停诊": Me.Height = 4350
        lbl停诊原因.Caption = "停诊原因"
        lbl停诊时间.Caption = "停诊时间"
        cbo替诊医生.Visible = False
        cbo停诊原因.Enabled = False
        dtpStart.Enabled = False: dtpEnd.Enabled = False
    Case F_替诊
        Me.Caption = "替诊": Me.Height = 5250
        lbl停诊原因.Caption = "替诊原因"
        lbl停诊时间.Caption = "替诊时间"
        dtpStart.Enabled = True: dtpEnd.Enabled = True
    Case F_取消替诊
        Me.Caption = "取消替诊": Me.Height = 5250
        lbl停诊原因.Caption = "替诊原因"
        lbl停诊时间.Caption = "替诊时间"
        cbo停诊原因.Enabled = False
        cbo替诊医生.Enabled = False
        dtpStart.Enabled = False: dtpEnd.Enabled = False
        idkDoctor.Enabled = False
    End Select
    cmdHelp.Top = Me.ScaleHeight - cmdHelp.Height - 300
End Sub

Private Function InitData() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng号源Id As Long, rsSignalSource As ADODB.Recordset
    Dim dtCurrent As Date
    
    Err = 0: On Error GoTo errHandle
    If mbytFun = F_停诊 Or mbytFun = F_替诊 Then
        strSQL = "Select 编码, 名称, 简码, Nvl(缺省标志,0) As 缺省 From 常用停诊原因"
        Set mrsStopReason = zlDatabase.OpenSQLRecord(strSQL, "获取常用停诊原因")
        With cbo停诊原因
            .Clear
            Do While Not mrsStopReason.EOF
                .AddItem Nvl(mrsStopReason!编码) & "-" & Nvl(mrsStopReason!名称)
                If Val(Nvl(mrsStopReason!缺省)) = 1 Then .ListIndex = .NewIndex
                mrsStopReason.MoveNext
            Loop
        End With
    End If
    
    If mbytFun = F_替诊 Or mbytFun = F_取消替诊 Then
        If mblnOnly院内医生 Then
            idkDoctor.IDkindStr = "医生|医生|0|0|0|0|0||0|0|0"
            idkDoctor.ToolTipText = "只能选院内建档医生"
        Else
            idkDoctor.IDkindStr = "院内医生|院内医生|0|0|0|0|0||0|0|0;院外医生|院外医生|0|0|0|0||0|0|0"
        End If
    End If
    
    '加载安排信息
    strSQL = "Select a.Id, a.号源id, a.出诊日期, a.上班时段, a.开始时间, a.终止时间," & vbNewLine & _
            "        a.停诊开始时间, a.停诊终止时间, a.停诊原因," & vbNewLine & _
            "        a.替诊开始时间, a.替诊终止时间, a.替诊医生id, a.替诊医生姓名, b.专业技术职务" & vbNewLine & _
            " From 临床出诊记录 A, 人员表 B" & vbNewLine & _
            " Where a.替诊医生id = b.ID(+) And  a.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记录ID)
    If rsTemp.BOF Then Exit Function
    With rsTemp
        lng号源Id = Nvl(!号源ID)
        txt出诊日期.Text = Format(Nvl(!出诊日期), "yyyy-mm-dd")
        zlControl.CboLocate cbo停诊原因, Nvl(!停诊原因)
        If cbo停诊原因.ListIndex = -1 And Nvl(!停诊原因) <> "" Then cbo停诊原因.AddItem Nvl(!停诊原因): cbo停诊原因.ListIndex = cbo停诊原因.NewIndex
        '替诊医生的加载移到后面
        txt上班时段.Text = Nvl(!上班时段)
        If mbytFun = F_取消停诊 Then
            dtpStart.Value = Format(Nvl(!停诊开始时间, "00:00:00"), "hh:mm:ss")
            dtpEnd.Value = Format(Nvl(!停诊终止时间, "00:00:00"), "hh:mm:ss")
        ElseIf mbytFun = F_取消替诊 Then
            dtpStart.Value = Format(Nvl(!替诊开始时间, "00:00:00"), "hh:mm:ss")
            dtpEnd.Value = Format(Nvl(!替诊终止时间, "00:00:00"), "hh:mm:ss")
        Else
            '如果停诊/替诊当前已处于挂号的安排，则以当前时间+1分钟为缺省时间.
            dtCurrent = zlDatabase.Currentdate
            If dtCurrent >= Nvl(!开始时间, "00:00:00") Then
                dtpStart.Value = Format(DateAdd("n", 1, dtCurrent), "hh:mm:ss")
            Else
                dtpStart.Value = Format(Nvl(!开始时间, "00:00:00"), "hh:mm:ss")
            End If
            dtpStart.Tag = Format(Nvl(!开始时间, "00:00:00"), "yyyy-MM-dd hh:mm:ss")
            dtpEnd.Value = Format(Nvl(!终止时间, "00:00:00"), "hh:mm:ss")
            dtpEnd.Tag = Format(Nvl(!终止时间, "00:00:00"), "yyyy-MM-dd hh:mm:ss")
        End If
    End With
    
    '号源信息
    strSQL = "Select a.号类, a.号码, a.科室ID, b.名称 As 科室, c.名称 As 收费项目, a.医生姓名," & vbNewLine & _
            "        Decode(Nvl(a.假日控制状态, 0), 1, '开放预约', 2, '禁止预约', 3, '受节假日设置控制', '不上班') As 假日控制," & vbNewLine & _
            "        Nvl(a.是否建病案, 0) As 病案" & vbNewLine & _
            " From 临床出诊号源 A, 部门表 B, 收费项目目录 C" & vbNewLine & _
            " Where a.科室id = b.Id And a.项目id = c.Id And a.Id = [1]"
    Set rsSignalSource = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng号源Id)
    If rsSignalSource.BOF Then Exit Function
    With rsSignalSource
        txtSignalNO.Text = Nvl(!号码)
        txt号类.Text = Nvl(!号类)
        txtDept.Text = Nvl(!科室)
        txtDept.Tag = Nvl(!科室ID)
        txtItem.Text = Nvl(!收费项目)
        txtDoctor.Text = Nvl(!医生姓名)
        txt假日控制.Text = Nvl(!假日控制)
        chk建档.Value = Val(Nvl(!病案))
    End With
    
    '加载替诊医生
    If mbytFun = F_替诊 Or mbytFun = F_取消替诊 Then
        If mbytFun = F_替诊 Then
            Call LoadDoctor(Val(txtDept.Tag))
        Else
            With rsTemp
                If Val(Nvl(rsTemp!替诊医生id)) = 0 Then idkDoctor.IDKind = 2
                zlControl.CboLocate cbo替诊医生, Nvl(!替诊医生id), True
                If cbo替诊医生.ListIndex = -1 And Nvl(!替诊医生姓名) <> "" Then
                    cbo替诊医生.AddItem Nvl(!替诊医生姓名) & IIf(Nvl(!专业技术职务) = "", "", "(" & Nvl(!专业技术职务) & ")")
                    cbo替诊医生.ItemData(cbo替诊医生.NewIndex) = Val(Nvl(!替诊医生id))
                    cbo替诊医生.ListIndex = cbo替诊医生.NewIndex
                End If
            End With
        End If
    End If
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDoctor(Optional ByVal lngSectID As Long = 0, Optional ByVal strWhere As String, _
    Optional ByVal blnNotEqualID As Boolean, _
    Optional ByVal blnGetSql As Boolean, Optional ByRef strSQL As String) As ADODB.Recordset
    '得到指定科室下的所有医生并返回
    '入参：
    '   blnNotEqualID - 不等于ID
    '   strWhere - 最多只能包含"[1]"
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        "Select Distinct c.id,c.编号,c.姓名,c.简码,c.专业技术职务" & vbNewLine & _
        " From 人员性质说明 a, 部门人员 b ,人员表 c" & vbNewLine & _
        " Where b.人员id=c.id And b.人员id=a.人员id  And  a.人员性质=[2]" & vbNewLine & _
        "       And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) " & vbNewLine & _
        "       And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & vbNewLine & _
                IIf(lngSectID = 0, "", IIf(blnNotEqualID, "  And b.部门id <> [3]", "  And b.部门id = [3]")) & vbNewLine & _
                strWhere & vbNewLine & _
        " Order By c.姓名"
        
    If blnGetSql Then
        strSQL = Replace(strSQL, "[2]", "'医生'")
        strSQL = Replace(strSQL, "[3]", lngSectID)
        Exit Function
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", "", "医生", lngSectID)
    Set GetDoctor = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    mblnOnly院内医生 = Val(zlDatabase.GetPara("只允许选院内医生", glngSys, mlngModule, "0")) = 1
    mbln替诊医生级别检查 = Val(zlDatabase.GetPara("替诊医生级别检查", glngSys, mlngModule, "0")) = 1
    mbyt预约清单控制方式 = Val(zlDatabase.GetPara("预约清单控制方式", glngSys, mlngModule, "0"))
    mbyt预约清单打印方式 = Val(zlDatabase.GetPara("预约清单打印方式", glngSys, mlngModule, "0"))
    
    Call InitFace
    If InitData() = False Then Unload Me: Exit Sub
    Call SetEnabledBackColor(Me.Controls)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsStopReason = Nothing
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String
    Dim arrTime As Variant
    Dim dtStartTime As Date, dtEndTime As Date
    Dim lngDoctor As Long
    
    Err = 0: On Error GoTo errHandle
    '停诊时间
    If mbytFun = F_停诊 Or mbytFun = F_替诊 Then
        dtStartTime = GetWorkTrueDate(dtpStart.Tag, Format(dtpStart.Tag, "yyyy-mm-dd ") & Format(dtpStart.Value, "hh:mm:ss"), True, False)
        dtEndTime = GetWorkTrueDate(dtpStart.Tag, Format(dtpStart.Tag, "yyyy-mm-dd ") & Format(dtpEnd.Value, "hh:mm:ss"))
    End If
    
    Select Case mbytFun
    Case F_停诊
        'Zl_临床出诊记录_Stopvisit
        strSQL = "Zl_临床出诊记录_Stopvisit("
        '  记录id_In   Varchar2,
        strSQL = strSQL & "" & mlng记录ID & ","
        '  开始时间_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(dtStartTime, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  终止时间_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(dtEndTime, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  停诊原因_In Varchar2 := Null,
        strSQL = strSQL & "'" & NeedName(cbo停诊原因.Text) & "',"
        '  操作员_In   Varchar2 := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  操作时间_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  取消替诊_In Number:=0
        strSQL = strSQL & "" & 0 & ")"
    Case F_取消停诊
        'Zl_临床出诊记录_Stopvisit
        strSQL = "Zl_临床出诊记录_Stopvisit("
        '  记录id_In   Varchar2,
        strSQL = strSQL & "" & mlng记录ID & ","
        '  开始时间_In Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  终止时间_In Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  停诊原因_In Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  操作员_In   Varchar2 := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  操作时间_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  取消替诊_In Number:=0
        strSQL = strSQL & "" & 1 & ")"
    Case F_替诊
        If cbo替诊医生.ListIndex <> -1 And mblnOnly院内医生 Then
            lngDoctor = cbo替诊医生.ItemData(cbo替诊医生.ListIndex)
        End If
        'Zl_临床出诊记录_Replacedoctor
        strSQL = "Zl_临床出诊记录_Replacedoctor("
        '  记录id_In       Varchar2,
        strSQL = strSQL & "" & mlng记录ID & ","
        '  开始时间_In     Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(dtStartTime, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  终止时间_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(dtEndTime, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  停诊原因_In Varchar2 := Null,
        strSQL = strSQL & "'" & NeedName(cbo停诊原因.Text) & "',"
        '  替诊医生id_In   Number := Null,
        strSQL = strSQL & "" & ZVal(lngDoctor) & ","
        '  替诊医生姓名_In Varchar2 := Null,
        strSQL = strSQL & "'" & GetReplaceDoctor(Trim(cbo替诊医生.Text)) & "',"
        '  操作员姓名_In   临床出诊停诊记录.申请人%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  操作员编号_In   人员表.编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作时间_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  取消替诊_In Number:=0
        strSQL = strSQL & "" & 0 & ")"
    Case F_取消替诊
        'Zl_临床出诊记录_Replacedoctor
        strSQL = "Zl_临床出诊记录_Replacedoctor("
        '  记录id_In       Varchar2,
        strSQL = strSQL & "" & mlng记录ID & ","
        '  开始时间_In     Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  终止时间_In     Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  停诊原因_In     Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  替诊医生id_In   Number := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  替诊医生姓名_In Varchar2 := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  操作员姓名_In   临床出诊停诊记录.申请人%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  操作员编号_In   人员表.编号%Type := Null,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作时间_In Varchar2 := Null,
        strSQL = strSQL & "To_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  取消替诊_In Number:=0
        strSQL = strSQL & "" & 1 & ")"
    End Select
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetReplaceDoctor(ByVal strIn As String, Optional strSplit As String = "(") As String
    '分离出替诊医生姓名
    GetReplaceDoctor = Mid(strIn, 1, IIf(InStr(strIn, strSplit) = 0, Len(strIn), InStr(strIn, strSplit) - 1))
End Function

Private Sub idkDoctor_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Err = 0: On Error GoTo errHandle
    mblnOnly院内医生 = index = 1
    cbo替诊医生.Clear
    If mblnOnly院内医生 = False Then Exit Sub
    
    Call LoadDoctor(Val(txtDept.Tag))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDoctor(Optional ByVal lng科室ID As Long)
    '根据科室ID加载医生
    '说明：
    '   科室ID为0是加载所有医生
    Dim strPersons As String
    Dim rsDoctor As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandle
    
    cbo替诊医生.Clear
    Set rsDoctor = GetDoctor(lng科室ID)
    If rsDoctor Is Nothing Then Exit Sub
    Do While Not rsDoctor.EOF
        If InStr("," & strPersons & ",", "," & Nvl(rsDoctor!ID) & ",") = 0 Then
            strPersons = strPersons & "," & Nvl(rsDoctor!ID)
            cbo替诊医生.AddItem Nvl(rsDoctor!姓名) & IIf(Nvl(rsDoctor!专业技术职务) = "", "", "(" & Nvl(rsDoctor!专业技术职务) & ")")
            cbo替诊医生.ItemData(cbo替诊医生.NewIndex) = Val(Nvl(rsDoctor!ID))
        End If
        rsDoctor.MoveNext
    Loop
    If lng科室ID <> 0 Then
        cbo替诊医生.AddItem "其他科室医生..."
        cbo替诊医生.ItemData(cbo替诊医生.NewIndex) = -1
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub idkDoctor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

