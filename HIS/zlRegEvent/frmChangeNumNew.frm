VERSION 5.00
Begin VB.Form frmChangeNumNew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "病人换号"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra2 
      Caption         =   "更换号别"
      Height          =   1065
      Left            =   165
      TabIndex        =   20
      Top             =   2070
      Width           =   6840
      Begin VB.ComboBox cmbDiagRoom2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2925
         TabIndex        =   25
         Text            =   "cmbDiagRoom2"
         Top             =   600
         Width           =   1620
      End
      Begin VB.ComboBox cmbSect2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   780
         TabIndex        =   23
         Top             =   615
         Width           =   1620
      End
      Begin VB.ComboBox cmbDocTor2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5040
         TabIndex        =   27
         Text            =   "cmbDocTor2"
         Top             =   600
         Width           =   1620
      End
      Begin VB.CommandButton cmdItemSel 
         Caption         =   "&P"
         Height          =   300
         Left            =   6285
         TabIndex        =   21
         ToolTipText     =   "选择新号别"
         Top             =   210
         Width           =   375
      End
      Begin VB.Label lblItem2 
         Alignment       =   1  'Right Justify
         Height          =   180
         Left            =   210
         TabIndex        =   31
         Top             =   255
         Width           =   5910
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   180
         Left            =   375
         TabIndex        =   22
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         Height          =   180
         Left            =   4635
         TabIndex        =   26
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "诊室"
         Height          =   180
         Left            =   2520
         TabIndex        =   24
         Top             =   675
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7155
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7155
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   735
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7155
      TabIndex        =   28
      Top             =   255
      Width           =   1100
   End
   Begin VB.Frame fra1 
      Caption         =   "原挂号单信息"
      Height          =   1815
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   6825
      Begin VB.TextBox txtSect 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   780
         TabIndex        =   15
         Top             =   1290
         Width           =   1620
      End
      Begin VB.TextBox txtDiagRoom 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2925
         TabIndex        =   17
         Top             =   1290
         Width           =   1620
      End
      Begin VB.TextBox txtTime 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5025
         TabIndex        =   13
         Top             =   915
         Width           =   1620
      End
      Begin VB.TextBox txtExes 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2925
         TabIndex        =   11
         Top             =   900
         Width           =   1620
      End
      Begin VB.TextBox txtOutNum 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   780
         TabIndex        =   8
         Top             =   900
         Width           =   1620
      End
      Begin VB.TextBox txtOld 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5025
         TabIndex        =   7
         Top             =   525
         Width           =   1620
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2925
         TabIndex        =   5
         Top             =   510
         Width           =   1620
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   780
         TabIndex        =   3
         Top             =   510
         Width           =   1620
      End
      Begin VB.ComboBox cmbDoctor 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5025
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1305
         Width           =   1650
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   180
         Left            =   375
         TabIndex        =   14
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label lblOldItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号单【】"
         Height          =   180
         Left            =   315
         TabIndex        =   1
         Top             =   255
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "时间"
         Height          =   180
         Left            =   4635
         TabIndex        =   12
         Top             =   990
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "费别"
         Height          =   180
         Left            =   2535
         TabIndex        =   10
         Top             =   975
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         Height          =   180
         Left            =   4620
         TabIndex        =   18
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "诊室"
         Height          =   180
         Left            =   2520
         TabIndex        =   16
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "门诊号"
         Height          =   180
         Left            =   210
         TabIndex        =   9
         Top             =   975
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   180
         Left            =   4635
         TabIndex        =   6
         Top             =   585
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Left            =   2535
         TabIndex        =   4
         Top             =   585
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Left            =   390
         TabIndex        =   2
         Top             =   585
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmChangeNumNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STR_COMP = "|',~" '分隔字符串

Private strSQL As String
Private i As Long
Private mblnCancel As Boolean
Private mlng挂号ID As Long
Private mstrNo As String
Private mstr号别 As String
Private mrsDoctor As New ADODB.Recordset
Private mlng出诊记录ID As Long

Public Function ShowMe(ByVal lng挂号ID As String, frmParent As Form) As Boolean
'显示本窗体并返回选择的是否正确
    On Error GoTo errHandle
    Dim rsTmp As ADODB.Recordset
    Dim strDoctor As String, lng执行部门ID As Long
    
    mlng挂号ID = lng挂号ID
    mblnCancel = False
    
    '读出以前的病人挂号记录
    strSQL = _
        " Select A.NO,X.号别,X.姓名,X.性别,X.年龄,X.门诊号," & _
        " A.费别,A.发生时间,X.诊室,X.执行人,X.执行部门ID," & _
        " D.号类,C.名称 as 收费项目名称,B.名称 as 执行部门名称,D1.ID As 出诊记录ID" & _
        " From 门诊费用记录 A,部门表 B,收费项目目录 C,临床出诊号源 D,临床出诊记录 D1,病人挂号记录 X" & _
        " Where A.记录性质=4 And A.记录状态=1 And A.序号=1 And A.NO=X.NO" & _
        " AND X.记录状态=1 and x.记录性质=1 And A.收费细目ID=C.ID And X.出诊记录ID=D1.ID And D.ID=D1.号源ID And X.执行部门ID=B.ID" & _
        " And X.ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人换号", lng挂号ID)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Me.lblOldItem.Caption = "挂号单【" & rsTmp!NO & "】   号类:" & Nvl(rsTmp!号类) & _
            "    号别:" & Nvl(rsTmp!号别) & "    挂号项目:" & rsTmp!收费项目名称
        mstrNo = rsTmp!NO
        mstr号别 = zlCommFun.Nvl(rsTmp!号别)
        Me.txtName = zlCommFun.Nvl(rsTmp!姓名)
        Me.txtSex = zlCommFun.Nvl(rsTmp!性别)
        Me.txtOld = zlCommFun.Nvl(rsTmp!年龄)
        Me.txtOutNum = zlCommFun.Nvl(rsTmp!门诊号)
        Me.txtExes = zlCommFun.Nvl(rsTmp!费别)
        Me.txtTime = Format(Nvl(rsTmp!发生时间), "YYYY-MM-DD HH:MM:SS")
        Me.txtSect = zlCommFun.Nvl(rsTmp!执行部门名称)
        Me.txtDiagRoom = zlCommFun.Nvl(rsTmp!诊室)
        
        '设置默认医生
        strDoctor = "" & rsTmp!执行人
        lng执行部门ID = Val("" & rsTmp!执行部门id)
        strSQL = "SELECT b.ID,b.简码,b.姓名 FROM 部门人员 a,人员表 b,人员性质说明 c  " & _
            " WHERE b.id=a.人员ID AND b.id=c.人员id AND c.人员性质='医生' AND a.部门ID=[1]" & _
            " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
            " And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人换号", lng执行部门ID)
        Me.cmbDoctor.Clear
        Me.cmbDoctor.AddItem "W-无" & String(400, " ") & STR_COMP
        Me.cmbDoctor.ListIndex = 0
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                Me.cmbDoctor.AddItem rsTmp!简码 & "-" & rsTmp!姓名 & String(400, " ") & STR_COMP & rsTmp!ID
                rsTmp.MoveNext
            Next
            If Trim(strDoctor) <> "" Then
                For i = 0 To Me.cmbDoctor.ListCount - 1
                    If Me.cmbDoctor.List(i) Like "*-" & strDoctor & " *" Then
                        Me.cmbDoctor.ListIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
    Else
        MsgBox "无该病人挂号信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Me.Show 1, frmParent
    If mblnCancel = False Then
        ShowMe = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetRoom(lng记录ID As Long) As String
'功能：根据号别的分诊方式获取号别的诊室
    Dim strSQL As String, strRoomIDs As String
    Dim rsTmp As ADODB.Recordset, rsRoom As ADODB.Recordset
    On Error GoTo errH
    
    strSQL = "Select ID,Nvl(分诊方式,0) as 分诊 From 临床出诊记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人换号", lng记录ID)
    If rsTmp.EOF Then Exit Function
    If rsTmp!分诊 = 0 Then Exit Function '不分诊
    
    '处理分诊
    If rsTmp!分诊 = 1 Then
        '指定诊室
        strSQL = "Select A.名称 As 门诊诊室 From 门诊诊室 A,临床出诊诊室记录 B Where B.记录ID=[1] And A.ID=B.诊室ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人换号", lng记录ID)
        If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 2 Then
        '动态分诊：该个号别当天挂号未诊数最少的诊室   //todo未考虑预约挂号
        strSQL = _
            " Select 门诊诊室,Sum(NUM) as NUM From (" & _
                " Select B.名称 As 门诊诊室,0 as NUM From 临床出诊诊室记录 A,门诊诊室 B Where A.记录ID=[1] And A.诊室ID=B.ID" & _
                " Union ALL" & _
                " Select 诊室,Count(诊室) as NUM From 病人挂号记录" & _
                " Where Nvl(执行状态,0)=0 And 出诊记录ID=[1]" & _
                " And 记录性质=1 and 记录状态=1 and 发生时间 Between Trunc(Sysdate) And  Sysdate" & _
                " And 诊室 IN (Select B.名称 From 临床出诊诊室记录 A,门诊诊室 B Where A.记录ID=[1] And A.诊室ID=B.ID)" & _
                " Group by 诊室)" & _
            " Group by 门诊诊室" & _
            " Order by Num"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人换号", lng记录ID)
        If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 3 Then
        '平均分诊：当前分配=1表示下次应取的当前诊室,(下面的语句,必须要用*,否则更新会出错)
        strSQL = "Select * From 临床出诊诊室记录 Where 记录ID=" & rsTmp!ID
        '返回可更新记录集
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption, adOpenStatic, adLockOptimistic)
        
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!当前分配), 0, rsTmp!当前分配) = 1 Then
                    strRoomIDs = rsTmp!诊室ID
                    rsTmp!当前分配 = 0

                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!当前分配 = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '处理第一次平均分配
            If strRoomIDs = "" Then
                rsTmp.MoveFirst
                strRoomIDs = rsTmp!诊室ID
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!当前分配 = 1
                rsTmp.Update
            End If
        End If
        If strRoomIDs <> "" Then
            strSQL = "Select 名称 From 门诊诊室 Where ID = [1]"
            Set rsRoom = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRoomIDs)
            If Not rsRoom.EOF Then
                GetRoom = rsRoom!名称
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdItemSel_Click()
On Error GoTo errHandle
'选择号别
    Dim strReturn As String
    '号表ID,项目ID,医生ID,医生,科室ID,科室,号类,号别,出诊记录ID
    If frmNumSortSelNew.ShowMe(mlng挂号ID, strReturn, Me) Then
        '号类及号别
        lblItem2.Caption = "号类:" & Trim(Split(strReturn, ",")(6)) & "   号别:" & Trim(Split(strReturn, ",")(7))
        lblItem2.Tag = Trim(Split(strReturn, ",")(7))
        mstr号别 = Trim(Split(strReturn, ",")(7))
        mlng出诊记录ID = Val(Split(strReturn, ",")(8))
        '执行部门
        Me.cmbSect2.Text = Trim(Split(strReturn, ",")(5))
        Me.cmbSect2.Tag = CLng(Trim(Split(strReturn, ",")(4)))
        
        '找到诊室
        Me.cmbDiagRoom2.Text = GetRoom(mlng出诊记录ID)
        
        '读出医生
        If Trim(Split(strReturn, ",")(3)) = "" Then
            Me.cmbDocTor2.Text = "无"
        Else
            Me.cmbDocTor2.Text = Trim(Split(strReturn, ",")(3)) & String(400, " ") & STR_COMP & Trim(Split(strReturn, ",")(2))
        End If
    
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOk_Click()
Dim strDoctor  As String
On Error GoTo errHandle

    'NO_IN          病人挂号记录.NO%TYPE:=NULL,
    '号别_IN        病人挂号记录.号别%TYPE:=NULL,
    '诊室_IN        病人挂号记录.诊室%TYPE:=NULL,
    '执行部门ID_IN  病人挂号记录.执行部门ID%TYPE:=NULL,
    '医生_IN        病人挂号记录.执行人%TYPE:=NULL,
    '医生ID_IN      病人挂号汇总.医生ID%TYPE:=NULL,
    '医生2_IN       病人挂号记录.执行人%TYPE:=NULL,
    '医生ID2_IN     病人挂号汇总.医生ID%TYPE:=NULL
    If Trim(lblItem2.Tag) = "" Then MsgBox "请选择一个号别！", vbInformation, gstrSysName: Exit Sub
    If Trim(zlCommFun.GetNeedName(Trim(Split(cmbDoctor.Text, STR_COMP)(0)))) = "无" Then
        strDoctor = "'',null"
    Else
        strDoctor = "'" & zlCommFun.GetNeedName(Trim(Split(cmbDoctor.Text, STR_COMP)(0))) & "'," & Trim(Split(cmbDoctor.Text, STR_COMP)(1))
    End If
    If Trim(cmbDocTor2.Text) = "无" Then
        strSQL = "'',null"
    Else
        strSQL = "'" & Trim(Split(cmbDocTor2.Text, STR_COMP)(0)) & "'," & Trim(Split(cmbDocTor2.Text, STR_COMP)(1))
    End If
    If ExcPlugInFun(1, mlng挂号ID, Trim(Split(cmbDocTor2.Text, STR_COMP)(0)), Me.cmbDiagRoom2.Text, lblItem2.Tag, mlng出诊记录ID) = False Then Exit Sub
    
    strSQL = "ZL_病人挂号记录_换号('" & mstrNo & "','" & lblItem2.Tag & "','" & _
            Me.cmbDiagRoom2.Text & "'," & Me.cmbSect2.Tag & "," & strDoctor & "," & strSQL & "," & mlng出诊记录ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    lblItem2.Tag = ""
    Me.cmbSect2.Tag = 0
    Me.cmbDiagRoom2.Text = ""
    Me.cmbDocTor2.Text = ""
End Sub
