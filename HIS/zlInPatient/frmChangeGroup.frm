VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmChangeGroup 
   Caption         =   "病人转医疗小组"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   Icon            =   "frmChangeGroup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   6135
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraGroup 
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   5940
      Begin VB.ComboBox cbo住院医师 
         Height          =   300
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   1830
      End
      Begin VB.ComboBox cbo医疗小组 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   1890
      End
      Begin VB.ComboBox cbo主治医师 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   555
         Width           =   1890
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3900
         TabIndex        =   3
         Top             =   555
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生效时间"
         Height          =   180
         Left            =   3120
         TabIndex        =   26
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院医师"
         Height          =   180
         Left            =   3120
         TabIndex        =   25
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主治医师"
         Height          =   180
         Left            =   210
         TabIndex        =   24
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新医疗小组"
         Height          =   180
         Left            =   30
         TabIndex        =   23
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3495
      TabIndex        =   4
      Top             =   2670
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4755
      TabIndex        =   5
      Top             =   2670
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   5940
      Begin VB.TextBox txt床号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1890
      End
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1830
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         Width           =   1890
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1890
      End
      Begin VB.TextBox txt科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   630
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原医疗小组"
         Height          =   180
         Left            =   2940
         TabIndex        =   19
         Top             =   1065
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前床位"
         Height          =   180
         Left            =   195
         TabIndex        =   18
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4605
         TabIndex        =   11
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   3000
         TabIndex        =   9
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   540
         TabIndex        =   7
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   180
         Left            =   375
         TabIndex        =   17
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前科室"
         Height          =   180
         Left            =   3120
         TabIndex        =   16
         Top             =   690
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   315
      TabIndex        =   6
      Top             =   2670
      Width           =   1100
   End
End
Attribute VB_Name = "frmChangeGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstrPirvs As String
Private mlngUnit As Long

Private mstrUnit As String
Private mrsPatiInfo As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cbo医疗小组_Click()
    Dim strSQL As String, strSQL医疗小组 As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lng医师 As Long
    
    On Error GoTo errHandle:
    
    '如果为病人指定了医疗小组，则"住院医师、主治医师"都从对应医疗小组中的医生中选择
    strSQL医疗小组 = "Select Distinct A.ID, A.编号, A.简码, A.姓名" & vbNewLine & _
                    " From 人员表 A, 人员性质说明 B, 部门人员 C, 医疗小组人员 D" & vbNewLine & _
                    " Where A.ID = B.人员id And A.ID = C.人员id And a.id = d.人员id And B.人员性质 = '医生' And d.小组id = [1] And" & vbNewLine & _
                    "   (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
                    "   (Instr(',' || [2] || ',', ',' || C.部门id || ',') > 0 Or a.姓名=[3]) And Instr(',' || [4] || ',', ',' || A.专业技术职务 || ',') > 0" & vbNewLine & _
                    "   And (A.站点=[5] Or A.站点 is Null)" & vbNewLine & _
                    " Order By A.简码"
    strSQL = "Select Distinct A.ID, A.编号, A.简码, A.姓名" & vbNewLine & _
                        " From 人员表 A, 人员性质说明 B, 部门人员 C" & vbNewLine & _
                        " Where A.ID = B.人员id And A.ID = C.人员id And B.人员性质 = '医生' And" & vbNewLine & _
                        "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
                        "      (Instr(',' || [1] || ',', ',' || C.部门id || ',') > 0 Or A.姓名=[2]) And Instr(',' || [3] || ',', ',' || A.专业技术职务 || ',') > 0" & vbNewLine & _
                        "      And (A.站点=[4] Or A.站点 is Null)" & _
                        " Order By A.简码"
    If cbo医疗小组.ListIndex <> -1 And cbo医疗小组.ListIndex <> cbo医疗小组.ListCount - 1 Then
        If Val(cbo医疗小组.ItemData(cbo医疗小组.ListIndex)) > 0 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL医疗小组, Me.Caption, Val(cbo医疗小组.ItemData(cbo医疗小组.ListIndex)), mstrUnit, CStr("" & mrsPatiInfo!住院医师), "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
            If Not rsTmp.RecordCount > 0 Then
                '如果小组未设置医生，则保持以前的科内选择范围
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPatiInfo!住院医师), "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
            End If
            
            If cbo住院医师.ListIndex <> -1 Then
                lng医师 = cbo住院医师.ItemData(cbo住院医师.ListIndex)
            Else
                lng医师 = 0
            End If
            cbo住院医师.Clear
            Do Until rsTmp.EOF
                cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
            '105133:当住院医师在所选医疗小组时不改变住院医师
            If lng医师 <> 0 Then Call cbo.SetIndex(cbo住院医师.hWnd, cbo.FindIndex(cbo住院医师, lng医师))
        
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL医疗小组, Me.Caption, Val(cbo医疗小组.ItemData(cbo医疗小组.ListIndex)), mstrUnit, CStr("" & mrsPatiInfo!主治医师), "主任医师,副主任医师,主治医师", gstrNodeNo)
            
            If Not rsTmp.RecordCount > 0 Then
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPatiInfo!主治医师), "主任医师,副主任医师,主治医师", gstrNodeNo)
            End If
            If cbo主治医师.ListIndex <> -1 Then
                lng医师 = cbo主治医师.ItemData(cbo主治医师.ListIndex)
            Else
                lng医师 = 0
            End If
            cbo主治医师.Clear
            Do Until rsTmp.EOF
                cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
             '105133:当主治医师在所选医疗小组时不改变主治医师
            If lng医师 <> 0 Then Call cbo.SetIndex(cbo主治医师.hWnd, cbo.FindIndex(cbo主治医师, lng医师))
        End If
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPatiInfo!住院医师), "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
        cbo住院医师.Clear
        Do Until rsTmp.EOF
            cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPatiInfo!主治医师), "主任医师,副主任医师,主治医师", gstrNodeNo)
        cbo主治医师.Clear
        Do Until rsTmp.EOF
            cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
    End If
    cbo住院医师.AddItem "其它..."
    cbo主治医师.AddItem "其它..."
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo主治医师_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    On Error GoTo errHandle:
    
    If cbo主治医师.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("医生", "主任医师,副主任医师,主治医师", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo主治医师.ListCount - 1
                If cbo主治医师.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo主治医师.ListIndex = i: Exit Sub
                End If
            Next
            cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo主治医师.ListCount - 1
            cbo主治医师.ListIndex = cbo主治医师.NewIndex
            cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
        Else
            cbo主治医师.ListIndex = -1
        End If
    Else
        '主治医师选择时医疗小组以住院医师优先
        '105133:加载数据完毕之前不应该根据住院医师来调整医疗小组
        If cbo医疗小组.ListCount <= 1 Or Not Me.Visible Then Exit Sub
        strSQL = "Select ID,名称,说明 From 临床医疗小组 A, 医疗小组人员 B " & _
                "Where a.id=b.小组id And b.人员id=[1] And a.科室id=[2] And (撤档时间 Is NULL Or Trunc(撤档时间)=To_Date('3000-01-01','YYYY-MM-DD')) Order By ID"
        If cbo住院医师.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo住院医师.ItemData(cbo住院医师.ListIndex)), Val("" & txt科室.Tag))
            Do While Not rsTmp.EOF
                If cbo医疗小组.Text = Nvl(rsTmp!ID) & "-" & Nvl(rsTmp!名称) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo.FindIndex(cbo医疗小组, Nvl(rsTmp!名称), True))
                Exit Sub
            End If
        End If
        If cbo主治医师.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo主治医师.ItemData(cbo主治医师.ListIndex)), Val("" & txt科室.Tag))
            Do While Not rsTmp.EOF
                If cbo医疗小组.Text = Nvl(rsTmp!ID) & "-" & Nvl(rsTmp!名称) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo.FindIndex(cbo医疗小组, Nvl(rsTmp!名称), True))
            Else
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo医疗小组.ListCount - 1)
            End If
        Else
            Call cbo.SetIndex(cbo医疗小组.hWnd, cbo医疗小组.ListCount - 1)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo住院医师_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    
    On Error GoTo errHandle:
    
    If cbo住院医师.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("医生", "主任医师,副主任医师,主治医师,医师,医士", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo住院医师.ListCount - 1
                If cbo住院医师.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo住院医师.ListIndex = i: Exit Sub
                End If
            Next
            cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo住院医师.ListCount - 1
            cbo住院医师.ListIndex = cbo住院医师.NewIndex
            cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!上级ID
        Else
            cbo住院医师.ListIndex = -1
        End If
    Else
        '105133:加载数据完毕之前不应该根据住院医师来调整医疗小组
        If cbo医疗小组.ListCount <= 1 Or Not Me.Visible Then Exit Sub
        
        strSQL = "Select ID,名称,说明 From 临床医疗小组 A, 医疗小组人员 B " & _
                "Where a.id=b.小组id And b.人员id=[1] And a.科室id=[2] And (撤档时间 Is NULL Or Trunc(撤档时间)=To_Date('3000-01-01','YYYY-MM-DD')) Order By ID"
        If cbo住院医师.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo住院医师.ItemData(cbo住院医师.ListIndex)), Val("" & txt科室.Tag))
            Do While Not rsTmp.EOF
                If cbo医疗小组.Text = Nvl(rsTmp!ID) & "-" & Nvl(rsTmp!名称) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo.FindIndex(cbo医疗小组, Nvl(rsTmp!名称), True))
                Exit Sub
            End If
        End If
        If cbo主治医师.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo主治医师.ItemData(cbo主治医师.ListIndex)), Val("" & txt科室.Tag))
            Do While Not rsTmp.EOF
                If cbo医疗小组.Text = Nvl(rsTmp!ID) & "-" & Nvl(rsTmp!名称) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo.FindIndex(cbo医疗小组, Nvl(rsTmp!名称), True))
            Else
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo医疗小组.ListCount - 1)
            End If
        Else
            Call cbo.SetIndex(cbo医疗小组.hWnd, cbo医疗小组.ListCount - 1)
        End If
    End If
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo住院医师_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo住院医师.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cbo住院医师.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cbo住院医师.ListIndex = lngIdx
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQL医疗小组 As String
    Dim str床号 As String
    Dim i As Integer, lngLevel As Long
    
    
    gblnOK = False
    
    strSQL = "Select NVl(A.姓名,D.姓名) 姓名,NVL(A.性别,D.性别) 性别, A.年龄,To_Char(A.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间,E.名称 as 当前科室,A.出院科室id as 当前科室ID,H.名称 当前病区,A.当前病区Id, A.医疗小组id, g.名称 as 医疗小组, " & vbNewLine & _
            "A.住院号,A.责任护士, A.门诊医师, A.住院医师, B.信息值 主治医师, C.信息值 主任医师, A.费别, A.婚姻状况, A.学历," & vbNewLine & _
            "       A.职业, A.当前病况, A.单位地址, A.单位邮编, A.单位电话, A.家庭地址, A.家庭电话, A.联系人地址," & vbNewLine & _
            "       A.联系人电话, A.联系人姓名, A.联系人关系, A.再入院, A.病人性质, A.险类, D.身份证号, D.区域, D.出生地点," & vbNewLine & _
            "       D.出生日期, A.入院属性,D.合同单位id, F.名称 As 护理等级,Nvl(A.病人类型,Decode(A.险类,Null,'普通病人','医保病人')) 病人类型,A.入院方式,A.备注,A.是否陪伴" & vbNewLine & _
            "From 病案主页 A, 病案主页从表 B, 病案主页从表 C, 病人信息 D,部门表 E,部门表 H,收费项目目录 F, 临床医疗小组 G " & vbNewLine & _
            "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id(+) And A.主页id = B.主页id(+) And A.病人id = C.病人id(+) And" & vbNewLine & _
            "      A.主页id = C.主页id(+) And A.医疗小组id = G.id(+) And B.信息名(+) = '主治医师' And C.信息名(+) = '主任医师' And A.病人id = D.病人id And A.出院科室id = E.id And A.当前病区Id=h.id(+)" & vbNewLine & _
            " And A.护理等级id = F.ID(+)"
    Set mrsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    gstrSQL = "Select 名称 as 医疗小组 From 临床医疗小组 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val("" & mrsPatiInfo!医疗小组id))
    If Not rsTmp.EOF Then
        txtPre.Text = Nvl(rsTmp!医疗小组)
    End If
    Set rsTmp = GetPatiBeds(mlng病人ID)
    If rsTmp.RecordCount = 0 Then
        str床号 = "家庭病床"
    Else
        Do While Not rsTmp.EOF
            str床号 = str床号 & "," & rsTmp!床号
            rsTmp.MoveNext
        Loop
        str床号 = Mid(str床号, 2)
    End If
    txt床号.Text = str床号
    
    With mrsPatiInfo
       txt姓名.Text = !姓名
       txt性别.Text = "" & !性别
       txt年龄.Text = "" & !年龄
       txt住院号.Text = "" & !住院号
       txt科室.Text = "" & !当前科室
       txt科室.Tag = "" & !当前科室id
    End With

    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    mstrUnit = Get科室IDs(mlngUnit) & "," & mlngUnit
    
    '初始化医疗小组
    strSQL = "Select ID,名称,说明,建档时间,撤档时间 From 临床医疗小组 Where 科室id=[1] " & _
            " And (撤档时间 Is NULL Or Trunc(撤档时间) = To_Date('3000-01-01','YYYY-MM-DD')) Order By Id "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & txt科室.Tag))
    
    If Not rsTmp.EOF Then
        cbo医疗小组.Clear
        Do Until rsTmp.EOF
            cbo医疗小组.AddItem rsTmp!ID & "-" & rsTmp!名称
            cbo医疗小组.ItemData(cbo医疗小组.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        cbo医疗小组.AddItem "": cbo医疗小组.ItemData(cbo医疗小组.NewIndex) = 0: cbo医疗小组.ListIndex = -1
    Else
        MsgBox "该科室未设置医疗小组,请先到【临床医疗小组】中设置！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    '初始化住院医师，主治医师
'    strSQL医疗小组 = "Select Distinct A.ID, A.编号, A.简码, A.姓名" & vbNewLine & _
'                    " From 人员表 A, 人员性质说明 B, 部门人员 C, 医疗小组人员 D" & vbNewLine & _
'                    " Where A.ID = B.人员id And A.ID = C.人员id And a.id = d.人员id And B.人员性质 = '医生' And d.小组id = [1] And" & vbNewLine & _
'                    "   (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
'                    "   (Instr(',' || [2] || ',', ',' || C.部门id || ',') > 0 Or a.姓名=[3]) And Instr(',' || [4] || ',', ',' || A.专业技术职务 || ',') > 0" & vbNewLine & _
'                    "   And (A.站点=[5] Or A.站点 is Null)" & vbNewLine & _
'                    " Order By A.简码"
    strSQL = "Select Distinct A.ID, A.编号, A.简码, A.姓名" & vbNewLine & _
                        " From 人员表 A, 人员性质说明 B, 部门人员 C" & vbNewLine & _
                        " Where A.ID = B.人员id And A.ID = C.人员id And B.人员性质 = '医生' And" & vbNewLine & _
                        "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
                        "      (Instr(',' || [1] || ',', ',' || C.部门id || ',') > 0 Or A.姓名=[2]) And Instr(',' || [3] || ',', ',' || A.专业技术职务 || ',') > 0" & vbNewLine & _
                        "      And (A.站点=[4] Or A.站点 is Null)" & _
                        " Order By A.简码"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit & "," & mlngUnit, CStr("" & mrsPatiInfo!住院医师), "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
    cbo住院医师.Clear
    Do Until rsTmp.EOF
        cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit & "," & mlngUnit, CStr("" & mrsPatiInfo!主治医师), "主任医师,副主任医师,主治医师", gstrNodeNo)
    cbo主治医师.Clear
    Do Until rsTmp.EOF
        cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    cbo住院医师.AddItem "其它..."
    cbo主治医师.AddItem "其它..."
    
    '光标定位
    cbo医疗小组.ListIndex = cbo.FindIndex(cbo医疗小组, IIf(IsNull(mrsPatiInfo!医疗小组), "", mrsPatiInfo!医疗小组), True)
    cbo住院医师.ListIndex = cbo.FindIndex(cbo住院医师, IIf(IsNull(mrsPatiInfo!住院医师), "", mrsPatiInfo!住院医师), True)
    cbo主治医师.ListIndex = cbo.FindIndex(cbo主治医师, IIf(IsNull(mrsPatiInfo!主治医师), "", mrsPatiInfo!主治医师), True)
    
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPirvs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '卸载消息对象
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim dMax As Date, strSQL As String
    Dim Curdate As Date
    Dim blnTrue As Boolean
    
    If cbo医疗小组.ListIndex = -1 Then
        MsgBox "请选择新的医疗小组！", vbInformation, gstrSysName
        cbo医疗小组.SetFocus: Exit Sub
    End If
    
    blnTrue = (Val(zlDatabase.GetPara("入住指定医疗小组", glngSys, glngModul, 0)) = 1)
    If cbo医疗小组.ItemData(cbo医疗小组.ListIndex) = 0 And blnTrue = True Then
        MsgBox "由于勾选了参数[入住指定医疗小组],必须选择一个医疗小组，请选择！", vbInformation, gstrSysName
        If cbo医疗小组.Enabled And cbo医疗小组.Visible Then cbo医疗小组.SetFocus
        Exit Sub
    End If
    
    If cbo住院医师.ListIndex = -1 Then
        MsgBox "请选择住院医师！", vbInformation, gstrSysName
        cbo医疗小组.SetFocus: Exit Sub
    End If
    
    If cbo主治医师.ListIndex = -1 Then
        MsgBox "请选择主治医师！", vbInformation, gstrSysName
        cbo医疗小组.SetFocus: Exit Sub
    End If
    
    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入合法的生效时间！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetMaxDate(mlng病人ID, mlng主页ID)
    If CDate(txtDate.Text) <= dMax Then
        MsgBox "生效时间必须大于该病人上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    '时间不能超过当前时间太长(一个月)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 30 Then
            MsgBox "生效时间比当前时间大得过多,请检查！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("生效时间大于了当前系统时间,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
        
    strSQL = "zl_病人变动记录_ChangeGroup(" & mlng病人ID & "," & mlng主页ID & "," & _
            cbo医疗小组.ItemData(cbo医疗小组.ListIndex) & ",'" & zlCommFun.GetNeedName(cbo住院医师.Text) & "','" & zlCommFun.GetNeedName(cbo主治医师.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gblnOK = True
    
    On Error Resume Next
     '住院医师变动后触发消息
    If zlCommFun.GetNeedName(cbo住院医师.Text) <> Nvl(mrsPatiInfo!住院医师) Or zlCommFun.GetNeedName(cbo主治医师.Text) <> Nvl(mrsPatiInfo!主治医师) Then
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '清除缓存中的XML
            '--进行消息组装
            '病人信息
            mclsXML.AppendNode "in_patient"
            'patient_id      病人id  1   N
            mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
            'page_id     主页id  1   N
            mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
            'patient_name        姓名    1   S
            mclsXML.appendData "patient_name", txt姓名.Text, xsString '姓名
            'patient_sex     性别    0..1    S
            mclsXML.appendData "patient_sex", txt性别.Text, xsString '性别
            'in_number       住院号  1   S
            mclsXML.appendData "in_number", txt住院号.Text, xsString  '住院号
            mclsXML.AppendNode "in_patient", True
            
            '当前情况
            'current_state       当前情况    1
            mclsXML.AppendNode "current_state"
            'current_area_id     当前病区id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPatiInfo!当前病区ID)), xsNumber
            'current_area_title      当前病区    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPatiInfo!当前病区), xsString
            'current_dept_id     当前科室id  1   N
            mclsXML.appendData "current_dept_id", Val(txt科室.Tag), xsNumber
            'current_dept_title      当前科室    1   S
            mclsXML.appendData "current_dept_title", txt科室.Text, xsString
            'curren_in_doctor        住院医师    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPatiInfo!住院医师), xsString
            'curren_director_doctor      主任医师    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPatiInfo!主任医师), xsString
            'curren_treat_doctor     主治医师    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPatiInfo!主治医师), xsString
            'curren_duty_nurse       责任护士    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPatiInfo!责任护士), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID 变动id,开始时间 变动时间 From 病人变动记录 Where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] And 开始时间=[4] And NVL(附加床位,0)=0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID, 14, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
            '变更信息
            'change_state        变更信息    1
            mclsXML.AppendNode "change_state"
            'change_id       变更id  1   N
            mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
            'change_date     变更时间    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        住院医师    1   S
            mclsXML.appendData "change_in_doctor", zlCommFun.GetNeedName(cbo住院医师.Text), xsString
            'change_director_doctor      主任医师    1   S
            mclsXML.appendData "change_director_doctor", Nvl(mrsPatiInfo!主任医师), xsString
            'change_treat_doctor     主治医师    1   S
            mclsXML.appendData "change_treat_doctor", zlCommFun.GetNeedName(cbo主治医师.Text), xsString
            'change_duty_nurse       责任护士    1   S
            mclsXML.appendData "change_duty_nurse", Nvl(mrsPatiInfo!责任护士), xsString
            'change_operator         操作员      1   S
            mclsXML.appendData "change_operator", UserInfo.姓名, xsString
            mclsXML.AppendNode "change_state", True
    
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_007", mclsXML.XmlText
        End If
    End If
    
    If Err <> 0 Then Err.Clear
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mstrPirvs = strPrivs
    mlngUnit = lngUnit
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    
    
    Me.Show 1, frmParent
    ShowMe = gblnOK
End Function
