VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm批量打印明细帐 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "批量打印明细帐"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "Frm批量打印明细帐.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.TreeView Tvw 
      Height          =   2505
      Left            =   1620
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4419
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgTree"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm批量打印明细帐.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm批量打印明细帐.frx":1E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm批量打印明细帐.frx":3B60
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.Animation Avi 
      Height          =   1005
      Left            =   4890
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1773
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   85
      FullHeight      =   67
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   4980
      TabIndex        =   16
      Top             =   2610
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4980
      TabIndex        =   15
      Top             =   870
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   4980
      TabIndex        =   14
      Top             =   390
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置"
      Height          =   2775
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4665
      Begin VB.ComboBox cbo包含未审核单据 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1920
         Width           =   1755
      End
      Begin MSComCtl2.DTPicker Dtp开始日期 
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   1140
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   275709955
         CurrentDate     =   37648
      End
      Begin VB.ComboBox Cbo库房 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1755
      End
      Begin VB.ComboBox Cbo单位 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   750
         Width           =   1755
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "…"
         Height          =   300
         Left            =   3810
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2310
         Width           =   300
      End
      Begin VB.TextBox Txt用途分类 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1440
         TabIndex        =   12
         Top             =   2310
         Width           =   2385
      End
      Begin MSComCtl2.DTPicker Dtp结束日期 
         Height          =   300
         Left            =   1440
         TabIndex        =   8
         Top             =   1530
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   275709955
         CurrentDate     =   37648
      End
      Begin VB.Label lbl包含未审核单据 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "含未审单据(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   1980
         Width           =   1170
      End
      Begin VB.Label Lbl结束日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   420
         TabIndex        =   7
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label Lbl开始日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   420
         TabIndex        =   5
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label lbl6库房 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "库房(&K)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   780
         TabIndex        =   1
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Lbl单位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   780
         TabIndex        =   3
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Lbl用途分类 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "用途分类(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   420
         TabIndex        =   11
         Top             =   2370
         Width           =   990
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "打印(&P)"
      Visible         =   0   'False
      Begin VB.Menu mnuPrintSET 
         Caption         =   "预览(&V)"
         Index           =   1
      End
      Begin VB.Menu mnuPrintSET 
         Caption         =   "打印(&P)"
         Index           =   2
      End
      Begin VB.Menu mnuPrintSET 
         Caption         =   "输出到&Excel"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Frm批量打印明细帐"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPath As String
Private intState As Integer '1=开始打印;2-暂停;3-继续
Private intPrint As Integer '打印模式
Private blnStart As Boolean
Private rs用途分类 As New ADODB.Recordset
Private rs药品 As New ADODB.Recordset

Private Sub cmdCancel_Click()
    If cmdCancel.Caption = "退出(&X)" Or cmdCancel.Caption = "取消(&C)" Then
        Unload Me
        Exit Sub
    End If
    intState = 2
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHand
    If rs药品.State = 0 Then
        If Val(Txt用途分类.Tag) = 0 Then
            MsgBox "请选择一种药品用途分类！", vbInformation, gstrSysName
            CmdSelect.SetFocus
            Exit Sub
        End If
        CmdSelect.Enabled = False
        intPrint = 1
        Call PopupMenu(mnuPrint, 2)
        
        '打开药品记录集
        gstrSQL = "Select '['||F.编码||']'||F.名称 名称,A.药品ID  " & _
                 " From 药品规格 A,收费项目目录 F, " & _
                 "  (Select ID 药名ID  " & _
                 "  From 诊疗项目目录   " & _
                 "  Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And 分类ID IN  " & _
                 "      (Select ID From 诊疗分类目录 " & _
                 "      where 类型 in (1,2,3) " & _
                 "      Start With ID=[1] Connect By Prior ID=上级ID)) B " & _
                 " Where (F.站点 = [2] Or f.站点 is Null) And A.药名ID=B.药名ID And A.药品ID=F.ID " & _
                 " Order by F.编码"
        Set rs药品 = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Txt用途分类.Tag), gstrNodeNo)
    End If
    
    intState = 1
    Me.cmdCancel.Caption = "暂停(&A)"
    Me.cmdOK.Enabled = False
    On Error Resume Next
    Avi.AutoPlay = True
    Avi.Play
    Err = 0
    On Error GoTo ErrHand
    Do While Not rs药品.EOF
        DoEvents
        If intState = 2 Then
            Me.cmdCancel.Caption = "退出(&X)"
            Me.cmdOK.Caption = "继续(&P)"
            Me.cmdOK.Enabled = True
            On Error Resume Next
            Avi.AutoPlay = False
            Avi.Open strPath & "\附加文件\打印.avi"
            Exit Sub
        End If
        
        '打印
        '最后附加参数:[0]=正常(含报表及预览),1=直接到预览,2=直接打印,3=输出到Excel
        If cbo库房.Text = "所有库房" Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1309_2", "ZL8_INSIDE_1309_2"), Me, "药品=" & rs药品!名称 & "|" & rs药品!药品id, "库房=所有库房|is not null", "单位=" & Choose(cbo单位.ListIndex + 1, "售价单位", "门诊单位", "药库单位", "住院单位") & "|" & Choose(cbo单位.ListIndex + 1, 1, 3, 2, 4), "开始日期=" & Format(Me.dtp开始日期.Value, "yyyy-MM-DD"), "结束日期=" & Format(Me.dtp结束日期.Value, "yyyy-MM-DD"), "包含未审核单据=" & IIf(cbo包含未审核单据.ListIndex = 0, " And 1=1", " And A.审核人 Is Not NULL"), intPrint)
        Else
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1309_2", "ZL8_INSIDE_1309_2"), Me, "药品=" & rs药品!名称 & "|" & rs药品!药品id, "库房=" & cbo库房.Text & "|=  " & cbo库房.ItemData(cbo库房.ListIndex), "单位=" & Choose(cbo单位.ListIndex + 1, "售价单位", "门诊单位", "药库单位", "住院单位") & "|" & Choose(cbo单位.ListIndex + 1, 1, 3, 2, 4), "开始日期=" & Format(Me.dtp开始日期.Value, "yyyy-MM-DD"), "结束日期=" & Format(Me.dtp结束日期.Value, "yyyy-MM-DD"), "包含未审核单据=" & IIf(cbo包含未审核单据.ListIndex = 0, " And 1=1", " And A.审核人 Is Not NULL"), intPrint)
        End If
        
        rs药品.MoveNext
        DoEvents
    Loop
    
    If rs药品.EOF Then
        '表示打印结束
        On Error Resume Next
        Avi.AutoPlay = False
        Avi.Open strPath & "\附加文件\打印.avi"
        Err = 0
        Unload Me
        Exit Sub
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub CmdSelect_Click()
    With Tvw
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    If Not blnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab): Exit Sub
    If KeyCode = vbKeyEscape Then intState = 2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim Str材质 As String
    intState = 0
    blnStart = False
    strPath = gstrAviPath
    Set rs药品 = New ADODB.Recordset
    
    On Error GoTo errHandle
    With cbo单位
        .Clear
        .AddItem "1-售价单位"
        .AddItem "2-门诊单位"
        .AddItem "3-药库单位"
        .AddItem "4-住院单位"
        .ListIndex = 0
    End With
    With cbo包含未审核单据
        .Clear
        .AddItem "含未审核单据"
        .AddItem "仅查询已审核单据"
        .ListIndex = 1
    End With
    
    gstrSQL = "Select distinct a.ID,a.编码,a.名称 From 部门表 a,部门性质说明 b,部门性质分类 C " & _
              "Where (a.站点 = [2] Or a.站点 is Null) And a.id=b.部门id And b.工作性质=c.名称 And Instr('HIJKLMN',c.编码,1)>0 " & _
              IIf(zlStr.IsHavePrivs(gstrStockSearchPrivs, "所有库房"), "", " And A.id In (Select 部门ID From 部门人员 Where 人员ID=[1]) ") & _
              "  and (to_char(a.撤档时间,'yyyy-mm-dd')='3000-01-01' or a.撤档时间 is null) "
    Set rs用途分类 = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.用户ID, gstrNodeNo)
       
    With rs用途分类
        If .RecordCount = 0 Then
            MsgBox "请初始化药品库房！[部门管理]", vbInformation, gstrSysName
            Exit Sub
        End If
        
        cbo库房.Clear
        If zlStr.IsHavePrivs(gstrStockSearchPrivs, "所有库房") Then cbo库房.AddItem "所有库房"
        Do While Not .EOF
            cbo库房.AddItem !名称
            cbo库房.ItemData(cbo库房.NewIndex) = !id
            .MoveNext
        Loop
        cbo库房.ListIndex = 0
    End With
    
    Me.dtp开始日期.Value = Format(DateAdd("m", -1, Sys.Currentdate), "yyyy年MM月dd日")
    Me.dtp结束日期.Value = Format(Sys.Currentdate, "yyyy年MM月dd日")
    dtp开始日期.MaxDate = Format(Sys.Currentdate, "yyyy年MM月dd日")
    dtp结束日期.MaxDate = Format(Sys.Currentdate, "yyyy年MM月dd日")
    Me.Txt用途分类 = ""
    Me.Txt用途分类.Tag = ""
    
    '打开药品用途分类
    gstrSQL = "": Str材质 = ""
    If zlStr.IsHavePrivs(gstrStockSearchPrivs, "西成药") Then
        Str材质 = Str材质 & ",'西成药'"
        gstrSQL = gstrSQL & _
                  " Select -1 id,'1' 编码,'西成药' 名称,to_number(NULL,0) 上级ID,-1 末级 from dual" & _
                  " Union all"
    End If
    If zlStr.IsHavePrivs(gstrStockSearchPrivs, "中成药") Then
        Str材质 = Str材质 & ",'中成药'"
        gstrSQL = gstrSQL & _
                " Select -2 id,'3' 编码,'中成药' 名称,to_number(NULL,0) 上级ID,-1 末级 from dual" & _
                " Union all"
    End If
    If zlStr.IsHavePrivs(gstrStockSearchPrivs, "中草药") Then
        Str材质 = Str材质 & ",'中草药'"
        gstrSQL = gstrSQL & _
                " Select -3 id,'2' 编码,'中草药' 名称,to_number(NULL,0) 上级ID,-1 末级 from dual" & _
                " Union all"
    End If
    If Str材质 = "" Then
        MsgBox "你没有权限使用批量打印明细帐！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Str材质 = Mid(Str材质, 2)
    gstrSQL = gstrSQL & _
        " Select ID,编码,名称,Decode(上级ID,null,-1*类型,Nvl(上级ID,0)) 上级ID,1" & _
        " From 诊疗分类目录 A" & _
        " Where A.类型 In (1,2,3)" & _
        " Start With Nvl(上级ID,0)=0" & _
        " Connect By Prior ID=上级ID"
    Set rs用途分类 = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-药品用途分类")
    
    With rs用途分类
        If .RecordCount = 0 Then
            MsgBox "请初始化药品用途分类！[药品用途分类]", vbInformation, gstrSysName
            Exit Sub
        End If
        Call LoadTvw
    End With
    
    On Error Resume Next
    With Avi
        .AutoPlay = False
        .Open strPath & "\打印.avi"
    End With
    
    blnStart = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTvw()
    Tvw.Nodes.Clear
    If rs用途分类.RecordCount = 0 Then Exit Sub
    
    With rs用途分类
        Do While Not .EOF
            If IsNull(!上级ID) Then
                If !末级 = 1 Then
                    Tvw.Nodes.Add , , "K_" & !id, "[" & !编码 & "]" & !名称, 2, 2
                Else
                    Tvw.Nodes.Add , , "K_" & !id, "[" & !编码 & "]" & !名称, 3, 3
                End If
            Else
                If !末级 = 1 Then
                    Tvw.Nodes.Add "K_" & !上级ID, 4, "K_" & !id, "[" & !编码 & "]" & !名称, 2, 2
                Else
                    Tvw.Nodes.Add "K_" & !上级ID, 4, "K_" & !id, "[" & !编码 & "]" & !名称, 3, 3
                End If
            End If
            Tvw.Nodes("K_" & !id).Tag = !末级
            .MoveNext
        Loop
    End With
    Tvw.Nodes(1).Selected = True
End Sub

Private Sub mnuPrintSET_Click(Index As Integer)
    intPrint = Index
End Sub

Private Sub Tvw_DblClick()
    If Tvw.SelectedItem.Tag = -1 Then Exit Sub
    Txt用途分类.Text = Tvw.SelectedItem.Text
    Txt用途分类.Tag = Mid(Tvw.SelectedItem.Key, 3)
    cmdOK.SetFocus
End Sub

Private Sub Tvw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call Tvw_DblClick
End Sub

Private Sub Tvw_LostFocus()
    Tvw.Visible = False
End Sub


