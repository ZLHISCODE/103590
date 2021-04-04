VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPrintList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "批量打印明细帐"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "FrmPrintList.frx":0000
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
            Picture         =   "FrmPrintList.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintList.frx":1E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintList.frx":3B60
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
         Format          =   114491395
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
      Begin VB.TextBox Txt分类 
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
         Format          =   114491395
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
         Caption         =   "分类(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   780
         TabIndex        =   11
         Top             =   2370
         Width           =   630
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
Attribute VB_Name = "FrmPrintList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPath As String
Private mintState As Integer '1=开始打印;2-暂停;3-继续
Private mintPrint As Integer '打印模式
Private mblnStart As Boolean
Private mrs分类 As New ADODB.Recordset
Private mrs材料 As New ADODB.Recordset
Public mstrPrivs As String

Private Sub cmdCancel_Click()
    If CmdCancel.Caption = "退出(&X)" Or CmdCancel.Caption = "取消(&C)" Then
        Unload Me
        Exit Sub
    End If
    mintState = 2
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrHand
    If mrs材料.State = 0 Then
        If Val(Txt分类.Tag) = 0 Then
            MsgBox "请选择一种卫材分类！", vbInformation, gstrSysName
            CmdSelect.SetFocus
            Exit Sub
        End If
        CmdSelect.Enabled = False
        mintPrint = 1
        Call PopupMenu(mnuPrint, 2)
        
        '打开药品记录集
        gstrSQL = "" & _
            " Select '['||d.编码||']'||d.名称 as 名称,A.材料ID " & _
            " From 材料特性 A,诊疗项目目录 c,收费项目目录 d" & _
            " Where a.诊疗id=c.id and a.材料id=d.id And (d.站点=[2] or d.站点 is null) " & _
            "       and c.分类ID in (Select ID From 诊疗分类目录 where 类型=7 start with ID = [1] connect by prior id=上级id)" & _
            " Order by d.编码"
        Set mrs材料 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Txt分类.Tag), gstrNodeNo)
    End If
    
    mintState = 1
    Me.CmdCancel.Caption = "暂停(&A)"
    Me.cmdOk.Enabled = False
    On Error Resume Next
    Avi.AutoPlay = True
    Avi.Play
    err = 0
    On Error GoTo ErrHand
    Do While Not mrs材料.EOF
        DoEvents
        If mintState = 2 Then
            Me.CmdCancel.Caption = "退出(&X)"
            Me.cmdOk.Caption = "继续(&P)"
            Me.cmdOk.Enabled = True
            On Error Resume Next
            Avi.AutoPlay = False
            Avi.Open mstrPath & "\附加文件\打印.avi"
            Exit Sub
        End If
        
        '打印
        '最后附加参数:[0]=正常(含报表及预览),1=直接到预览,2=直接打印,3=输出到Excel
        If cbo库房.Text = "所有库房" Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_2", Me, "材料=" & mrs材料!名称 & "|" & mrs材料!材料ID, "库房=所有库房|is not null", "单位=" & IIf(cbo单位.ListIndex = 0, "散装单位", "包装单位") & "|" & cbo单位.ListIndex, "开始日期=" & Format(Me.dtp开始日期.Value, "yyyy-MM-DD"), "结束日期=" & Format(Me.dtp结束日期.Value, "yyyy-MM-DD"), "包含未审核单据=" & IIf(cbo包含未审核单据.ListIndex = 0, " And 1=1", " And A.审核人 Is Not NULL"), mintPrint)
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_2", Me, "材料=" & mrs材料!名称 & "|" & mrs材料!材料ID, "库房=" & cbo库房.Text & "|=  " & cbo库房.ItemData(cbo库房.ListIndex), "单位=" & IIf(cbo单位.ListIndex = 0, "散装单位", "包装单位") & "|" & cbo单位.ListIndex, "开始日期=" & Format(Me.dtp开始日期.Value, "yyyy-MM-DD"), "结束日期=" & Format(Me.dtp结束日期.Value, "yyyy-MM-DD"), "包含未审核单据=" & IIf(cbo包含未审核单据.ListIndex = 0, " And 1=1", " And A.审核人 Is Not NULL"), mintPrint)
        End If
        
        mrs材料.MoveNext
        DoEvents
    Loop
    
    If mrs材料.EOF Then
        '表示打印结束
        On Error Resume Next
        Avi.AutoPlay = False
        Avi.Open mstrPath & "\附加文件\打印.avi"
        err = 0
        Unload Me
        Exit Sub
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelect_Click()
    With tvw
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab): Exit Sub
    If KeyCode = vbKeyEscape Then mintState = 2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim str材质 As String
    mintState = 0
    mblnStart = False
    mstrPath = gstrAviPath
    Set mrs材料 = New ADODB.Recordset
    
    On Error GoTo ErrHandle
    With cbo单位
        .Clear
        .AddItem "1-散装单位"
        .AddItem "2-包装单位"
        .ListIndex = 0
    End With
    With cbo包含未审核单据
        .Clear
        .AddItem "含未审核单据"
        .AddItem "仅查询已审核单据"
        .ListIndex = 1
    End With
    
    gstrSQL = "Select distinct a.ID,a.编码,a.名称 From 部门表 a,部门性质说明 b,部门性质分类 C " & _
             " Where a.id=b.部门id And b.工作性质=c.名称 And C.名称 In('制剂室','卫材库','发料部门','虚拟库房') and (a.站点=[2] or a.站点 is null) " & _
             IIf(InStr(1, mstrPrivs, "所有库房") <> 0, "", " And A.id In (Select 部门ID From 部门人员 Where 人员ID=[1])") & _
             "   and (to_char(a.撤档时间,'yyyy-mm-dd')='3000-01-01' or a.撤档时间 is null) "
    Set mrs分类 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.Id, gstrNodeNo)
    
    With mrs分类
        If .RecordCount = 0 Then
            MsgBox "请初始化卫材库房或卫材发料部门！[部门管理]", vbInformation, gstrSysName
            Exit Sub
        End If
        
        cbo库房.Clear
        If InStr(1, mstrPrivs, "所有库房") <> 0 Then cbo库房.AddItem "所有库房"
        Do While Not .EOF
            cbo库房.AddItem !名称
            cbo库房.ItemData(cbo库房.NewIndex) = !Id
            .MoveNext
        Loop
        cbo库房.ListIndex = 0
    End With
    
    Me.dtp开始日期.Value = Format(DateAdd("m", -1, Sys.Currentdate), "yyyy年MM月dd日")
    Me.dtp结束日期.Value = Format(Sys.Currentdate, "yyyy年MM月dd日")
    dtp开始日期.MaxDate = Format(Sys.Currentdate, "yyyy年MM月dd日")
    dtp结束日期.MaxDate = Format(Sys.Currentdate, "yyyy年MM月dd日")
    Me.Txt分类 = ""
    Me.Txt分类.Tag = ""
    
    '打开药品用途分类
        gstrSQL = ""
        gstrSQL = "" & _
            "   Select ID,编码,名称,上级ID,1 as 末级" & _
            "   From 诊疗分类目录 " & _
            "   Where 类型=7" & _
            "   Start With 上级ID is null  Connect By Prior ID=上级ID"
    
        zlDatabase.OpenRecordset mrs分类, gstrSQL, Me.Caption
    With mrs分类
        If .RecordCount = 0 Then
            MsgBox "请初始化卫材分类！[卫材目录管理]", vbInformation, gstrSysName
            Exit Sub
        End If
        Call LoadTvw
    End With
    
    On Error Resume Next
    With Avi
        .AutoPlay = False
        .Open mstrPath & "\打印.avi"
    End With
    
    mblnStart = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTvw()
    tvw.Nodes.Clear
    If mrs分类.RecordCount = 0 Then Exit Sub
    
    With mrs分类
        Do While Not .EOF
            If IsNull(!上级ID) Then
                    tvw.Nodes.Add , , "K_" & !Id, "[" & !编码 & "]" & !名称, 2, 2
            Else
                If !末级 = 1 Then
                    tvw.Nodes.Add "K_" & !上级ID, 4, "K_" & !Id, "[" & !编码 & "]" & !名称, 2, 2
                Else
                    tvw.Nodes.Add "K_" & !上级ID, 4, "K_" & !Id, "[" & !编码 & "]" & !名称, 3, 3
                End If
            End If
            tvw.Nodes("K_" & !Id).Tag = !末级
            .MoveNext
        Loop
    End With
    tvw.Nodes(1).Selected = True
End Sub

Private Sub mnuPrintSET_Click(Index As Integer)
    mintPrint = Index
End Sub

Private Sub tvw_DblClick()
    If tvw.SelectedItem.Tag = -1 Then Exit Sub
    Txt分类.Text = tvw.SelectedItem.Text
    Txt分类.Tag = Mid(tvw.SelectedItem.Key, 3)
    cmdOk.SetFocus
End Sub

Private Sub Tvw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call tvw_DblClick
End Sub

Private Sub Tvw_LostFocus()
    tvw.Visible = False
End Sub
