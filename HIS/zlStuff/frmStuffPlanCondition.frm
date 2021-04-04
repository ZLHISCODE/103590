VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStuffPlanCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "条件设置"
   ClientHeight    =   4920
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7560
   Icon            =   "frmStuffPlanCondition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList img16 
      Left            =   6360
      Top             =   1245
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCondition.frx":000C
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCondition.frx":0EE6
            Key             =   "Folder1"
            Object.Tag             =   "Folder1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCondition.frx":1338
            Key             =   "Card"
            Object.Tag             =   "Card"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCondition.frx":178A
            Key             =   "Folder"
            Object.Tag             =   "Folder"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   6360
      TabIndex        =   22
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   21
      Top             =   855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   20
      Top             =   435
      Width           =   1100
   End
   Begin TabDlg.SSTab stb 
      Height          =   4725
      Left            =   45
      TabIndex        =   24
      Top             =   105
      Width           =   6045
      _ExtentX        =   10668
      _ExtentY        =   8340
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "计划(&J)"
      TabPicture(0)   =   "frmStuffPlanCondition.frx":1BDC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl库房"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra辅助条件"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra计划类型"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra计划方法"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbo库房"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Chk仅提取低取下限的材料"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Chk不产生计划数量"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "用途分类(Y)"
      TabPicture(1)   =   "frmStuffPlanCondition.frx":1BF8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tvw用途"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "供应商(&F)"
      TabPicture(2)   =   "frmStuffPlanCondition.frx":1C14
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chk中标单位"
      Tab(2).Control(1)=   "tvw供货单位"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox Chk不产生计划数量 
         Appearance      =   0  'Flat
         Caption         =   "不产生计划数量"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   192
         TabIndex        =   25
         Top             =   600
         Width           =   1560
      End
      Begin VB.CheckBox chk中标单位 
         Caption         =   "无上次供应商以中标单位为准(&W)"
         Enabled         =   0   'False
         Height          =   240
         Left            =   -74520
         TabIndex        =   19
         Top             =   4305
         Width           =   2985
      End
      Begin VB.CheckBox Chk仅提取低取下限的材料 
         Appearance      =   0  'Flat
         Caption         =   "仅提取低于下限的材料(&Q)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   924
         TabIndex        =   16
         Top             =   4245
         Width           =   2895
      End
      Begin VB.ComboBox cbo库房 
         Height          =   276
         Left            =   924
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3795
         Width           =   4848
      End
      Begin VB.Frame fra计划方法 
         Caption         =   "编制方法"
         Height          =   1695
         Left            =   192
         TabIndex        =   4
         Top             =   1875
         Width           =   2640
         Begin VB.OptionButton opt方法 
            Caption         =   "部门申购参照法(&5)"
            Height          =   270
            Index           =   4
            Left            =   240
            TabIndex        =   9
            Top             =   1230
            Width           =   2370
         End
         Begin VB.OptionButton opt方法 
            Caption         =   "卫材日销售量参照法(&4)"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   8
            Top             =   1020
            Width           =   2190
         End
         Begin VB.OptionButton opt方法 
            Caption         =   "材料储备定额参照法(&3)"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   795
            Width           =   2190
         End
         Begin VB.OptionButton opt方法 
            Caption         =   "临近期间平均参照法(&2)"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   570
            Width           =   2190
         End
         Begin VB.OptionButton opt方法 
            Caption         =   "往年同期线性参照法(&1)"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   330
            Value           =   -1  'True
            Width           =   2190
         End
      End
      Begin VB.Frame fra计划类型 
         Caption         =   "计划类型"
         Height          =   765
         Left            =   192
         TabIndex        =   0
         Top             =   960
         Width           =   5580
         Begin VB.OptionButton opt计划 
            Caption         =   "周度计划(&W)"
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Value           =   -1  'True
            Width           =   1296
         End
         Begin VB.OptionButton opt计划 
            Caption         =   "年度计划(&Y)"
            Enabled         =   0   'False
            Height          =   210
            Index           =   2
            Left            =   4116
            TabIndex        =   3
            Top             =   375
            Width           =   1290
         End
         Begin VB.OptionButton opt计划 
            Caption         =   "季度计划(&B)"
            Height          =   210
            Index           =   1
            Left            =   2786
            TabIndex        =   2
            Top             =   375
            Width           =   1290
         End
         Begin VB.OptionButton opt计划 
            Caption         =   "月度计划(&A)"
            Height          =   210
            Index           =   0
            Left            =   1456
            TabIndex        =   1
            Top             =   375
            Width           =   1290
         End
      End
      Begin VB.Frame fra辅助条件 
         Caption         =   "辅助条件"
         Enabled         =   0   'False
         Height          =   1680
         Left            =   3387
         TabIndex        =   23
         Top             =   1875
         Width           =   2385
         Begin VB.TextBox txt下限天数 
            Height          =   300
            Left            =   1185
            TabIndex        =   13
            Top             =   885
            Width           =   900
         End
         Begin VB.TextBox txt上限天数 
            Height          =   300
            Left            =   1185
            TabIndex        =   11
            Top             =   525
            Width           =   900
         End
         Begin VB.Label lbl下限天数 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "下限天数(&T)"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   150
            TabIndex        =   12
            Top             =   945
            Width           =   990
         End
         Begin VB.Label lbl上限天数 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "上限天数(&X)"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   150
            TabIndex        =   10
            Top             =   585
            Width           =   990
         End
      End
      Begin MSComctlLib.TreeView tvw供货单位 
         Height          =   3780
         Left            =   -74925
         TabIndex        =   18
         Top             =   465
         Width           =   5805
         _ExtentX        =   10245
         _ExtentY        =   6668
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvw用途 
         Height          =   4230
         Left            =   -74925
         TabIndex        =   17
         Top             =   420
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   7451
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VB.Label lbl库房 
         AutoSize        =   -1  'True
         Caption         =   "库房(&K)"
         Height          =   180
         Left            =   192
         TabIndex        =   14
         Top             =   3840
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmStuffPlanCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean

Private mstr分类ID As String
Private mstr供货商ID As String
Private mbln中标单位 As Boolean
Private mlng库房id As Long
Private mint计划类型 As Integer
Private mint编制方法 As Integer
Private mbln下限 As Boolean
Private mint上限 As Integer
Private mint下限 As Integer
Private mfrmMain As Form
Private mbln计划数量 As Boolean
Private Const mlngModule = 1724

Public Function GetCondition(frmMain As Form, ByRef str分类ID, ByRef lng库房ID As Long, _
    ByRef int计划类型 As Integer, ByRef int编制方法 As Integer, ByRef bln下限 As Boolean, _
    ByRef int上限 As Integer, ByRef int下限 As Integer, _
    ByRef str供货商ID As String, ByRef bln中标单位 As Boolean, ByRef bln计划数量 As Boolean) As Boolean
    
    mstr分类ID = ""
    mlng库房id = 0
    mint计划类型 = 0
    mint编制方法 = 0
    mblnSelect = False
    
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain
    GetCondition = mblnSelect
    
    bln中标单位 = mbln中标单位
    str分类ID = mstr分类ID
    str供货商ID = mstr供货商ID
    lng库房ID = mlng库房id
    int计划类型 = mint计划类型
    int编制方法 = mint编制方法
    bln下限 = mbln下限
    int上限 = mint上限
    int下限 = mint下限
    bln计划数量 = mbln计划数量
End Function

Private Sub cbo库房_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Chk仅提取低取下限的材料_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        stb.Tab = 1
        If tvw用途.Enabled And tvw用途.Visible Then
            tvw用途.SetFocus
        Else
            OS.PressKey vbKeyTab
        End If
    End If
End Sub

Private Sub chk中标单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnSelect = False
    Hide
    Unload Me
End Sub


Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub
Private Function ISValid() As Boolean
    '验证数据
    ISValid = False
    
    If opt方法(3).Value Then
        '库存上限天数不能小于库存下限天数
        '库存上限天数与库存下限天数不能为零
        If Trim(txt上限天数.Text) = "" Then
            MsgBox "请输入库存上限天数！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Function
        End If
        If Trim(txt下限天数.Text) = "" Then
            MsgBox "请输入库存下限天数！", vbInformation, gstrSysName
            txt下限天数.SetFocus
            Exit Function
        End If
        If Not IsNumeric(txt上限天数.Text) Then
            MsgBox "库存上限天数中含有非法字符！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Function
        End If
        If Not IsNumeric(txt下限天数.Text) Then
            MsgBox "库存下限天数中含有非法字符！", vbInformation, gstrSysName
            txt下限天数.SetFocus
            Exit Function
        End If
        If Val(txt上限天数.Text) <= 0 Then
            MsgBox "库存上限天数不能小于零！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Function
        End If
        If Val(txt下限天数.Text) <= 0 Then
            MsgBox "库存下限天数不能小于零！", vbInformation, gstrSysName
            txt下限天数.SetFocus
            Exit Function
        End If
        If Val(txt上限天数.Text) < Val(txt下限天数.Text) Then
            MsgBox "库存上限天数不能小于库存下限天数！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Function
        End If
        If Val(txt上限天数.Text) > 300 Then
            MsgBox "库存上限天数不能大于300天！", vbInformation, gstrSysName
            txt上限天数.SetFocus
            Exit Function
        End If
    End If
    ISValid = True
    
End Function
Private Sub cmdOK_Click()
    Dim intIndex As Integer
    Dim i As Integer
    Dim Str期间 As String
    Dim intMonth As Integer
    
    If ISValid() = False Then Exit Sub
    
    mstr分类ID = ""
    For i = 1 To tvw用途.Nodes.Count
        If tvw用途.Nodes(i).Key <> "Root" And _
            tvw用途.Nodes(i).Checked Then
            mstr分类ID = mstr分类ID & "," & Mid(tvw用途.Nodes(i).Key, 2)
        End If
    Next
    mstr供货商ID = ""
    For i = 1 To tvw供货单位.Nodes.Count
        If tvw供货单位.Nodes(i).Key <> "Root" And _
            tvw供货单位.Nodes(i).Checked Then
            If tvw供货单位.Nodes(i).Tag = "1" Then
                mstr供货商ID = mstr供货商ID & "," & Mid(tvw供货单位.Nodes(i).Key, 2)
            End If
        End If
    Next
    mint上限 = Val(txt上限天数.Text)
    mint下限 = Val(txt下限天数.Text)
    
    If mstr供货商ID <> "" Then mstr供货商ID = Mid(mstr供货商ID, 2)
    If chk中标单位.Value = 1 And chk中标单位.Enabled Then
        If mstr供货商ID = "" Then
            ShowMsgBox "在没有选择供货单位时，不能选择“无上次供应商以中标单位为准”"
            Me.stb.Tab = 2
            If tvw供货单位.Enabled Then tvw供货单位.SetFocus
            Exit Sub
        End If
    End If
    
    mbln中标单位 = chk中标单位.Value = 1
    
    If mbln中标单位 And mstr供货商ID = "" Then
        ShowMsgBox "未选择供应商,不能设置“无上次供应商以中标单位为准”"
        stb.Tab = 2
        If chk中标单位.Enabled And chk中标单位.Visible Then chk中标单位.SetFocus
        Exit Sub
    End If
    
    If mstr分类ID <> "" Then
        mstr分类ID = Mid(mstr分类ID, 2)
    End If
    
    mlng库房id = cbo库房.ItemData(cbo库房.ListIndex)
    frmStuffPlanCard.LblTitle.Tag = cbo库房.Text
    
    For i = 0 To opt计划.Count - 1
       If opt计划(i).Value Then
           frmStuffPlanCard.txt计划类型.Caption = Mid(opt计划(i).Caption, 1, InStr(1, opt计划(i).Caption, "(") - 1)
           mint计划类型 = i + 1
           Exit For
       End If
    Next
    
    For i = 0 To opt方法.Count - 1
       If opt方法(i).Value Then
           frmStuffPlanCard.txt编制方法.Caption = Mid(opt方法(i).Caption, 1, InStr(1, opt方法(i).Caption, "(") - 1)
           mint编制方法 = i + 1
           Exit For
       End If
    Next
    mbln下限 = (Chk仅提取低取下限的材料.Value = 1)
    mbln计划数量 = (Chk不产生计划数量.Value <> 1)
    
    mblnSelect = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If tvw用途.Visible Then
            tvw用途.Visible = False
        Else
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strReg As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As Node
    Dim strIco As String
    Dim i As Integer
    Dim strSelectStock As String
    
    On Error GoTo errH
    strReg = IIf(Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModule, "0")) = 1, 1, 0)
    strSelectStock = Val(strReg)
    
    stb.Tab = 0
    Call opt方法_Click(0)
    
    With mfrmMain.cboStock
        cbo库房.Clear
        For i = 0 To .ListCount - 1
            cbo库房.AddItem .List(i)
            cbo库房.ItemData(cbo库房.NewIndex) = .ItemData(i)
        Next
        cbo库房.ListIndex = .ListIndex
    End With
    
    If InStr(1, gstrPrivs, "所有库房") <> 0 Then
        If strSelectStock = "0" Then
            cbo库房.Enabled = False
        Else
            cbo库房.Enabled = True
        End If
    Else
        cbo库房.Enabled = False
    End If
        

    gstrSQL = "" & _
        "   Select Level as 层,ID,上级ID,名称 " & _
        "   From 诊疗分类目录" & _
        "   where 类型=7 " & _
        "   Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        "   Order by Level"
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    Set objNode = tvw用途.Nodes.Add(, , "Root", "所有材料分类", "Item")
    
    Do While Not rsTemp.EOF
        If rsTemp!层 = 1 Then
            Set objNode = tvw用途.Nodes.Add("Root", 4, "_" & rsTemp!Id, rsTemp!名称, "Item")
        Else
            Set objNode = tvw用途.Nodes.Add("_" & rsTemp!上级ID, 4, "_" & rsTemp!Id, rsTemp!名称, "Item")
        End If
        rsTemp.MoveNext
    Loop
    tvw用途.Nodes("Root").Selected = True
    tvw用途.Nodes("Root").Expanded = True
    
    gstrSQL = "" & _
        "   Select Level as 层,ID,上级ID,编码||'-'||名称 名称,末级 " & _
        "   From 供应商" & _
        "   where (substr(类型,5,1)=1 and (站点=[1] or 站点 is null) Or Nvl(末级,0)=0) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null)" & _
        "   Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        "   Order by Level"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrNodeNo)
    
    tvw供货单位.Nodes.Clear
    Set objNode = tvw供货单位.Nodes.Add(, , "Root", "所有卫材供货商", "Folder")
    objNode.Sorted = True
    Do While Not rsTemp.EOF
        strIco = IIf(Val(NVL(rsTemp!末级)) = 1, "Card", "Folder")
        If rsTemp!层 = 1 Then
            Set objNode = tvw供货单位.Nodes.Add("Root", 4, "_" & rsTemp!Id, rsTemp!名称, strIco)
        Else
            Set objNode = tvw供货单位.Nodes.Add("_" & rsTemp!上级ID, 4, "_" & rsTemp!Id, rsTemp!名称, strIco)
        End If
        If strIco = "Card" Then
            objNode.Tag = "1"
        End If
        objNode.Sorted = True
        rsTemp.MoveNext
    Loop
    tvw供货单位.Nodes("Root").Selected = True
    tvw供货单位.Nodes("Root").Expanded = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub
Private Sub opt方法_Click(Index As Integer)
    fra辅助条件.Enabled = False
    Chk仅提取低取下限的材料.Enabled = True
    tvw供货单位.Enabled = True
    chk中标单位.Enabled = True
    If Index = 0 Then
        If opt计划(2).Value = True Then
            opt计划(2).Value = False
            opt计划(0).Value = True
        End If
        If opt计划(3).Value = True Then
            opt计划(3).Value = False
            opt计划(0).Value = True
        End If
        opt计划(2).Enabled = False
        opt计划(3).Enabled = False
    ElseIf Index = 3 Then
        fra辅助条件.Enabled = True
'        opt计划(0).Value = True
'        opt计划(1).Value = False
'        opt计划(2).Value = False
        opt计划(2).Enabled = False
        opt计划(3).Enabled = False
        If opt计划(2).Value = True Then
            opt计划(2).Value = False
            opt计划(0).Value = True
        End If
        If opt计划(1).Value = True Then
            opt计划(1).Value = False
            opt计划(0).Value = True
        End If
        If opt计划(3).Value = True Then
            opt计划(3).Value = False
            opt计划(0).Value = True
        End If
        
    ElseIf Index = 4 Then
        '根据部门申购编制计划
        fra辅助条件.Enabled = False
        Chk仅提取低取下限的材料.Enabled = False
        tvw供货单位.Enabled = False
        chk中标单位.Enabled = False
        opt计划(2).Enabled = True
        opt计划(3).Enabled = False
        If opt计划(3).Value = True Then
            opt计划(3).Value = False
            opt计划(0).Value = True
        End If
    Else
        opt计划(2).Enabled = True
        opt计划(3).Enabled = True
    End If
End Sub

Private Sub opt方法_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub opt计划_Click(Index As Integer)
    If opt计划(0).Value = False Then
        If opt方法(3).Value Then
            opt方法(0).Value = True
        End If
    End If
End Sub

Private Sub opt计划_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub



Private Sub stb_Click(PreviousTab As Integer)
    Select Case stb.Tab
    Case 1
       If tvw用途.Visible And tvw用途.Enabled Then tvw用途.SetFocus
    Case 2
       If tvw供货单位.Visible And tvw供货单位.Enabled Then tvw供货单位.SetFocus
    End Select
    
End Sub
 
Private Sub tvw供货单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub tvw供货单位_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked, False
    chk中标单位.Enabled = IsHavingCheck(Node)
End Sub
Private Function IsHavingCheck(ByVal objNode As Node) As Boolean
    '功能:检查是否存在Node被选择了的
    Dim objNode1 As Node
    If Not objNode Is Nothing Then
        If objNode.Checked = True Then IsHavingCheck = True: Exit Function
    End If
    For Each objNode1 In tvw供货单位.Nodes
        If objNode1.Checked Then IsHavingCheck = True: Exit Function
    Next
End Function
Private Sub tvw用途_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        stb.Tab = 2
        If tvw供货单位.Enabled And tvw供货单位.Visible Then
            tvw供货单位.SetFocus
        Else
            OS.PressKey vbKeyTab
        End If

    End If
End Sub

Private Sub tvw用途_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked
End Sub

Private Sub SetParentNode(ByVal Node As MSComctlLib.Node, blnCheck As Boolean, Optional blnTvw用途 As Boolean = True)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If blnTvw用途 = True Then
                    If tvw用途.Nodes(intIdx).Checked = False Then
                        Node.Parent.Checked = False
                        Exit Do
                    End If
                    intIdx = tvw用途.Nodes(intIdx).Next.Index
                Else
                    If tvw供货单位.Nodes(intIdx).Checked = False Then
                        Node.Parent.Checked = False
                        Exit Do
                    End If
                    intIdx = tvw供货单位.Nodes(intIdx).Next.Index
                End If
            Loop
            If intIdx = Node.LastSibling.Index Then
                If blnTvw用途 = True Then
                       If tvw供货单位.Nodes(intIdx).Checked = True Then
                           Node.Parent.Checked = True
                       End If
                Else
                       If tvw供货单位.Nodes(intIdx).Checked = True Then
                           Node.Parent.Checked = True
                       End If
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode Node, blnCheck, blnTvw用途
        End If
    End If
End Sub


Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Function CheckCount() As Integer
    Dim i As Integer
    For i = 1 To tvw用途.Nodes.Count
        If tvw用途.Nodes(i).Checked Then CheckCount = CheckCount + 1
    Next
End Function
Private Sub txt上限天数_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txt上限天数_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt上限天数, KeyAscii, m数字式
End Sub
Private Sub txt下限天数_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txt下限天数_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt下限天数, KeyAscii, m数字式
End Sub
