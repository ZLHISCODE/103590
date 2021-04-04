VERSION 5.00
Object = "{5C493D4E-FD57-4FF4-9BA4-C6C670BFF9A7}#70.0#0"; "zl9PacsControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOpenStudyList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打开检查"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpenStudyList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin zl9PacsCapture.TranControl tcFrmQuery 
      Height          =   5085
      Left            =   3405
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   8969
      Begin VB.PictureBox picQuery 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4590
         Left            =   390
         ScaleHeight     =   4560
         ScaleWidth      =   5040
         TabIndex        =   10
         Top             =   120
         Width           =   5070
         Begin VB.CommandButton cmdNotOk 
            Caption         =   "取 消(&Q)"
            Height          =   420
            Left            =   3300
            TabIndex        =   15
            Top             =   3765
            Width           =   1215
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "确 定(&O)"
            Height          =   420
            Left            =   1485
            TabIndex        =   14
            Top             =   3765
            Width           =   1215
         End
         Begin VB.TextBox txtQueryValue 
            Height          =   390
            Left            =   1515
            TabIndex        =   13
            Top             =   2985
            Width           =   3060
         End
         Begin VB.ComboBox cbxQueryType 
            Height          =   330
            ItemData        =   "frmOpenStudyList.frx":000C
            Left            =   1530
            List            =   "frmOpenStudyList.frx":0028
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2505
            Width           =   3045
         End
         Begin VB.TextBox txtName 
            Height          =   390
            Left            =   1530
            TabIndex        =   11
            Top             =   420
            Width           =   3060
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   390
            Left            =   1530
            TabIndex        =   16
            Top             =   1485
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   688
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   115736579
            CurrentDate     =   41535.6989930556
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   390
            Left            =   1530
            TabIndex        =   17
            Top             =   945
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   688
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   115736579
            CurrentDate     =   41534.6979166667
         End
         Begin VB.Label labName 
            Caption         =   "查 询 值"
            Height          =   270
            Index           =   4
            Left            =   435
            TabIndex        =   22
            Top             =   3060
            Width           =   870
         End
         Begin VB.Label labName 
            Caption         =   "查询号别"
            Height          =   270
            Index           =   3
            Left            =   435
            TabIndex        =   21
            Top             =   2565
            Width           =   870
         End
         Begin VB.Label labName 
            Caption         =   "结束日期"
            Height          =   270
            Index           =   2
            Left            =   405
            TabIndex        =   20
            Top             =   1575
            Width           =   870
         End
         Begin VB.Label labName 
            Caption         =   "开始日期"
            Height          =   270
            Index           =   1
            Left            =   405
            TabIndex        =   19
            Top             =   1020
            Width           =   870
         End
         Begin VB.Label labName 
            Caption         =   "姓    名"
            Height          =   270
            Index           =   0
            Left            =   405
            TabIndex        =   18
            Top             =   465
            Width           =   870
         End
      End
   End
   Begin VB.PictureBox picPanel 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   12090
      TabIndex        =   0
      Top             =   6315
      Width           =   12090
      Begin VB.PictureBox picInf 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   45
         ScaleHeight     =   1065
         ScaleWidth      =   6600
         TabIndex        =   5
         Top             =   45
         Visible         =   0   'False
         Width           =   6600
         Begin VB.Label labAdviceInf 
            Height          =   645
            Left            =   1455
            TabIndex        =   8
            Top             =   360
            Width           =   5040
            WordWrap        =   -1  'True
         End
         Begin VB.Label labMoneyState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "收"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   42
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   855
            Left            =   15
            TabIndex        =   7
            Top             =   105
            Width           =   870
         End
         Begin VB.Label labAdviceContext 
            Caption         =   "医嘱内容："
            Height          =   255
            Left            =   1035
            TabIndex        =   6
            Top             =   105
            Width           =   1140
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取 消(&C)"
         Height          =   975
         Left            =   10950
         Picture         =   "frmOpenStudyList.frx":007C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   105
         Width           =   1080
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "确 定(&S)"
         Height          =   975
         Left            =   9885
         Picture         =   "frmOpenStudyList.frx":0570
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   105
         Width           =   1080
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查 找(&F)"
         Height          =   975
         Left            =   8820
         Picture         =   "frmOpenStudyList.frx":15B2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   105
         Width           =   1080
      End
   End
   Begin MSComctlLib.ImageList Imglist 
      Left            =   510
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":1AAB
            Key             =   "住院"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":2385
            Key             =   "病人"
            Object.Tag             =   "2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwStudy 
      Height          =   6240
      Left            =   45
      TabIndex        =   3
      Top             =   15
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11007
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Imglist"
      SmallIcons      =   "Imglist"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "姓名"
         Text            =   "姓名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "性别"
         Text            =   "性别"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "年龄"
         Text            =   "年龄"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "检查号"
         Text            =   "检查号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "标识号"
         Text            =   "标识号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "影像类别"
         Text            =   "影像类别"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "医嘱内容"
         Text            =   "医嘱内容"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "检查状态"
         Text            =   "检查状态"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "病人科室"
         Text            =   "病人科室"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "申请时间"
         Text            =   "申请时间"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "报到时间"
         Text            =   "报到时间"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Key             =   "申请人"
         Text            =   "申请人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Key             =   "报到人"
         Text            =   "报到人"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmOpenStudyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public mlngModule As Long
Public blnOk As Boolean


Public Function ShowStudyWindow(ByRef lngAdviceId As Long, lngSendNo As Long, ByRef lngStudyState As Long, objOwner As Object) As Boolean
'显示检查窗口
    blnOk = False
    
    Me.Show 1, objOwner
    
    If Me.blnOk Then
        lngAdviceId = Nvl(Me.lvwStudy.SelectedItem.Tag)
        lngSendNo = Nvl(Me.lvwStudy.SelectedItem.ListSubItems(1).Tag)
        lngStudyState = Nvl(Me.lvwStudy.SelectedItem.ListSubItems(2).Tag)
    End If
    
    ShowStudyWindow = blnOk
End Function

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    blnOk = False
    Call Me.Hide
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdFind_Click()
On Error GoTo errHandle
    Call ShowQueryWindow
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ShowQueryWindow()
    tcFrmQuery.Left = 0
    tcFrmQuery.Top = 0
    tcFrmQuery.Width = Me.ScaleWidth
    tcFrmQuery.Height = Me.ScaleHeight
    
    picQuery.Left = (tcFrmQuery.Width - picQuery.Width) / 2
    picQuery.Top = (tcFrmQuery.Height - picQuery.Height) / 2
    
    dtpStart.value = Now - 7
    dtpEnd.value = Now
    cbxQueryType.ListIndex = 0
    
    tcFrmQuery.Visible = True
    tcFrmQuery.Translucence
End Sub

Private Sub CloseQueryWindow()
    tcFrmQuery.Visible = False
End Sub

Private Sub cmdNotOk_Click()
On Error GoTo errHandle
    Call CloseQueryWindow
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdOk_Click()
On Error GoTo errHandle
    Dim strFilter As String
    Dim strQueryType As String
    
    If Trim(txtQueryValue.Text) <> "" Then
        Select Case cbxQueryType.Text
            Case "检 查 号"
                strQueryType = "c.检查号"
            Case "门 诊 号"
                strQueryType = "d.门诊号"
            Case "住 院 号"
                strQueryType = "d.住院号"
            Case "体 检 号"
                strQueryType = "d.健康号"
            Case "就诊卡号"
                strQueryType = "d.就诊卡号"
            Case "IC 卡 号"
                strQueryType = "d.IC卡号"
            Case "医 保 号"
                strQueryType = "d.医保号"
        End Select
        
        strFilter = strQueryType & "='" & txtQueryValue.Text & "'"
    Else
        strFilter = " a.姓名 like '" & txtName.Text & "%' and b.发送时间 between " & To_Date(dtpStart.value) & " and " & To_Date(dtpEnd.value)
    End If
    
    Call LoadStudyData(strFilter)
    
    Call CloseQueryWindow
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    If lvwStudy.ListItems.Count <= 0 Then
        Call MsgboxCus("没有可进行采集的检查数据。", vbOKOnly, G_STR_HINT_TITLE)
        Exit Sub
    End If
    
    If lvwStudy.SelectedItem Is Nothing Then
        Call MsgboxCus("请选择需要进行采集的检查数据。", vbOKOnly, gstrSysName)
        Exit Sub
    End If

    blnOk = True
    Call Me.Hide
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
    
    Call zlCL_RestoreWinState(Me, App.ProductName)
    
    Call LoadStudyData

End Sub

Private Function GetColumnIndex(ByVal strColumnCaption As String) As Long
    Dim i As Long
    
    For i = 1 To lvwStudy.ColumnHeaders.Count
        If UCase(lvwStudy.ColumnHeaders(i).Text) = UCase(strColumnCaption) Then Exit For
    Next i
    
    GetColumnIndex = i - 1
End Function

Private Sub LoadStudyData(Optional strFilter As String = "")
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strCurFilter As String
    
    strCurFilter = strFilter
    If Trim(strCurFilter) = "" Then
        strCurFilter = "b.执行过程 in(2,3) and b.首次时间 between sysdate - 3 and sysdate"
    End If
    
    strSQL = "select /*+ Rule*/ a.id,b.发送号, a.姓名, a.性别, a.年龄, a.病人来源, e.名称 as 病人科室, a.医嘱内容, " & _
                    "Decode(a.病人来源,3,a.开嘱医生,b.发送人) 申请人, b.发送时间 as 申请时间, c.报到人, b.首次时间 as 报到时间, " & _
                    "Decode(a.病人来源,2,d.住院号,d.门诊号) 标识号, c.影像类别,c.检查号, nvl(b.执行过程,0) as 检查过程 " & _
            "from 病人医嘱记录 a, 病人医嘱发送 b, 影像检查记录 c, 病人信息 d, 部门表 e " & _
            "where a.ID=b.医嘱id and b.医嘱id=c.医嘱Id(+) and a.病人id=d.病人id and a.病人科室id=e.id and a.相关ID is null and a.执行科室ID=" & glngDepartId & IIf(strCurFilter <> "", " and ", "") & strCurFilter
            
                
    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, Me.Caption)
    
    lvwStudy.ListItems.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    While Not rsData.EOF
        Call SetListItemData(rsData)
        Call rsData.MoveNext
    Wend
    
End Sub

Private Sub SetListItemData(rsCurData As ADODB.Recordset)
    Dim objNewItem As ListItem
    Dim objNewSubItem As ListSubItem
 
    Set objNewItem = lvwStudy.ListItems.Add
    
    objNewItem.Text = Nvl(rsCurData!姓名)
    objNewItem.Icon = 2
    objNewItem.SmallIcon = 2

    objNewItem.Tag = Nvl(rsCurData!ID)

    
    objNewItem.SubItems(GetColumnIndex("性别")) = Nvl(rsCurData!性别)
    objNewItem.SubItems(GetColumnIndex("年龄")) = Nvl(rsCurData!年龄)
    
    objNewItem.SubItems(GetColumnIndex("检查号")) = Nvl(rsCurData!检查号)
    
    objNewItem.SubItems(GetColumnIndex("标识号")) = Nvl(rsCurData!标识号)
    If Nvl(rsCurData!病人来源) = 2 Then
        objNewItem.ListSubItems(GetColumnIndex("标识号")).ReportIcon = 1
    End If
    
    objNewItem.SubItems(GetColumnIndex("影像类别")) = Nvl(rsCurData!影像类别)
    objNewItem.SubItems(GetColumnIndex("医嘱内容")) = Nvl(rsCurData!医嘱内容)
    objNewItem.SubItems(GetColumnIndex("检查状态")) = Decode(Nvl(rsCurData!检查过程), -1, "已驳回", 0, "已登记", 1, "已登记", 2, "已报到", 3, "已检查", 4, "审核中", 5, "已审核", "已完成")
    objNewItem.SubItems(GetColumnIndex("病人科室")) = Nvl(rsCurData!病人科室)
    objNewItem.SubItems(GetColumnIndex("申请时间")) = Format(Nvl(rsCurData!申请时间), "yyyy-mm-dd hh:mm:ss")
    objNewItem.SubItems(GetColumnIndex("报到时间")) = Format(Nvl(rsCurData!报到时间), "yyyy-mm-dd hh:mm:ss")
    objNewItem.SubItems(GetColumnIndex("申请人")) = Nvl(rsCurData!申请人)
    objNewItem.SubItems(GetColumnIndex("报到人")) = Nvl(rsCurData!报到人)

    objNewItem.ListSubItems(1).Tag = Nvl(rsCurData!发送号)
    objNewItem.ListSubItems(2).Tag = Nvl(rsCurData!检查过程)
End Sub


Private Sub Form_Resize()
On Error GoTo errHandle
    lvwStudy.Left = 120
    lvwStudy.Top = 120
    lvwStudy.Height = Me.ScaleHeight - picPanel.Height - 120
    lvwStudy.Width = Me.ScaleWidth - 240

    cmdCancel.Left = picPanel.Width - cmdCancel.Width - 120
    cmdSure.Left = cmdCancel.Left - cmdSure.Width + 10
    cmdFind.Left = cmdSure.Left - cmdFind.Width + 20
    
    Exit Sub
errHandle:
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call zlCL_SaveWinState(Me, App.ProductName)
End Sub


Private Sub lvwStudy_DblClick()
On Error GoTo errHandle
    If lvwStudy.SelectedItem Is Nothing Then Exit Sub
    
    Call cmdSure_Click
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOK, G_STR_HINT_TITLE
End Sub

Private Sub lvwStudy_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo errHandle
    Dim lngCurAdviceId As Long
    
    lngCurAdviceId = Item.Tag
    
    Call ConfigAdviceInf(lngCurAdviceId, Item.ListSubItems(GetColumnIndex("医嘱内容")).Text)
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOK, G_STR_HINT_TITLE
End Sub


Private Sub ConfigAdviceInf(ByVal lngAdviceId As Long, ByVal strAdviceContext As String)
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngMoneyState As Long
    
    strSQL = "select a.id as 医嘱ID, nvl(a.相关Id, 0) as 相关ID, b.计费状态,b.记录性质 " & _
            " from 病人医嘱记录 a, 病人医嘱发送 b " & _
            " where a.Id=b.医嘱ID and (a.id=[1] or a.相关ID=[1])"
    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, Me.Caption, lngAdviceId)
    
    picInf.Visible = False
        
    If rsData.RecordCount <= 0 Then Exit Sub
    
    lngMoneyState = GetMoneyState(rsData)
    Select Case lngMoneyState
        Case 0
            labMoneyState.Caption = "欠"
            labMoneyState.ForeColor = &H80FF&
        Case 1
            labMoneyState.Caption = "收"
            labMoneyState.ForeColor = &HC000&
        Case 2
            labMoneyState.Caption = "无"
            labMoneyState.ForeColor = &HFF0000
        Case 3
    End Select
    
    labAdviceInf.Caption = strAdviceContext
    
    picInf.Visible = True
End Sub



Private Function GetMoneyState(rsData As ADODB.Recordset) As Long
    '判断是否已经收费
    '"病人医嘱发送.记录性质"--- 1是收费的，2是记帐的。

    '通过"病人医嘱发送.计费状态"直接判断,原有值：-1-无须计费;0-未计费;1-已计费，对于记帐单（包括门诊记帐单），保持原有值不变。
    '对于收费单的发送记录，增加两种状态：2-部分收费，3-全部收费

    '没有对应费用的医嘱有两种情况，一种是"-1-无须计费"，即没有设置收费对照，一种是"0-未计费"，即虽然设置了收费对照，但设置为发送后手工计费，即在医技科室去生成。
    '"1-已计费"就是发送时生成了费用的。但生成了费用单据不表示收费了，生成可能是记帐划价单，或收费划价单，其中收费划价单就多两种状态。
    '"2-部分收费"表示部分收费和部分退费的情况，反正没收得完。

    '已收费显示状态：已收费；无费用；未收费：
    '未收费----
    '1、主医嘱是收费单的，满足以下条件算未收费
    '   (1)有一条主医嘱和部位医嘱的 计费状态 in (1,2)算未收费 ------“记录性质=1 and 计费状态 in (1,2)”
    '已收费：
    '1、主医嘱是记账的算收费-------“记录性质=2”
    '2、主医嘱是收费单的，满足以下条件算收费
    '   (1)排除未收费后，有一条主医嘱和部位医嘱的 计费状态 =3 算收费-----“记录性质=1 and 计费状态 = 3”
    '无费用
    '1、主医嘱是收费单的，满足以下条件算无费用
    '   (1)所有主医嘱和部位医嘱的 计费状态 in (-1,0)算无费用 ------“记录性质=1 and 计费状态 in (-1,0)”


    ' intCharged  '0--未收费；1--已收费；2--无费用
    Dim lngTempCharged As Long

    lngTempCharged = 2 '无费用
    
    rsData.Filter = "相关Id = 0"

    If Nvl(rsData!记录性质, 2) = 2 Then
        '住院登记的病人，如果没有计费，则归为无费用
        If Nvl(rsData!计费状态, -1) = 0 Then
            lngTempCharged = 2
        Else
            lngTempCharged = 1  '已收费
        End If
    Else
        If Nvl(rsData!计费状态, -1) = 1 Or Nvl(rsData!计费状态, -1) = 2 Then
            lngTempCharged = 0      '未收费
        Else        '主医嘱的计费状态是 -1,0,3  （3--已收费；-1，0--无费用）
            '查询主医嘱未计费或者已经收费了，还要查部位医嘱的收费情况，所有医嘱都已经收费，才算是收费

            '如果主费用是已收费的，先记录成已收费
            If Nvl(rsData!计费状态, -1) = 3 Then
                lngTempCharged = 1      '已收费
            End If

            rsData.Filter = "相关ID <> 0 "
            Do While rsData.EOF = False
                If Nvl(rsData!计费状态, -1) = 1 Or Nvl(rsData!计费状态, -1) = 2 Then
                    lngTempCharged = 0      '未收费

                    Exit Do
                ElseIf Nvl(rsData!计费状态, -1) = 3 Then
                    lngTempCharged = 1      '已收费
                End If

                rsData.MoveNext
            Loop

        End If
    End If

    GetMoneyState = lngTempCharged
End Function





















