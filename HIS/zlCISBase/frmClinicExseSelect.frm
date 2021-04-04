VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicExseSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择器"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   Icon            =   "frmClinicExseSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2700
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleMode       =   0  'User
      ScaleWidth      =   33.75
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   45
   End
   Begin VB.CheckBox ChkDown 
      Caption         =   "显示下级目录内容"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   2025
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   8160
      TabIndex        =   3
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   300
      Left            =   6840
      TabIndex        =   2
      Top             =   4440
      Width           =   1100
   End
   Begin MSComctlLib.ListView LivMain 
      Height          =   4035
      Left            =   3360
      TabIndex        =   0
      Top             =   330
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   7117
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "规格"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "产地"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "记算单位"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "售价"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.TreeView LvwMain 
      Height          =   4035
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   7117
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2700
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicExseSelect.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择一个项目，然后点击确定"
      Height          =   180
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   2520
   End
End
Attribute VB_Name = "frmClinicExseSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MouseStartX As Single                       '移动前鼠标的位置
Dim NowIndex As Long                            '当前位置
Private mstr服务对象 As String

Sub LoadTree()
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select level as 级数,8 类型,Id, 上级id, 编码, 名称" & _
             " From 收费分类目录 " & _
             " Start With 上级id Is Null " & _
             " Connect By Prior Id = 上级id " & _
             " Union " & _
             " Select level as 级数,类型,Id, 上级id, 编码, 名称 " & _
             " From 诊疗分类目录 " & _
             " Where 类型 in(1,2,3,7) " & _
             " Start With 上级id Is Null " & _
             " Connect By Prior Id = 上级id " & _
             " Order By 级数,类型,编码 "

    zlDatabase.OpenRecordset rsTmp, gstrSql, "选择器"
    
    '根结点
    LvwMain.Nodes.Clear
    LvwMain.Nodes.Add , , "Root", "所有收费项目", 1, 1
    LvwMain.Nodes.Add "Root", tvwChild, "C1", "[1]西成药", 1, 1
    LvwMain.Nodes.Add "Root", tvwChild, "C2", "[2]中成费", 1, 1
    LvwMain.Nodes.Add "Root", tvwChild, "C3", "[3]中草药", 1, 1
    LvwMain.Nodes.Add "Root", tvwChild, "C7", "[7]卫生材料", 1, 1
    
    With rsTmp
        Do While Not .EOF
            If IsNull(!上级ID) Then
                If !类型 <> 8 Then
                    LvwMain.Nodes.Add "C" & Val(!类型), tvwChild, "C" & Val(!类型) & Val(!ID), "[" & !编码 & "]" & !名称, 1, 1
                Else
                    LvwMain.Nodes.Add "Root", tvwChild, "C" & Val(!类型) & Val(!ID), "[" & !编码 & "]" & !名称, 1, 1
                End If
            Else
                LvwMain.Nodes.Add "C" & Val(!类型) & Val(!上级ID), tvwChild, "C" & Val(!类型) & Val(!ID), "[" & !编码 & "]" & !名称, 1, 1
            End If
            .MoveNext
        Loop
    End With

    rsTmp.Close
    
    Dim nod As Node
    On Error Resume Next
    Set nod = LvwMain.Nodes(strKey)
    If Err <> 0 Then
        Set nod = LvwMain.Nodes("Root").Child
        nod.Selected = True
        nod.Expanded = True
        LvwMain_NodeClick nod
        NowIndex = nod.Index
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ChkDown_Click()
    Dim nod As Node
    Set nod = Me.LvwMain.Nodes(Me.LvwMain.SelectedItem.Index)
    LvwMain_NodeClick nod
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHandle
    If Me.LivMain.ListItems.Count > 0 Then
        gstrSql = "select * from 收费项目目录 where id = " & Mid(Me.LivMain.SelectedItem.Key, 2)
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Me.LivMain.SelectedItem.Key, 2)))
        Set frmClinicExse.rsSelect = rsTmp
    End If
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    LoadTree
End Sub

Private Sub Form_Resize()
    
    'LvwMain
    Me.LvwMain.Width = Me.picSplit.Left
    
    'picSplit
    Me.picSplit.Top = Me.LvwMain.Top
    Me.picSplit.Height = Me.LvwMain.Height
    
    'Livmain
    Me.LivMain.Top = Me.LvwMain.Top
    Me.LivMain.Left = Me.picSplit.Left + Me.picSplit.Width
    Me.LivMain.Height = Me.LvwMain.Height
    Me.LivMain.Width = Me.ScaleWidth - Me.picSplit.Left - Me.picSplit.Width
    
End Sub
Private Sub LivMain_DblClick()
    cmdOK_Click
End Sub

Private Sub LvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    Dim str类别 As String
    Dim str分类 As String
    Dim strTemp As String
    Dim i As Integer
    
    If Node.Key = "Root" Then Exit Sub
    On Error GoTo errHandle
    str类别 = Mid(Node.Key, 2, 1)
    str分类 = Mid(Node.Key, 3)
    
    
    If str类别 <> "8" Then
        For i = 0 To UBound(Split(mstr服务对象, ","))
            strTemp = IIf(strTemp = "", "(A.服务对象=" & Split(mstr服务对象, ",")(i), strTemp & " or A.服务对象=" & Split(mstr服务对象, ",")(i))
        Next
        strTemp = strTemp & ")"
        If Len(Node.Key) <= 2 Then Exit Sub
        If Me.ChkDown.Value = 0 Then
            gstrSql = "Select A.Id,C.分类id,A.编码, A.名称, A.规格,  A.产地, A.计算单位," & _
                    " Decode(Nvl(A.是否变价,0),0,ltrim(rtrim(to_char(nvl(D.现价,0),'9999999990.0000'))),'时价') As 售价 " & _
                    " From 收费项目目录 A," & IIf(str类别 <> "7", "药品规格", "材料特性") & " B,诊疗项目目录 C,收费价目 D " & _
                    " Where A.ID=D.收费细目ID(+) And a.ID = b." & IIf(str类别 <> "7", "药品id", "材料id") & " And b." & IIf(str类别 <> "7", "药名id", "诊疗id") & " = C.ID And C.分类id = [1] " & _
                    " and (A.撤档时间 is null or A.撤档时间 = to_date('3000-01-01', 'YYYY-MM-DD')) " & _
                    " And D.执行日期 <= SYSDATE AND (D.终止日期 > SYSDATE OR D.终止日期 IS NULL)  And D.价格等级 Is Null " & _
                    " And " & strTemp & _
                    " Order By A.编码"
        Else
            gstrSql = "Select a.Id, c.分类id, a.编码, A.名称, A.规格, A.产地, a.计算单位, " & _
                    " Decode(Nvl(A.是否变价,0),0,ltrim(rtrim(to_char(nvl(e.现价,0),'9999999990.0000'))),'时价') As 售价 " & _
                    " From 收费项目目录 a," & IIf(str类别 <> "7", "药品规格", "材料特性") & " B,诊疗项目目录 c,收费价目 e, " & _
                    " (Select * From 诊疗分类目录 " & _
                    " Start With 上级id = [1] " & _
                    " Connect By Prior Id = 上级id " & _
                    " Union " & _
                    " Select * From 诊疗分类目录 Where Id = [1]) d " & _
                    " Where A.ID=e.收费细目ID(+) And a.ID = b." & IIf(str类别 <> "7", "药品id", "材料id") & " And b." & IIf(str类别 <> "7", "药名id", "诊疗id") & " = c.ID And c.分类id = d.ID " & _
                    " and (A.撤档时间 is null or A.撤档时间 = to_date('3000-01-01', 'YYYY-MM-DD')) " & _
                    " And e.执行日期 <= SYSDATE AND (e.终止日期 > SYSDATE OR e.终止日期 IS NULL)   And e.价格等级 Is Null " & _
                    " And " & strTemp & _
                    " Order By a.编码 "
        End If
        

    Else
        For i = 0 To UBound(Split(mstr服务对象, ","))
            strTemp = IIf(strTemp = "", "(I.服务对象=" & Split(mstr服务对象, ",")(i), strTemp & " or I.服务对象=" & Split(mstr服务对象, ",")(i))
        Next
        strTemp = strTemp & ")"
        If Me.ChkDown.Value = 0 Then
            gstrSql = " Select 1 As 末级,I.ID,分类ID,I.编码,I.名称,I.规格,I.产地, I.计算单位," & _
                    " Decode(Nvl(I.是否变价,0),0,ltrim(rtrim(to_char(Sum(nvl(D.现价,0)),'9999999990.0000'))),ltrim(rtrim(to_char(Sum(nvl(D.缺省价格,0)),'9999999990.0000')))) As 售价 " & _
                    " from 收费项目目录 I,收费价目 D " & _
                    " where I.ID=D.收费细目ID(+) And I.类别 not in ('1','J')" & _
                    " and 分类ID = [1] " & _
                    " and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
                    " And D.执行日期 <= SYSDATE AND (D.终止日期 > SYSDATE OR D.终止日期 IS NULL)  And D.价格等级 Is Null " & _
                    " And " & strTemp & _
                    " Group By i.Id, 分类id, i.编码, i.名称, i.规格, i.产地, i.计算单位, i.是否变价 " & _
                    " Order By 编码 "
        Else
            gstrSql = "select b.id , b.编码 , b.名称 ,b.规格,b.产地, b.计算单位,b.售价 " & _
                    " From " & _
                    "     (select * from 收费分类目录 " & _
                    "        start with 上级id = [1] " & _
                    "        connect by prior id = 上级id " & _
                    "      Union " & _
                    "      select * from 收费分类目录 " & _
                    "        where id = [1] )  a , " & _
                    " (Select 1 As 末级,I.ID,分类ID,I.编码,I.名称,I.规格,I.产地, I.计算单位, " & _
                    " Decode(Nvl(I.是否变价,0),0,ltrim(rtrim(to_char(Sum(nvl(D.现价,0)),'9999999990.0000'))),ltrim(rtrim(to_char(Sum(nvl(D.缺省价格,0)),'9999999990.0000')))) As 售价 " & _
                    " from 收费项目目录 I,收费价目 D " & _
                    " Where I.ID=D.收费细目ID(+) And " & _
                    " I.类别 not in ('1','J') and " & _
                    " (I.撤档时间 is null or I.撤档时间 = to_date('3000-01-01', 'YYYY-MM-DD')) And D.执行日期 <= SYSDATE AND (D.终止日期 > SYSDATE OR D.终止日期 IS NULL)  And D.价格等级 Is Null  " & _
                    " And " & strTemp & _
                    " Group By i.Id, 分类id, i.编码, i.名称, i.规格, i.产地, i.计算单位, i.是否变价 ) b " & _
                    " Where b.分类id = a.ID  Order By b.编码 "
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "选择器", str分类, mstr服务对象, gstrPriceClass)
    
    '清空
    Me.LivMain.ListItems.Clear
    
    Me.MousePointer = 11
     
    
    Do Until rsTmp.EOF
        Set ItmX = Me.LivMain.ListItems.Add(, "A" & rsTmp("ID"), rsTmp("编码"), 1, 1)
        ItmX.SubItems(1) = Nvl(rsTmp("名称"))
        ItmX.SubItems(2) = Nvl(rsTmp("规格"))
        ItmX.SubItems(3) = Nvl(rsTmp("产地"))
        ItmX.SubItems(4) = Nvl(rsTmp("计算单位"))
        ItmX.SubItems(5) = Nvl(rsTmp("售价"))
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    Me.MousePointer = 1
    
    NowIndex = Node.Index
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MouseStartX = X
    End If
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MoveTmp As Single
    '暂时屏蔽方便查问题
    On Error Resume Next
    If Button = 1 Then
        
        '得到移动后的位置
        MoveTmp = Me.picSplit.Left + X - MouseStartX
        
        '超过最大或最小宽度时退出
        If MoveTmp <= 2000 Or Me.ScaleWidth - MoveTmp <= 2000 Then Exit Sub
        
        'picSplit
        picSplit.Left = MoveTmp
        
        'LvwMain
        Me.LvwMain.Width = Me.picSplit.Left
        
        'LivMain
        Me.LivMain.Left = Me.picSplit.Left + Me.picSplit.Width
        Me.LivMain.Width = Me.ScaleWidth - Me.picSplit.Left - Me.picSplit.Width
    End If
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal str服务对象 As String)
    mstr服务对象 = str服务对象
    
    Me.Show 1, frmParent
End Sub
