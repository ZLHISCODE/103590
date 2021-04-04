VERSION 5.00
Begin VB.Form frmPresFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "人员查找"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "frmPresFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra匹配 
      Caption         =   "匹配方式"
      Height          =   1515
      Left            =   3030
      TabIndex        =   14
      Top             =   120
      Width           =   1500
      Begin VB.OptionButton optMatch 
         Caption         =   "从左匹配"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   450
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "任意匹配"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5160
      TabIndex        =   13
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   5160
      TabIndex        =   12
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "定位(&L)"
      Height          =   350
      Left            =   5160
      TabIndex        =   11
      Top             =   180
      Width           =   1100
   End
   Begin VB.Frame fra条件 
      Caption         =   "查找条件"
      Height          =   2685
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   2760
      Begin VB.ComboBox cmb性别 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   870
         MaxLength       =   255
         TabIndex        =   1
         Top             =   330
         Width           =   1725
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   870
         MaxLength       =   255
         TabIndex        =   3
         Top             =   720
         Width           =   1725
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   870
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1110
         Width           =   1725
      End
      Begin VB.ComboBox cmb学历 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1890
         Width           =   1755
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "学历(&D)"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   8
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "性别(&X)"
         Height          =   180
         Index           =   3
         Left            =   180
         TabIndex        =   6
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "编号(&C)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   630
      End
   End
   Begin VB.Label lbl结果 
      BackStyle       =   0  'Transparent
      Caption         =   " 请输入查找条件"
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   3000
      TabIndex        =   10
      Top             =   1920
      Width           =   3315
   End
End
Attribute VB_Name = "frmPresFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintfun As Integer  '0-人员查找,1-部门查找
Private mbln是否显示停用 As Boolean
Private mblnViewDel As Boolean
Dim mrsFind As New ADODB.Recordset
Private mint模式 As Integer '1-按层次显示，2-按性质显示


Private Sub cmb性别_Click()
    If mrsFind.State = 1 Then mrsFind.Close
    lbl结果.Caption = "  条件已改变，请重新定位"
    lbl结果.ForeColor = &H8000&
    If Not cmdFind.Enabled Then cmdFind.Enabled = True
End Sub

Private Sub cmb性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb学历_Click()
    If mrsFind.State = 1 Then mrsFind.Close
    lbl结果.Caption = "  条件已改变，请重新定位"
    lbl结果.ForeColor = &H8000&
    If Not cmdFind.Enabled Then cmdFind.Enabled = True
End Sub

Private Sub cmb学历_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    Select Case mintfun
    Case 0   '默认按人员查找
        Dim rsTemp As New ADODB.Recordset
        
        cmb性别.AddItem " "
        gstrSQL = "Select '性别' As 类别, 名称,编码 From 性别 Union All Select '学历' As 类别, 名称,编码 From 学历 Order By 类别,编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        rsTemp.Filter = "类别='性别'"
        Do Until rsTemp.EOF
            cmb性别.AddItem rsTemp("名称")
            rsTemp.MoveNext
        Loop
        
        cmb学历.AddItem " "
        rsTemp.Filter = "类别='学历'"
        Do Until rsTemp.EOF
            cmb学历.AddItem rsTemp("名称")
            rsTemp.MoveNext
        Loop
        
        rsTemp.Close
        cmdFind.Enabled = False
    Case 1
        frmPresFind.Caption = "部门查找"
        cmb学历.Visible = False
        cmb性别.Visible = False
        lbl(1).Caption = "名称"
        lbl(3).Visible = False
        lbl(4).Visible = False
        fra条件.Height = fra匹配.Height
        lbl结果.Left = fra条件.Left
        lbl结果.Width = fra匹配.Left - fra条件.Left + fra匹配.Width
        'lbl结果.Height = lbl结果.Height / 2
        'Me.Height = Me.Height - lbl结果.Height
        
        cmdFind.Enabled = False
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowOfType(frmParent As Object, intType As Integer, Optional blnShowStop As Boolean = False, Optional blnShowDel As Boolean = False, Optional int模式 As Integer)
    mintfun = intType
    mbln是否显示停用 = blnShowStop
    mblnViewDel = blnShowDel
    mint模式 = int模式
    
    frmPresFind.Show vbModal, frmParent
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    Set mrsFind = Nothing
End Sub

Private Sub cmdFind_Click()
    On Error GoTo ErrHandle
    If mrsFind.State = 1 Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocateItem
        Exit Sub
    End If
    If IsValid = False Then Exit Sub
    gstrSQL = ""
    
    If txtEdit(0).Text <> "" Then
        gstrSQL = "and upper(" & Choose(mintfun + 1, "A.编号", "a.编码") & ") like [1]  "
    End If
    If txtEdit(1).Text <> "" Then
        gstrSQL = gstrSQL & " and upper(" & Choose(mintfun + 1, "A.姓名", "a.名称") & ") like [2] "
    End If
    
    If txtEdit(2).Text <> "" Then
        gstrSQL = gstrSQL & "and upper(A.简码) like [3]  "
    End If
    
    If mintfun = 0 Then
        If Trim(cmb性别.Text) <> "" Then
            gstrSQL = gstrSQL & "and A.性别=[4] "
        End If
        If Trim(cmb学历.Text) <> "" Then
            gstrSQL = gstrSQL & "and A.学历=[5] "
        End If
    End If
    
    If gstrSQL = "" Then
'        gstrSQL = Mid(gstrSQL, 1, Len(gstrSQL) - 4)
'    Else
        MsgBox "请输入查找条件。", vbExclamation, gstrSysName
        txtEdit(0).SetFocus
        Exit Sub
    End If
        
    Select Case mintfun
    Case 0  '查找人员
        If InStr(frmPresManage.mstrPrivs, "所有部门") = 0 Then
            gstrSQL = "Select a.Id, a.姓名, b.部门id" & vbNewLine & _
                      "From 人员表 A, 部门人员 B " & vbNewLine & _
                      "Where a.Id = b.人员id And b.部门id In (Select Distinct ID" & vbNewLine & _
                      "    From 部门表 A" & vbNewLine & _
                      "    Start With ID In (Select 部门id From 部门人员 Where 人员id = [6])" & vbNewLine & _
                      "    Connect By Prior ID = 上级id) And b.缺省 = 1 " & gstrSQL
        Else
            gstrSQL = "select A.ID,A.姓名,B.部门ID " & _
                       " from 人员表 A,部门人员 B " & _
                       " where A.ID =B.人员ID and B.缺省=1  " & gstrSQL
        End If
        If Not mbln是否显示停用 Then
            gstrSQL = gstrSQL & " and (a.撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or a.撤档时间 is null ) "
        End If
    Case 1 '查找部门
        gstrSQL = "Select a.id,a.上级id,a.名称,a.编码 ,c.编码 as 性质 From 部门表 A, 部门性质说明 B,部门性质分类 c Where b.工作性质=c.名称 " & _
                " and A.ID=B.部门ID " & gstrSQL
        
        If Not mbln是否显示停用 Then
            gstrSQL = gstrSQL & " and (a.撤档时间 = to_date('3000-01-01','YYYY-MM-DD') or a.撤档时间 is null ) "
        End If
        If Not mblnViewDel Then
            gstrSQL = gstrSQL & " and substr(a.编码,1,1)<>'-' "
        End If
    End Select
    
    Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIF(optMatch(1).Value = True, "%", "") & UCase(txtEdit(0).Text) & "%", _
        IIF(optMatch(1).Value = True, "%", "") & UCase(txtEdit(1).Text) & "%", _
        IIF(optMatch(1).Value = True, "%", "") & UCase(txtEdit(2).Text) & "%", _
        cmb性别.Text, cmb学历.Text, glngUserId)

    If mrsFind.State = 1 Then
        Call LocateItem
    End If
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LocateItem()
    Dim strTemp As String
    
    If mrsFind.RecordCount = 0 Then
        lbl结果.Caption = " 没有找到符合条件的信息!"
        lbl结果.ForeColor = &HFF&
        Beep
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        lbl结果.Caption = " 已经定位完所有找到的信息，请重新输入条件"
        lbl结果.ForeColor = &HFF&
        Beep
        Exit Sub
    End If
    lbl结果.Caption = "  找到" & mrsFind.RecordCount & "条符合条件的信息。" & vbCrLf & "当前是第" & mrsFind.AbsolutePosition & _
                    "条，" & Choose(mintfun + 1, "姓名：", "名称：") & mrsFind(Choose(mintfun + 1, "姓名", "名称"))
    lbl结果.ForeColor = &H8000000D
    
    If mrsFind.RecordCount > 0 Then
        If mrsFind.RecordCount <> mrsFind.AbsolutePosition Then
            cmdFind.Caption = "下一个(&L)"
        Else
            cmdFind.Caption = "定位(&L)"
            cmdFind.Enabled = False
            lbl结果.Caption = lbl结果.Caption & vbCrLf & "已经定位到最后一条信息，请重新输入条件"
        End If
    End If
    
    Select Case mintfun
    Case 0  '查找人员
        With frmPresManage.tvwMain_S
            .Nodes("C" & mrsFind("部门ID")).Selected = True
            .SelectedItem.EnsureVisible
            frmPresManage.FillList "C" & mrsFind("部门ID")
        End With
            
        With frmPresManage.lvwMain
            .ListItems("C" & mrsFind("ID")).Selected = True
            .SelectedItem.EnsureVisible
            frmPresManage.lvwMain_ItemClick .SelectedItem
        End With
    Case 1 '查找部门
        With frmDeptManage.tvwMain_S
            If IsNull(mrsFind("上级ID")) Then
                .Nodes("C" & mrsFind("ID")).Selected = True
                .SelectedItem.EnsureVisible
                frmDeptManage.tvwMain_S_NodeClick .SelectedItem
            Else
                If mint模式 = 1 Then
                    .Nodes("C" & mrsFind("上级ID")).Selected = True
                    .Nodes("C" & mrsFind("上级ID")).Expanded = True
                Else
                    strTemp = mrsFind!性质 & "|" & mrsFind!ID
                    .Nodes("C" & strTemp).Selected = True
                    .Nodes("C" & strTemp).Expanded = True
                End If
                .SelectedItem.EnsureVisible
                frmDeptManage.tvwMain_S_NodeClick .SelectedItem
                
                If mint模式 = 1 Then
                    frmDeptManage.lvwMain.ListItems("C" & mrsFind("ID")).Selected = True
                    frmDeptManage.lvwMain.SelectedItem.EnsureVisible
                    frmDeptManage.lvwMain_ItemClick frmDeptManage.lvwMain.SelectedItem
                End If
            End If
        End With
    End Select
End Sub

Private Function IsValid() As Boolean
'功能:分析输入的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 2
        strTemp = Trim(txtEdit(i).Text)
        If InStr(strTemp, "'") > 0 Then
            MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
    Next
    IsValid = True
End Function

Private Sub optMatch_Click(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
End Sub

Private Sub optMatch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    lbl结果.Caption = "  条件已改变，请重新定位"
    cmdFind.Enabled = True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdFind.SetFocus
        Call cmdFind_Click
'          OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 1 Then
        OS.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    OS.OpenIme False
End Sub
