VERSION 5.00
Begin VB.Form frmDefQueryPage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "页面编辑"
   ClientHeight    =   5700
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9180
   Icon            =   "frmDefQueryPage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1110
      MaxLength       =   100
      TabIndex        =   32
      Top             =   4800
      Width           =   7995
   End
   Begin VB.Frame fra 
      Caption         =   "基本信息"
      Height          =   4665
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   4920
      Begin VB.CommandButton cmdOpen 
         Caption         =   "…"
         Height          =   240
         Index           =   2
         Left            =   4425
         TabIndex        =   10
         Top             =   1425
         Width           =   285
      End
      Begin VB.TextBox txtEdit 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1395
         Width           =   4005
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   720
         MaxLength       =   30
         TabIndex        =   5
         Top             =   660
         Width           =   4005
      End
      Begin VB.ListBox lst 
         Height          =   1320
         Left            =   720
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   1755
         Width           =   4005
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3225
         Width           =   4005
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "试听(&H)"
         Height          =   350
         Left            =   3615
         TabIndex        =   15
         Top             =   3585
         Width           =   1100
      End
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   960
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "11111"
         Top             =   345
         Width           =   900
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   720
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1020
         Width           =   4005
      End
      Begin VB.TextBox txtTemp 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   720
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "编码"
         Text            =   "1111111111"
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "风格(&S)"
         Height          =   180
         Left            =   75
         TabIndex        =   11
         Top             =   1815
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Left            =   75
         TabIndex        =   4
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "音乐(&M)"
         Height          =   180
         Left            =   75
         TabIndex        =   13
         Top             =   3285
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "编码(&B)"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "简码(&Y)"
         Height          =   180
         Index           =   2
         Left            =   75
         TabIndex        =   6
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "上级(&U)"
         Height          =   180
         Index           =   3
         Left            =   75
         TabIndex        =   8
         Top             =   1455
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "宣传标语"
      Height          =   2085
      Index           =   2
      Left            =   4965
      TabIndex        =   21
      Top             =   2640
      Width           =   4110
      Begin VB.CommandButton cmdPos 
         Height          =   345
         Index           =   1
         Left            =   3630
         Picture         =   "frmDefQueryPage.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "显示宣传标语在查询中的位置"
         Top             =   1470
         Width           =   345
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   345
         Index           =   1
         Left            =   3630
         Picture         =   "frmDefQueryPage.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "选择宣传标语"
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton cmdClear 
         Height          =   345
         Index           =   1
         Left            =   3630
         Picture         =   "frmDefQueryPage.frx":0643
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "清除宣传标语"
         Top             =   615
         Width           =   345
      End
      Begin zl9NewQuery.ctlPicture UsrPic 
         Height          =   1590
         Index           =   1
         Left            =   75
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   225
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   2805
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   30
         Top             =   1845
         Width           =   810
      End
   End
   Begin VB.Frame fra 
      Caption         =   "背景图片"
      Height          =   2535
      Index           =   1
      Left            =   4965
      TabIndex        =   16
      Top             =   60
      Width           =   4110
      Begin VB.CommandButton cmdPos 
         Height          =   345
         Index           =   0
         Left            =   3615
         Picture         =   "frmDefQueryPage.frx":06E9
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "显示背景图片在查询中的位置"
         Top             =   1905
         Width           =   345
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   345
         Index           =   0
         Left            =   3615
         Picture         =   "frmDefQueryPage.frx":0C73
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "选择背景图片"
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton cmdClear 
         Height          =   345
         Index           =   0
         Left            =   3615
         Picture         =   "frmDefQueryPage.frx":0D20
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "清除背景图片"
         Top             =   615
         Width           =   345
      End
      Begin zl9NewQuery.ctlPicture UsrPic 
         Height          =   2010
         Index           =   0
         Left            =   75
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   3545
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   29
         Top             =   2280
         Width           =   810
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   75
      TabIndex        =   28
      Top             =   5250
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7995
      TabIndex        =   27
      Top             =   5250
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6765
      TabIndex        =   26
      Top             =   5250
      Width           =   1100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "命令参数(&A)"
      Height          =   180
      Left            =   75
      TabIndex        =   31
      Top             =   4845
      Width           =   990
   End
End
Attribute VB_Name = "frmDefQueryPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFirst As Boolean
Private mKey As Long
Private mOK As Boolean

Private mvarSvrPicRange As String           '保存增加图片的范围
Private mvarSvrPicType As String            '保存增加图片的类型

Private mlngKey As Long
Private mlngUpKey As Long
Private mstr上级ID As String
Private mstr上级编码 As String
Private mstr编码 As String
Const mlng编码长度 = 10

Private Sub GetTreeCode(ByVal lngUpKey As Long)
    '获取树型结构的编码规则,包括上级编码,本级编码
    
    If lngUpKey = 0 Then
        '如果没有指定上级
        mstr上级编码 = ""
        txtTemp.Text = ""
        
        txtEdit(3).Text = "无"
        
        '取得上级编码，本级编码长度等值
        txtTemp.MaxLength = GetLocalCodeLength("", "咨询页面目录")
        
    Else
        '指定了上级
        gstrSQL = "select 编码 as 上级编码,页面名称 as 上级名称,页面序号 as 上级ID from 咨询页面目录 where 页面序号=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUpKey)
                        
        mstr上级ID = IIf(IsNull(gRs("上级ID")), "", gRs("上级ID"))
        mstr上级编码 = IIf(IsNull(gRs("上级编码")), "", gRs("上级编码"))
        txtEdit(3).Text = IIf(IsNull(gRs("上级名称")), "无", gRs("上级名称"))
        txtEdit(3).Tag = lngUpKey
        txtTemp.MaxLength = 0
        txtTemp.Text = mstr上级编码
        
        '判断编码是否满了
        If Len(mstr上级编码) >= mlng编码长度 Then
            MsgBox "不能再增加子分类了，编码长度已经用尽。", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        '取得上级编码，本级编码长度等值
        txtTemp.MaxLength = GetLocalCodeLength(mstr上级ID, "咨询页面目录")
    End If
        
    txtEdit(0).MaxLength = IIf(txtTemp.MaxLength = 0, mlng编码长度, txtTemp.MaxLength) - Len(mstr上级编码)
    txtEdit(0).Text = Mid(txtEdit(0).Text, Len(txtTemp.Text) + 1)
    
    If mKey = 0 Then txtEdit(0).Text = GetMaxLocalCode(mstr上级ID, "咨询页面目录")
End Sub

Public Function ShowPageEdit(frmMain As Object, ByVal Key As Long, ByVal lngUpKey As Long) As Boolean

    mKey = Key
    mlngUpKey = lngUpKey

    frmDefQueryPage.Show 1, frmMain
    ShowPageEdit = mOK
End Function

Private Sub cbo_Click()
    cmdOK.Tag = "1"
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdOK.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click(Index As Integer)
    If UsrPic(Index).Tag <> "" Then
        UsrPic(Index).Tag = ""
        UsrPic(Index).Cls
        cmdOK.Tag = "1"
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mOK = True
        If mKey = 0 Then
            txt.Text = ""
            txtEdit(0).Text = ""
            txtEdit(2).Text = ""
            txtEdit(1).Text = ""
            
            UsrPic(0).Tag = ""
            UsrPic(1).Tag = ""
            lblSize(0).Caption = ""
            lblSize(1).Caption = ""
            UsrPic(0).Cls
            UsrPic(1).Cls
            txtEdit(0).Text = GetMaxLocalCode(txtEdit(3).Tag, "咨询页面目录")
            cmdOK.Tag = ""
            txtEdit(0).SetFocus
        Else
            cmdOK.Tag = ""
            Unload Me
        End If
    End If
End Sub


Private Sub cmdOpen_Click(Index As Integer)
    Dim lngKey As Long
    Dim strFilter As String
    Dim strTitle As String
            
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim strRerurnID As String
    Dim str编码 As String
    Dim int编码  As Integer
            
            
    If Index = 2 Then
        
        strSQL = "Select 页面序号 AS id,上级序号 AS 上级id,页面名称 AS 名称,编码,0 as 末级 From 咨询页面目录 Where (末级 IS NULL OR 末级=0)  Start with 上级序号 is null connect by prior 页面序号 =上级序号 "
        
        strID = txtEdit(3).Tag
        str名称 = txtEdit(3).Text
        str编码 = txtTemp.Text & txtEdit(0).Text
            
        blnRe = frm树型选择.ShowTree(strSQL, strID, str名称, mstr上级编码, "", Me.Caption, "所有页面分类", , mstr编码)
    
        If blnRe Then       '新的本级的宽度
            
            mlngUpKey = Val(strID)
            txtEdit(3).Tag = strID
            txtEdit(3).Text = str名称
            Call GetTreeCode(mlngUpKey)
            txtEdit(0).Text = GetMaxLocalCode(strID, "咨询页面目录")
            cmdOK.Tag = "1"
        End If
    Else
        strFilter = IIf(Index = 0, "4;0;1;2;3;9", "1;0;2;3;4;9")
        Select Case Index
        Case 0
            strTitle = "添加页面背景"
        Case 1
            strTitle = "添加页面宣传标语"
        End Select
        If frmPicSelect.OpenPictureBox(Me, strTitle, strFilter, lngKey, mvarSvrPicRange, mvarSvrPicType) Then
            '更新图片显示
            Call ShowPicture(lngKey, Index)
            UsrPic(Index).Tag = lngKey
            cmdOK.Tag = "1"
        End If
    End If
End Sub

Private Sub cmdPos_Click(Index As Integer)
    Select Case Index
    Case 0
        Call frmPosSample.ShowPageSample("页面背景")
    Case 1
        Call frmPosSample.ShowPageSample("宣传标语")
    End Select
End Sub

Private Sub cmdTest_Click()
    Dim vFileData As New FileSystemObject
    Dim strFile As String
    
    Call MusicClose
    
    
    If cbo.ListIndex < 0 Then Exit Sub
    If cbo.ItemData(cbo.ListIndex) <= 0 Then Exit Sub
    
    '1.检查图形目录是否存在
    On Error Resume Next
    vFileData.CreateFolder App.Path & "\图形"
    
    '2.检查本系统中可能使用到的图片
    gstrSQL = "select 序号,类型,名称,修改日期 from 咨询图片元素 where 序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(cbo.ItemData(cbo.ListIndex)))
    If gRs.BOF Then Exit Sub
    
    strFile = IIf(IsNull(gRs!名称), "", gRs!名称)
    If strFile <> "" Then Call CheckFileNew(strFile, IIf(IsNull(gRs!类型), 0, gRs!类型), gRs!序号, gRs!修改日期, vFileData)
            
    Call MusicPlay(strFile)
End Sub

Private Sub Command1_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    DoEvents
    
    '初始化过程
    lst.AddItem "0-标准"
    lst.ItemData(lst.NewIndex) = 0
    Call SelectListItem(0)
    
    cbo.AddItem "[无]"
    gstrSQL = "select 序号,名称 from 咨询图片元素 where 类型=3"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            cbo.AddItem IIf(IsNull(gRs!名称), "", gRs!名称)
            cbo.ItemData(cbo.NewIndex) = IIf(IsNull(gRs!序号), 0, gRs!序号)
            gRs.MoveNext
        Wend
    End If
    cbo.ListIndex = 0
    
    If mKey <> 0 Then
        If frmDefQuery.lvw.SelectedItem.Tag = "1" Then
            txt.Enabled = False
            lst.Enabled = False
            If frmDefQuery.lvw.SelectedItem.Text <> "专家介绍" And mKey > 0 Then
                'tbs.TabEnabled(1) = False
                Fra(1).Enabled = False
            End If
        End If
                
        gstrSQL = "select A.命令参数,A.编码,A.简码,A.页面名称,A.页面风格,A.宣传标语,A.页面背景,A.宣传标语,A.页面背景,背景音乐 from 咨询页面目录 A where A.页面序号=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mKey)
        If gRs.BOF = False Then
            txt.Text = IIf(IsNull(gRs!页面名称), "", gRs!页面名称)
            Call SelectListItem(IIf(IsNull(gRs!页面风格), 0, gRs!页面风格))
            
            Call ShowPicture(IIf(IsNull(gRs!页面背景), 0, gRs!页面背景), 0)
            Call ShowPicture(IIf(IsNull(gRs!宣传标语), 0, gRs!宣传标语), 1)
            
            UsrPic(0).Tag = IIf(IsNull(gRs!页面背景), 0, gRs!页面背景)
            UsrPic(1).Tag = IIf(IsNull(gRs!宣传标语), 0, gRs!宣传标语)
                        
            cbo.ListIndex = FindCboIndex(cbo, IIf(IsNull(gRs!背景音乐), 0, gRs!背景音乐))
            txtEdit(0).Text = IIf(IsNull(gRs!编码), "", gRs!编码)
            txtEdit(2).Text = IIf(IsNull(gRs!简码), "", gRs!简码)
            
            txtEdit(1).Text = IIf(IsNull(gRs!命令参数), "", gRs!命令参数)
            
            mstr编码 = txtEdit(0).Text
        End If
    End If
    
    Call GetTreeCode(mlngUpKey)
    
    cmdOK.Tag = ""
    
    txtEdit(2).Enabled = txt.Enabled
    
    DoEvents
    
    txtEdit(0).SetFocus
    Call SelAll(txtEdit(0))
    
    mblnFirst = False
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mOK = False
        
    lblSize(0).Caption = ""
    lblSize(1).Caption = ""
                
    mvarSvrPicRange = ""
    mvarSvrPicType = ""
    
End Sub

Private Function CheckValid() As Boolean
    txtEdit(0).Text = Trim(txtEdit(0).Text)

    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(0).Text) = 0 Then
            MsgBox "编码不能为空。", vbExclamation, gstrSysName
            txtEdit(0).SetFocus
            Exit Function
        End If
    Else
        If Len(txtEdit(0).Text) < txtEdit(0).MaxLength Then
            MsgBox "编码的长度不够。", vbExclamation, gstrSysName
            txtEdit(0).SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txtEdit(0).Text) Or InStr(txtEdit(0).Text, ",") > 0 Or InStr(txtEdit(0).Text, ".") > 0 Or InStr(txtEdit(0).Text, "-") > 0 Then
        MsgBox "编码应由数字组成。", vbExclamation, gstrSysName
        txtEdit(0).SetFocus
        Exit Function
    End If
    If Len(Trim(txt.Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        txt.Text = ""
        txt.SetFocus
        Exit Function
    End If
    
    CheckValid = True
End Function


Private Function SaveData() As Boolean
    Dim lng序号 As Long
    Dim LngStyle As Long
    Dim i As Long
        
    If cmdOK.Tag <> "" Then
        
        If CheckValid = False Then Exit Function
        
        For i = 0 To lst.ListCount - 1
            If lst.Selected(i) Then
                LngStyle = lst.ItemData(i)
                Exit For
            End If
        Next
        If mKey = 0 Then
            lng序号 = NextValue("咨询页面目录", "页面序号")
            gstrSQL = "zl_咨询页面目录_insert(" & lng序号 & ",'" & txt.Text & "',0," & LngStyle & "," & IIf(Val(UsrPic(1).Tag) = 0, "NULL", Val(UsrPic(1).Tag)) & "," & IIf(Val(UsrPic(0).Tag) = 0, "NULL", Val(UsrPic(0).Tag)) & "," & IIf(cbo.ItemData(cbo.ListIndex) = 0, "NULL", cbo.ItemData(cbo.ListIndex)) & "," & IIf(Val(txtEdit(3).Tag) = 0, "NULL", Val(txtEdit(3).Tag)) & ",1,'" & txtTemp.Text & txtEdit(0).Text & "','" & txtEdit(2).Text & "','" & txtEdit(1).Text & "')"
        Else
            lng序号 = mKey
            gstrSQL = "zl_咨询页面目录_update(" & mKey & ",'" & txt.Text & "'," & LngStyle & "," & IIf(Val(UsrPic(1).Tag) = 0, "NULL", Val(UsrPic(1).Tag)) & "," & IIf(Val(UsrPic(0).Tag) = 0, "NULL", Val(UsrPic(0).Tag)) & "," & IIf(cbo.ItemData(cbo.ListIndex) = 0, "NULL", cbo.ItemData(cbo.ListIndex)) & "," & IIf(Val(txtEdit(3).Tag) = 0, "NULL", Val(txtEdit(3).Tag)) & ",'" & txtTemp.Text & txtEdit(0).Text & "','" & txtEdit(2).Text & "'," & Len(mstr编码) + 1 & ",'" & txtEdit(1).Text & "')"
        End If
                        
        On Error GoTo errHand
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        
        Call frmDefQuery.RefreshPage(CStr(lng序号))
        
    End If
    
    SaveData = True
    Exit Function
errHand:
    If ErrCenter() = -1 Then Resume
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Cancel = mblnFirst
    If Cancel Then Exit Sub
    
    Call MusicClose
    If cmdOK.Tag = "1" Then
        If MsgBox("查询页面已经更改，确认不保存就退出吗？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True
    End If
End Sub

Private Sub lst_ItemCheck(Item As Integer)
    Call SelectListItem(lst.ItemData(Item))
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
End Sub

Private Sub txt_Change()
    cmdOK.Tag = "1"
End Sub

Private Sub SelectListItem(ByVal Key As Long)
    Dim i As Long
    
    For i = 0 To lst.ListCount - 1
        If lst.ItemData(i) = Key Then
            lst.Selected(i) = True
        Else
            lst.Selected(i) = False
        End If
    Next
End Sub

Private Sub ShowPicture(ByVal PicNo As Long, ByVal Index As Long)
    Dim rs As New ADODB.Recordset
    gstrSQL = "select 序号,宽度,高度,类型 from 咨询图片元素 where 序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PicNo)
    If rs.BOF = False Then
        Call UsrPic(Index).ShowPictureByFieldNew(rs!序号, rs!宽度 * Screen.TwipsPerPixelX, rs!高度 * Screen.TwipsPerPixelY, IIf(IsNull(rs!类型), 0, rs!类型))
        lblSize(Index).Caption = "宽度:" & Format(rs!宽度 * Screen.TwipsPerPixelX / 567, "0.0(厘米)") & " 高度:" & Format(rs!高度 * Screen.TwipsPerPixelY / 567, "0.0(厘米)")
    End If
    CloseRecord rs
End Sub

Private Sub txt_GotFocus()
    Call SelAll(txt)
    zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    Else
        txtEdit(2).Text = zlCommFun.SpellCode(txt.Text)
    End If
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    cmdOK.Tag = "1"
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Call SelAll(txtEdit(Index))
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        If Index = 3 Then SendKeys "{TAB}"
    Else
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Index = 3 And Chr(KeyAscii) = "*" Then Call cmdOpen_Click(2)
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txtEdit(Index).Text, txtEdit(Index).MaxLength)
End Sub

Private Sub txtTemp_Change()
    txtEdit(0).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(0).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub
