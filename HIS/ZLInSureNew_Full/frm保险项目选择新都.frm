VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm保险项目选择新都 
   AutoRedraw      =   -1  'True
   Caption         =   "医保项目选择"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm保险项目选择新都.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7845
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7845
      TabIndex        =   4
      Top             =   4350
      Width           =   7845
      Begin VB.CommandButton Cmd更新 
         Caption         =   "更新项目"
         Height          =   345
         Left            =   4020
         TabIndex        =   11
         Top             =   150
         Width           =   915
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印列表"
         Height          =   350
         Left            =   2790
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Top             =   175
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   7
         Top             =   150
         Width           =   1100
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "明细查找(&F)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   930
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   4050
      Left            =   3060
      TabIndex        =   2
      Top             =   270
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   7144
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
         Text            =   "项目编码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "项目名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "一级医院价格"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "二级医院价格"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "三级医院价格"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "自付比例"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3915
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
            Picture         =   "frm保险项目选择新都.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择新都.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4050
      Left            =   0
      TabIndex        =   10
      Top             =   270
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7144
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目大类(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2970
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目明细(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   1
      Top             =   15
      Width           =   4710
   End
End
Attribute VB_Name = "frm保险项目选择新都"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private mstrCode As String
Private mstrName As String
Private mdbl医院价格 As Double
Private mobjStream As TextStream
Private mobjFileSystem As New FileSystemObject
Private mblnOK As Boolean
Private Const strFile = "C:\XDYB_YH\ERR.LOG"
Private mErrFile As TextStream

Private Declare Function MakeTxt Lib "yhybReckoning.dll" Alias "_MakeTxt@8" (ByVal str服务目录文件 As String, _
        ByVal str病种目录文件 As String) As String

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If lvwDetail.SelectedItem Is Nothing Then
        MsgBox "没有选择项目！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    
    '返回选择项目编码
    mstrCode = lvwDetail.SelectedItem.Text
    mstrName = lvwDetail.SelectedItem.ListSubItems(1)
    
    mblnOK = True
    Unload Me
End Sub

Public Function GetCode(strCode As String, STRNAME As String, ByVal dbl医院价格 As Double, ByVal int险类 As Integer) As Boolean
'功能：获得一个收费项目的医保编码
'参数：strCode 既作为输入参数，又输出
'返回：成功返回True
    tvwClass.Nodes.Clear
    tvwClass.Nodes.Add , , "YAO", "药品项目", 2
    tvwClass.Nodes.Add , , "ZEN", "诊疗项目", 2
    
    mint险类 = int险类
    
    Set tvwClass.SelectedItem = tvwClass.Nodes("YAO")
    FillList
    
    frm保险项目选择新都.Show vbModal, frm保险项目
    '返回值
    If mblnOK = True Then
        strCode = mstrCode
        STRNAME = mstrName
    End If
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillList()
'功能：显示当前类别下的医保明细
    Dim rsTemp As New ADODB.Recordset, rsTmp As New ADODB.Recordset, strTemp As String
    Dim cn新都 As New ADODB.Connection
    Dim itmTemp As ListItem
    
    lvwDetail.ListItems.Clear
    On Error Resume Next

    On Error GoTo errHandle

    

    If cn新都.State Then cn新都.Close
    cn新都.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\YWCS.MDB;Persist Security Info=True;Jet OLEDB:Database Password=yhybv1.1cdb"
    cn新都.CursorLocation = adUseClient
    cn新都.Open
    
    If tvwClass.SelectedItem.Key = "YAO" Then
        strTemp = "Select ybxmbm As 医保编码,ybxmmc As 项目名称,ybxmrj as 项目简码,zgxj as 限价,zfbl1 As 自付比例,fzxbh as 分中心编号 From KYH904 order by ybxmbm"
    Else
        strTemp = "Select ybxmbm As 医保编码,ybxmmc As 项目名称,zgxj As 一级医院价格,zgxj1 As 二级医院价格,zgxj2 As 三级医院价格,zfbl1 As 自付比例 From KYH100 "
    End If
    Set rsTemp = cn新都.Execute(strTemp)
    
    If tvwClass.SelectedItem.Key <> "YAO" Then
        lvwDetail.ColumnHeaders(3).Text = "一级医院价格"
        lvwDetail.ColumnHeaders(4).Text = "二级医院价格"
        lvwDetail.ColumnHeaders(5).Text = "三级医院价格"
        lvwDetail.ColumnHeaders(6).Text = "自付比例"
    Else
        lvwDetail.ColumnHeaders(3).Text = "项目简码"
        lvwDetail.ColumnHeaders(4).Text = "限    价"
        lvwDetail.ColumnHeaders(5).Text = "自付比例"
        lvwDetail.ColumnHeaders(6).Text = "分中心编号"
    End If

    While Not rsTemp.EOF
        Set itmTemp = lvwDetail.ListItems.Add(, , rsTemp(0), 1)
        itmTemp.ListSubItems.Add , , rsTemp(1)
        itmTemp.ListSubItems.Add , , rsTemp(2)
'        If tvwClass.SelectedItem.Key <> "YAO" Then
        itmTemp.ListSubItems.Add , , rsTemp(3)
        itmTemp.ListSubItems.Add , , rsTemp(4)
        itmTemp.ListSubItems.Add , , rsTemp(5)
'        End If
        rsTemp.MoveNext
    Wend
    
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    
    objPrint.Title.Text = "保险项目"
    Set objPrint.Body.objData = lvwDetail
    objPrint.UnderAppItems.Add "医保大类：" & tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
    End Select

End Sub

Private Sub Cmd更新_Click()
    Dim rsTemp As New ADODB.Recordset, rsTmp As New ADODB.Recordset, strTemp As String
    Dim cn新都 As New ADODB.Connection
    Dim i As Integer

    On Error Resume Next

    On Error GoTo errHandle

    
    cn新都.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\YYXX.MDB;Persist Security Info=True;Jet OLEDB:Database Password=yhybv1.1cdb"
    cn新都.CursorLocation = adUseClient
    cn新都.Open
    strTemp = "Select fzx1 as 序号,fzx2 as 编号,fzx3 as 名称 From KYH999"
    Set rsTemp = cn新都.Execute(strTemp)
    
    i = 0
    '自动更新医保分中心表
    While Not rsTemp.EOF
        i = i + 1
        Me.Caption = "正在更新医保分中心:第" & i & "条,共" & rsTemp.RecordCount & "条!"
        gstrSQL = "zl_医保分中心_成都_INSERT(" & rsTemp!序号 & ",'" & rsTemp!编号 & "','" & rsTemp!名称 & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        rsTemp.MoveNext
    Wend
    
    If cn新都.State Then cn新都.Close
    cn新都.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\YWCS.MDB;Persist Security Info=True;Jet OLEDB:Database Password=yhybv1.1cdb"
    cn新都.CursorLocation = adUseClient
    cn新都.Open
        
    strTemp = "Select ybxmbm As 医保编码,ybxmmc As 项目名称,ybxmrj as 项目简码,zgxj as 限价,zfbl1 As 自付比例,fzxbh as 分中心编号 From KYH904 order by ybxmbm"

    Set rsTemp = cn新都.Execute(strTemp)
    
    i = 0
    While Not rsTemp.EOF
        '自动更新医保项目表
        i = i + 1
        Me.Caption = "正在更新医保项目:第" & i & "条,共" & rsTemp.RecordCount & "条!"
        gstrSQL = "zl_医保项目表_成都_INSERT('" & rsTemp!医保编码 & "','" & rsTemp!项目名称 & "','" & rsTemp!项目简码 & "'," & rsTemp!限价 & "," & rsTemp!自付比例 & ",'" & rsTemp!分中心编号 & "'," & mint险类 & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        rsTemp.MoveNext
    Wend
    
    Me.Caption = "医保项目选择"
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = tvwClass.Width
    
    On Error Resume Next
    
    tvwClass.Left = 0: tvwClass.Top = lblClass.Top + lblClass.Height
    tvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = tvwClass.Top
    picSplit.Left = tvwClass.Left + tvwClass.Width
    picSplit.Height = tvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If tvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
    
    lvwDetail.Top = tvwClass.Top
    lvwDetail.Left = lblDetail.Left
    lvwDetail.Width = lblDetail.Width
    lvwDetail.Height = tvwClass.Height
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwDetail_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvwClass.Width + x < 1000 Or lvwDetail.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        tvwClass.Width = tvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        lvwDetail.Left = lvwDetail.Left + x
        lvwDetail.Width = lvwDetail.Width - x
    End If
End Sub

Private Sub lvwdetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwDetail.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwDetail.SortOrder = lvwDescending
    Else
        lvwDetail.SortOrder = lvwAscending
    End If
    lvwDetail.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwDetail.SelectedItem Is Nothing Then lvwDetail.SelectedItem.EnsureVisible
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillList
End Sub

Private Sub txtFind_Change()
'功能：根据用户输入的内容查找匹配的内容
    Dim lst As ListItem, lngIndex As Long, lngSubItems As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    If lvwDetail.ListItems.Count = 0 Then Exit Sub

    Set lst = lvwDetail.FindItem(strFind, lvwText, , lvwPartial)
    If Not lst Is Nothing Then
        lst.Selected = True
        lst.EnsureVisible
    Else
        '非文本不能做到部分匹配
        lngSubItems = lvwDetail.ColumnHeaders.Count - 1
        For Each lst In lvwDetail.ListItems
            For lngIndex = 1 To lngSubItems
                If lst.SubItems(lngIndex) Like strFind & "*" Then
                    lst.Selected = True
                    lst.EnsureVisible
                    Exit Sub
                End If
            Next
            
        Next
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Function ReplaceStr(ByVal StrInput As String) As String
    ReplaceStr = Trim(Replace(StrInput, "'", ""))
    ReplaceStr = Replace(ReplaceStr, """", "")
End Function
