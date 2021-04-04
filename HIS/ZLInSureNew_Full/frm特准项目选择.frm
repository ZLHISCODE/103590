VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm特准项目选择 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "特准项目选择"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frm特准项目选择.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdDel 
      Caption         =   "删除(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5190
      TabIndex        =   11
      Top             =   2430
      Width           =   1100
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "加入(&A)"
      Height          =   350
      Left            =   5190
      TabIndex        =   10
      Top             =   1980
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5190
      TabIndex        =   14
      Top             =   270
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5190
      TabIndex        =   13
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5190
      TabIndex        =   12
      Top             =   3330
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   1950
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "条件(&R)"
      Height          =   1695
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4875
      Begin VB.ComboBox cbo性质 
         Height          =   300
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1140
         Width           =   2685
      End
      Begin VB.CommandButton Cmd结束编码 
         Caption         =   "…"
         Height          =   300
         Left            =   4050
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   750
         Width           =   285
      End
      Begin VB.TextBox Txt结束编码 
         Height          =   300
         Left            =   1650
         TabIndex        =   5
         Top             =   750
         Width           =   2415
      End
      Begin VB.CommandButton Cmd开始编码 
         Caption         =   "…"
         Height          =   300
         Left            =   4050
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox Txt开始编码 
         Height          =   300
         Left            =   1650
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lbl性质 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性质(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   960
         TabIndex        =   7
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lbl结束编码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束编码(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         TabIndex        =   4
         Top             =   810
         Width           =   990
      End
      Begin VB.Label lbl开始编码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始编码(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         TabIndex        =   1
         Top             =   420
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   5370
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特准项目选择.frx":1CFA
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特准项目选择.frx":2014
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特准项目选择.frx":232E
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特准项目选择.frx":2648
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特准项目选择.frx":2962
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特准项目选择.frx":2EFC
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm特准项目选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmParent As Object
Public lng险类 As Long
Public bln明细 As Boolean
Private strSelect As String

Private Sub CmdAdd_Click()
    Dim str开始编码 As String, str结束编码 As String, strSql As String
    Dim lvsItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    '将用户输入的范围内的项目，加入到列表框中
    str开始编码 = Trim(Txt开始编码.Tag)
    str结束编码 = Trim(Txt结束编码.Tag)
    
    '产生SQL
    If str开始编码 <> "" And str结束编码 <> "" Then
        strSql = " And A.编码 Between '" & str开始编码 & "' And '" & str结束编码 & "'"
    Else
        If str开始编码 <> "" Then
            strSql = " And A.编码>='" & str开始编码 & "'"
        ElseIf str结束编码 <> "" Then
            strSql = " And A.编码<='" & str结束编码 & "'"
        Else
            MsgBox "请输入开始编码：", vbInformation, gstrSysName
            Txt开始编码.SetFocus
        End If
    End If
    
    If bln明细 Then
        gstrSQL = "Select Distinct Rownum as 序号, A.ID,A.编码,B.名称,B.简码,A.类别,A.规格 " & _
                 "   FROM 收费细目 A,收费别名 B WHERE A.ID=B.收费细目ID " & strSql
    Else
        gstrSQL = "Select A.ID,A.编码,A.名称,A.简码 " & _
                 "   FROM 保险支付大类 A WHERE 险类=" & lng险类 & strSql
    End If
    Call OpenRecordset(rsTemp, "按用户设定的条件搜索记录")
    
    Do While Not rsTemp.EOF
        If InStr(1, strSelect & "|", "|" & rsTemp!ID & "|") = 0 Then
            strSelect = strSelect & "|" & rsTemp!ID
            Call addLvw(rsTemp)
        End If
        rsTemp.MoveNext
    Loop
    
    cmdOK.Enabled = (lvwDetail.ListItems.Count <> 0)
    CmdDel.Enabled = cmdOK.Enabled
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDel_Click()
    Dim lngItem As Long
    
    With lvwDetail
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        
        strSelect = strSelect & "|"
        For lngItem = 1 To .ListItems.Count
            If lngItem > .ListItems.Count Then Exit For
            If .ListItems(lngItem).Selected Then
                strSelect = Replace(strSelect, "|" & Mid(.ListItems(lngItem).Key, 2) & "|", "|")
                .ListItems.Remove .ListItems(lngItem).Key
                lngItem = lngItem - 1
            End If
        Next
        If Len(strSelect) > 1 Then
            strSelect = Mid(strSelect, 1, Len(strSelect) - 1)
        Else
            strSelect = ""
        End If
        
        If .ListItems.Count <> 0 Then .ListItems(1).Selected = True
        cmdOK.Enabled = (.ListItems.Count <> 0)
        CmdDel.Enabled = cmdOK.Enabled
    End With
End Sub

Private Sub cmdOK_Click()
    Dim strExist As String
    Dim objLvw As ListView, objItem As ListItem
    '更新主窗体中的数据
    If bln明细 Then
        Set objLvw = frmParent.Lvw明细
    Else
        Set objLvw = frmParent.lvw大类
    End If
    
    '获取已存在项目的ID串
    strExist = ""
    With objLvw
        For Each objItem In .ListItems
            strExist = strExist & "|" & Mid(objItem.Key, 2)
        Next
    End With
    
    '加入当前选择的项目
    With lvwDetail
        For Each objItem In .ListItems
            If InStr(1, strExist & "|", "|" & Mid(objItem.Key, 2) & "|") = 0 Then
                With objLvw
                    .ListItems.Add , objItem.Key, "[" & objItem.Text & "]" & objItem.SubItems(1), IIf(bln明细, "Fix", "Limit"), IIf(bln明细, "Fix", "Limit")
                    If bln明细 Then
                        .ListItems(objItem.Key).SubItems(1) = objItem.SubItems(2)
                        .ListItems(objItem.Key).SubItems(2) = objItem.SubItems(3)
                    Else
                        .ListItems(objItem.Key).SubItems(1) = objItem.SubItems(2)
                    End If
                End With
                strExist = strExist & "|" & Mid(objItem.Key, 2)
            End If
        Next
    End With
    
    Unload Me
End Sub

Private Sub Cmd结束编码_Click()
    Dim strID As String, str编码 As String, str名称 As String
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If bln明细 Then
        If frm收费细目选择.ShowTree(strID, str编码, str名称) = True Then
            If Get项目(strID, rsTemp) = False Then Exit Sub
            Txt结束编码.Text = "[" & rsTemp!编码 & "]" & rsTemp!名称
            Txt结束编码.Tag = rsTemp!编码
        End If
    Else
        gstrSQL = "Select Distinct Rownum as 序号, A.ID,A.编码,A.名称,A.简码 " & _
                 "   FROM 保险支付大类 A Where  险类=" & lng险类
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.RecordCount > 0 Then
            '出现选择器
            If rsTemp.RecordCount > 1 Then
                '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
                blnReturn = frmListSel.ShowSelect(lng险类, rsTemp, "序号", "选择器", "请选择项目：")
            Else
                blnReturn = True
            End If
        End If
        
        If blnReturn = False Then
            '记录集中没有可选择的数据
            Exit Sub
        Else
            Txt结束编码.Text = "[" & rsTemp!编码 & "]" & rsTemp!名称
            Txt结束编码.Tag = rsTemp!编码
        End If
    End If
End Sub

Private Sub Cmd开始编码_Click()
    Dim strID As String, str编码 As String, str名称 As String
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If bln明细 Then
        If frm收费细目选择.ShowTree(strID, str编码, str名称) = True Then
            If Get项目(strID, rsTemp) = False Then Exit Sub
            Txt开始编码.Text = "[" & rsTemp!编码 & "]" & rsTemp!名称
            Txt开始编码.Tag = rsTemp!编码
        End If
    Else
        gstrSQL = "Select Distinct Rownum as 序号, A.ID,A.编码,A.名称,A.简码 " & _
                 "   FROM 保险支付大类 A  Where 险类=" & lng险类
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.RecordCount > 0 Then
            '出现选择器
            If rsTemp.RecordCount > 1 Then
                '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
                blnReturn = frmListSel.ShowSelect(lng险类, rsTemp, "序号", "选择器", "请选择项目：")
            Else
                blnReturn = True
            End If
        End If
        
        If blnReturn = False Then
            '记录集中没有可选择的数据
            Exit Sub
        Else
            Txt开始编码.Text = "[" & rsTemp!编码 & "]" & rsTemp!名称
            Txt开始编码.Tag = rsTemp!编码
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With cbo性质
        .Clear
        .AddItem "0-不限"
        .AddItem "1-允许"
        .AddItem "2-排斥"
        .ListIndex = 0
    End With
    
    strSelect = ""
    Call initLvw
End Sub

Private Sub initLvw()
    lvwDetail.ColumnHeaders.Clear
    If bln明细 = False Then
        lvwDetail.ColumnHeaders.Add , "K1", "编码", 800
        lvwDetail.ColumnHeaders.Add , "K2", "名称", 2000
        lvwDetail.ColumnHeaders.Add , "K3", "性质", 1000
    Else
        lvwDetail.ColumnHeaders.Add , "K1", "编码", 800
        lvwDetail.ColumnHeaders.Add , "K2", "名称", 2000
        lvwDetail.ColumnHeaders.Add , "K3", "规格", 1000
        lvwDetail.ColumnHeaders.Add , "K4", "性质", 1000
    End If
End Sub

Private Sub addLvw(ByVal rsTemp As ADODB.Recordset)
    With lvwDetail
        .ListItems.Add , "K" & rsTemp!ID, rsTemp!编码, IIf(bln明细, "Fix", "Limit"), IIf(bln明细, "Fix", "Limit")
        .ListItems("K" & rsTemp!ID).SubItems(1) = rsTemp!名称
        If bln明细 Then
            .ListItems("K" & rsTemp!ID).SubItems(2) = Nvl(rsTemp!规格)
            .ListItems("K" & rsTemp!ID).SubItems(3) = cbo性质.Text
        Else
            .ListItems("K" & rsTemp!ID).SubItems(2) = cbo性质.Text
        End If
    End With
End Sub

Private Function Get项目(ByVal strID As String, rsTemp As ADODB.Recordset) As Boolean
'功能：根据项目ID，得到项目内容
    On Error GoTo errHandle
    
    If Trim(strID) = "" Then Exit Function
    If bln明细 Then
        gstrSQL = "Select ID,编码,名称,规格,类别 From 收费细目 Where ID=" & strID
    Else
        gstrSQL = "Select ID,编码,名称 From 保险支付大类 Where ID=" & strID
    End If
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Get项目 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Txt结束编码_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Txt结束编码.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = Txt结束编码.Text
    If bln明细 Then
        gstrSQL = "Select Distinct Rownum as 序号, A.ID,A.编码,B.名称,B.简码,A.类别,A.规格 " & _
                 "   FROM 收费细目 A,收费别名 B WHERE A.ID=B.收费细目ID And (" & _
                    zlCommFun.GetLike("A", "编码", strText) & " or " & zlCommFun.GetLike("B", "名称", strText) & " or " & zlCommFun.GetLike("B", "简码", strText) & ")"
    Else
        gstrSQL = "Select Distinct Rownum as 序号, A.ID,A.编码,A.名称,A.简码 " & _
                 "   FROM 保险支付大类 A WHERE 险类=" & lng险类 & " And (" & _
                    zlCommFun.GetLike("A", "编码", strText) & " or " & zlCommFun.GetLike("B", "名称", strText) & " or " & zlCommFun.GetLike("B", "简码", strText) & ")"
    End If
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(lng险类, rsTemp, "序号", "选择器", "请选择项目：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        zlControl.TxtSelAll Txt结束编码
        Exit Sub
    Else
        Txt结束编码.Text = "[" & rsTemp!编码 & "]" & rsTemp!名称
        Txt结束编码.Tag = rsTemp!编码
    End If
    zlControl.TxtSelAll Txt结束编码
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Txt开始编码_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Txt开始编码.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = Txt开始编码.Text
    If bln明细 Then
        gstrSQL = "Select Distinct Rownum as 序号, A.ID,A.编码,B.名称,B.简码,A.类别,A.规格 " & _
                 "   FROM 收费细目 A,收费别名 B WHERE A.ID=B.收费细目ID And (" & _
                    zlCommFun.GetLike("A", "编码", strText) & " or " & zlCommFun.GetLike("B", "名称", strText) & " or " & zlCommFun.GetLike("B", "简码", strText) & ")"
    Else
        gstrSQL = "Select Distinct Rownum as 序号, A.ID,A.编码,A.名称,A.简码 " & _
                 "   FROM 保险支付大类 A WHERE 险类=" & lng险类 & " And (" & _
                    zlCommFun.GetLike("A", "编码", strText) & " or " & zlCommFun.GetLike("B", "名称", strText) & " or " & zlCommFun.GetLike("B", "简码", strText) & ")"
    End If
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(lng险类, rsTemp, "序号", "选择器", "请选择项目：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        zlControl.TxtSelAll Txt开始编码
        Exit Sub
    Else
        Txt开始编码.Text = "[" & rsTemp!编码 & "]" & rsTemp!名称
        Txt开始编码.Tag = rsTemp!编码
    End If
    zlControl.TxtSelAll Txt开始编码
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
