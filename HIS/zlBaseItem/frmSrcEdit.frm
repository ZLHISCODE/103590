VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSrcEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数据源编辑"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmSrcEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6345
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdData 
      Caption         =   "设置(&S)"
      Height          =   350
      Left            =   2740
      TabIndex        =   5
      ToolTipText     =   "设置数据导入方式"
      Top             =   5070
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择ODBC数据源"
      Height          =   3135
      Left            =   720
      TabIndex        =   15
      Top             =   1680
      Width           =   5415
      Begin MSComctlLib.ListView lvwDSN 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&A)"
         Height          =   350
         Left            =   1960
         TabIndex        =   9
         ToolTipText     =   "增加ODBC数据源"
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   4200
         TabIndex        =   11
         ToolTipText     =   "删除ODBC数据源"
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdModi 
         Caption         =   "修改(&M)"
         Height          =   350
         Left            =   3075
         TabIndex        =   10
         ToolTipText     =   "设置ODBC数据源"
         Top             =   200
         Width           =   1100
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3855
      TabIndex        =   6
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   270
      Picture         =   "frmSrcEdit.frx":058A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4980
      TabIndex        =   7
      Top             =   5070
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   4920
      Width           =   6315
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   585
      Width           =   5505
   End
   Begin VB.TextBox txt说明 
      Height          =   555
      Left            =   1425
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1110
      Width           =   4575
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   1425
      MaxLength       =   50
      TabIndex        =   1
      Top             =   720
      Width           =   4590
   End
   Begin VB.Label lblNote 
      Caption         =   "系统通过本机ODBC数据源可实现各类医价数据的导入，用户可在此设定系统与医价数据文件的连接和导入方式。"
      Height          =   345
      Left            =   735
      TabIndex        =   13
      Top             =   120
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmSrcEdit.frx":06D4
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl说明 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "说明(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   2
      Top             =   1125
      Width           =   630
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   780
      Width           =   630
   End
End
Attribute VB_Name = "frmSrcEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSourceSQL As String, strDestFields As String, ifDeleteData As Boolean
Private strNewName As String

Private ifOK As Boolean
Private OldSourceName As String, OldDSN As String
Public Function EditSource(ByVal frmParent As Object, ByRef SourceName As String) As Boolean
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    On Error Resume Next
    EditSource = False
    OldSourceName = SourceName
    strSourceSQL = "": strDestFields = "": ifDeleteData = False: OldDSN = "": strNewName = ""
    
    If Len(SourceName) > 0 Then
        Me.txt名称 = SourceName
        Me.txt说明 = GetSetting("ZLSOFT", "医价数据\" & SourceName, "说明", "")
        OldDSN = GetSetting("ZLSOFT", "医价数据\" & SourceName, "ODBC", "")
        
        strSourceSQL = GetSetting("ZLSOFT", "医价数据\" & SourceName, "数据源", "")
        strDestFields = GetSetting("ZLSOFT", "医价数据\" & SourceName, "字段", "")
        ifDeleteData = GetSetting("ZLSOFT", "医价数据\" & SourceName, "清除数据", "false")
    End If
    
    ListSource True
    '显示窗体
    Me.Show 1, frmParent
    EditSource = ifOK: If EditSource Then SourceName = strNewName
End Function

Private Sub cmdAdd_Click()
    Dim curIndex As Long
    If CreateDataSource(Me.hWnd) Then
        curIndex = lvwDSN.SelectedItem.Index
        Call ListSource: lvwDSN.SelectedItem = lvwDSN.ListItems(curIndex): lvwDSN.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdDel_Click()
    Dim curIndex As Long
    If lvwDSN.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("是否删除数据源：" + lvwDSN.SelectedItem.Text + "？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    If RemoveDataSource(Me.hWnd, lvwDSN.SelectedItem.Text, lvwDSN.SelectedItem.SubItems(1)) Then
        curIndex = lvwDSN.SelectedItem.Index
        Call ListSource
        
        On Error Resume Next
        If curIndex > lvwDSN.ListItems.Count - 1 Then curIndex = curIndex - 1
        If curIndex > -1 Then lvwDSN.SelectedItem = lvwDSN.ListItems(curIndex)
        lvwDSN.SetFocus
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    If Trim(Me.txt名称.Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或" & CInt(Me.txt名称.MaxLength / 2) & "个汉字）！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    End If
    If Me.lvwDSN.SelectedItem Is Nothing Then
        MsgBox "请选择一个ODBC数据源！", vbInformation, gstrSysName: Exit Sub
    End If
    If Len(Trim(strSourceSQL)) = 0 Then
        If MsgBox("未设置数据导入方式，是否继续？", vbDefaultButton2 + vbYesNo + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    '判断名称是否重复
    If UCase(Trim(OldSourceName)) <> UCase(Trim(Me.txt名称)) And _
        GetSetting("ZLSOFT", "医价数据", UCase(Trim(Me.txt名称)), Chr(0)) <> Chr(0) Then _
            MsgBox "数据源：" & Me.txt名称 & "已存在，请重新命名", vbInformation, gstrSysName: _
            Me.txt名称.SetFocus: Exit Sub
    
    If Len(OldSourceName) > 0 Then
        Call DeleteSetting("ZLSOFT", "医价数据", UCase(OldSourceName))
        Call DeleteSetting("ZLSOFT", "医价数据\" & UCase(OldSourceName))
    End If
    Call SaveSetting("ZLSOFT", "医价数据", UCase(Trim(Me.txt名称)), "1")
    Call SaveSetting("ZLSOFT", "医价数据\" & UCase(Trim(Me.txt名称)), "说明", Me.txt说明)
    Call SaveSetting("ZLSOFT", "医价数据\" & UCase(Trim(Me.txt名称)), "ODBC", Me.lvwDSN.SelectedItem.Text)
    Call SaveSetting("ZLSOFT", "医价数据\" & UCase(Trim(Me.txt名称)), "数据源", strSourceSQL)
    Call SaveSetting("ZLSOFT", "医价数据\" & UCase(Trim(Me.txt名称)), "字段", strDestFields)
    Call SaveSetting("ZLSOFT", "医价数据\" & UCase(Trim(Me.txt名称)), "清除数据", CStr(ifDeleteData))
    
    ifOK = True: strNewName = UCase(Me.txt名称)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdModi_Click()
    Dim curIndex As Long
    If lvwDSN.SelectedItem Is Nothing Then Exit Sub
    
    If ConfigDataSource(Me.hWnd, lvwDSN.SelectedItem.Text, lvwDSN.SelectedItem.SubItems(1)) Then
        curIndex = lvwDSN.SelectedItem.Index
        Call ListSource: lvwDSN.SelectedItem = lvwDSN.ListItems(curIndex): lvwDSN.SetFocus
    End If
End Sub

Private Sub cmdData_Click()
    Dim strSQL As String, strFlds As String, ifClear As Boolean
    
    If lvwDSN.SelectedItem Is Nothing Then Exit Sub
    
    frmDataSet.ShowMe Me, lvwDSN.SelectedItem.Text, strSQL, strFlds, ifClear
    If Len(Trim(strSQL)) > 0 Then
        strSourceSQL = strSQL: strDestFields = strFlds: ifDeleteData = ifClear
    End If
End Sub

Private Sub Form_Activate()
    '提取执行项目的信息
    Err = 0: On Error GoTo ErrHand
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_Load()
    ifOK = False
    With lvwDSN.ColumnHeaders
        .Add , "_Name", "数据源", 1800
        .Add , "_Desc", "说明", 3000
    End With
    lvwDSN.Sorted = True
End Sub

Private Sub ListSource(Optional ByVal ifInit As Boolean = False)
    Dim strDrivers As String, aDrivers() As String
    Dim i As Integer, tmpItem As ListItem, aSourceInfo() As String
    lvwDSN.ListItems.Clear
    
    strDrivers = GetODBCSources
    If Len(strDrivers) > 0 Then
        aDrivers = Split(strDrivers, Chr(0) + Chr(0))

        For i = 0 To UBound(aDrivers, 1)
            aSourceInfo = Split(aDrivers(i), Chr(0))
            Set tmpItem = lvwDSN.ListItems.Add(, "_" & i, aSourceInfo(0))
            tmpItem.SubItems(1) = aSourceInfo(1)
            
            If ifInit And UCase(aSourceInfo(0)) = UCase(OldDSN) Then tmpItem.Selected = True
        Next
    End If
End Sub

Private Sub lvwDSN_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwDSN
        .SortKey = ColumnHeader.Index - 1: .SortOrder = (.SortOrder + 1) Mod 2: .Sorted = True
    End With
End Sub

Private Sub lvwDSN_DblClick()
    cmdModi_Click
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_LostFocus()
    Me.txt说明.Text = Replace(Me.txt说明, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub


