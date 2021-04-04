VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDoctorEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择医生"
   ClientHeight    =   4260
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6780
   Icon            =   "frmDoctorEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ListView lvw 
      Height          =   3030
      Left            =   90
      TabIndex        =   2
      Top             =   375
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   5345
      View            =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "医生名称"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "编号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "性别"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "民族"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "职务"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "列表(&L)"
      Height          =   350
      Index           =   0
      Left            =   4575
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2370
      Width           =   1100
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "帮助(&H)"
      Height          =   350
      Index           =   1
      Left            =   4620
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2775
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3795
      TabIndex        =   8
      Top             =   3855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3795
      TabIndex        =   7
      Top             =   3465
      Width           =   1100
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      ItemData        =   "frmDoctorEdit.frx":000C
      Left            =   1140
      List            =   "frmDoctorEdit.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3855
      Width           =   2565
   End
   Begin VB.TextBox txt 
      Height          =   300
      Left            =   1140
      TabIndex        =   4
      Top             =   3465
      Width           =   2550
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      ItemData        =   "frmDoctorEdit.frx":0010
      Left            =   1065
      List            =   "frmDoctorEdit.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Width           =   3015
   End
   Begin MSComctlLib.Toolbar tbrThis 
      Height          =   345
      Left            =   4155
      TabIndex        =   9
      Top             =   15
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   609
      ButtonWidth     =   1349
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ils16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "列表"
            Key             =   "列表"
            Object.Tag             =   "列表"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List0"
                  Text            =   "大图标"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List1"
                  Text            =   "小图标"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List2"
                  Text            =   "列表"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List3"
                  Text            =   "详细资料"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "帮助"
            Key             =   "帮助"
            Object.Tag             =   "帮助"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   5370
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorEdit.frx":0014
            Key             =   "person"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   6135
      Top             =   3705
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
            Picture         =   "frmDoctorEdit.frx":0330
            Key             =   "person"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorEdit.frx":064C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorEdit.frx":09E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "医生名称(&N)"
      Height          =   210
      Left            =   75
      TabIndex        =   3
      Top             =   3525
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "职务类别(&T)"
      Height          =   210
      Left            =   75
      TabIndex        =   5
      Top             =   3900
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "部门名称(D)"
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   1245
   End
End
Attribute VB_Name = "frmDoctorEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintColumn As Integer
Private mblnFirst As Boolean
Private mOK As Boolean

Private mvarDefaultDept As String
Private mvarDefaultDuty As String

Public Function OpenDoctorDialog(frmMain As Object, DefaultDept As String, DefaultDuty As String) As Boolean
    
    mvarDefaultDept = DefaultDept
    mvarDefaultDuty = DefaultDuty
    
    frmDoctorEdit.Show 1, frmMain
    
    DefaultDept = mvarDefaultDept
    DefaultDuty = mvarDefaultDuty
    OpenDoctorDialog = mOK
    
End Function

Private Sub cbo_Click()
    Call cboDept_Click
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboDept_Click()
    If mblnFirst Then Exit Sub
    
    Call LoadDoctorList(cboDept.ItemData(cboDept.ListIndex), cbo.Text)
    If Not (lvw.SelectedItem Is Nothing) Then
        lvw.ListItems(1).Selected = True
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdCancel_Click()
    mvarDefaultDept = cboDept.Text
    mvarDefaultDuty = cbo.Text
    Unload Me
End Sub

Private Sub cmdMenu_Click(Index As Integer)
    Select Case Index
    Case 0
        Call tbrThis_ButtonClick(tbrThis.Buttons("列表"))
        lvw.SetFocus
    Case 1
        Call tbrThis_ButtonClick(tbrThis.Buttons("帮助"))
        lvw.SetFocus
    End Select
End Sub

Private Sub cmdOK_Click()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If SaveData Then
        Call SetLvwItemForeColor(lvw.SelectedItem, &HFF0000)
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    
    DoEvents
    
    '装载部门
    cbo.AddItem "所有职务"
    cbo.ListIndex = 0
    
    cboDept.AddItem "所有部门"
    cboDept.ListIndex = 0
    
    Call LoadDuty
    Call LoadDept
    
    mblnFirst = False
    
    Call cboDept_Click
    
End Sub

Private Sub Form_Load()
    mblnFirst = True
    RestoreWinState Me, App.ProductName
    mOK = False
End Sub

Private Sub LoadDept()
    On Error GoTo errHand
    gstrSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码 " & _
                " from 部门表 A,部门性质说明 B " & _
                " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " and B.部门ID=A.ID and B.服务对象 IN(1,2,3) And " & GetNodeCheckSQL("a.站点") & " " & _
                " Order by A.编码"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            cboDept.AddItem gRs!编码 & "-" & gRs!名称
            cboDept.ItemData(cboDept.ListCount - 1) = gRs!ID
            gRs.MoveNext
        Wend
    End If
    
    On Error Resume Next
    If mvarDefaultDept <> "" Then cboDept.Text = mvarDefaultDept
    
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDoctorList(ByVal lngDept As Long, ByVal strType As String)
    '加载指定部门的医生列表
    Dim Itmx As ListItem
    
    On Error GoTo errHand
    
    lvw.ListItems.Clear
    txt.Text = ""
    
    If strType = "所有职务" Then
        gstrSQL = "select B.部门id,A.ID,A.编号,A.姓名,A.性别,A.民族,A.专业技术职务,D.人员id from 人员表 A,部门人员 B,人员性质说明 C,咨询专家清单 D where D.人员id(+)=A.id and B.缺省=1 and A.id=B.人员id and C.人员id=A.id And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And " & GetNodeCheckSQL("a.站点") & " and C.人员性质='医生'" & IIf(lngDept = 0, "", " and B.部门id=[1]")
    Else
        gstrSQL = "select B.部门id,A.ID,A.编号,A.姓名,A.性别,A.民族,A.专业技术职务,D.人员id from 人员表 A,部门人员 B,人员性质说明 C,咨询专家清单 D where D.人员id(+)=A.id and B.缺省=1 and A.id=B.人员id and C.人员id=A.id And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And " & GetNodeCheckSQL("a.站点") & " and C.人员性质='医生'" & IIf(lngDept = 0, "", " and B.部门id=[1]") & " and A.专业技术职务=[2]"
    End If
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDept, strType)
    If gRs.BOF = False Then
        While Not gRs.EOF
            Set Itmx = lvw.ListItems.Add(, "K" & gRs!ID, IIf(IsNull(gRs!姓名), "", gRs!姓名), "person", "person")
            Itmx.Tag = IIf(IsNull(gRs!部门ID), 0, gRs!部门ID)
            Itmx.SubItems(1) = IIf(IsNull(gRs!编号), "", gRs!编号)
            Itmx.SubItems(2) = IIf(IsNull(gRs!性别), "", gRs!性别)
            Itmx.SubItems(3) = IIf(IsNull(gRs!民族), "", gRs!民族)
            Itmx.SubItems(4) = IIf(IsNull(gRs!专业技术职务), "", gRs!专业技术职务)
            
            If IsNull(gRs!人员ID) = False Then Call SetLvwItemForeColor(Itmx, &HFF0000)
            
            gRs.MoveNext
        Wend
    End If
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDuty()
    On Error GoTo errHand
    gstrSQL = "Select 编码,名称,简码,是否选择 from 专业技术职务 where 是否选择=1"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            cbo.AddItem IIf(IsNull(gRs!名称), "", gRs!名称)
            gRs.MoveNext
        Wend
    End If
    
    On Error Resume Next
    If mvarDefaultDuty <> "" Then cbo.Text = mvarDefaultDuty
    
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Function SaveData() As Boolean
        
    If lvw.SelectedItem Is Nothing Then Exit Function
    
    gstrSQL = "zl_咨询专家清单_insert(" & NextValue("咨询专家清单", "序号") & "," & Val(Mid(lvw.SelectedItem.Key, 2)) & "," & Val(lvw.SelectedItem.Tag) & ")"
    
    On Error GoTo errHand
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    
    '刷新父窗体
    Call frmDoctor.AddLvwItem(Val(Mid(lvw.SelectedItem.Key, 2)))
    
    SaveData = True
    
    Exit Function
errHand:
    If ErrCenter() = -1 Then Resume
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvw.SortKey = mintColumn
        lvw.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw_DblClick()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    Call cmdOK_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txt.Text = Item.Text
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
   
    Select Case Button.Key
    Case "列表"
        If lvw.View = 3 Then
            lvw.View = 0
        Else
            lvw.View = lvw.View + 1
        End If
    Case "帮助"
        ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
    End Select
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "List0"
        lvw.View = 0
    Case "List1"
        lvw.View = 1
    Case "List2"
        lvw.View = 2
    Case "List3"
        lvw.View = 3
    End Select
End Sub

Private Sub txt_GotFocus()
    SelAll txt
    zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim intLen As Long
    
    
    For i = 1 To lvw.ListItems.Count
        intLen = Len(txt.Text)
        If Mid(lvw.ListItems(i).Text, 1, intLen) = txt.Text Then
            lvw.ListItems(i).Selected = True
            Exit Sub
        End If
    Next
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub SetLvwItemForeColor(ByVal Itmx As ListItem, ByVal Color As Long)
'设置ListView项的前景色

    Dim i As Long
    
    Itmx.ForeColor = Color
    For i = 1 To Itmx.ListSubItems.Count
        Itmx.ListSubItems(i).ForeColor = Color
    Next
    
End Sub




