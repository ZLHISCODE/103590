VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeptCharacter 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "部门性质分类"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList ils16 
      Left            =   6750
      Top             =   1170
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
            Picture         =   "frmDeptCharacter.frx":0000
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptCharacter.frx":031C
            Key             =   "Dept_No"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOrien 
      Caption         =   "定位(&L)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   1590
      TabIndex        =   4
      Top             =   4950
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3075
      Left            =   270
      TabIndex        =   3
      Top             =   930
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   5424
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "名称"
         Object.Tag             =   "名称"
         Text            =   "名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "编码"
         Object.Tag             =   "编码"
         Text            =   "编码"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "服务对象"
         Object.Tag             =   "服务对象"
         Text            =   "服务对象"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   4335
      Left            =   90
      TabIndex        =   2
      Top             =   240
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   7646
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   180
      TabIndex        =   1
      Top             =   4950
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   6630
      TabIndex        =   0
      Top             =   4950
      Width           =   1100
   End
End
Attribute VB_Name = "frmDeptCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintColumn As Integer '
Dim mstrTab As String '当前的工作性质

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOrien_Click()
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    frmDeptManage.tvwMain_S.Nodes(lvwMain.SelectedItem.Key).Selected = True
    frmDeptManage.tvwMain_S_NodeClick frmDeptManage.tvwMain_S.SelectedItem
    frmDeptManage.tvwMain_S.Nodes(lvwMain.SelectedItem.Key).EnsureVisible
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub cmdCancel_Click()
    mintColumn = 0
    mstrTab = ""
    Unload Me
End Sub


Public Sub 显示部门()
    
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockReadOnly
    
    gstrSQL = "select 名称 from 部门性质分类"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    tabMain.Tabs.Clear
    Do Until rsTemp.EOF
        tabMain.Tabs.Add , rsTemp("名称"), rsTemp("名称")
        rsTemp.MoveNext
    Loop
    
    frmDeptCharacter.Show , frmDeptManage
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    lvwMain.View = lvwReport
End Sub

Private Sub Form_Activate()
    tabMain.Tabs(1).Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdCancel.Left = ScaleWidth - cmdOrien.Width - 250
    cmdCancel.Top = ScaleHeight - cmdOrien.Height - 250
    cmdOrien.Top = ScaleHeight - cmdOrien.Height - 250
    cmdHelp.Top = ScaleHeight - cmdOrien.Height - 250
    
    tabMain.Top = 100
    tabMain.Left = 100
    tabMain.Height = cmdCancel.Top - 200 - tabMain.Top
    tabMain.Width = ScaleWidth - 100 - tabMain.Left
    
    lvwMain.Top = tabMain.ClientTop
    lvwMain.Left = tabMain.ClientLeft
    lvwMain.Height = tabMain.ClientHeight
    lvwMain.Width = tabMain.ClientWidth
    
    
End Sub

Private Sub tabMain_Click()
    Dim lngCol  As Long
    Dim varValue As Variant
    Dim lst As ListItem
    Dim str停用 As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If tabMain.SelectedItem.Caption = mstrTab Then Exit Sub
    '刷新
    mstrTab = tabMain.SelectedItem.Caption
    If frmDeptManage.mnuViewShowStop.Checked = False Then
        str停用 = " (A.撤档时间 is null or A.撤档时间 = to_date('3000-01-01','YYYY-MM-DD')) and "
    End If
    gstrSQL = "select A.ID,A.编码,A.名称,A.撤档时间,B.服务对象 from 部门表 A,部门性质说明 B where " & str停用 _
        & " A.ID=B.部门ID and B.工作性质=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrTab)
        
    lvwMain.ListItems.Clear
    Do Until rsTemp.EOF
        If CDate(IIF(IsNull(rsTemp("撤档时间")), CDate("3000/1/1"), rsTemp("撤档时间"))) = CDate("3000/1/1") Then
            str停用 = "Dept"
        Else
            str停用 = "Dept_No"
        End If
        
        Set lst = lvwMain.ListItems.Add(, "C" & rsTemp("ID"), rsTemp("名称"), str停用, str停用)
        If str停用 = "Dept_No" Then lst.ForeColor = RGB(255, 0, 0)
        
        '根据ListView的列名从数据库取数
        For lngCol = 2 To lvwMain.ColumnHeaders.Count
            varValue = rsTemp(lvwMain.ColumnHeaders(lngCol).Text).Value
            If lvwMain.ColumnHeaders(lngCol).Text = "服务对象" Then
                Select Case varValue
                    Case 1
                       varValue = "门诊病人"
                    Case 2
                       varValue = "住院病人"
                    Case 3
                       varValue = "门诊和住院病人"
                    Case Else
                       varValue = "不服务于病人"
                End Select
            End If
            lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
            If str停用 = "Dept_No" Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
        Next
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
