VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCooperateReg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicUnit 
      BorderStyle     =   0  'None
      Height          =   4365
      Left            =   360
      ScaleHeight     =   4365
      ScaleWidth      =   2700
      TabIndex        =   9
      Top             =   2520
      Width           =   2700
      Begin VB.ListBox lstUnits 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2910
         ItemData        =   "frmCooperateReg.frx":0000
         Left            =   960
         List            =   "frmCooperateReg.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label lblUnitTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合作单位"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   1020
      End
   End
   Begin VB.PictureBox picPlan 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   7905
      Left            =   2640
      ScaleHeight     =   7905
      ScaleWidth      =   8025
      TabIndex        =   0
      Top             =   0
      Width           =   8025
      Begin VB.CheckBox chkDisable 
         Caption         =   "本合作单位禁用该号别"
         Height          =   330
         Left            =   2400
         TabIndex        =   18
         Top             =   -15
         Width           =   2265
      End
      Begin VB.Frame fraLimit 
         Height          =   645
         Left            =   135
         TabIndex        =   15
         Top             =   1710
         Width           =   7845
         Begin VB.TextBox txtLimit 
            Height          =   345
            Left            =   1965
            TabIndex        =   17
            Top             =   195
            Width           =   1665
         End
         Begin VB.Label lblLimit 
            Caption         =   "当日合作单位限号数"
            Height          =   255
            Left            =   225
            TabIndex        =   16
            Top             =   240
            Width           =   1830
         End
      End
      Begin VB.Frame fraMove 
         Height          =   6495
         Left            =   3600
         TabIndex        =   1
         Top             =   360
         Width           =   975
         Begin VB.CommandButton cmdMoveSource 
            Caption         =   ">"
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton cmdMoveAllSource 
            Caption         =   ">>"
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton cmdMoveUnit 
            Caption         =   "<"
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1920
            Width           =   735
         End
         Begin VB.CommandButton cmdMoveAllUnit 
            Caption         =   "<<"
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2760
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView lvwSource 
         Height          =   6255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "序号"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "时间段"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.TabStrip tbWeekTime 
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   503
         Style           =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwReg 
         Height          =   6375
         Left            =   4440
         TabIndex        =   8
         Top             =   1560
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   11245
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "序号"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "时间段"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lbl已分配 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "已分配序号"
         Height          =   240
         Left            =   4560
         TabIndex        =   14
         Top             =   1320
         Width           =   3120
      End
      Begin VB.Label lbl未分配 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "未分配序号"
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   1320
         Width           =   3120
      End
      Begin VB.Label lblUnitRegTitle 
         Caption         =   "***:序号分配"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmCooperateReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const conPane_Info = 1
Private Const conPane_Plan = 3
Private Const conPane_Unit = 2
Private mlngPriItem As Long
Private mlng安排ID              As Long
Private mrs限号                 As ADODB.Recordset
Private mrs安排                 As ADODB.Recordset
Private mstr排班                As String '周日|全日||周一|白天||…………
Private mblnUnload As Boolean
Private mbln时段                As Boolean '如果安排设置了时段则严格按照时段来分配
Private mrs时间段               As ADODB.Recordset
Private mrsUnits              As ADODB.Recordset
Private mstrKey     As String
Private mrsSource   As ADODB.Recordset
Private mrsUnitsReg As ADODB.Recordset
Private mrsLimit As ADODB.Recordset
Private mrsDisable As ADODB.Recordset
Private mblnNoManual As Boolean

Public Event frmUnload(ByVal blnCancel As Boolean)
Private Sub cmdCancel_Click()
     RaiseEvent frmUnload(True)
End Sub


Private Sub chkDisable_Click()
    With chkDisable
        If .Value = 1 Then
            fraMove.Enabled = False
            fraLimit.Enabled = False
            lvwReg.Enabled = False
            lvwSource.Enabled = False
        Else
            fraMove.Enabled = True
            fraLimit.Enabled = True
            lvwReg.Enabled = True
            lvwSource.Enabled = True
        End If
        With mrsDisable
            .Filter = "合作单位='" & lstUnits.Text & "'"
            If .RecordCount <> 0 Then
                .MoveFirst
                .Delete adAffectCurrent
                .Update
            End If
            If chkDisable.Value = 1 Then
                .AddNew
                !合作单位 = lstUnits.Text
                .Update
            End If
        End With
    End With
End Sub

Private Sub cmdMoveAllSource_Click()
    Dim i As Long
    Dim lvwItem As ListItem
    Dim lvwitem1 As ListItem
    For i = 1 To lvwSource.ListItems.Count
        Set lvwItem = lvwSource.ListItems(i)
        
        Set lvwitem1 = lvwReg.ListItems.Add(, lvwItem.Key, lvwItem.Text)
        lvwitem1.SubItems(1) = lvwItem.SubItems(1)
        mrsSource.Filter = "限制项目='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and 序号=" & Val(lvwItem.Text)
        If mrsSource.RecordCount > 0 Then
            mrsSource.Delete adAffectCurrent
            mrsSource.Update
        End If
        mrsSource.Filter = 0
        InsertUnitReg lstUnits.Text, Val(lvwitem1.Text), mlng安排ID, Mid(Me.tbWeekTime.SelectedItem.Key, 2), 1, lvwitem1.SubItems(1)
    Next
    lvwSource.ListItems.Clear
    ClearLimit
End Sub

Private Sub cmdMoveAllUnit_Click()
     Dim lvwItem  As ListItem
    Dim lvwitem1 As ListItem
    Dim i        As Long
    If mrsSource Is Nothing Then
        With mrsSource
           Set mrsSource = New ADODB.Recordset
           mrsSource.Fields.Append "安排ID", adBigInt
           mrsSource.Fields.Append "限制项目", adVarChar, 10
           mrsSource.Fields.Append "序号", adBigInt, 18
           mrsSource.Fields.Append "数量", adBigInt, 18
           mrsSource.Fields.Append "时间段", adVarChar, 60
           mrsSource.CursorLocation = adUseClient
           mrsSource.LockType = adLockOptimistic
           mrsSource.CursorType = adOpenStatic
           mrsSource.Open
         End With
    End If
    For i = 1 To lvwReg.ListItems.Count
        Set lvwItem = lvwReg.ListItems(i)
        Set lvwitem1 = lvwSource.ListItems.Add(, lvwItem.Key, lvwItem.Text)
        mrsUnitsReg.Filter = "限制项目='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and 序号=" & Val(lvwItem.Text)
        With mrsSource
           .AddNew
           !安排ID = mrsUnitsReg!安排ID
           !限制项目 = mrsUnitsReg!限制项目
           !序号 = Val(lvwItem.Text)
           !时间段 = Nvl(mrsUnitsReg!时间段)
           !数量 = Val(mrsUnitsReg!数量)
           .Update
        End With
        mrsUnitsReg.Delete adAffectCurrent
        mrsUnitsReg.Update
        lvwitem1.SubItems(1) = lvwItem.SubItems(1)
        'UnitRegToSource Mid(Me.tbWeekTime.SelectedItem.Key, 2), Val(lvwitem1.Text), mlng安排ID, lvwitem1.SubItems(1), 1
    Next
    lvwReg.ListItems.Clear
    mrsSource.Filter = 0
    mrsUnitsReg.Filter = 0
    ClearLimit
End Sub

Private Sub cmdMoveSource_Click()

    Dim lvwItem  As ListItem
    Dim lvwitem1 As ListItem
    If lvwSource.SelectedItem Is Nothing Then Exit Sub
    
    Set lvwItem = lvwSource.SelectedItem
    Set lvwitem1 = lvwReg.ListItems.Add(, lvwItem.Key, lvwItem.Text)
    lvwitem1.SubItems(1) = lvwItem.SubItems(1)
    lvwSource.ListItems.Remove lvwItem.index
    mrsSource.Filter = "限制项目='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and 序号=" & Val(lvwItem.Text)
    If mrsSource.RecordCount > 0 Then
        mrsSource.Delete adAffectCurrent
        mrsSource.Update
    End If
    mrsSource.Filter = 0
    InsertUnitReg lstUnits.Text, Val(lvwitem1.Text), mlng安排ID, Mid(Me.tbWeekTime.SelectedItem.Key, 2), 1, lvwitem1.SubItems(1)
                
'    For i = 1 To lvwSource.ListItems.Count
'        If i <= lvwSource.ListItems.Count Then
'            Set lvwItem = lvwSource.ListItems(i)
'            If lvwItem.Checked Then
'                Set lvwitem1 = lvwReg.ListItems.Add(, lvwItem.Key, lvwItem.Text)
'                lvwitem1.SubItems(1) = lvwItem.SubItems(1)
'                lvwSource.ListItems.Remove lvwItem.Index
'                mrsSource.Filter = "限制项目='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and 序号=" & Val(lvwItem.Text)
'                If mrsSource.RecordCount > 0 Then
'                    mrsSource.Delete adAffectCurrent
'                    mrsSource.Update
'                End If
'                mrsSource.Filter = 0
'                i = i - 1
'                InsertUnitReg lstUnits.Text, Val(lvwitem1.Text), mlng安排ID, Mid(Me.tbWeekTime.SelectedItem.Key, 2), 1, lvwitem1.SubItems(1)
'            End If
'        End If
'
'    Next
    ClearLimit
End Sub

Private Sub cmdMoveUnit_Click()
    Dim lvwItem  As ListItem
    Dim lvwitem1 As ListItem
    Dim i        As Long
    If lvwReg.SelectedItem Is Nothing Then Exit Sub
    If mrsSource Is Nothing Then
        With mrsSource
           Set mrsSource = New ADODB.Recordset
           mrsSource.Fields.Append "安排ID", adBigInt
           mrsSource.Fields.Append "限制项目", adVarChar, 10
           mrsSource.Fields.Append "序号", adBigInt, 18
           mrsSource.Fields.Append "数量", adBigInt, 18
           mrsSource.Fields.Append "时间段", adVarChar, 60
           mrsSource.CursorLocation = adUseClient
           mrsSource.LockType = adLockOptimistic
           mrsSource.CursorType = adOpenStatic
           mrsSource.Open
         End With
     End If
     Set lvwItem = lvwReg.SelectedItem
     mrsUnitsReg.Filter = "限制项目='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and 序号=" & Val(lvwItem.Text)
     With mrsSource
        .AddNew
        !安排ID = mrsUnitsReg!安排ID
        !限制项目 = mrsUnitsReg!限制项目
        !序号 = Val(lvwItem.Text)
        !时间段 = Nvl(mrsUnitsReg!时间段)
        !数量 = Val(mrsUnitsReg!数量)
        .Update
     End With
     mrsUnitsReg.Delete adAffectCurrent
     mrsUnitsReg.Update
     lvwReg.ListItems.Remove lvwItem.index
     Set lvwitem1 = lvwSource.ListItems.Add(, lvwItem.Key, lvwItem.Text)
      
     mrsSource.Filter = 0
     mrsUnitsReg.Filter = 0
     ClearLimit
End Sub

 

Public Function SaveData() As Boolean
    Dim i       As Long
    Dim strSQL  As String
    Dim strTmp  As String
    Dim strLimit As String
    Dim strDisable As String
    Dim colExec As New Collection
    Dim str合作单位 As String
    Dim strInput As String
    Dim lngPosition As Long
    Dim strDivide As String
    Call txtLimit_Validate(False)
    Do While Not mrsUnits.EOF
        strSQL = "Zl_合作单位安排控制_Delete(" & mlng安排ID & ",'" & mrsUnits!名称 & "')"
        zlAddArray colExec, strSQL
        With mrsUnitsReg
                strTmp = ""
                strLimit = ""
                strDisable = ""
                mrsUnitsReg.Filter = "合作单位='" & mrsUnits!名称 & "' And 数量>0"
                Do While Not mrsUnitsReg.EOF
                    If strTmp <> "" Then strTmp = strTmp & "|"
                    strTmp = strTmp & !限制项目 & "," & !序号 & "," & !数量
                    mrsUnitsReg.MoveNext
                Loop
                mrsLimit.Filter = "合作单位='" & mrsUnits!名称 & "'"
                Do While Not mrsLimit.EOF
                    If strLimit <> "" Then strLimit = strLimit & "|"
                    strLimit = strLimit & mrsLimit!限制项目 & "," & mrsLimit!限制数量
                    mrsLimit.MoveNext
                Loop
                mrsDisable.Filter = "合作单位='" & mrsUnits!名称 & "'"
                If mrsDisable.RecordCount <> 0 Then
                    For i = 1 To tbWeekTime.Tabs.Count
                        If strDisable <> "" Then strDisable = strDisable & "|"
                        strDisable = strDisable & Mid(tbWeekTime.Tabs.Item(i).Key, 2)
                    Next i
                End If
                If strDisable <> "" Then
                    strSQL = "Zl_合作单位安排控制_Insert(" & mlng安排ID & ",'" & mrsUnits!名称 & "',Null,Null,'" & strDisable & "')"
                    zlAddArray colExec, strSQL
                Else
                    If strTmp <> "" Or strLimit <> "" Then
                        If zlCommFun.ActualLen(strTmp) > 3800 Then
                            strInput = strTmp
                            Do While zlCommFun.ActualLen(strTmp) > 3800
                                lngPosition = 2000 + InStr(Mid(strTmp, 2000), "|")
                                strDivide = Mid(strTmp, 1, lngPosition - 1)
                                strTmp = Mid(strTmp, lngPosition)
                                strSQL = "Zl_合作单位安排控制_Insert(" & mlng安排ID & ",'" & mrsUnits!名称 & "'," & IIf(strDivide = "", "Null,", "'" & strDivide & "',") & IIf(strLimit = "", "Null)", "'" & strLimit & "')")
                                zlAddArray colExec, strSQL
                            Loop
                            If strTmp <> "" Then
                                strSQL = "Zl_合作单位安排控制_Insert(" & mlng安排ID & ",'" & mrsUnits!名称 & "'," & IIf(strTmp = "", "Null,", "'" & strTmp & "',") & IIf(strLimit = "", "Null)", "'" & strLimit & "')")
                                zlAddArray colExec, strSQL
                            End If
                        Else
                            strSQL = "Zl_合作单位安排控制_Insert(" & mlng安排ID & ",'" & mrsUnits!名称 & "'," & IIf(strTmp = "", "Null,", "'" & strTmp & "',") & IIf(strLimit = "", "Null)", "'" & strLimit & "')")
                            zlAddArray colExec, strSQL
                        End If
                    End If
                End If
        End With
        mrsUnits.MoveNext
    Loop
    mrsUnitsReg.Filter = 0
    mrsDisable.Filter = 0
    strDisable = ""
    Do While Not mrsDisable.EOF
        strDisable = strDisable & "|" & mrsDisable!合作单位
        mrsDisable.MoveNext
    Loop
    If strDisable <> "" Then strDisable = Mid(strDisable, 2)
    zlDatabase.SetPara "禁用合作单位", strDisable, glngSys, 1110
    On Error GoTo Hd
    If colExec.Count > 0 Then zlExecuteProcedureArrAy colExec, Me.Caption
    SaveData = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
 
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case conPane_Plan
            Item.Handle = picPlan.Hwnd
        Case conPane_Unit
            Item.Handle = picUnit.Hwnd
    End Select
    
End Sub

Private Sub Form_Activate()
    If mblnUnload Then mblnUnload = False: Unload Me
End Sub

Private Sub Form_Load()
    mblnUnload = False
   ' Call InitPancel
End Sub

Public Function frmInit(ByVal lng安排ID As Long) As Boolean
    mlng安排ID = lng安排ID

    If InitData() = False Then Exit Function
    If InitRs() = False Then Exit Function
    If InitUntils() = False Then Exit Function
    If InitPage() = False Then Exit Function
End Function

Private Function InitPage() As Boolean
    Dim i         As Long
    Dim strList() As String
    If mstr排班 = "" Then Exit Function
    strList = Split(mstr排班, "||")
    tbWeekTime.Tabs.Clear
    For i = 0 To UBound(strList)
        tbWeekTime.Tabs.Add , "K" & Split(strList(i), "|")(0), Split(strList(i), "|")(0) & "(" & Split(strList(i), "|")(1) & ")"
    Next
    If tbWeekTime.Tabs.Count > 0 Then tbWeekTime.Tabs(1).Selected = True
    InitPage = True
End Function

Private Function InitUntils() As Boolean
    Dim strSQL As String
    Dim rsTmp  As ADODB.Recordset
    lstUnits.Clear
    strSQL = "Select 编码, 名称, 简码, 缺省标志 From 挂号合作单位 Order By 缺省标志 Desc"
    On Error GoTo Hd
    Set mrsUnits = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsUnits.EOF Then Exit Function
    
    Do While Not mrsUnits.EOF
        lstUnits.AddItem Nvl(mrsUnits!名称)
        mrsUnits.MoveNext
    Loop
    If lstUnits.ListCount > 0 Then lstUnits.Selected(0) = True
    mrsUnits.MoveFirst
    InitUntils = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'Private Sub InitPancel()
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:区哉设置
'    '编制:
'    '日期:2009-09-14 18:06:29
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim sngWidth As Single
'    Dim strReg   As String
'    Dim panThis  As Pane
'    Set panThis = dkpMan.CreatePane(conPane_Unit, 160, 600, DockBottomOf, panThis)
'    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption  'Or PaneNoHideable
'    panThis.Title = "合作单位"
'    panThis.Tag = conPane_Unit
'    panThis.Handle = PicUnit.hWnd
'    dkpMan.Options.ThemedFloatingFrames = False
'    dkpMan.Options.HideClient = False
'
'    Set panThis = dkpMan.CreatePane(conPane_Plan, 740, 600, DockRightOf, panThis)
'    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'    panThis.Title = "挂号安排"
'    panThis.Tag = conPane_Plan
'    panThis.Handle = picPlan.hWnd
'    dkpMan.Options.ThemedFloatingFrames = False
'    dkpMan.Options.HideClient = False
'    '    Set panThis = dkpMan.CreatePane(conPane_Plan, 250, 580, DockBottomOf, panThis)
'    '    panThis.Title = ""
'    '    panThis.Tag = conPane_Plan
'    '    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'    '    panThis.Handle = picPage.hWnd
'    '    dkpMan.Options.ThemedFloatingFrames = True
'    '    dkpMan.Options.HideClient = True
'    ' zlRestoreDockPanceToReg Me, dkpMan, "区域"
'
'End Sub

'------------------------------------------------------------------------
'页面调用过程与方法
'------------------------------------------------------------------------
Public Function InitData() As Boolean

    Dim strSQL As String
    Dim lng安排ID       As Long
    Dim i       As Long
    Dim strTemp As String
    If mlng安排ID = -1 Then Exit Function
    lng安排ID = mlng安排ID

    On Error GoTo Hd

    strSQL = " " & "   Select A.Id as 安排ID,0 as 计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id," & _
    "          A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六,nvl(A.默认时段间隔,5) As 默认时段间隔, " & "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室 " & "   From 挂号安排 A,收费项目目录 B,部门表 D " & "   Where A.项目id=b.Id(+) And A.科室id =d.Id(+) " & "         And A.Id=[1]"
    Set mrs安排 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
         
    If mrs安排.EOF Then
        ShowMsgbox "未找到指定的号别,请检查!"
        Exit Function
    End If
        
    mstr排班 = ""
    For i = 0 To 6
        strTemp = Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
        If Nvl(mrs安排("周" & strTemp)) <> "" Then
            If mstr排班 <> "" Then mstr排班 = mstr排班 & "||"
            mstr排班 = mstr排班 & "周" & strTemp & "|" & Nvl(mrs安排("周" & strTemp))
        End If
    Next
        
    strSQL = "" & "   Select decode(星期,'周日',1,'周一',2,'周二',3,'周三',4,'周四',5,'周五',6,7) as 排序,星期,to_char(开始时间,'HH24')||':00' as 时点,序号,to_char(开始时间,'hh24:mi')||'-' ||to_char(结束时间,'hh24:mi') as 时间范围, " & "               限制数量,是否预约" & "   From  挂号安排时段 " & "   Where 安排ID=[1]" & "   Order by 排序,时点,序号"
    Set mrs时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
 
    If Not mrs时间段.EOF Then mbln时段 = True
    '挂号安排限制
    strSQL = "Select 限制项目,限号数,  限约数,限制项目 as 星期 From  挂号安排限制 where 安排ID=[1]  Order BY 限制项目      "
    Set mrs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
    InitData = True

    Exit Function

Hd:

    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
 
Private Sub InsertUnitReg(ByVal str合作单位 As String, ByVal lng序号 As Long, ByVal lng安排ID As Long, ByVal str限制项目 As String, ByVal lng数量 As Long, Optional ByVal str时间段 As String = "")

    If mrsUnitsReg Is Nothing Then
        Set mrsUnitsReg = New ADODB.Recordset
        mrsUnitsReg.Fields.Append "合作单位", adVarChar, 50
        mrsUnitsReg.Fields.Append "安排ID", adBigInt
        mrsUnitsReg.Fields.Append "限制项目", adVarChar, 10
        mrsUnitsReg.Fields.Append "序号", adBigInt, 18
        mrsUnitsReg.Fields.Append "数量", adBigInt, 18
        mrsUnitsReg.Fields.Append "时间段", adVarChar, 60
        mrsUnitsReg.CursorLocation = adUseClient
        mrsUnitsReg.LockType = adLockOptimistic
        mrsUnitsReg.CursorType = adOpenStatic
      
        mrsUnitsReg.Open
    End If

    mrsUnitsReg.Filter = "合作单位='" & str合作单位 & "' and 序号=" & lng序号 & " and  安排ID=" & lng安排ID & " And 限制项目='" & str限制项目 & "' And 数量=" & lng数量
    
    If mrsUnitsReg.RecordCount > 0 Then
        mrsUnitsReg.Filter = 0

        Exit Sub

    End If

    mrsUnitsReg.Filter = 0

    With mrsUnitsReg
        .Filter = 0
        .AddNew
        !合作单位 = str合作单位
        !安排ID = lng安排ID
        !限制项目 = str限制项目
        !序号 = lng序号
        !数量 = lng数量
        !时间段 = str时间段
        .Update
    End With

End Sub

Private Sub UnitRegToSource(ByVal str限制项目 As String, ByVal lng序号 As Long, ByVal lng安排ID As Long, ByVal str时间段 As String, ByVal lng数量 As Long)
     
    mrsUnitsReg.Filter = "限制项目='" & str限制项目 & "' and 序号=" & lng序号

    If mrsUnitsReg.RecordCount = 0 Then mrsUnitsReg.Filter = 0: Exit Sub
    mrsUnitsReg.Delete adAffectCurrent
    mrsUnitsReg.Update
    mrsUnitsReg.Filter = 0
     
    With mrsSource
        .AddNew
        !安排ID = lng安排ID
        !限制项目 = str限制项目
        !时间段 = str时间段
        !序号 = lng序号
        !数量 = lng数量
        .Update
    End With
    
End Sub

Private Function InitRs() As Boolean
    Dim i         As Long
    Dim j         As Long
    Dim strList() As String
    Dim lng限号数   As Long
    Dim lng限约数   As Long
    Dim rsTmp  As ADODB.Recordset
    Dim strSQL As String

    '初始化 数据集
    With mrsUnitsReg
        Set mrsUnitsReg = New ADODB.Recordset
        mrsUnitsReg.Fields.Append "合作单位", adVarChar, 50
        mrsUnitsReg.Fields.Append "安排ID", adBigInt
        mrsUnitsReg.Fields.Append "限制项目", adVarChar, 10
        mrsUnitsReg.Fields.Append "序号", adBigInt, 18
        mrsUnitsReg.Fields.Append "数量", adBigInt, 18
        mrsUnitsReg.Fields.Append "时间段", adVarChar, 60
        mrsUnitsReg.CursorLocation = adUseClient
        mrsUnitsReg.LockType = adLockOptimistic
        mrsUnitsReg.CursorType = adOpenStatic
        mrsUnitsReg.Open
    End With

    With mrsSource
        Set mrsSource = New ADODB.Recordset
        mrsSource.Fields.Append "安排ID", adBigInt
        mrsSource.Fields.Append "限制项目", adVarChar, 10
        mrsSource.Fields.Append "序号", adBigInt, 18
        mrsSource.Fields.Append "数量", adBigInt, 18
        mrsSource.Fields.Append "时间段", adVarChar, 60
        mrsSource.CursorLocation = adUseClient
        mrsSource.LockType = adLockOptimistic
        mrsSource.CursorType = adOpenStatic
        mrsSource.Open
    End With
    
    With mrsLimit
        Set mrsLimit = New ADODB.Recordset
        mrsLimit.Fields.Append "合作单位", adVarChar, 50
        mrsLimit.Fields.Append "限制项目", adVarChar, 10
        mrsLimit.Fields.Append "限制数量", adBigInt, 18
        mrsLimit.CursorLocation = adUseClient
        mrsLimit.LockType = adLockOptimistic
        mrsLimit.CursorType = adOpenStatic
        mrsLimit.Open
    End With
    
    With mrsDisable
        Set mrsDisable = New ADODB.Recordset
        mrsDisable.Fields.Append "合作单位", adVarChar, 50
        mrsDisable.CursorLocation = adUseClient
        mrsDisable.LockType = adLockOptimistic
        mrsDisable.CursorType = adOpenStatic
        mrsDisable.Open
    End With
    
    If mstr排班 = "" Then Exit Function
    strList = Split(mstr排班, "||")
    If mbln时段 Then
         '如果是分时段
         
        For i = 0 To UBound(strList)
            mrs时间段.Filter = "星期='" & Split(strList(i), "|")(0) & "' and 是否预约=1"
            If mrs时间段.RecordCount = 0 Then mrs时间段.Filter = "星期='" & Split(strList(i), "|")(0) & "'"
            
            If mrs时间段.RecordCount = 0 Then
               '如果没有设置时间段 不填写时间段
               mrs限号.Filter = "限制项目='" & Split(strList(i), "|")(0) & "'"

               If mrs限号.RecordCount = 0 Then
                   mrs限号.Filter = 0
               Else
                   lng限号数 = Val(Nvl(mrs限号!限号数))
                   lng限约数 = Val(Nvl(mrs限号!限约数))
                   If lng限约数 = 0 Then lng限约数 = lng限号数

                   '加载初始化数据
                   For j = 1 To lng限号数

                       With mrsSource
                           .AddNew
                           !安排ID = mlng安排ID
                           !限制项目 = Split(strList(i), "|")(0)
                           !序号 = j
                           !数量 = 1
                           .Update
                       End With

                   Next

               End If 'mrs限号.recourdcount
               
            Else    'mrs时间段.recordCount=0
                Do While Not mrs时间段.EOF
                    With mrsSource
                        .AddNew
                        !安排ID = mlng安排ID
                        !限制项目 = Split(strList(i), "|")(0)
                        !序号 = Val(Nvl(mrs时间段!序号))
                        !数量 = 1
                        !时间段 = mrs时间段!时间范围
                        .Update
                    End With
                    mrs时间段.MoveNext
                Loop
            End If
        Next
        mrs时间段.Filter = 0
    Else
    
        For i = 0 To UBound(strList)
           '如果没有设置时间段 不填写时间段
            mrs限号.Filter = "限制项目='" & Split(strList(i), "|")(0) & "'"
    
            If mrs限号.RecordCount = 0 Then
                mrs限号.Filter = 0
            Else
                lng限号数 = Val(Nvl(mrs限号!限号数))
                lng限约数 = Val(Nvl(mrs限号!限约数))
                If lng限约数 = 0 Then lng限约数 = lng限号数
                '加载初始化数据
                For j = 1 To lng限号数
                    With mrsSource
                        .AddNew
                        !安排ID = mlng安排ID
                        !限制项目 = Split(strList(i), "|")(0)
                        !序号 = j
                        !数量 = 1
                        .Update
                    End With
    
                Next
    
            End If 'mrs限号.recourdcount
        Next
    End If
    
    '已经分配序号
    strSQL = "Select 合作单位, 安排id, 限制项目, 序号, 数量 From 合作单位安排控制  Where 安排ID=[1] And 序号 <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)

    If rsTmp.RecordCount > 0 Then

        Do While Not rsTmp.EOF
            mrsSource.Filter = "限制项目='" & rsTmp!限制项目 & "' and 序号=" & rsTmp!序号

            With mrsUnitsReg
                .AddNew
                !合作单位 = Nvl(rsTmp!合作单位)
                !安排ID = mlng安排ID
                !限制项目 = Nvl(rsTmp!限制项目)
                !序号 = Val(Nvl(rsTmp!序号))
                !数量 = Val(Nvl(rsTmp!数量))

                If mrsSource.RecordCount > 0 Then
                    !时间段 = mrsSource!时间段
                    mrsSource.Delete
                    mrsSource.Update
                End If

                .Update
            End With

            mrsSource.Filter = 0
            rsTmp.MoveNext
        Loop
    
    End If
    
    strSQL = "Select 合作单位, 安排id, 限制项目, 序号, 数量 From 合作单位安排控制  Where 安排ID=[1] And 序号 = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
    
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            With mrsLimit
                .AddNew
                !合作单位 = Nvl(rsTmp!合作单位)
                !限制项目 = Nvl(rsTmp!限制项目)
                !限制数量 = Val(Nvl(rsTmp!数量))
                .Update
            End With
            rsTmp.MoveNext
        Loop
    End If
    
    strSQL = "Select Distinct 合作单位 From 合作单位安排控制  Where 安排ID=[1] And 数量 = 0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
    
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            With mrsDisable
                .AddNew
                !合作单位 = Nvl(rsTmp!合作单位)
                .Update
            End With
            rsTmp.MoveNext
        Loop
    End If
    
    InitRs = True
End Function

Private Sub Form_Resize()
     Err.Number = 0
     On Error Resume Next
     With Me.picUnit
         .Left = Me.ScaleLeft
         .Top = Me.ScaleTop
         .Height = Me.ScaleHeight
     End With
     With Me.picPlan
         .Left = picUnit.Left + picUnit.Width + 1 * Screen.TwipsPerPixelX
         .Top = Me.ScaleTop
         .Height = Me.ScaleHeight
         .Width = Me.ScaleWidth - .Left
     End With

End Sub

Private Sub lstUnits_Click()

    Static strUnits As String

    If lstUnits.Text = strUnits Then Exit Sub
    strUnits = lstUnits.Text
    lblUnitRegTitle.Caption = strUnits & ":预约分配"
    Call tbWeekTime_Click
End Sub

 

Private Sub picPlan_Resize()

    On Error Resume Next
    
    lblUnitRegTitle.Move 0, 0, picPlan.ScaleWidth, lblUnitRegTitle.Height
    chkDisable.Left = picPlan.ScaleWidth - chkDisable.Width
    Me.tbWeekTime.Move 0, lblUnitRegTitle.Height + Screen.TwipsPerPixelY, picPlan.ScaleWidth, Me.tbWeekTime.Height
    
    With lbl未分配
        .Left = Screen.TwipsPerPixelX * 2
        .Top = tbWeekTime.Height + tbWeekTime.Top + Screen.TwipsPerPixelY * 4
        .Width = lvwSource.Width
    End With
    
    With lvwSource
        .Left = Screen.TwipsPerPixelX * 2
        .Top = lbl未分配.Height + lbl未分配.Top ' + Screen.TwipsPerPixelY * 4
        .Height = Me.picPlan.ScaleHeight - lbl未分配.Height - lbl未分配.Top - Screen.TwipsPerPixelY * 2 - 700
    End With

    With fraMove
        .Left = lvwSource.Left + lvwSource.Width + Screen.TwipsPerPixelX * 2
        .Top = lvwSource.Top
        .Height = lvwSource.Top + lvwSource.Height - lbl未分配.Top - lbl未分配.Height
    End With
    
     With lbl已分配
        .Left = fraMove.Left + fraMove.Width + Screen.TwipsPerPixelX * 2
        .Top = lbl未分配.Top
        .Width = picPlan.ScaleWidth - fraMove.Width - lvwSource.Width - Screen.TwipsPerPixelX * 8
    End With
    
    With lvwReg
        .Left = fraMove.Left + fraMove.Width + Screen.TwipsPerPixelX * 2
        .Top = lvwSource.Top
        .Width = picPlan.ScaleWidth - fraMove.Width - lvwSource.Width - Screen.TwipsPerPixelX * 8
        .Height = lvwSource.Height
    End With
    
    With fraLimit
        .Left = lvwSource.Left
        .Top = lvwSource.Top + lvwSource.Height + 30
        .Width = picPlan.ScaleWidth - .Left - 15
    End With
    
End Sub

Private Sub PicUnit_Resize()
    On Error Resume Next
    lblUnitTitle.Move 0, 0, picUnit.ScaleWidth, lblUnitTitle.Height
    Me.lstUnits.Move 0, lblUnitTitle.Height, picUnit.ScaleWidth, picUnit.ScaleHeight - lblUnitTitle.Height
End Sub
 
Private Sub tbWeekTime_Click()
    Dim lvwItem    As ListItem
    Static strUnit As String
    If Not mrsLimit Is Nothing Then
        With mrsLimit
            .Filter = "限制项目='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "' and  合作单位='" & lstUnits.Text & "'"
            If .RecordCount = 0 Then
                mblnNoManual = True
                txtLimit.Text = ""
                mblnNoManual = False
            Else
                mblnNoManual = True
                txtLimit.Text = !限制数量
                mblnNoManual = False
            End If
        End With
    End If
    
    If Not mrsDisable Is Nothing Then
        With mrsDisable
            .Filter = "合作单位='" & lstUnits.Text & "'"
            If .RecordCount = 0 Then
                chkDisable.Value = 0
            Else
                chkDisable.Value = 1
            End If
        End With
    End If
    
    If mstrKey = tbWeekTime.SelectedItem.Key And strUnit = lstUnits.Text Then Exit Sub
    mstrKey = tbWeekTime.SelectedItem.Key
    strUnit = lstUnits.Text
        
        '如果序号设置了时段
        lvwSource.ListItems.Clear
        mrsSource.Filter = "限制项目='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "'"

        Do While Not mrsSource.EOF

            With lvwSource
                Set lvwItem = .ListItems.Add(, "k" & mrsSource!序号, mrsSource!序号)
                lvwItem.SubItems(1) = Nvl(mrsSource!时间段)
            End With

            mrsSource.MoveNext
        Loop

        mrsSource.Filter = 0
        lvwReg.ListItems.Clear
        mrsUnitsReg.Filter = "限制项目='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "' and  合作单位='" & lstUnits.Text & "'"

        Do While Not mrsUnitsReg.EOF

            With lvwReg
                Set lvwItem = .ListItems.Add(, "k" & mrsUnitsReg!序号, mrsUnitsReg!序号)
                lvwItem.SubItems(1) = Nvl(mrsUnitsReg!时间段)
            End With

            mrsUnitsReg.MoveNext
        Loop

        mrsUnitsReg.Filter = 0
End Sub

Private Sub txtLimit_Change()
    Dim lvwItem  As ListItem
    Dim lvwitem1 As ListItem
    Dim i        As Long
    If mblnNoManual Then Exit Sub
    If mrsSource Is Nothing Then
        With mrsSource
           Set mrsSource = New ADODB.Recordset
           mrsSource.Fields.Append "安排ID", adBigInt
           mrsSource.Fields.Append "限制项目", adVarChar, 10
           mrsSource.Fields.Append "序号", adBigInt, 18
           mrsSource.Fields.Append "数量", adBigInt, 18
           mrsSource.Fields.Append "时间段", adVarChar, 60
           mrsSource.CursorLocation = adUseClient
           mrsSource.LockType = adLockOptimistic
           mrsSource.CursorType = adOpenStatic
           mrsSource.Open
         End With
    End If
    For i = 1 To lvwReg.ListItems.Count
        Set lvwItem = lvwReg.ListItems(i)
        Set lvwitem1 = lvwSource.ListItems.Add(, lvwItem.Key, lvwItem.Text)
        mrsUnitsReg.Filter = "限制项目='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and 序号=" & Val(lvwItem.Text)
        With mrsSource
           .AddNew
           !安排ID = mrsUnitsReg!安排ID
           !限制项目 = mrsUnitsReg!限制项目
           !序号 = Val(lvwItem.Text)
           !时间段 = Nvl(mrsUnitsReg!时间段)
           !数量 = Val(mrsUnitsReg!数量)
           .Update
        End With
        mrsUnitsReg.Delete adAffectCurrent
        mrsUnitsReg.Update
        lvwitem1.SubItems(1) = lvwItem.SubItems(1)
        'UnitRegToSource Mid(Me.tbWeekTime.SelectedItem.Key, 2), Val(lvwitem1.Text), mlng安排ID, lvwitem1.SubItems(1), 1
    Next
    lvwReg.ListItems.Clear
    mrsSource.Filter = 0
    mrsUnitsReg.Filter = 0
End Sub

Private Sub txtLimit_GotFocus()
    zlControl.TxtSelAll txtLimit
End Sub

Private Sub txtLimit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub ClearLimit()
    mblnNoManual = True
    txtLimit.Text = ""
    mblnNoManual = False
    With mrsLimit
        .Filter = "限制项目='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "' and  合作单位='" & lstUnits.Text & "'"
        If .RecordCount <> 0 Then
            .MoveFirst
            .Delete adAffectCurrent
        End If
    End With
End Sub

Private Sub txtLimit_Validate(Cancel As Boolean)
    If mrs限号 Is Nothing Then Exit Sub
    mrs限号.Filter = "限制项目='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "'"
    If mrs限号.RecordCount = 0 Then Exit Sub
    If Val(txtLimit.Text) > mrs限号!限号数 Then
        MsgBox "设置的合作单位限号数不能超过当前安排的限号数！", vbInformation, gstrSysName
        Cancel = True
        If txtLimit.Enabled And txtLimit.Visible Then txtLimit.SetFocus
        Exit Sub
    End If
    With mrsLimit
        .Filter = "限制项目='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "' and  合作单位='" & lstUnits.Text & "'"
        If .RecordCount <> 0 Then
            .MoveFirst
            .Delete adAffectCurrent
            .Update
        End If
        If Val(Nvl(txtLimit.Text)) <> 0 Then
            .AddNew
            !合作单位 = lstUnits.Text
            !限制项目 = Mid(tbWeekTime.SelectedItem.Key, 2)
            !限制数量 = Val(txtLimit.Text)
            .Update
        End If
    End With
End Sub
