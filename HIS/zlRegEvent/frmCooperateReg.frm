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
   StartUpPosition =   3  '����ȱʡ
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
            Name            =   "����"
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
         Caption         =   "������λ"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��������λ���øúű�"
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
            Caption         =   "���պ�����λ�޺���"
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
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "���"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ʱ���"
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
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "���"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ʱ���"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lbl�ѷ��� 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�ѷ������"
         Height          =   240
         Left            =   4560
         TabIndex        =   14
         Top             =   1320
         Width           =   3120
      End
      Begin VB.Label lblδ���� 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "δ�������"
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   1320
         Width           =   3120
      End
      Begin VB.Label lblUnitRegTitle 
         Caption         =   "***:��ŷ���"
         BeginProperty Font 
            Name            =   "����"
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
Private mlng����ID              As Long
Private mrs�޺�                 As ADODB.Recordset
Private mrs����                 As ADODB.Recordset
Private mstr�Ű�                As String '����|ȫ��||��һ|����||��������
Private mblnUnload As Boolean
Private mblnʱ��                As Boolean '�������������ʱ�����ϸ���ʱ��������
Private mrsʱ���               As ADODB.Recordset
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
            .Filter = "������λ='" & lstUnits.Text & "'"
            If .RecordCount <> 0 Then
                .MoveFirst
                .Delete adAffectCurrent
                .Update
            End If
            If chkDisable.Value = 1 Then
                .AddNew
                !������λ = lstUnits.Text
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
        mrsSource.Filter = "������Ŀ='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and ���=" & Val(lvwItem.Text)
        If mrsSource.RecordCount > 0 Then
            mrsSource.Delete adAffectCurrent
            mrsSource.Update
        End If
        mrsSource.Filter = 0
        InsertUnitReg lstUnits.Text, Val(lvwitem1.Text), mlng����ID, Mid(Me.tbWeekTime.SelectedItem.Key, 2), 1, lvwitem1.SubItems(1)
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
           mrsSource.Fields.Append "����ID", adBigInt
           mrsSource.Fields.Append "������Ŀ", adVarChar, 10
           mrsSource.Fields.Append "���", adBigInt, 18
           mrsSource.Fields.Append "����", adBigInt, 18
           mrsSource.Fields.Append "ʱ���", adVarChar, 60
           mrsSource.CursorLocation = adUseClient
           mrsSource.LockType = adLockOptimistic
           mrsSource.CursorType = adOpenStatic
           mrsSource.Open
         End With
    End If
    For i = 1 To lvwReg.ListItems.Count
        Set lvwItem = lvwReg.ListItems(i)
        Set lvwitem1 = lvwSource.ListItems.Add(, lvwItem.Key, lvwItem.Text)
        mrsUnitsReg.Filter = "������Ŀ='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and ���=" & Val(lvwItem.Text)
        With mrsSource
           .AddNew
           !����ID = mrsUnitsReg!����ID
           !������Ŀ = mrsUnitsReg!������Ŀ
           !��� = Val(lvwItem.Text)
           !ʱ��� = Nvl(mrsUnitsReg!ʱ���)
           !���� = Val(mrsUnitsReg!����)
           .Update
        End With
        mrsUnitsReg.Delete adAffectCurrent
        mrsUnitsReg.Update
        lvwitem1.SubItems(1) = lvwItem.SubItems(1)
        'UnitRegToSource Mid(Me.tbWeekTime.SelectedItem.Key, 2), Val(lvwitem1.Text), mlng����ID, lvwitem1.SubItems(1), 1
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
    mrsSource.Filter = "������Ŀ='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and ���=" & Val(lvwItem.Text)
    If mrsSource.RecordCount > 0 Then
        mrsSource.Delete adAffectCurrent
        mrsSource.Update
    End If
    mrsSource.Filter = 0
    InsertUnitReg lstUnits.Text, Val(lvwitem1.Text), mlng����ID, Mid(Me.tbWeekTime.SelectedItem.Key, 2), 1, lvwitem1.SubItems(1)
                
'    For i = 1 To lvwSource.ListItems.Count
'        If i <= lvwSource.ListItems.Count Then
'            Set lvwItem = lvwSource.ListItems(i)
'            If lvwItem.Checked Then
'                Set lvwitem1 = lvwReg.ListItems.Add(, lvwItem.Key, lvwItem.Text)
'                lvwitem1.SubItems(1) = lvwItem.SubItems(1)
'                lvwSource.ListItems.Remove lvwItem.Index
'                mrsSource.Filter = "������Ŀ='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and ���=" & Val(lvwItem.Text)
'                If mrsSource.RecordCount > 0 Then
'                    mrsSource.Delete adAffectCurrent
'                    mrsSource.Update
'                End If
'                mrsSource.Filter = 0
'                i = i - 1
'                InsertUnitReg lstUnits.Text, Val(lvwitem1.Text), mlng����ID, Mid(Me.tbWeekTime.SelectedItem.Key, 2), 1, lvwitem1.SubItems(1)
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
           mrsSource.Fields.Append "����ID", adBigInt
           mrsSource.Fields.Append "������Ŀ", adVarChar, 10
           mrsSource.Fields.Append "���", adBigInt, 18
           mrsSource.Fields.Append "����", adBigInt, 18
           mrsSource.Fields.Append "ʱ���", adVarChar, 60
           mrsSource.CursorLocation = adUseClient
           mrsSource.LockType = adLockOptimistic
           mrsSource.CursorType = adOpenStatic
           mrsSource.Open
         End With
     End If
     Set lvwItem = lvwReg.SelectedItem
     mrsUnitsReg.Filter = "������Ŀ='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and ���=" & Val(lvwItem.Text)
     With mrsSource
        .AddNew
        !����ID = mrsUnitsReg!����ID
        !������Ŀ = mrsUnitsReg!������Ŀ
        !��� = Val(lvwItem.Text)
        !ʱ��� = Nvl(mrsUnitsReg!ʱ���)
        !���� = Val(mrsUnitsReg!����)
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
    Dim str������λ As String
    Dim strInput As String
    Dim lngPosition As Long
    Dim strDivide As String
    Call txtLimit_Validate(False)
    Do While Not mrsUnits.EOF
        strSQL = "Zl_������λ���ſ���_Delete(" & mlng����ID & ",'" & mrsUnits!���� & "')"
        zlAddArray colExec, strSQL
        With mrsUnitsReg
                strTmp = ""
                strLimit = ""
                strDisable = ""
                mrsUnitsReg.Filter = "������λ='" & mrsUnits!���� & "' And ����>0"
                Do While Not mrsUnitsReg.EOF
                    If strTmp <> "" Then strTmp = strTmp & "|"
                    strTmp = strTmp & !������Ŀ & "," & !��� & "," & !����
                    mrsUnitsReg.MoveNext
                Loop
                mrsLimit.Filter = "������λ='" & mrsUnits!���� & "'"
                Do While Not mrsLimit.EOF
                    If strLimit <> "" Then strLimit = strLimit & "|"
                    strLimit = strLimit & mrsLimit!������Ŀ & "," & mrsLimit!��������
                    mrsLimit.MoveNext
                Loop
                mrsDisable.Filter = "������λ='" & mrsUnits!���� & "'"
                If mrsDisable.RecordCount <> 0 Then
                    For i = 1 To tbWeekTime.Tabs.Count
                        If strDisable <> "" Then strDisable = strDisable & "|"
                        strDisable = strDisable & Mid(tbWeekTime.Tabs.Item(i).Key, 2)
                    Next i
                End If
                If strDisable <> "" Then
                    strSQL = "Zl_������λ���ſ���_Insert(" & mlng����ID & ",'" & mrsUnits!���� & "',Null,Null,'" & strDisable & "')"
                    zlAddArray colExec, strSQL
                Else
                    If strTmp <> "" Or strLimit <> "" Then
                        If zlCommFun.ActualLen(strTmp) > 3800 Then
                            strInput = strTmp
                            Do While zlCommFun.ActualLen(strTmp) > 3800
                                lngPosition = 2000 + InStr(Mid(strTmp, 2000), "|")
                                strDivide = Mid(strTmp, 1, lngPosition - 1)
                                strTmp = Mid(strTmp, lngPosition)
                                strSQL = "Zl_������λ���ſ���_Insert(" & mlng����ID & ",'" & mrsUnits!���� & "'," & IIf(strDivide = "", "Null,", "'" & strDivide & "',") & IIf(strLimit = "", "Null)", "'" & strLimit & "')")
                                zlAddArray colExec, strSQL
                            Loop
                            If strTmp <> "" Then
                                strSQL = "Zl_������λ���ſ���_Insert(" & mlng����ID & ",'" & mrsUnits!���� & "'," & IIf(strTmp = "", "Null,", "'" & strTmp & "',") & IIf(strLimit = "", "Null)", "'" & strLimit & "')")
                                zlAddArray colExec, strSQL
                            End If
                        Else
                            strSQL = "Zl_������λ���ſ���_Insert(" & mlng����ID & ",'" & mrsUnits!���� & "'," & IIf(strTmp = "", "Null,", "'" & strTmp & "',") & IIf(strLimit = "", "Null)", "'" & strLimit & "')")
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
        strDisable = strDisable & "|" & mrsDisable!������λ
        mrsDisable.MoveNext
    Loop
    If strDisable <> "" Then strDisable = Mid(strDisable, 2)
    zlDatabase.SetPara "���ú�����λ", strDisable, glngSys, 1110
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

Public Function frmInit(ByVal lng����ID As Long) As Boolean
    mlng����ID = lng����ID

    If InitData() = False Then Exit Function
    If InitRs() = False Then Exit Function
    If InitUntils() = False Then Exit Function
    If InitPage() = False Then Exit Function
End Function

Private Function InitPage() As Boolean
    Dim i         As Long
    Dim strList() As String
    If mstr�Ű� = "" Then Exit Function
    strList = Split(mstr�Ű�, "||")
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
    strSQL = "Select ����, ����, ����, ȱʡ��־ From �Һź�����λ Order By ȱʡ��־ Desc"
    On Error GoTo Hd
    Set mrsUnits = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsUnits.EOF Then Exit Function
    
    Do While Not mrsUnits.EOF
        lstUnits.AddItem Nvl(mrsUnits!����)
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
'    '����:��������
'    '����:
'    '����:2009-09-14 18:06:29
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim sngWidth As Single
'    Dim strReg   As String
'    Dim panThis  As Pane
'    Set panThis = dkpMan.CreatePane(conPane_Unit, 160, 600, DockBottomOf, panThis)
'    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption  'Or PaneNoHideable
'    panThis.Title = "������λ"
'    panThis.Tag = conPane_Unit
'    panThis.Handle = PicUnit.hWnd
'    dkpMan.Options.ThemedFloatingFrames = False
'    dkpMan.Options.HideClient = False
'
'    Set panThis = dkpMan.CreatePane(conPane_Plan, 740, 600, DockRightOf, panThis)
'    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'    panThis.Title = "�ҺŰ���"
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
'    ' zlRestoreDockPanceToReg Me, dkpMan, "����"
'
'End Sub

'------------------------------------------------------------------------
'ҳ����ù����뷽��
'------------------------------------------------------------------------
Public Function InitData() As Boolean

    Dim strSQL As String
    Dim lng����ID       As Long
    Dim i       As Long
    Dim strTemp As String
    If mlng����ID = -1 Then Exit Function
    lng����ID = mlng����ID

    On Error GoTo Hd

    strSQL = " " & "   Select A.Id as ����ID,0 as �ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id," & _
    "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,nvl(A.Ĭ��ʱ�μ��,5) As Ĭ��ʱ�μ��, " & "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ���� " & "   From �ҺŰ��� A,�շ���ĿĿ¼ B,���ű� D " & "   Where A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & "         And A.Id=[1]"
    Set mrs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
         
    If mrs����.EOF Then
        ShowMsgbox "δ�ҵ�ָ���ĺű�,����!"
        Exit Function
    End If
        
    mstr�Ű� = ""
    For i = 0 To 6
        strTemp = Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
        If Nvl(mrs����("��" & strTemp)) <> "" Then
            If mstr�Ű� <> "" Then mstr�Ű� = mstr�Ű� & "||"
            mstr�Ű� = mstr�Ű� & "��" & strTemp & "|" & Nvl(mrs����("��" & strTemp))
        End If
    Next
        
    strSQL = "" & "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & "               ��������,�Ƿ�ԤԼ" & "   From  �ҺŰ���ʱ�� " & "   Where ����ID=[1]" & "   Order by ����,ʱ��,���"
    Set mrsʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
 
    If Not mrsʱ���.EOF Then mblnʱ�� = True
    '�ҺŰ�������
    strSQL = "Select ������Ŀ,�޺���,  ��Լ��,������Ŀ as ���� From  �ҺŰ������� where ����ID=[1]  Order BY ������Ŀ      "
    Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    InitData = True

    Exit Function

Hd:

    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
 
Private Sub InsertUnitReg(ByVal str������λ As String, ByVal lng��� As Long, ByVal lng����ID As Long, ByVal str������Ŀ As String, ByVal lng���� As Long, Optional ByVal strʱ��� As String = "")

    If mrsUnitsReg Is Nothing Then
        Set mrsUnitsReg = New ADODB.Recordset
        mrsUnitsReg.Fields.Append "������λ", adVarChar, 50
        mrsUnitsReg.Fields.Append "����ID", adBigInt
        mrsUnitsReg.Fields.Append "������Ŀ", adVarChar, 10
        mrsUnitsReg.Fields.Append "���", adBigInt, 18
        mrsUnitsReg.Fields.Append "����", adBigInt, 18
        mrsUnitsReg.Fields.Append "ʱ���", adVarChar, 60
        mrsUnitsReg.CursorLocation = adUseClient
        mrsUnitsReg.LockType = adLockOptimistic
        mrsUnitsReg.CursorType = adOpenStatic
      
        mrsUnitsReg.Open
    End If

    mrsUnitsReg.Filter = "������λ='" & str������λ & "' and ���=" & lng��� & " and  ����ID=" & lng����ID & " And ������Ŀ='" & str������Ŀ & "' And ����=" & lng����
    
    If mrsUnitsReg.RecordCount > 0 Then
        mrsUnitsReg.Filter = 0

        Exit Sub

    End If

    mrsUnitsReg.Filter = 0

    With mrsUnitsReg
        .Filter = 0
        .AddNew
        !������λ = str������λ
        !����ID = lng����ID
        !������Ŀ = str������Ŀ
        !��� = lng���
        !���� = lng����
        !ʱ��� = strʱ���
        .Update
    End With

End Sub

Private Sub UnitRegToSource(ByVal str������Ŀ As String, ByVal lng��� As Long, ByVal lng����ID As Long, ByVal strʱ��� As String, ByVal lng���� As Long)
     
    mrsUnitsReg.Filter = "������Ŀ='" & str������Ŀ & "' and ���=" & lng���

    If mrsUnitsReg.RecordCount = 0 Then mrsUnitsReg.Filter = 0: Exit Sub
    mrsUnitsReg.Delete adAffectCurrent
    mrsUnitsReg.Update
    mrsUnitsReg.Filter = 0
     
    With mrsSource
        .AddNew
        !����ID = lng����ID
        !������Ŀ = str������Ŀ
        !ʱ��� = strʱ���
        !��� = lng���
        !���� = lng����
        .Update
    End With
    
End Sub

Private Function InitRs() As Boolean
    Dim i         As Long
    Dim j         As Long
    Dim strList() As String
    Dim lng�޺���   As Long
    Dim lng��Լ��   As Long
    Dim rsTmp  As ADODB.Recordset
    Dim strSQL As String

    '��ʼ�� ���ݼ�
    With mrsUnitsReg
        Set mrsUnitsReg = New ADODB.Recordset
        mrsUnitsReg.Fields.Append "������λ", adVarChar, 50
        mrsUnitsReg.Fields.Append "����ID", adBigInt
        mrsUnitsReg.Fields.Append "������Ŀ", adVarChar, 10
        mrsUnitsReg.Fields.Append "���", adBigInt, 18
        mrsUnitsReg.Fields.Append "����", adBigInt, 18
        mrsUnitsReg.Fields.Append "ʱ���", adVarChar, 60
        mrsUnitsReg.CursorLocation = adUseClient
        mrsUnitsReg.LockType = adLockOptimistic
        mrsUnitsReg.CursorType = adOpenStatic
        mrsUnitsReg.Open
    End With

    With mrsSource
        Set mrsSource = New ADODB.Recordset
        mrsSource.Fields.Append "����ID", adBigInt
        mrsSource.Fields.Append "������Ŀ", adVarChar, 10
        mrsSource.Fields.Append "���", adBigInt, 18
        mrsSource.Fields.Append "����", adBigInt, 18
        mrsSource.Fields.Append "ʱ���", adVarChar, 60
        mrsSource.CursorLocation = adUseClient
        mrsSource.LockType = adLockOptimistic
        mrsSource.CursorType = adOpenStatic
        mrsSource.Open
    End With
    
    With mrsLimit
        Set mrsLimit = New ADODB.Recordset
        mrsLimit.Fields.Append "������λ", adVarChar, 50
        mrsLimit.Fields.Append "������Ŀ", adVarChar, 10
        mrsLimit.Fields.Append "��������", adBigInt, 18
        mrsLimit.CursorLocation = adUseClient
        mrsLimit.LockType = adLockOptimistic
        mrsLimit.CursorType = adOpenStatic
        mrsLimit.Open
    End With
    
    With mrsDisable
        Set mrsDisable = New ADODB.Recordset
        mrsDisable.Fields.Append "������λ", adVarChar, 50
        mrsDisable.CursorLocation = adUseClient
        mrsDisable.LockType = adLockOptimistic
        mrsDisable.CursorType = adOpenStatic
        mrsDisable.Open
    End With
    
    If mstr�Ű� = "" Then Exit Function
    strList = Split(mstr�Ű�, "||")
    If mblnʱ�� Then
         '����Ƿ�ʱ��
         
        For i = 0 To UBound(strList)
            mrsʱ���.Filter = "����='" & Split(strList(i), "|")(0) & "' and �Ƿ�ԤԼ=1"
            If mrsʱ���.RecordCount = 0 Then mrsʱ���.Filter = "����='" & Split(strList(i), "|")(0) & "'"
            
            If mrsʱ���.RecordCount = 0 Then
               '���û������ʱ��� ����дʱ���
               mrs�޺�.Filter = "������Ŀ='" & Split(strList(i), "|")(0) & "'"

               If mrs�޺�.RecordCount = 0 Then
                   mrs�޺�.Filter = 0
               Else
                   lng�޺��� = Val(Nvl(mrs�޺�!�޺���))
                   lng��Լ�� = Val(Nvl(mrs�޺�!��Լ��))
                   If lng��Լ�� = 0 Then lng��Լ�� = lng�޺���

                   '���س�ʼ������
                   For j = 1 To lng�޺���

                       With mrsSource
                           .AddNew
                           !����ID = mlng����ID
                           !������Ŀ = Split(strList(i), "|")(0)
                           !��� = j
                           !���� = 1
                           .Update
                       End With

                   Next

               End If 'mrs�޺�.recourdcount
               
            Else    'mrsʱ���.recordCount=0
                Do While Not mrsʱ���.EOF
                    With mrsSource
                        .AddNew
                        !����ID = mlng����ID
                        !������Ŀ = Split(strList(i), "|")(0)
                        !��� = Val(Nvl(mrsʱ���!���))
                        !���� = 1
                        !ʱ��� = mrsʱ���!ʱ�䷶Χ
                        .Update
                    End With
                    mrsʱ���.MoveNext
                Loop
            End If
        Next
        mrsʱ���.Filter = 0
    Else
    
        For i = 0 To UBound(strList)
           '���û������ʱ��� ����дʱ���
            mrs�޺�.Filter = "������Ŀ='" & Split(strList(i), "|")(0) & "'"
    
            If mrs�޺�.RecordCount = 0 Then
                mrs�޺�.Filter = 0
            Else
                lng�޺��� = Val(Nvl(mrs�޺�!�޺���))
                lng��Լ�� = Val(Nvl(mrs�޺�!��Լ��))
                If lng��Լ�� = 0 Then lng��Լ�� = lng�޺���
                '���س�ʼ������
                For j = 1 To lng�޺���
                    With mrsSource
                        .AddNew
                        !����ID = mlng����ID
                        !������Ŀ = Split(strList(i), "|")(0)
                        !��� = j
                        !���� = 1
                        .Update
                    End With
    
                Next
    
            End If 'mrs�޺�.recourdcount
        Next
    End If
    
    '�Ѿ��������
    strSQL = "Select ������λ, ����id, ������Ŀ, ���, ���� From ������λ���ſ���  Where ����ID=[1] And ��� <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)

    If rsTmp.RecordCount > 0 Then

        Do While Not rsTmp.EOF
            mrsSource.Filter = "������Ŀ='" & rsTmp!������Ŀ & "' and ���=" & rsTmp!���

            With mrsUnitsReg
                .AddNew
                !������λ = Nvl(rsTmp!������λ)
                !����ID = mlng����ID
                !������Ŀ = Nvl(rsTmp!������Ŀ)
                !��� = Val(Nvl(rsTmp!���))
                !���� = Val(Nvl(rsTmp!����))

                If mrsSource.RecordCount > 0 Then
                    !ʱ��� = mrsSource!ʱ���
                    mrsSource.Delete
                    mrsSource.Update
                End If

                .Update
            End With

            mrsSource.Filter = 0
            rsTmp.MoveNext
        Loop
    
    End If
    
    strSQL = "Select ������λ, ����id, ������Ŀ, ���, ���� From ������λ���ſ���  Where ����ID=[1] And ��� = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            With mrsLimit
                .AddNew
                !������λ = Nvl(rsTmp!������λ)
                !������Ŀ = Nvl(rsTmp!������Ŀ)
                !�������� = Val(Nvl(rsTmp!����))
                .Update
            End With
            rsTmp.MoveNext
        Loop
    End If
    
    strSQL = "Select Distinct ������λ From ������λ���ſ���  Where ����ID=[1] And ���� = 0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            With mrsDisable
                .AddNew
                !������λ = Nvl(rsTmp!������λ)
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
    lblUnitRegTitle.Caption = strUnits & ":ԤԼ����"
    Call tbWeekTime_Click
End Sub

 

Private Sub picPlan_Resize()

    On Error Resume Next
    
    lblUnitRegTitle.Move 0, 0, picPlan.ScaleWidth, lblUnitRegTitle.Height
    chkDisable.Left = picPlan.ScaleWidth - chkDisable.Width
    Me.tbWeekTime.Move 0, lblUnitRegTitle.Height + Screen.TwipsPerPixelY, picPlan.ScaleWidth, Me.tbWeekTime.Height
    
    With lblδ����
        .Left = Screen.TwipsPerPixelX * 2
        .Top = tbWeekTime.Height + tbWeekTime.Top + Screen.TwipsPerPixelY * 4
        .Width = lvwSource.Width
    End With
    
    With lvwSource
        .Left = Screen.TwipsPerPixelX * 2
        .Top = lblδ����.Height + lblδ����.Top ' + Screen.TwipsPerPixelY * 4
        .Height = Me.picPlan.ScaleHeight - lblδ����.Height - lblδ����.Top - Screen.TwipsPerPixelY * 2 - 700
    End With

    With fraMove
        .Left = lvwSource.Left + lvwSource.Width + Screen.TwipsPerPixelX * 2
        .Top = lvwSource.Top
        .Height = lvwSource.Top + lvwSource.Height - lblδ����.Top - lblδ����.Height
    End With
    
     With lbl�ѷ���
        .Left = fraMove.Left + fraMove.Width + Screen.TwipsPerPixelX * 2
        .Top = lblδ����.Top
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
            .Filter = "������Ŀ='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "' and  ������λ='" & lstUnits.Text & "'"
            If .RecordCount = 0 Then
                mblnNoManual = True
                txtLimit.Text = ""
                mblnNoManual = False
            Else
                mblnNoManual = True
                txtLimit.Text = !��������
                mblnNoManual = False
            End If
        End With
    End If
    
    If Not mrsDisable Is Nothing Then
        With mrsDisable
            .Filter = "������λ='" & lstUnits.Text & "'"
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
        
        '������������ʱ��
        lvwSource.ListItems.Clear
        mrsSource.Filter = "������Ŀ='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "'"

        Do While Not mrsSource.EOF

            With lvwSource
                Set lvwItem = .ListItems.Add(, "k" & mrsSource!���, mrsSource!���)
                lvwItem.SubItems(1) = Nvl(mrsSource!ʱ���)
            End With

            mrsSource.MoveNext
        Loop

        mrsSource.Filter = 0
        lvwReg.ListItems.Clear
        mrsUnitsReg.Filter = "������Ŀ='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "' and  ������λ='" & lstUnits.Text & "'"

        Do While Not mrsUnitsReg.EOF

            With lvwReg
                Set lvwItem = .ListItems.Add(, "k" & mrsUnitsReg!���, mrsUnitsReg!���)
                lvwItem.SubItems(1) = Nvl(mrsUnitsReg!ʱ���)
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
           mrsSource.Fields.Append "����ID", adBigInt
           mrsSource.Fields.Append "������Ŀ", adVarChar, 10
           mrsSource.Fields.Append "���", adBigInt, 18
           mrsSource.Fields.Append "����", adBigInt, 18
           mrsSource.Fields.Append "ʱ���", adVarChar, 60
           mrsSource.CursorLocation = adUseClient
           mrsSource.LockType = adLockOptimistic
           mrsSource.CursorType = adOpenStatic
           mrsSource.Open
         End With
    End If
    For i = 1 To lvwReg.ListItems.Count
        Set lvwItem = lvwReg.ListItems(i)
        Set lvwitem1 = lvwSource.ListItems.Add(, lvwItem.Key, lvwItem.Text)
        mrsUnitsReg.Filter = "������Ŀ='" & Mid(Me.tbWeekTime.SelectedItem.Key, 2) & "' and ���=" & Val(lvwItem.Text)
        With mrsSource
           .AddNew
           !����ID = mrsUnitsReg!����ID
           !������Ŀ = mrsUnitsReg!������Ŀ
           !��� = Val(lvwItem.Text)
           !ʱ��� = Nvl(mrsUnitsReg!ʱ���)
           !���� = Val(mrsUnitsReg!����)
           .Update
        End With
        mrsUnitsReg.Delete adAffectCurrent
        mrsUnitsReg.Update
        lvwitem1.SubItems(1) = lvwItem.SubItems(1)
        'UnitRegToSource Mid(Me.tbWeekTime.SelectedItem.Key, 2), Val(lvwitem1.Text), mlng����ID, lvwitem1.SubItems(1), 1
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
        .Filter = "������Ŀ='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "' and  ������λ='" & lstUnits.Text & "'"
        If .RecordCount <> 0 Then
            .MoveFirst
            .Delete adAffectCurrent
        End If
    End With
End Sub

Private Sub txtLimit_Validate(Cancel As Boolean)
    If mrs�޺� Is Nothing Then Exit Sub
    mrs�޺�.Filter = "������Ŀ='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "'"
    If mrs�޺�.RecordCount = 0 Then Exit Sub
    If Val(txtLimit.Text) > mrs�޺�!�޺��� Then
        MsgBox "���õĺ�����λ�޺������ܳ�����ǰ���ŵ��޺�����", vbInformation, gstrSysName
        Cancel = True
        If txtLimit.Enabled And txtLimit.Visible Then txtLimit.SetFocus
        Exit Sub
    End If
    With mrsLimit
        .Filter = "������Ŀ='" & Mid(tbWeekTime.SelectedItem.Key, 2) & "' and  ������λ='" & lstUnits.Text & "'"
        If .RecordCount <> 0 Then
            .MoveFirst
            .Delete adAffectCurrent
            .Update
        End If
        If Val(Nvl(txtLimit.Text)) <> 0 Then
            .AddNew
            !������λ = lstUnits.Text
            !������Ŀ = Mid(tbWeekTime.SelectedItem.Key, 2)
            !�������� = Val(txtLimit.Text)
            .Update
        End If
    End With
End Sub
