VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeItemFind 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "�շ���Ŀ����"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ImageList ils16 
      Left            =   1980
      Top             =   3870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":0000
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":0458
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":08AC
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":0D00
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItemFind.frx":1B52
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra���� 
      Caption         =   "��������"
      Height          =   765
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   7365
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   2835
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "A.��ʶ����"
         Top             =   300
         Width           =   1020
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   765
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "A.����"
         Top             =   300
         Width           =   1020
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   4530
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "B.����"
         Top             =   300
         Width           =   1020
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   6210
         MaxLength       =   12
         TabIndex        =   8
         Tag             =   "B.����"
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ʶ����(&P)"
         Height          =   180
         Index           =   0
         Left            =   1845
         TabIndex        =   3
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   3885
         TabIndex        =   5
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   3
         Left            =   5595
         TabIndex        =   7
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame fra�߼� 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   -60
      TabIndex        =   24
      Top             =   900
      Width           =   7635
      Begin VB.CheckBox chkCase 
         Caption         =   "���ִ�Сд(&E)"
         Height          =   255
         Left            =   2940
         TabIndex        =   17
         Top             =   1170
         Width           =   1485
      End
      Begin VB.CheckBox chkStop 
         Caption         =   "������ͣ����Ŀ(&T)"
         Enabled         =   0   'False
         Height          =   225
         Left            =   4740
         TabIndex        =   18
         Top             =   1170
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Frame fra��Χ 
         Caption         =   "���ҷ�Χ"
         Height          =   1035
         Left            =   2820
         TabIndex        =   13
         Top             =   30
         Width           =   4755
         Begin VB.OptionButton optScope 
            Caption         =   "��ǰ�����µ�������Ŀ(&2)"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   15
            Top             =   660
            Width           =   2385
         End
         Begin VB.OptionButton optScope 
            Caption         =   "�������(&1)"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   14
            Top             =   285
            Width           =   1335
         End
         Begin VB.OptionButton optScope 
            Caption         =   "��ǰ�����µ�ֱ����Ŀ(&3)"
            Height          =   195
            Index           =   3
            Left            =   2280
            TabIndex        =   16
            Top             =   315
            Width           =   2385
         End
      End
      Begin VB.Frame fra��ʽ 
         Caption         =   "ƥ�䷽ʽ"
         Height          =   1305
         Left            =   210
         TabIndex        =   9
         Top             =   30
         Width           =   2295
         Begin VB.OptionButton optMode 
            Caption         =   "������������(&C)"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   12
            Top             =   990
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "���������ݿ�ͷ(&B)"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   11
            Top             =   660
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "��ȫ��ͬ(&A)"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   330
            Width           =   1845
         End
      End
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "��λ(&L)"
      Height          =   350
      Left            =   7770
      TabIndex        =   21
      Top             =   680
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   7770
      TabIndex        =   22
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7770
      TabIndex        =   23
      Top             =   2640
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3735
      Left            =   120
      TabIndex        =   20
      Top             =   2430
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   6588
      View            =   3
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "_����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "_��ʶ����"
         Object.Tag             =   "��ʶ����"
         Text            =   "��ʶ����"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "_��ʶ����"
         Object.Tag             =   "��ʶ����"
         Text            =   "��ʶ����"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "_����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "_����"
         Text            =   "����"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   7770
      TabIndex        =   19
      Top             =   180
      Width           =   1100
   End
   Begin VB.Label lbl���� 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7770
      TabIndex        =   26
      Top             =   4860
      Width           =   1100
   End
   Begin VB.Label lbl 
      Caption         =   "���ҽ����"
      Height          =   180
      Left            =   7770
      TabIndex        =   25
      Top             =   4560
      Width           =   900
   End
End
Attribute VB_Name = "frmChargeItemFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintColumn As Integer
Dim mblnItem As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdLocate_Click()
    Dim strKey As String
    Dim strClass As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    On Error Resume Next
    With lvwMain.SelectedItem
        strClass = Mid(.Tag, 2, 1)
        If .SubItems(3) <> "δ����" Then
            strKey = "R" & .ListSubItems(1).Tag
            frmChargeManage.tvwMainItem.Nodes(strKey).Selected = True
            frmChargeManage.tvwMainItem.Nodes(strKey).EnsureVisible
            frmChargeManage.tvwMainItem_NodeClick frmChargeManage.tvwMainItem.SelectedItem
            Err.Clear
            frmChargeManage.lvwMain_S.ListItems(.Tag).Selected = True
            frmChargeManage.lvwMain_S.ListItems(.Tag).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "���ҵ���������¼�����ѱ�ɾ����ͣ�ã���ˢ���б�", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            frmChargeManage.lvwMain_S_ItemClick frmChargeManage.lvwMain_S.SelectedItem
        Else
            frmChargeManage.tvwMainItem.Nodes("Root").Selected = True
            frmChargeManage.tvwMainItem.Nodes(strKey).EnsureVisible
            frmChargeManage.tvwMainItem_NodeClick frmChargeManage.tvwMainItem.SelectedItem
            Err.Clear
            frmChargeManage.lvwMain_S.ListItems(.Tag).Selected = True
            frmChargeManage.lvwMain_S.ListItems(.Tag).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "���ҵ���������¼�����ѱ�ɾ����ͣ�ã���ˢ���б�", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            frmChargeManage.lvwMain_S_ItemClick frmChargeManage.lvwMain_S.SelectedItem
        End If
    End With
    Err.Clear
End Sub

Private Sub cmdFind_Click()
    On Error GoTo ErrHandle
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strTable As String
    Dim strWhere As String
    Dim str����ID As String
    Dim i As Long
    Dim str���� As String
    Dim str��ʶ���� As String
    Dim str���� As String
    Dim str���� As String
    
    str���� = IIF(chkCase.Value = 0, UCase(txtEdit(0).Text), txtEdit(0).Text)
    str��ʶ���� = IIF(chkCase.Value = 0, UCase(txtEdit(1).Text), txtEdit(1).Text)
    str���� = IIF(chkCase.Value = 0, UCase(txtEdit(2).Text), txtEdit(2).Text)
    str���� = IIF(chkCase.Value = 0, UCase(txtEdit(3).Text), txtEdit(3).Text)
    
    For i = 0 To 3
        If zlCommFun.StrIsValid(txtEdit(i).Text) = False Then
            txtEdit(i).SetFocus
            Exit Sub
        End If
    Next
    With frmChargeManage.tvwMainItem
        If .SelectedItem Is Nothing Then Exit Sub
        If .SelectedItem.Key = "Root" Then
            str����ID = ""
        Else
            If .SelectedItem.Key <> "Root" Then
                str����ID = .SelectedItem.Tag
                If str����ID = "0" Then
                    str����ID = ""
                End If
            Else
                str����ID = ""
            End If
        End If
    End With
    
    If chkStop.Value = 0 Then
        strWhere = " and (����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or ����ʱ�� is null) "
    End If
    '���ҷ�Χ
    If optScope(0).Value = True Then
        strTable = "select ID,���,����ID,����,����,��ʶ����,��ʶ����,����ʱ�� from �շ���ĿĿ¼ where ���<>'5' and ���<>'6' and ���<>'7'" & strWhere
    ElseIf optScope(2).Value = True Then
        If str����ID = "" Then
            strTable = "select ID,���,����ID,����,����,��ʶ����,��ʶ����,����ʱ�� from �շ���ĿĿ¼  where ���<>'5' and ���<>'6' and ���<>'7' " & strWhere & vbCrLf & _
            " and (����id IN (SELECT id FROM �շѷ���Ŀ¼  START WITH �ϼ�id is null  CONNECT BY PRIOR id=�ϼ�id) OR ����id  is null ) "   ' start with ���='" & str��� & "'and �ϼ�ID is null connect by prior ID=�ϼ�ID"
        Else
            strTable = "select ID,���,����ID,����,����,��ʶ����,��ʶ����,����ʱ�� from �շ���ĿĿ¼   where  ���<>'5' and ���<>'6' and ���<>'7' " & strWhere & vbCrLf & _
            " and (����id IN (SELECT id FROM �շѷ���Ŀ¼  START WITH �ϼ�ID=[1] CONNECT BY PRIOR id=�ϼ�id) OR ����id=[1] ) "
        End If
    Else
        If str����ID = "" Then
            strTable = "select ID,���,����ID,����,����,��ʶ����,��ʶ����,����ʱ�� from �շ���ĿĿ¼ " & _
            "where  ���<>'5' and ���<>'6' and ���<>'7'and ����ID is null " & strWhere
        Else
            strTable = "select ID,���,����ID,����,����,��ʶ����,��ʶ����,����ʱ�� from �շ���ĿĿ¼ " & _
            "where   ���<>'5' and ���<>'6' and ���<>'7'and ����ID=[1] " & strWhere
        End If
    End If
    '�ȽϷ�ʽ
    strWhere = ""
    If optmode(0).Value = True Then
        For i = 0 To 3
            If txtEdit(i).Text <> "" Then
                strWhere = strWhere & " and " & IIF(chkCase.Value = 0, "Upper(", "") & txtEdit(i).Tag & IIF(chkCase.Value = 0, ")", "") & "=[" & i + 2 & "] "
            End If
        Next
    ElseIf optmode(1).Value = True Then
        For i = 0 To 3
            If txtEdit(i).Text <> "" Then
                strWhere = strWhere & " and " & IIF(chkCase.Value = 0, "Upper(", "") & txtEdit(i).Tag & IIF(chkCase.Value = 0, ")", "") & " like [" & i + 2 & "] "
            End If
        Next
        str���� = str���� & "%"
        str��ʶ���� = str��ʶ���� & "%"
        str���� = str���� & "%"
        str���� = str���� & "%"
    Else
        For i = 0 To 3
            If txtEdit(i).Text <> "" Then
                strWhere = strWhere & " and " & IIF(chkCase.Value = 0, "Upper(", "") & txtEdit(i).Tag & IIF(chkCase.Value = 0, ")", "") & " like [" & i + 2 & "] "
            End If
        Next
        str���� = "%" & str���� & "%"
        str��ʶ���� = "%" & str��ʶ���� & "%"
        str���� = "%" & str���� & "%"
        str���� = "%" & str���� & "%"
    End If
    
    '�õ�SQL���
    gstrSQL = "select distinct A.ID,A.���,B.����,A.����,A.��ʶ����,A.��ʶ����,B.����,C.���� as ����,A.����ID,A.����ʱ�� from (" & _
        strTable & ") A,(Select A.�շ�ϸĿid, A.����, A.���� || '/' || B.���� As ����" & _
        " From �շ���Ŀ���� A, �շ���Ŀ���� B " & _
        " Where A.�շ�ϸĿid = B.�շ�ϸĿid And A.���� = 1 And B.���� = 2) B,�շѷ���Ŀ¼ C where A.����id=c.id(+) And  A.ID=B.�շ�ϸĿID and C.���� is not NULL " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����ID, str����, str��ʶ����, str����, str����)
    
    Me.MousePointer = 11
    zlControl.FormLock lvwMain.hwnd
    With lvwMain.ListItems
        .Clear
        i = 1
        Do Until rsTemp.EOF
            '�ó���ȷ��ͼ��
            strWhere = "Item"
            If Not CDate(IIF(IsNull(rsTemp("����ʱ��")), CDate("3000/1/1"), rsTemp("����ʱ��"))) = CDate("3000/1/1") Then
                strWhere = strWhere & "No"
            End If
            '��ӽڵ�
            Set lst = .Add(, "C" & i, rsTemp("����"), strWhere, strWhere)
            If InStr(strWhere, "No") > 0 Then lst.ForeColor = RGB(255, 0, 0)
            
            Dim lngCol  As Long
            Dim varValue As Variant
            '����ListView�����������ݿ�ȡ��
            For lngCol = 2 To lvwMain.ColumnHeaders.Count
                varValue = rsTemp(lvwMain.ColumnHeaders(lngCol).Text).Value
                lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
                
                lst.ListSubItems(1).Tag = IIF(IsNull(rsTemp("����ID")), "", rsTemp("����ID"))
                lst.Tag = "C" & rsTemp("���") & rsTemp("id")
                If InStr(strWhere, "No") > 0 Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
            Next
            rsTemp.MoveNext
            i = i + 1
        Loop
        If .Count > 0 Then .Item(1).Selected = True
        lbl����.Caption = "��" & .Count & "��"
    End With
    
    zlControl.FormLock 0
    Me.MousePointer = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlControl.FormLock 0
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim intSel As Integer
    
    RestoreWinState Me, App.ProductName
    
    intSel = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "ƥ�䷽ʽ", 2)
    If intSel > 2 Or intSel < 0 Then intSel = 2
    optmode(intSel).Value = True
    
    intSel = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "���ҷ�Χ", 0)
    If intSel > 3 Or intSel < 0 Then intSel = 0
    optScope(intSel).Value = True
    
    intSel = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "���ִ�Сд", 0)
    If intSel > 1 Or intSel < 0 Then intSel = 0
    chkCase.Value = intSel
    
    chkStop.Value = IIF(frmChargeManage.mnuViewShowStop.Checked, 1, 0)
    fra����.Caption = fra����.Caption & IIF(chkStop.Value, "(����ͣ����Ŀ)", "(����ͣ����Ŀ)")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intSel As Integer
    
    For intSel = 0 To 2
        If optmode(intSel).Value = True Then Exit For
    Next
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "ƥ�䷽ʽ", intSel)
    For intSel = 0 To 3
        If intSel <> 1 Then
            If optScope(intSel).Value = True Then Exit For
        End If
    Next
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "���ҷ�Χ", intSel)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "���ִ�Сд", chkCase.Value)
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim sngLeft As Single
    
    lvwMain.Top = IIF(fra�߼�.Visible = True, fra�߼�.Top + fra�߼�.Height, fra�߼�.Top)
    lvwMain.Height = Me.ScaleHeight - lvwMain.Top - 120
    
    sngLeft = ScaleWidth - cmdFind.Width - 200
    If sngLeft >= 7770 Then
        cmdFind.Left = sngLeft
    Else
        sngLeft = 7770
        cmdFind.Left = sngLeft
    End If
    cmdLocate.Left = cmdFind.Left
    cmdExit.Left = cmdFind.Left
    cmdHelp.Left = cmdFind.Left
    lbl.Left = cmdFind.Left
    lbl����.Left = cmdFind.Left
    lvwMain.Width = cmdFind.Left - lvwMain.Left - 245
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub chkCase_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chkStop_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True Then Call cmdLocate_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub optMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub optScope_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub
