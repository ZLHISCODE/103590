VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicFind 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "������Ŀ����"
   ClientHeight    =   6240
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   8925
   Icon            =   "frmClinicFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   7710
      TabIndex        =   23
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7710
      TabIndex        =   21
      Top             =   2670
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   7710
      TabIndex        =   20
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "��λ(&L)"
      Height          =   350
      Left            =   7710
      TabIndex        =   19
      Top             =   705
      Width           =   1100
   End
   Begin VB.Frame fra�߼� 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   -120
      TabIndex        =   7
      Top             =   930
      Width           =   7635
      Begin VB.Frame fra��ʽ 
         Caption         =   "ƥ�䷽ʽ"
         Height          =   1305
         Left            =   210
         TabIndex        =   15
         Top             =   30
         Width           =   2295
         Begin VB.OptionButton optMode 
            Caption         =   "��ȫ��ͬ(&A)"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   18
            Top             =   330
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "���������ݿ�ͷ(&B)"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   17
            Top             =   660
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "������������(&C)"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   16
            Top             =   990
            Value           =   -1  'True
            Width           =   1845
         End
      End
      Begin VB.Frame fra��Χ 
         Caption         =   "���ҷ�Χ"
         Height          =   1035
         Left            =   2600
         TabIndex        =   11
         Top             =   30
         Width           =   4995
         Begin VB.OptionButton optScope 
            Caption         =   "��ǰ�����µ�ֱ����Ŀ(&3)"
            Height          =   195
            Index           =   3
            Left            =   2280
            TabIndex        =   14
            Top             =   315
            Width           =   2385
         End
         Begin VB.OptionButton optScope 
            Caption         =   "�������(&1)"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   13
            Top             =   285
            Width           =   1335
         End
         Begin VB.OptionButton optScope 
            Caption         =   "��ǰ�����µ�������Ŀ(&2)"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   12
            Top             =   660
            Width           =   2385
         End
      End
      Begin VB.CheckBox chkStop 
         Caption         =   "������ͣ����Ŀ(&T)"
         Enabled         =   0   'False
         Height          =   225
         Left            =   4200
         TabIndex        =   10
         Top             =   1170
         Width           =   1845
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "���ִ�Сд(&E)"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1170
         Width           =   1485
      End
      Begin VB.CheckBox chkAlias 
         Caption         =   "����������(&I)"
         Height          =   180
         Left            =   6120
         TabIndex        =   8
         Top             =   1170
         Width           =   1575
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "��������"
      Height          =   765
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7365
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   5610
         MaxLength       =   12
         TabIndex        =   3
         Top             =   300
         Width           =   1620
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   3210
         MaxLength       =   40
         TabIndex        =   2
         Top             =   300
         Width           =   1620
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   765
         MaxLength       =   10
         TabIndex        =   1
         Top             =   300
         Width           =   1620
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   3
         Left            =   4985
         TabIndex        =   6
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   2565
         TabIndex        =   5
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1920
      Top             =   3900
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
            Picture         =   "frmClinicFind.frx":058A
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicFind.frx":09E2
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicFind.frx":0E36
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicFind.frx":128A
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicFind.frx":20DC
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3735
      Left            =   60
      TabIndex        =   22
      Top             =   2460
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   4763
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "_����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "_����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "_����"
         Text            =   "����"
         Object.Width           =   2648
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "���ҽ����"
      Height          =   180
      Left            =   7710
      TabIndex        =   25
      Top             =   4590
      Width           =   900
   End
   Begin VB.Label lbl���� 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7710
      TabIndex        =   24
      Top             =   4890
      Width           =   1095
   End
End
Attribute VB_Name = "frmClinicFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintColumn As Integer
Dim mblnItem As Boolean

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdLocate_Click()
    Dim strkey As String
    Dim strItemKey As String
    
    If LvwMain.SelectedItem Is Nothing Then Exit Sub
    On Error Resume Next
    With LvwMain.SelectedItem
        strkey = "_" & .ListSubItems(1).Tag
        strItemKey = "_" & .Tag
        If .SubItems(3) <> "δ����" Then
            frmClinicLists.tvwClass.Nodes(strkey).Selected = True
            frmClinicLists.tvwClass.Nodes(strkey).EnsureVisible
            frmClinicLists.tvwClass_NodeClick frmClinicLists.tvwClass.SelectedItem
            Err.Clear
            frmClinicLists.lvwItems.ListItems(strItemKey).Selected = True
            frmClinicLists.lvwItems.ListItems(strItemKey).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "���ҵ���������¼�����ѱ�ɾ����ͣ�ã���ˢ���б�", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            frmClinicLists.lvwItems_ItemClick frmClinicLists.lvwItems.SelectedItem
        Else
            frmClinicLists.tvwClass.Nodes("Root").Selected = True
            frmClinicLists.tvwClass.Nodes(strkey).EnsureVisible
            frmClinicLists.tvwClass_NodeClick frmClinicLists.tvwClass.SelectedItem
            Err.Clear
            frmClinicLists.lvwItems.ListItems(strItemKey).Selected = True
            frmClinicLists.lvwItems.ListItems(strItemKey).EnsureVisible
            If Err.Number = 35601 Then
                MsgBox "���ҵ���������¼�����ѱ�ɾ����ͣ�ã���ˢ���б�", vbInformation, gstrSysName
                Err.Clear
                Exit Sub
            End If
            frmClinicLists.lvwItems_ItemClick frmClinicLists.lvwItems.SelectedItem
        End If
    End With
    Err.Clear
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strTable As String
    Dim strWhere As String
    Dim strWhereAlias As String
    Dim strAlaisSql As String
    Dim str����ID As String
    Dim i As Long
    Dim str���� As String
    Dim str���� As String
    Dim str���� As String
    
    On Error GoTo errHandle
    
    str���� = IIf(chkCase.Value = 0, UCase(txtEdit(0).Text), txtEdit(0).Text)
    str���� = IIf(chkCase.Value = 0, UCase(txtEdit(1).Text), txtEdit(1).Text)
    str���� = IIf(chkCase.Value = 0, UCase(txtEdit(2).Text), txtEdit(2).Text)
    
    For i = 0 To 2
        If zlCommFun.StrIsValid(txtEdit(i).Text) = False Then
            txtEdit(i).SetFocus
            Exit Sub
        End If
    Next
    With frmClinicLists.tvwClass
        If .SelectedItem Is Nothing Then Exit Sub
        If .SelectedItem.Key = "Root" Then
            str����ID = ""
        Else
            If .SelectedItem.Key <> "Root" Then
                str����ID = Val(Mid(.SelectedItem.Key, 2))
                If str����ID = "0" Then
                    str����ID = ""
                End If
            Else
                str����ID = ""
            End If
        End If
    End With
    
    strWhere = " And ���" & Me.Tag
    If chkStop.Value = 0 Then
        strWhere = strWhere & " and (����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or ����ʱ�� is null) "
    End If
    '���ҷ�Χ
    If optScope(0).Value = True Then
        strTable = "select ID,���,����ID,����,����,����ʱ�� from ������ĿĿ¼ where ��� Not In('4','5','6','7') " & strWhere
    ElseIf optScope(2).Value = True Then
        If str����ID = "" Then
            strTable = "select ID,���,����ID,����,����,����ʱ�� from ������ĿĿ¼  where ��� Not In('4','5','6','7') " & strWhere & vbCrLf & _
            " and (����id IN (SELECT id FROM ���Ʒ���Ŀ¼  START WITH �ϼ�id is null  CONNECT BY PRIOR id=�ϼ�id) OR ����id  is null ) "   ' start with ���='" & str��� & "'and �ϼ�ID is null connect by prior ID=�ϼ�ID"
        Else
            strTable = "select ID,���,����ID,����,����,����ʱ�� from ������ĿĿ¼   where  ��� Not In('4','5','6','7') " & strWhere & vbCrLf & _
            " and (����id IN (SELECT id FROM ���Ʒ���Ŀ¼  START WITH �ϼ�ID=[1] CONNECT BY PRIOR id=�ϼ�id) OR ����id=[1] ) "
        End If
    Else
        If str����ID = "" Then
            strTable = "select ID,���,����ID,����,����,����ʱ�� from ������ĿĿ¼ " & _
            "where  ��� Not In('4','5','6','7') and ����ID is null " & strWhere
        Else
            strTable = "select ID,���,����ID,����,����,����ʱ�� from ������ĿĿ¼ " & _
            "where  ��� Not In('4','5','6','7') and ����ID=[1] " & strWhere
        End If
    End If
    '�ȽϷ�ʽ
    strWhere = ""
    If optMode(0).Value = True Then
        If Trim(txtEdit(0).Text) <> "" Then
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.����" & IIf(chkCase.Value = 0, ")", "") & "=[2] "
        End If
        
        If Trim(txtEdit(1).Text) <> "" Then
            strWhereAlias = strWhereAlias & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.����" & IIf(chkCase.Value = 0, ")", "") & "=[3] "
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "Nvl(B.����, A.����)" & IIf(chkCase.Value = 0, ")", "") & "=[3] "
        End If
        
        If Trim(txtEdit(2).Text) <> "" Then
            strWhereAlias = strWhereAlias & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "A.����" & IIf(chkCase.Value = 0, ")", "") & "=[4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.����" & IIf(chkCase.Value = 0, ")", "") & "=[4] " & ") "
            strWhere = strWhere & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "B.ƴ����" & IIf(chkCase.Value = 0, ")", "") & "=[4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.�����" & IIf(chkCase.Value = 0, ")", "") & "=[4] " & ") "
        End If
    ElseIf optMode(1).Value = True Then
        If Trim(txtEdit(0).Text) <> "" Then
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.����" & IIf(chkCase.Value = 0, ")", "") & " Like [2] "
        End If
        
        If Trim(txtEdit(1).Text) <> "" Then
            strWhereAlias = strWhereAlias & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.����" & IIf(chkCase.Value = 0, ")", "") & " Like [3] "
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "Nvl(B.����, A.����)" & IIf(chkCase.Value = 0, ")", "") & " Like [3] "
        End If
        
        If Trim(txtEdit(2).Text) <> "" Then
            strWhereAlias = strWhereAlias & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "A.����" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.����" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & ") "
            strWhere = strWhere & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "B.ƴ����" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.�����" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & ") "
        End If
        
        str���� = str���� & "%"
        str���� = str���� & "%"
        str���� = str���� & "%"
    Else
        If Trim(txtEdit(0).Text) <> "" Then
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.����" & IIf(chkCase.Value = 0, ")", "") & " Like [2] "
        End If
        
        If Trim(txtEdit(1).Text) <> "" Then
            strWhereAlias = strWhereAlias & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "A.����" & IIf(chkCase.Value = 0, ")", "") & " Like [3] "
            strWhere = strWhere & " and " & IIf(chkCase.Value = 0, "Upper(", "") & "Nvl(B.����, A.����)" & IIf(chkCase.Value = 0, ")", "") & " Like [3] "
        End If
        
        If Trim(txtEdit(2).Text) <> "" Then
            strWhereAlias = strWhereAlias & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "A.����" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.����" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & ") "
            strWhere = strWhere & " And (" & IIf(chkCase.Value = 0, "Upper(", "") & "B.ƴ����" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & _
                " Or " & IIf(chkCase.Value = 0, "Upper(", "") & "B.�����" & IIf(chkCase.Value = 0, ")", "") & " Like [4] " & ") "
        End If
        
        str���� = "%" & str���� & "%"
        str���� = "%" & str���� & "%"
        str���� = "%" & str���� & "%"
    End If
    
    '�õ�SQL���
    strAlaisSql = " (Select Distinct A.������Ŀid, A.����, A.���� As ƴ����, B.���� As �����, A.���� || '/' || B.���� As ����" & _
        " From ������Ŀ���� A, ������Ŀ���� B " & _
        " Where A.������Ŀid = B.������Ŀid And A.���� = 1 And B.���� = 2 " & strWhereAlias
    If chkAlias.Value = 1 Then
        '���ұ���
        strAlaisSql = strAlaisSql & " And A.���� = 9 And B.���� = 9 ) B "
    Else
        '����ͨ����
        strAlaisSql = strAlaisSql & " And A.���� = 1 And B.���� = 1 ) B "
    End If
    gstrSql = "select distinct A.ID,A.���,Nvl(B.����, A.����) As ����,A.����,B.����,C.���� as ����,A.����ID,A.����ʱ�� from (" & _
        strTable & ") A," & strAlaisSql & ",���Ʒ���Ŀ¼ C where A.����id=c.id(+) And  A.ID=B.������Ŀid(+) and C.���� is not NULL And C.���� In (4,5,6) " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(str����ID), str����, str����, str����)
    
    Me.MousePointer = 11
    zlControl.FormLock LvwMain.hWnd
    With LvwMain.ListItems
        .Clear
        i = 1
        Do Until rsTemp.EOF
            '�ó���ȷ��ͼ��
            strWhere = "Item"
            If Not CDate(IIf(IsNull(rsTemp("����ʱ��")), CDate("3000/1/1"), rsTemp("����ʱ��"))) = CDate("3000/1/1") Then
                strWhere = strWhere & "No"
            End If
            '��ӽڵ�
            Set lst = .Add(, "C" & i, rsTemp("����"), strWhere, strWhere)
            If InStr(strWhere, "No") > 0 Then lst.ForeColor = RGB(255, 0, 0)
            
            Dim lngCol  As Long
            Dim varValue As Variant
            '����ListView�����������ݿ�ȡ��
            For lngCol = 2 To LvwMain.ColumnHeaders.Count
                varValue = rsTemp(LvwMain.ColumnHeaders(lngCol).Text).Value
                lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
                
                lst.ListSubItems(1).Tag = IIf(IsNull(rsTemp("����ID")), "", rsTemp("����ID"))
                lst.Tag = rsTemp("id")
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
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlControl.FormLock 0
    Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
    Select Case Val(frmClinicLists.tvwClass.Tag)
    Case 0
        Me.Tag = ">='A'": Me.Caption = "������Ŀ����..."
    Case 1
        Me.Tag = "='8'": Me.Caption = "��ҩ�䷽����..."
    Case 2
        Me.Tag = "='9'": Me.Caption = "���׷�������..."
    End Select
End Sub

Private Sub Form_Load()
    Dim intSel As Integer
    
    RestoreWinState Me, App.ProductName
    
    intSel = Val(zlDatabase.GetPara("ƥ�䷽ʽ", glngSys, 1054, 2))
    
    If intSel > 2 Or intSel < 0 Then intSel = 2
    optMode(intSel).Value = True
    
    intSel = Val(zlDatabase.GetPara("���ҷ�Χ", glngSys, 1054, 0))
    If intSel > 3 Or intSel < 0 Then intSel = 0
    optScope(intSel).Value = True
    
    intSel = Val(zlDatabase.GetPara("���ִ�Сд", glngSys, 1054, 0))
    If intSel > 1 Or intSel < 0 Then intSel = 0
    chkCase.Value = intSel
    
    intSel = Val(zlDatabase.GetPara("���ұ���", glngSys, 1054, 0))
    If intSel > 1 Or intSel < 0 Then intSel = 0
    chkAlias.Value = intSel
    
    chkStop.Value = IIf(frmClinicLists.mnuViewStoped.Checked, 1, 0)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim intSel As Integer
    
    For intSel = 0 To 2
        If optMode(intSel).Value = True Then Exit For
    Next
    Call zlDatabase.SetPara("ƥ�䷽ʽ", intSel, glngSys, 1054)
    For intSel = 0 To 3
        If intSel <> 1 Then
            If optScope(intSel).Value = True Then Exit For
        End If
    Next
    Call zlDatabase.SetPara("���ҷ�Χ", intSel, glngSys, 1054)
    Call zlDatabase.SetPara("���ִ�Сд", chkCase.Value, glngSys, 1054)
    Call zlDatabase.SetPara("���ұ���", chkCase.Value, glngSys, 1054)
        
    SaveWinState Me, App.ProductName
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Dim sngLeft As Single
    
    LvwMain.Top = IIf(fra�߼�.Visible = True, fra�߼�.Top + fra�߼�.Height, fra�߼�.Top)
    LvwMain.Height = Me.ScaleHeight - LvwMain.Top - 120
    
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
    LvwMain.Width = cmdFind.Left - LvwMain.Left - 245
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        LvwMain.SortOrder = IIf(LvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        LvwMain.SortKey = mintColumn
        LvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub chkCase_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkStop_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True Then Call cmdLocate_Click
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnItem = False
End Sub

Private Sub optMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optScope_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub


