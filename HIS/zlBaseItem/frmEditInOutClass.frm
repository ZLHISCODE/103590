VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditInOutClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�༭������"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmEditInOutClass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList ImgLvwSmall 
      Left            =   7650
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -105
      TabIndex        =   13
      Top             =   4095
      Width           =   8595
   End
   Begin MSComctlLib.ImageList ImgLvw����Small 
      Left            =   4620
      Top             =   930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7110
      TabIndex        =   6
      Top             =   4260
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5850
      TabIndex        =   5
      Top             =   4260
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   180
      TabIndex        =   7
      Top             =   4260
      Width           =   1100
   End
   Begin MSComctlLib.ListView Lvw���ݷ����б� 
      Height          =   3060
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��������"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "˵��"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.ComboBox Cbo���� 
      Height          =   300
      Left            =   6015
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   150
      Width           =   2175
   End
   Begin VB.TextBox Txt���� 
      Height          =   300
      Left            =   2370
      MaxLength       =   20
      TabIndex        =   1
      Top             =   180
      Width           =   2145
   End
   Begin VB.TextBox Txt���� 
      Height          =   300
      Left            =   585
      MaxLength       =   2
      TabIndex        =   0
      Top             =   180
      Width           =   645
   End
   Begin MSComctlLib.ListView Lvw������ 
      Height          =   1755
      Left            =   5400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   3096
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ʹ�ø����ĵ��ݣ�"
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   675
      Width           =   1800
   End
   Begin VB.Label Lbl˵�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ע(�õ����Ѱ��������)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5400
      TabIndex        =   11
      Top             =   960
      Width           =   2160
   End
   Begin VB.Label Lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5565
      TabIndex        =   10
      Top             =   210
      Width           =   360
   End
   Begin VB.Label Lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1920
      TabIndex        =   9
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmEditInOutClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IntEditState As Integer             '1-������2-�޸�
Private BlnModifySuccess As Boolean         '�Ƿ�༭�ɹ�
Private BlnStartUp As Boolean
Private strSQL As String                    'Sql���
Private RecClass As New ADODB.Recordset     'ҩƷ���ݷ���
Private BlnRunTime As Boolean               '�Ƿ����ڶ�̬װ��
'----�޸�ʱ����
Private Lng���ID As Long                   '���ID
Private strCode As String                   '����
Private strName As String                   '����
Private StrInOut As Integer                 '���ϵ��
Private mstrKey As String                   '������¼������Ҫ���keyֵ

Public Property Get EditState() As Integer
    EditState = IntEditState
End Property

Public Property Let EditState(ByVal vNewValue As Integer)
    IntEditState = vNewValue
End Property

Public Property Get ���ID() As Long
    ���ID = Lng���ID
End Property

Public Property Let ���ID(ByVal vNewValue As Long)
    Lng���ID = vNewValue
End Property

Public Property Get ����() As String
    ���� = strCode
End Property

Public Property Let ����(ByVal vNewValue As String)
    strCode = vNewValue
End Property

Public Property Get ����() As String
    ���� = strName
End Property

Public Property Let ����(ByVal vNewValue As String)
    strName = vNewValue
End Property

Public Property Get ϵ��() As String
    ϵ�� = StrInOut
End Property

Public Property Let ϵ��(ByVal vNewValue As String)
    StrInOut = vNewValue
End Property

Private Sub Cbo����_Click()
    '����û�ѡ��������Ƿ���ȷ����ֹ���û����������ʱѡ�����ڳ�������ʱ����ѡ��ĵ��ݣ�Ȼ���û���Ϊ���Ᵽ������������ݣ�
    Dim ItemSelect As ListItem
    
    mstrKey = ""
    DependOnCheck
    LoadInLvw
    '���û���ѡ��ĵ����ٴ�ѡ�񣬿�����Ƿ�ѡ��
    For Each ItemSelect In Lvw���ݷ����б�.ListItems
        ItemSelect.Selected = True
        ItemSelect.Checked = False
        ItemSelect.Ghosted = Not CheckItemCheck
    Next
    Call RemoveList(mstrKey)
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    '--�Ϸ��Լ��--
    If Trim(Txt����) = "" Then
        MsgBox "���벻��Ϊ�գ�", vbInformation, gstrSysName
        Txt����.SetFocus
        Exit Sub
    End If
    If Trim(Txt����) = "" Then
        MsgBox "���Ʋ���Ϊ�գ�", vbInformation, gstrSysName
        Txt����.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Txt����) Then
        MsgBox "�����к��зǷ��ַ���", vbInformation, gstrSysName
        Txt����.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Txt����, vbFromUnicode)) > 20 Then
        MsgBox "���Ƴ����������20���ַ���10�����֣�", vbInformation, gstrSysName
        Txt����.SetFocus
        Exit Sub
    End If
    Txt���� = Trim(Txt����)
    If Len(Txt����) <> 3 Then Txt���� = String(3 - Len(Txt����), "0") & Txt����
    
    '--����--
    Dim ItemThis As ListItem
    On Error GoTo ErrHand
    
    gcnOracle.BeginTrans
    Select Case EditState
    Case 1
        
        '--����ҩƷ������--
        Lng���ID = zlDatabase.GetNextId("ҩƷ������")
        gstrSQL = "zl_ҩƷ������_insert (" & Lng���ID & ",'" & Txt���� & "','" & Txt���� & "'," & Me.Cbo����.ItemData(Me.Cbo����.ListIndex) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-����ҩƷ������")
    Case 2
        '--�޸�ҩƷ������--
        gstrSQL = "zl_ҩƷ������_update (" & Lng���ID & ",'" & Txt���� & "','" & Txt���� & "'," & Me.Cbo����.ItemData(Me.Cbo����.ListIndex) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-�޸�ҩƷ������")
        '--ɾ��ҩƷ��������--
        gstrSQL = "zl_ҩƷ��������_delete (" & Lng���ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ɾ��ҩƷ��������")
    End Select
        
    '--���β���ҩƷ��������--
    For Each ItemThis In Lvw���ݷ����б�.ListItems
        With ItemThis
            If .Checked Then
                gstrSQL = "zl_ҩƷ��������_insert (" & Lng���ID & "," & Mid(.Key, 3) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���β���ҩƷ��������")
            End If
        End With
    Next
    gcnOracle.CommitTrans
    
    BlnModifySuccess = True  '���ӳɹ�
    Call frmMedInOutClass.EditReturn(BlnModifySuccess)
    '--����Ϊ����״̬--
    If EditState = 1 Then
        ClearConsForAddNew
    Else
        Unload Me
    End If
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Call frmMedInOutClass.EditReturn(BlnModifySuccess)
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    Dim lng���볤�� As Long
    Dim lng���Ƴ��� As Long
    
    gstrSQL = "Select ����,���� From ҩƷ������ Where ID = 0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    lng���볤�� = rsTmp.Fields("����").DefinedSize
    lng���Ƴ��� = rsTmp.Fields("����").DefinedSize
    
    Txt����.MaxLength = lng���볤��
    Txt����.MaxLength = lng���Ƴ���
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strTemp As String
    
    BlnStartUp = False
    BlnModifySuccess = False
    
    Call GetDefineSize '����ֶγ���
    
    If DependOnCheck = False Then Exit Sub
    If LoadInIcon = False Then Exit Sub
    LoadInLvw
    
    With Me.Cbo����
        .Clear
        .AddItem "���"
        .ItemData(.NewIndex) = 1
        .AddItem "����"
        .ItemData(.NewIndex) = -1
        .ListIndex = 0
    End With
    
    If EditState = 1 Then
        Me.Txt���� = GetMaxCode()
        Me.Cbo����.ListIndex = IIF(ϵ�� = 1, 0, 1)
    Else
        SetSelect
    End If
    BlnStartUp = True
End Sub

Private Function LoadInIcon() As Boolean
    '--Ϊ���ؼ�װ��ͼ��--
    On Error Resume Next
    Err = 0
    LoadInIcon = False
    
    '--�б�Lvw��������--
    With ImgLvw����Small
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add , , LoadResPicture("BILL1", vbResIcon)
    End With
    With Lvw���ݷ����б�
        Set .SmallIcons = ImgLvw����Small
    End With
    
    '--�б�Lvw��������--
    With ImgLvwSmall
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add , , LoadResPicture("CLASS", vbResIcon)
    End With
    With Lvw������
        Set .SmallIcons = ImgLvwSmall
    End With
    
    If Err <> 0 Then
        MsgBox "�����Դ�ļ���ʧ�����������������ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function

Private Function DependOnCheck() As Boolean
    DependOnCheck = False
    '--�������ݼ��--
    On Error GoTo errHandle
'        If .State = 1 Then .Close
    strSQL = "Select ����,����,����,˵�� From ҩƷ���ݷ��� Order by ����"
        
'        Call SQLTest(App.Title, Me.Caption, strSQL)
    Set RecClass = zlDatabase.OpenSQLRecord(strSQL, "DependOnCheck")
'        Call SQLTest
    With RecClass
        If .EOF Then
            MsgBox "ҩƷ���ݷ������ݲ�ȫ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadInLvw()
    '--�����е��ݷ�������--
    
    Dim ItemThis As ListItem
    
    Lvw���ݷ����б�.ListItems.Clear
    With RecClass
        Do While Not .EOF
            Set ItemThis = Lvw���ݷ����б�.ListItems.Add(, "K_" & !����, !����, , 1)
            ItemThis.SubItems(1) = IIF(IsNull(!˵��), "", !˵��)
            ItemThis.Tag = IIF(IsNull(!����), 1, !����)
            
            .MoveNext
        Loop
    End With
    
    With Lvw���ݷ����б�
        .ListItems(1).Selected = True
        .SelectedItem.Selected = True
    End With
    Lvw���ݷ����б�_ItemClick Lvw���ݷ����б�.SelectedItem
End Function

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    If BlnRunTime Then
        MsgBox "���ڶ�̬װ�����ݣ����Ժ�...", vbInformation, gstrSysName
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub Lvw���ݷ����б�_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw���ݷ����б�
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw���ݷ����б�_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'δ��--����Ƿ���������Ϊָ�����ݵ�������--
    If Item.Ghosted Then Item.Checked = False
End Sub

Private Sub Lvw���ݷ����б�_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '--��ʾ��ѡ��ĵ��ݷ����Ѱ�����ҩƷ������--
    Call װ��������Ѱ�����������
End Sub

Private Sub Lvw���ݷ����б�_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '--ʹ��ǰ������ڵ�Item��Ϊѡ��--
    
    Dim ItemThis As ListItem
    If Button <> 1 And Button <> 2 Then Exit Sub
    On Error Resume Next
    Err = 0
    
    With Lvw���ݷ����б�
        Set ItemThis = .HitTest(X, Y)
        If Err <> 0 Then Exit Sub
        
        ItemThis.Selected = True
        .SelectedItem.Selected = True
    End With
    Lvw���ݷ����б�_ItemClick Lvw���ݷ����б�.SelectedItem
End Sub

Private Function GetMaxCode() As String
    '--��ȡ���ı���--
    Dim RecGetMaxCode As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    
'        If .State = 1 Then .Close
    strSQL = "Select Max(����) ���� From ҩƷ������"
        
'        Call SQLTest(App.Title, Me.Caption, strSQL)
    Set RecGetMaxCode = zlDatabase.OpenSQLRecord(strSQL, "GetMaxCode")
'        Call SQLTest
    With RecGetMaxCode
        If .EOF Then
            GetMaxCode = "01"
        Else
            If IsNull(!����) Then
                GetMaxCode = "01"
            Else
                GetMaxCode = CInt(!����) + 1
                If Len(GetMaxCode) > 2 Then
                    GetMaxCode = "01"
                Else
                    GetMaxCode = String(2 - Len(GetMaxCode), "0") & GetMaxCode
                End If
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetSelect()
    '--װ������--
    '--������ҩƷ������������Lvw���ݷ����б�����ѡ��״̬--
    
    Dim RecSetSelect As New ADODB.Recordset
    
    Me.Txt���� = ����
    Me.Txt���� = ����
    
    If ϵ�� = -1 Then Me.Cbo����.ListIndex = 1
    
    On Error GoTo errHandle
    strSQL = "Select ���� From ҩƷ�������� Where ���ID=[1] "
    Set RecSetSelect = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Lng���ID)

    With RecSetSelect
        If .EOF Then Exit Sub
        Do While Not .EOF
            With RecClass
                .MoveFirst
                .Find "����=" & RecSetSelect!����
                If Not .EOF Then Lvw���ݷ����б�.ListItems("K_" & RecSetSelect!����).Checked = True
            End With
            
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ClearConsForAddNew()
    '�����ؿؼ����ݣ�Ϊ������һ����׼��
    Dim ItemThis As ListItem
    Me.Txt���� = GetMaxCode()
    Me.Txt���� = ""
    
    For Each ItemThis In Lvw���ݷ����б�.ListItems
        ItemThis.Checked = False
    Next
    Call Cbo����_Click
    Me.Txt����.SetFocus
End Sub

Private Function CheckItemCheck() As Boolean
    '--����Ƿ���������Ϊָ�����ݵ�������--
    Dim RecCheck As New ADODB.Recordset
    Dim IntBillStyle As Integer
    
    CheckItemCheck = False
    
    On Error GoTo errHandle
    IntBillStyle = Lvw���ݷ����б�.SelectedItem.Tag
    If RecCheck.State = 1 Then RecCheck.Close
    
    Select Case IntBillStyle
    Case "1", "2"   'ֻ����һ�����/ֻ����һ�ֳ���
        If Me.Cbo����.ItemData(Me.Cbo����.ListIndex) = IIF(IntBillStyle = 1, -1, 1) Then
            mstrKey = mstrKey & "|" & Lvw���ݷ����б�.SelectedItem.Key
            Exit Function  'ֻ����һ�����ʱ����ǰ�ǳ������˳�����֮����Ȼ
        End If
        strSQL = "Select ���ID From ҩƷ�������� Where ����=[1]"
        Set RecCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(Lvw���ݷ����б�.SelectedItem.Key, 3)))

        With RecCheck
            If Not .EOF Then
                If Not IsNull(!���ID) Then
                    If EditState = 1 Then
                        mstrKey = mstrKey & "|" & Lvw���ݷ����б�.SelectedItem.Key
                        Exit Function
                    End If
                    If !���ID <> ���ID Then
                        mstrKey = mstrKey & "|" & Lvw���ݷ����б�.SelectedItem.Key
                        Exit Function
                    End If
                End If
            End If
        End With
        
    Case "3"    'ֻ����һ����⼰����
        strSQL = " Select ID,ϵ�� From ҩƷ������ Where ID IN " & _
                 " (Select ���ID From ҩƷ�������� Where ����=[1])"
        Set RecCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(Lvw���ݷ����б�.SelectedItem.Key, 3)))
        
        With RecCheck
            .Find "ϵ��=" & Me.Cbo����.ItemData(Me.Cbo����.ListIndex)
            If Not .EOF Then
                If Not IsNull(!ID) Then
                    If EditState = 1 Then
                        mstrKey = mstrKey & "|" & Lvw���ݷ����б�.SelectedItem.Key
                        Exit Function
                    End If
                    If !ID <> ���ID Then
                        mstrKey = mstrKey & "|" & Lvw���ݷ����б�.SelectedItem.Key
                        Exit Function
                    End If
                End If
            End If
        End With
        
    Case "4", "5"   '����������/������ֳ���
        If Me.Cbo����.ItemData(Me.Cbo����.ListIndex) = IIF(IntBillStyle = 4, -1, 1) Then
            mstrKey = mstrKey & "|" & Lvw���ݷ����б�.SelectedItem.Key
            Exit Function  'ֻ����һ�����ʱ����ǰ�ǳ������˳�����֮����Ȼ
        End If
    End Select
    
    CheckItemCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckData()
    '����û�ѡ��������Ƿ���ȷ����ֹ���û����������ʱѡ�����ڳ�������ʱ����ѡ��ĵ��ݣ�Ȼ���û���Ϊ���Ᵽ������������ݣ�
    Dim ItemSelect As ListItem
    
    '���û���ѡ��ĵ����ٴ�ѡ�񣬿�����Ƿ�ѡ��
    For Each ItemSelect In Lvw���ݷ����б�.ListItems
        If ItemSelect.Checked Then
            ItemSelect.Selected = True
            Call CheckItemCheck
        End If
    Next
    Call RemoveList(mstrKey)
End Function

Private Sub Lvw������_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw������
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw������_GotFocus()
    OS.OpenImeByName
End Sub

Private Sub Txt����_GotFocus()
    zlControl.TxtSelAll Txt����
End Sub

Private Sub Txt����_GotFocus()
    zlControl.TxtSelAll Txt����
    If GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", "") <> "" Then
        OS.OpenImeByName GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", "")
    End If
End Sub

Private Sub װ��������Ѱ�����������()
    Dim RecLoad As New ADODB.Recordset
    Dim ItemThis As ListItem
    Dim strBegin As String, StrMiddle As String, strEnd As String
    Dim StrLoad As String, str���� As String
    Dim StrShow As String '��ʾ������
    Dim IntStyle As Integer '���ϵ��
    
    On Error GoTo errHandle
    strBegin = " select '['||����||']'||���� ������,nvl(ϵ��,1) ϵ�� From ҩƷ������" & _
               " Where ID IN (select ���ID from ҩƷ�������� where ����=[1] "
    strEnd = " ) Order by ϵ�� Desc"
    
    Lvw������.ListItems.Clear
    With Lvw���ݷ����б�
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
    
        Set ItemThis = .SelectedItem
    End With
    
    StrShow = ""
    str���� = Mid(ItemThis.Key, 3)
        
    StrLoad = strBegin & strEnd
    Set RecLoad = zlDatabase.OpenSQLRecord(StrLoad, Me.Caption, Val(str����))
    
    With RecLoad
        Do While Not .EOF
            Set ItemThis = Lvw������.ListItems.Add(, , !������, , 1)
            ItemThis.SubItems(1) = IIF(!ϵ�� = 1, "���", "����")
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Txt����_LostFocus()
    If GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", "") <> "" Then OS.OpenImeByName
End Sub

Private Sub RemoveList(ByVal strKey As String)
    '�Ƴ��������������б�
    Dim i As Integer
    
    If strKey <> "" Then
        For i = 1 To UBound(Split(strKey, "|"))
            Lvw���ݷ����б�.ListItems.Remove (Split(strKey, "|")(i))
        Next
    End If
End Sub
