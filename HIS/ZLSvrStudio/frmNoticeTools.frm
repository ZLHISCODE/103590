VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNoticeTools 
   BackColor       =   &H8000000E&
   Caption         =   "�Զ�����"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmNoticeTools.frx":0000
   ScaleHeight     =   7365
   ScaleWidth      =   10290
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDetail 
      Caption         =   "��ϸ����(&D)"
      Height          =   350
      Left            =   4410
      TabIndex        =   8
      Top             =   4080
      Width           =   1230
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1530
      Top             =   6090
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNoticeTools.frx":803A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   3225
      TabIndex        =   6
      Top             =   4080
      Width           =   1100
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   1725
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   750
      Width           =   4185
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   930
      TabIndex        =   4
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸�(&M)"
      Height          =   350
      Left            =   2085
      TabIndex        =   5
      Top             =   4080
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6960
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2505
      Left            =   945
      TabIndex        =   3
      Top             =   1485
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   4419
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��������"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��������"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "���ѱ���"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "���Ѵ���"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�������"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "��������"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "��ʼʱ��"
         Object.Width           =   2963
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "��ֹʱ��"
         Object.Width           =   2963
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Զ����ѹ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   7
      Top             =   135
      Width           =   1440
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��ϵͳ"
      Height          =   180
      Left            =   945
      TabIndex        =   0
      Top             =   810
      Width           =   720
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   195
      Picture         =   "frmNoticeTools.frx":9D44
      Top             =   675
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ�б�"
      Height          =   180
      Index           =   1
      Left            =   945
      TabIndex        =   2
      Top             =   1230
      Width           =   1080
   End
End
Attribute VB_Name = "frmNoticeTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnStartUp As Boolean
Private mstr������ As String
Private mlngSys As Long
Private mintColumn As Integer
Private mlngLoop As Long

Private Sub AdjustMenuEnabled()
    
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdDelete.Enabled = True
    cmdDetail.Enabled = True
    
    If lvw.SelectedItem Is Nothing Then
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
        cmdDetail.Enabled = False
    End If
End Sub

Private Function GetWaveName(ByVal lngNo As Long) As String
    
    Select Case lngNo
    Case 101
        GetWaveName = "����"
    Case 102
        GetWaveName = "����ռ�"
    Case 103
        GetWaveName = "�绰����1"
    Case 104
        GetWaveName = "�绰����2"
    Case 105
        GetWaveName = "�绰��"
    Case 106
        GetWaveName = "������"
    Case 107
        GetWaveName = "����"
    Case 108
        GetWaveName = "����"
    Case 109
        GetWaveName = "��ʾ"
    Case 110
        GetWaveName = "����Ϣ"
    End Select
        
End Function

Private Function CalcTimeUnit(ByVal lngData As Long, Optional ByVal strParam As String = "") As String
    
    Dim strNumber As String
    Dim strUnit As String
    
    If lngData = 0 Then Exit Function
    
    If lngData / (24 * 60) >= 1 Then
        strNumber = lngData / (24 * 60)
        strUnit = "��"
    ElseIf (lngData / 60) >= 1 Then
        strNumber = (lngData / 60)
        strUnit = "Сʱ"
    Else
        strNumber = lngData
        strUnit = "����"
    End If
    
    Select Case strParam
    Case "����"
        CalcTimeUnit = strNumber
    Case "ʱ�䵥λ"
        CalcTimeUnit = strUnit
    Case ""
        CalcTimeUnit = strNumber & strUnit
    End Select
    
End Function

Private Function Nvl(ByVal varOld As Variant, Optional ByVal varNew As Variant = "") As Variant
    If IsNull(varOld) Then
        Nvl = varNew
    Else
        Nvl = varOld
    End If
    
End Function

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

Private Sub cboSystem_Click()
    Dim rs As New ADODB.Recordset
    Dim objItem As ListItem
    Dim strKey As String
    
    '����Ƿ����Ӧ��ϵͳ����������ȡӦ��ϵͳ��ϵͳ��ţ������˳�
    If mlngSys = cboSystem.ItemData(cboSystem.ListIndex) Then Exit Sub
    mlngSys = cboSystem.ItemData(cboSystem.ListIndex)
    
    Dim varOut As Variant
    If mlngSys <> 0 Then
        mstr������ = GetOwnerName(Val(mlngSys), gcnOracle)
    End If
    '�ȱ��浱ǰ��ѡ��״̬
    If Not (lvw.SelectedItem Is Nothing) Then strKey = lvw.SelectedItem.Key
    
    
    lvw.ListItems.Clear

    Set rs = OpenCursor(gcnOracle, "ZLTOOLS.B_Expert.Get_notices", 0, mlngSys)
 
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            Set objItem = lvw.ListItems.Add(, "K" & rs("���").Value, Nvl(rs("��������").Value), 1, 1)
            objItem.SubItems(1) = GetWaveName(Nvl(rs("��������").Value, 0))
            
            objItem.SubItems(2) = Nvl(rs("��������").Value)
            
            objItem.SubItems(3) = IIf(Nvl(rs("���Ѵ���").Value, 0) = 1, "��", "")
            
            objItem.SubItems(6) = Format(rs("��ʼʱ��").Value, "yyyy-MM-dd HH:mm")
            
            If IsNull(rs("��ֹʱ��").Value) = False Then objItem.SubItems(7) = Format(rs("��ֹʱ��").Value, "yyyy-MM-dd HH:mm")
            
            objItem.SubItems(4) = IIf(Nvl(rs("�������").Value, 0) = 0, "�������", CalcTimeUnit(Nvl(rs("�������").Value, 0)))
            objItem.SubItems(5) = CalcTimeUnit(Nvl(rs("��������").Value, 0))
            
            rs.MoveNext
        Loop
    End If
    
    
    '�ָ���ǰ��ѡ��״̬
    On Error Resume Next
    lvw.ListItems(strKey).Selected = True
    lvw.ListItems(strKey).EnsureVisible
    
    Call AdjustMenuEnabled
    
End Sub

Private Sub cmdAdd_Click()
    
    If cboSystem.ListIndex = -1 Then Exit Sub
    
    If frmNoticesEdit.ShowEdit(frmMDIMain, 0, mlngSys, mstr������) Then
        mlngSys = -1
        
        Call cboSystem_Click
        
    End If
End Sub

Private Sub CmdDelete_Click()
    Dim lngIndex As Long
    Dim strRemarks As String
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("���Ƿ����Ҫɾ��ѡ�е����ѣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '��֤��ݲ��������˵��
    If Not CheckAuditStatus("0504", "ɾ��", strRemarks) Then Exit Sub
    gstrSQL = "ZL_ZLNOTICES_DELETE(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
    
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    gcnOracle.CommitTrans
    '������Ҫ������־
    Call SaveAuditLog(3, "ɾ��", "ɾ�����ѡ�" & lvw.SelectedItem.Text & "��", strRemarks)
    lngIndex = lvw.SelectedItem.Index
    lvw.ListItems.Remove lvw.SelectedItem.Index
    Call NextLvwPos(lvw, lngIndex)
            
    Call AdjustMenuEnabled
    
    Exit Sub
    
errHand:
    gcnOracle.RollbackTrans
    MsgBox "ɾ������ʧ�ܣ�", vbInformation, gstrSysName
End Sub

Private Sub cmdDetail_Click()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If frmNoticesEdit.ShowEdit(frmMDIMain, Val(Mid(lvw.SelectedItem.Key, 2)), mlngSys, mstr������, True) Then
        
    End If
End Sub

Private Sub cmdModify_Click()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If frmNoticesEdit.ShowEdit(frmMDIMain, Val(Mid(lvw.SelectedItem.Key, 2)), mlngSys, mstr������) Then
        mlngSys = -1
        Call cboSystem_Click
    End If
End Sub

Private Sub Form_Activate()
    Dim rsTmp As New ADODB.Recordset
    Dim lngLoop As Long
    On Error GoTo ErrHandle
    If mblnStartUp = False Then Exit Sub
    
    '��ʾ����ʾ��ϵͳ
    mstr������ = UCase(gstrUserName)
    mlngSys = 0
                                
    cboSystem.AddItem "����ϵͳ����"
    cboSystem.ItemData(cboSystem.NewIndex) = 0
    cboSystem.ListIndex = 0
    
    If gblnDBA = True Then
        Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", UCase(mstr������))
    End If
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            cboSystem.AddItem rsTmp("����") & " v" & rsTmp("�汾��") & "��" & rsTmp("���") & "��"
            cboSystem.ItemData(cboSystem.NewIndex) = rsTmp("���")
            
            If gblnDBA = False Then
                If cboSystem.ListIndex < 1 Then
                    cboSystem.ListIndex = cboSystem.NewIndex
                    mlngSys = cboSystem.ItemData(cboSystem.ListIndex)
                End If
            End If
            
            rsTmp.MoveNext
            
        Loop
        
    Else
        cboSystem.Enabled = False
        lvw.Enabled = False
        cmdAdd.Enabled = False
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    End If
    
    mblnStartUp = False
    DoEvents
    
    If cboSystem.ListCount >= 0 Then
        mlngSys = -1
        Call cboSystem_Click
    End If
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Load()
    mblnStartUp = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With lvw
        .Width = Me.ScaleWidth - .Left - 120
        .Height = Me.ScaleHeight - .Top - 120 - cmdAdd.Height
    End With
    
    With cmdAdd
        .Left = lvw.Left
        .Top = lvw.Top + lvw.Height + 60
    End With
    
    With cmdModify
        .Left = cmdAdd.Left + .Width + 60
        .Top = cmdAdd.Top
    End With
    
    With cmdDelete
        .Left = cmdModify.Left + .Width + 60
        .Top = cmdAdd.Top
    End With
    
    With cmdDetail
        .Left = cmdDelete.Left + .Width + 60
        .Top = cmdAdd.Top
    End With
    
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
    
    If cmdDetail.Visible And cmdDetail.Enabled Then
        Call cmdDetail_Click
    End If
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call lvw_DblClick
    End If
End Sub
