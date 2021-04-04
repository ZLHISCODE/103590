VERSION 5.00
Begin VB.Form frmKindEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   3270
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6240
   Icon            =   "frmKindEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Height          =   3165
      Left            =   105
      TabIndex        =   15
      Top             =   0
      Width           =   4545
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2355
         Width           =   3255
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   255
         Index           =   0
         Left            =   4050
         TabIndex        =   12
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox txt 
         Height          =   825
         Index           =   3
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1080
         TabIndex        =   5
         Top             =   1035
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   645
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2730
         Width           =   3255
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&R)"
         Height          =   180
         Index           =   5
         Left            =   405
         TabIndex        =   8
         Top             =   2430
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ϼ�(&U)"
         Height          =   180
         Index           =   4
         Left            =   405
         TabIndex        =   10
         Top             =   2790
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "˵��(&T)"
         Height          =   180
         Index           =   3
         Left            =   405
         TabIndex        =   6
         Top             =   1470
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   2
         Left            =   405
         TabIndex        =   4
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   405
         TabIndex        =   2
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&B)"
         Height          =   180
         Index           =   0
         Left            =   405
         TabIndex        =   0
         Top             =   330
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4995
      TabIndex        =   14
      Top             =   765
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4995
      TabIndex        =   13
      Top             =   285
      Width           =   1100
   End
End
Attribute VB_Name = "frmKindEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mlngUpKey As Long

'�������Զ�����̻���************************************************************************************************

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long, ByVal lngUpKey As Long) As Boolean
    
    mblnStartUp = True
    mblnOK = False
    
    mlngKey = lngKey
    mlngUpKey = lngUpKey
        
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    If mlngKey > 0 Then
        
        '�޸Ĵ��ڵ���Ŀ
        If ReadData(mlngKey) = False Then Exit Function
    Else
        
        '����������,����ȱʡ�ı���
        
        txt(0).Text = NewDefaultCode(mlngUpKey)
        
    End If
    
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function NewDefaultCode(ByVal lngUpKey As Long) As String
    
    '------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ����
    '------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ����
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim intMaxLength As Integer
    Dim str������ As String
    Dim str�ϼ����� As String
    
    '��ȡ�ϼ�����
    strSQL = "SELECT B.���� AS �ϼ����� FROM ������� B WHERE B.���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUpKey)
    If rs.BOF Then Exit Function
    
    intMaxLength = rs.Fields(0).DefinedSize
    str�ϼ����� = zlCommFun.NVL(rs("�ϼ�����").Value)
            
    If intMaxLength = Len(str�ϼ�����) Then
        MsgBox "������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������", vbExclamation, gstrSysName
        Exit Function
    End If
        
    '��ȡͬ��������+1
    If lngUpKey = 0 Then
        strSQL = "SELECT MAX(B.����) AS ������ FROM ������� B WHERE B.ĩ��=1 AND B.�ϼ���� IS NULL "
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        strSQL = "SELECT MAX(B.����) AS ������ FROM ������� B WHERE B.ĩ��=1 AND B.�ϼ����=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUpKey)
    End If
    
    If rs.BOF Then Exit Function
    
    str������ = Trim(zlCommFun.NVL(rs("������").Value, ""))
  
    If str������ = "" Then
        str������ = str�ϼ����� & "001"
    Else
        str������ = Format(Val(str������) + 1, String(Len(str������), "0"))
    End If
    
    NewDefaultCode = str������
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT * FROM ������� WHERE ���=[1]"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        txt(0).Text = zlCommFun.NVL(rs("����").Value)
        txt(1).Text = zlCommFun.NVL(rs("����").Value)
        txt(2).Text = zlCommFun.NVL(rs("����").Value)
        txt(3).Text = zlCommFun.NVL(rs("˵��").Value)
        
        zlControl.CboLocate cbo, zlCommFun.NVL(rs("���÷�Χ").Value, 0), True
        
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
        
    txt(0).MaxLength = GetMaxLength("�������", "����")
    txt(1).MaxLength = GetMaxLength("�������", "����")
    txt(2).MaxLength = GetMaxLength("�������", "����")
    txt(3).MaxLength = GetMaxLength("�������", "˵��")
        
    gstrSQL = "SELECT '['||����||']'||���� AS ���� FROM ������� WHERE ���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngUpKey)
    If rs.BOF = False Then
        
        txt(4).Text = zlCommFun.NVL(rs("����"))
        cmd(0).Tag = mlngUpKey
        
    End If
    
    cbo.Clear
    
    cbo.AddItem "0-����"
    cbo.ItemData(cbo.NewIndex) = 0
    
    cbo.AddItem "1-����"
    cbo.ItemData(cbo.NewIndex) = 1
    
    cbo.AddItem "2-����"
    cbo.ItemData(cbo.NewIndex) = 2
    
    cbo.ListIndex = 0
    
    InitData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:У��༭���ݵ���Ч��
    '------------------------------------------------------------------------------------------------------------------
    If Trim(txt(0).Text) = "" Then
        ShowSimpleMsg "���벻��Ϊ��ֵ���������룡"
        LocationObj txt(0)
        Exit Function
    End If
    
    '�������Ƿ�Ϊ�����ַ�
    If CheckStrType(Trim(txt(0).Text), 99, "0123456789") = False Then
        ShowSimpleMsg "�������Ϊ�����ַ���"
        LocationObj txt(0)
        Exit Function
    End If
    
    If Trim(txt(1).Text) = "" Then
        ShowSimpleMsg "���Ʋ���Ϊ��ֵ���������룡"
        LocationObj txt(1)
        Exit Function
    End If
    
    ValidEdit = True
    
End Function

Private Function SaveEdit(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    If mlngKey = 0 Then
        '��������
        
        lngKey = GetMaxNo
        strSQL(ReDimArray(strSQL)) = "ZL_�������_INSERT(" & lngKey & ",'" & Trim(txt(0).Text) & "','" & txt(1).Text & "','" & txt(2).Text & "','" & txt(3).Text & "'," & cbo.ItemData(cbo.ListIndex) & "," & Val(cmd(0).Tag) & ",1)"
    Else
        '�޸�����
        lngKey = mlngKey
        strSQL(ReDimArray(strSQL)) = "ZL_�������_UPDATE(" & lngKey & ",'" & Trim(txt(0).Text) & "','" & txt(1).Text & "','" & txt(2).Text & "','" & txt(3).Text & "'," & cbo.ItemData(cbo.ListIndex) & "," & Val(cmd(0).Tag) & ")"
    End If
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function GetMaxNo() As Long
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT NVL(MAX(���),0)+1 AS ��� FROM �������"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then GetMaxNo = rs("���").Value
        
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub chk_Click()
    cmdOK.Tag = "Changed"
End Sub

Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo_Click()
    cmdOK.Tag = "Changed"
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    
    gstrSQL = "SELECT 0 As ĩ��,-1 AS ID,0 AS �ϼ�id,'���з���' AS ����,'' AS ���� FROM DUAL " & _
                "UNION ALL " & _
                "SELECT 0 As ĩ��,��� AS ID,DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�id,'['||����||']'||���� AS ����,���� FROM ������� WHERE ĩ��=0 START WITH �ϼ���� IS NULL AND ���<>" & mlngKey & " CONNECT BY PRIOR ���=�ϼ���� "

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Call ClientToScreen(txt(4).hWnd, objPoint)
    
    If frmSelectDialog.ShowSelect(Me, 1, rs, "", "�������ѡ��һ������", objPoint.X * 15 - 30, objPoint.Y * 15 + txt(4).Height - 30, txt(4).Width, 3900, txt(4).Height, mlngKey, Me.Name & "\������ͷ���ѡ��", , False) Then
    
        If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
            If zlCommFun.NVL(rs("ID")) = -1 Then
'                mstr�ϼ����� = ""
                txt(4).Text = ""
                cmd(0).Tag = ""
            Else
'                mstr�ϼ����� = zlCommFun.NVL(rs("����"))
                txt(4).Text = zlCommFun.NVL(rs("����"))
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                
            End If
                       
            cmdOK.Tag = "Changed"
        End If
    End If
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngKey As Long
        
    If cmdOK.Tag <> "" Then
            
        If ValidEdit() = False Then Exit Sub
        If SaveEdit(lngKey) = False Then Exit Sub
        mblnOK = True
        
        '���µ��ô����������ʾ
        
        Call mfrmMain.EditRefresh("�������", lngKey)
        
        If mlngKey = 0 Then
            
            txt(0).Text = NewDefaultCode(Val(cmd(0).Tag))
            txt(1).Text = ""
            txt(2).Text = ""
            txt(3).Text = ""
                    
            txt(0).SetFocus
            
            cmdOK.Tag = ""
            Exit Sub
        End If
        
    End If
    
    cmdOK.Tag = ""
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 1, 3
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
        If Index = 4 Then zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Index = 0 Then
            If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 3
        zlCommFun.OpenIme False
        If Index = 1 Then
            If InStr(txt(Index).Text, "'") = 0 Then txt(2).Text = zlGetSymbol(txt(Index).Text)
        End If
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub
