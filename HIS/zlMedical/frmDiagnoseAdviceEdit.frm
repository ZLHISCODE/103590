VERSION 5.00
Begin VB.Form frmDiagnoseAdviceEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ϱ༭"
   ClientHeight    =   5550
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7590
   Icon            =   "frmDiagnoseAdviceEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin zl9Medical.VsfGrid vsf 
      Height          =   2175
      Left            =   60
      TabIndex        =   21
      Top             =   3300
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   3836
   End
   Begin VB.Frame fra 
      Height          =   3285
      Left            =   60
      TabIndex        =   20
      Top             =   -30
      Width           =   6090
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   3465
         TabIndex        =   13
         Top             =   2865
         Width           =   330
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   5415
         TabIndex        =   16
         Top             =   2865
         Width           =   345
      End
      Begin VB.CheckBox chk 
         Caption         =   "�������(&3)"
         Height          =   240
         Index           =   2
         Left            =   4140
         TabIndex        =   15
         Top             =   2895
         Width           =   1320
      End
      Begin VB.CheckBox chk 
         Caption         =   "������(&2)"
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   12
         Top             =   2910
         Width           =   1320
      End
      Begin VB.CheckBox chk 
         Caption         =   "����(&1)"
         Height          =   240
         Index           =   0
         Left            =   1140
         TabIndex        =   11
         Top             =   2910
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   1110
         Index           =   6
         Left            =   1140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1320
         Width           =   4785
      End
      Begin VB.CommandButton cmd 
         Caption         =   "��"
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   10
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1140
         TabIndex        =   5
         Top             =   960
         Width           =   4785
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1140
         TabIndex        =   3
         Top             =   600
         Width           =   4785
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1140
         TabIndex        =   1
         Top             =   240
         Width           =   4785
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2490
         Width           =   4785
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   6
         Left            =   3810
         TabIndex        =   14
         Top             =   2925
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   5775
         TabIndex        =   17
         Top             =   2895
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��Ͻ���(&A)"
         Height          =   180
         Index           =   5
         Left            =   90
         TabIndex        =   6
         Top             =   1335
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��������(&U)"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   8
         Top             =   2550
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��ϼ���(&S)"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   1050
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�������(&N)"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   2
         Top             =   675
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��ϱ���(&B)"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   330
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6375
      TabIndex        =   19
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6375
      TabIndex        =   18
      Top             =   75
      Width           =   1100
   End
End
Attribute VB_Name = "frmDiagnoseAdviceEdit"
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


Private Enum mCol
    ���� = 1
    ����
    ���
    
End Enum

'�������Զ�����̻���************************************************************************************************

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ����Ƿ����ظ�����Ŀ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey And vsf.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

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
    strSQL = "SELECT B.���� AS �ϼ����� FROM �����Ͻ��� B WHERE B.���=[1]"
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
        strSQL = "SELECT MAX(B.����) AS ������ FROM �����Ͻ��� B WHERE B.ĩ��=1 AND B.�ϼ���� IS NULL "
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        strSQL = "SELECT MAX(B.����) AS ������ FROM �����Ͻ��� B WHERE B.ĩ��=1 AND B.�ϼ����=[1]"
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
    
    gstrSQL = "SELECT A.*,Decode(C.����,Null,'','��'||C.����||'��'||C.����) As �ϼ����� " & _
            "FROM �����Ͻ��� A,�����Ͻ��� C WHERE A.�ϼ����=C.���(+) And A.���=[1]"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        txt(0).Text = zlCommFun.NVL(rs("����").Value)
        txt(1).Text = zlCommFun.NVL(rs("����").Value)
        txt(2).Text = zlCommFun.NVL(rs("����").Value)
        
        txt(4).Text = zlCommFun.NVL(rs("�ϼ�����").Value)
        cmd(0).Tag = zlCommFun.NVL(rs("�ϼ����").Value)
        
        chk(0).Value = zlCommFun.NVL(rs("�Ƿ񼲲�").Value, 0)
        
        txt(6).Text = zlCommFun.NVL(rs("��Ͻ���").Value)
        
        txt(5).Text = zlCommFun.NVL(rs("������").Value)
        txt(3).Text = zlCommFun.NVL(rs("�������").Value)
        
        chk(1).Value = IIf(Val(txt(5).Text) > 0, 1, 0)
        chk(2).Value = IIf(Val(txt(3).Text) > 0, 1, 0)
        
        txt(5).Visible = (chk(1).Value = 1)
        lbl(6).Visible = (chk(1).Value = 1)
        
        txt(3).Visible = (chk(2).Value = 1)
        lbl(3).Visible = (chk(2).Value = 1)
        
    End If
    
    gstrSQL = "SELECT B.ID,B.����,b.����,c.���� As ��� FROM ���������� A,������ĿĿ¼ B,������Ŀ��� c WHERE A.������Ŀid=B.ID And A.������=[1] And c.����=b.���"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
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
        
    txt(0).MaxLength = GetMaxLength("�����Ͻ���", "����")
    txt(1).MaxLength = GetMaxLength("�����Ͻ���", "����")
    txt(2).MaxLength = GetMaxLength("�����Ͻ���", "����")
        
    txt(6).MaxLength = GetMaxLength("�����Ͻ���", "��Ͻ���")
        
    gstrSQL = "SELECT '['||����||']'||���� AS ���� FROM �����Ͻ��� WHERE ���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngUpKey)
    If rs.BOF = False Then
        
        txt(4).Text = zlCommFun.NVL(rs("����"))
        cmd(0).Tag = mlngUpKey
        
    End If

    With vsf
        .Cols = 0
        .NewColumn "", 255
        .NewColumn "����", 2400, 1, "...", 1
        .NewColumn "����", 1500, 1
        .NewColumn "���", 900, 1
        .FixedCols = 1
    End With
    
    
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
        strSQL(ReDimArray(strSQL)) = "ZL_�����Ͻ���_INSERT(" & lngKey & ",'" & Trim(txt(0).Text) & "','" & txt(1).Text & "','" & txt(2).Text & "'," & chk(0).Value & ",'" & txt(6).Text & "'," & IIf(chk(1).Value = 0, "NULL", Val(txt(5).Text)) & "," & IIf(chk(2).Value = 0, "NULL", Val(txt(3).Text)) & "," & Val(cmd(0).Tag) & ",1)"
    Else
        '�޸�����
        lngKey = mlngKey
        strSQL(ReDimArray(strSQL)) = "ZL_�����Ͻ���_UPDATE(" & lngKey & ",'" & Trim(txt(0).Text) & "','" & txt(1).Text & "','" & txt(2).Text & "'," & chk(0).Value & ",'" & txt(6).Text & "'," & IIf(chk(1).Value = 0, "NULL", Val(txt(5).Text)) & "," & IIf(chk(2).Value = 0, "NULL", Val(txt(3).Text)) & "," & Val(cmd(0).Tag) & ")"
    End If
    
    strSQL(ReDimArray(strSQL)) = "ZL_����������_Delete(" & lngKey & ")"
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            strSQL(ReDimArray(strSQL)) = "ZL_����������_Insert(" & lngKey & "," & Val(vsf.RowData(lngLoop)) & ")"
        End If
    Next
    
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
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function GetMaxNo() As Long
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT NVL(MAX(���),0)+1 AS ��� FROM �����Ͻ���"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then GetMaxNo = rs("���").Value
        
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub chk_Click(Index As Integer)
    cmdOK.Tag = "Changed"
    
    txt(5).Visible = (chk(1).Value = 1)
    lbl(6).Visible = txt(5).Visible
    
    txt(3).Visible = (chk(2).Value = 1)
    lbl(3).Visible = txt(3).Visible
    
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    
    
    If Index = 0 Then
        gstrSQL = "SELECT 0 As ĩ��,-1 AS ID,0 AS �ϼ�id,'���з���' AS ����,'' AS ���� FROM DUAL " & _
                    "UNION ALL " & _
                    "SELECT 0 As ĩ��,��� AS ID,DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�id,'��'||����||'��'||���� AS ����,���� FROM �����Ͻ��� WHERE ĩ��=0  START WITH �ϼ���� IS NULL AND ���<>" & mlngKey & " CONNECT BY PRIOR ���=�ϼ���� "
    
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        Call ClientToScreen(txt(4).hWnd, objPoint)
        
        If frmSelectDialog.ShowSelect(Me, 1, rs, "", "�������ѡ��һ������", objPoint.X * 15 - 30, objPoint.Y * 15 + txt(4).Height - 30, txt(4).Width, 3900, txt(4).Height, mlngKey, Me.Name & "\������ͷ���ѡ��", , False) Then
            If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
                If zlCommFun.NVL(rs("ID")) = -1 Then
                    txt(4).Text = ""
                    cmd(0).Tag = ""
                Else
                    txt(4).Text = zlCommFun.NVL(rs("����"))
                    cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                End If
                           
                cmdOK.Tag = "Changed"
            End If
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
        
        Call mfrmMain.EditRefresh("������Ŀ¼", lngKey)
        
        If mlngKey = 0 Then
            
            txt(0).Text = NewDefaultCode(Val(cmd(0).Tag))
            txt(1).Text = ""
            txt(2).Text = ""
            
            txt(6).Text = ""
                    
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
    Case 1, 6
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim rs As New ADODB.Recordset
    

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
        
        If Index = 4 Then zlCommFun.PressKey vbKeyTab
                
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        If Index = 0 Then If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
                
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 6
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
    If Cancel Then Exit Sub
    
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    cmdOK.Tag = "Changed"
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf.RowData(Row)) <= 0)
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Select Case Col
        Case mCol.����
            
            gstrSQL = GetPublicSQL(SQL.�����Ŀѡ��)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
            If ShowGrdSelect(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 5100) Then
                
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
                
                vsf.EditText = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
                vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                
                cmdOK.Tag = "Changed"
                
            End If

    End Select

End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim strTmp As String
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.����
                    
                    strText = UCase(vsf.EditText)
                    gstrSQL = GetPublicSQL(SQL.�����Ŀ����ѡ��, strText)
                    
                    If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If

                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp, 1, 2)
                    
                    If ShowGrdFilter(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀ����ѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 5100) Then

                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If

                        vsf.EditText = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
                        vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                        vsf.Cell(flexcpData, Row, Col) = vsf.TextMatrix(Row, Col)
                        vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                        
                        
                        cmdOK.Tag = "Changed"
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                        vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        cmdOK.Tag = "Changed"
    End If
End Sub
