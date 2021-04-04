VERSION 5.00
Begin VB.Form frmDiagnoseAdviceClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����Ϸ���"
   ClientHeight    =   2355
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5805
   Icon            =   "frmDiagnoseAdviceClass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk 
      Caption         =   "������ı��볤�ȣ������˵�����ͬ������(&L)"
      Height          =   285
      Left            =   210
      TabIndex        =   10
      Top             =   1875
      Width           =   4095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4590
      TabIndex        =   7
      Top             =   225
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4590
      TabIndex        =   8
      Top             =   705
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   1605
      Left            =   90
      TabIndex        =   9
      Top             =   105
      Width           =   4380
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   300
         Width           =   2100
      End
      Begin VB.TextBox txtParentCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   255
         Index           =   0
         Left            =   3930
         TabIndex        =   6
         Top             =   1065
         Width           =   255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   645
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1035
         Width           =   3255
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&B)"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   0
         Top             =   330
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   285
         TabIndex        =   2
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ϼ�(&S)"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   4
         Top             =   1110
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmDiagnoseAdviceClass"
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
Private mlngSvrMaxLen As Long

'Private usrSaveItem As Items

'�������Զ�����̻���************************************************************************************************

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long, ByVal lngUpKey As Long) As Boolean
    
    mblnStartUp = True
    mblnOK = False
    
    mlngUpKey = lngUpKey
    mlngKey = lngKey
        
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    If mlngKey > 0 Then
        
        '�޸Ĵ��ڵ���Ŀ
        If ReadData(mlngKey) = False Then Exit Function
        
    Else
        
        '����������,����ȱʡ�ı���
        If Not NewDefaultCode(mlngUpKey, txtParentCode, txt(0), chk) Then Exit Function
        
    End If
    
    Call AdjustCodePostion(txtParentCode, txt(0))
    
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function AnalyzeCode(ByVal lngKey As Long, ByRef objTxtParent As Object, ByRef objTxt As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:�ֽ����
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT B.���� AS �ϼ�����,A.���� FROM �����Ͻ��� A,�����Ͻ��� B WHERE A.�ϼ����=B.���(+) AND A.���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF Then Exit Function
    
    objTxtParent.Text = zlCommFun.NVL(rs("�ϼ�����").Value)
    objTxt.Text = zlCommFun.NVL(rs("����").Value)
    
    If Len(objTxt.Text) >= Len(objTxtParent.Text) Then objTxt.Text = Mid(objTxt.Text, Len(objTxtParent.Text) + 1)
    
    objTxt.MaxLength = Len(objTxt.Text)
    objTxt.Tag = rs.Fields(1).DefinedSize - Len(objTxtParent.Text)
    
    AnalyzeCode = True
End Function

Private Function NewDefaultCode(ByVal lngUpKey As Long, ByRef objTxtParent As Object, ByRef objTxt As Object, ByRef objChk As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ����
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim intMaxLength As Integer
    Dim str������ As String
    
    '��ȡ�ϼ�����
    strSQL = "SELECT B.���� AS �ϼ����� FROM �����Ͻ��� B WHERE B.���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUpKey)
    intMaxLength = rs.Fields(0).DefinedSize
    
    If rs.BOF = False Then
        objTxtParent.Text = zlCommFun.NVL(rs("�ϼ�����").Value)
    Else
        objTxtParent.Text = ""
    End If
    
    If intMaxLength = Len(objTxtParent.Text) Then
        MsgBox "������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������", vbExclamation, gstrSysName
        objTxt.Text = Space(objTxt.MaxLength)
        objChk.Value = 0
        objChk.Enabled = False
        Exit Function
    End If
        
    '��ȡͬ��������+1
    If lngUpKey = 0 Then
        strSQL = "SELECT MAX(B.����) AS ������ FROM �����Ͻ��� B WHERE B.ĩ��=0 AND B.�ϼ���� IS NULL "
         Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        strSQL = "SELECT MAX(B.����) AS ������ FROM �����Ͻ��� B WHERE B.ĩ��=0 AND B.�ϼ����=[1]"
         Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUpKey)
    End If
   
    If rs.BOF = False Then
        str������ = Trim(zlCommFun.NVL(rs("������").Value, ""))
    End If
    
    If str������ = "" Then
    
        objChk.Value = 1
        objTxt.Text = "01"
        objTxt.MaxLength = intMaxLength - Len(objTxtParent.Text)
        objTxt.Tag = objTxt.MaxLength
        objChk.Enabled = False
    Else
        
        objChk.Value = 0
        objTxt.MaxLength = Len(str������) - Len(objTxtParent.Text)    '��ǰ����ĳ���
        objTxt.Tag = intMaxLength - Len(objTxtParent.Text)              '�������ĳ���
        
        objChk.Enabled = True
        
        If Mid(str������, Len(objTxtParent.Text) + 1) = String(objTxt.MaxLength, "9") Then
            If objTxt.MaxLength >= intMaxLength Then
                MsgBox "������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������", vbExclamation, gstrSysName
                
                objChk.Value = 0
                objChk.Enabled = False
                objTxt.Text = Space(objTxt.MaxLength)
                                
                Exit Function
            Else
                MsgBox "�������Ѿ��ﵽ�������ƣ������������볤����������Ҫ", vbExclamation, gstrSysName
                
                objChk.Value = 1
                objTxt.Text = "1" & String(objTxt.MaxLength, "0")
                objTxt.MaxLength = objTxt.MaxLength + 1
                objTxt.Tag = intMaxLength - Len(objTxtParent.Text)
                
                
            End If
        Else
            objTxt.Text = Format(Mid(str������, Len(objTxtParent.Text) + 1) + 1, String(objTxt.MaxLength, "0"))
        End If
    End If
    
    NewDefaultCode = True
End Function

Private Function AdjustCodePostion(ByRef objTxtParent As Object, ByRef objTxt As Object) As Boolean
    
    objTxt.Top = objTxtParent.Top + 45
    objTxt.Left = objTxtParent.Left + TextWidth(objTxtParent.Text) + 60
    objTxt.Width = objTxtParent.Width - TextWidth(objTxtParent.Text) - 120
    objTxt.BackColor = objTxtParent.BackColor
    
    AdjustCodePostion = True
    
End Function
        
Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT * FROM �����Ͻ��� WHERE ���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    
    If rs.BOF = False Then
        'txt(0).Text = zlCommFun.NVL(rs("����").Value)
        txt(1).Text = zlCommFun.NVL(rs("����").Value)
    End If
    
    Call AnalyzeCode(lngKey, txtParentCode, txt(0))
    
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
        
    txtParentCode.MaxLength = GetMaxLength("�����Ͻ���", "����")
    txt(1).MaxLength = GetMaxLength("�����Ͻ���", "����")
            
    gstrSQL = "SELECT '['||����||']'||���� AS ���� FROM �����Ͻ��� WHERE ���=[1]"
   Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngUpKey)
    If rs.BOF = False Then
        
        txt(2).Text = zlCommFun.NVL(rs("����"))
        cmd(0).Tag = mlngUpKey
        
    End If
                        
    InitData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:У��༭���ݵ���Ч��
    '------------------------------------------------------------------------------------------------------------------
    
    If txt(0).MaxLength = 0 Then
        ShowSimpleMsg "�ϼ������Ѿ��ﵽ��󳤶ȣ����������¼���"
        cmdCancel.SetFocus
        Exit Function
    End If
    
    If chk.Value = 0 And Len(Trim(txt(0).Text)) <> txt(0).MaxLength Then
        ShowSimpleMsg "���볤�ȱ���Ϊ" & txt(0).MaxLength & "λ��������ѡ����ĳ���ѡ��"
        LocationObj txt(0)
        Exit Function
    End If
    
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
        strSQL(ReDimArray(strSQL)) = "ZL_�����Ͻ���_INSERT(" & lngKey & ",'" & Trim(txtParentCode.Text & txt(0).Text) & "','" & txt(1).Text & "',NULL,NULL,NULL,NULL,NULL," & Val(cmd(0).Tag) & ",0," & chk.Value & ")"
    Else
        '�޸�����
        lngKey = mlngKey
        strSQL(ReDimArray(strSQL)) = "ZL_�����Ͻ���_UPDATE(" & lngKey & ",'" & Trim(txtParentCode.Text & txt(0).Text) & "','" & txt(1).Text & "',NULL,NULL,0," & Val(cmd(0).Tag) & "," & chk.Value & ")"
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
    
    gstrSQL = "SELECT NVL(MAX(���),0)+1 AS ��� FROM �����Ͻ���"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then GetMaxNo = rs("���").Value
        
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub chk_Click()

    If chk.Value = 1 Then
        mlngSvrMaxLen = txt(0).MaxLength
        'txt(0).MaxLength = txtParentCode.MaxLength - Len(txtParentCode.Text)
        txt(0).MaxLength = Val(txt(0).Tag)
    Else
        txt(0).MaxLength = mlngSvrMaxLen
        txt(0).Text = Mid(txt(0).Text, 1, txt(0).MaxLength)
    End If
    
    cmdOK.Tag = "Changed"
    
    On Error Resume Next
    txt(0).SetFocus
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    
    gstrSQL = "SELECT 0 As ĩ��,-1 AS ID,0 AS �ϼ�id,'���з���' AS ����,'' AS ���� FROM DUAL " & _
                "UNION ALL " & _
                "SELECT 0 As ĩ��,��� AS ID,DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�id,'['||����||']'||���� AS ����,���� FROM �����Ͻ��� WHERE Nvl(ĩ��,0)=0 START WITH �ϼ���� IS NULL and ���<>" & mlngKey & " CONNECT BY PRIOR ���=�ϼ���� "

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Call ClientToScreen(txt(2).hWnd, objPoint)
    
    If frmSelectDialog.ShowSelect(Me, 1, rs, "", "�������ѡ��һ������", objPoint.X * 15 - 30, objPoint.Y * 15 + txt(2).Height - 30, txt(2).Width, 3900, txt(2).Height, mlngKey, Me.Name & "\�����Ϸ���ѡ��", , False) Then
                
        If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
            If zlCommFun.NVL(rs("ID")) = -1 Then
                txt(2).Text = ""
                cmd(0).Tag = ""
            Else
                txt(2).Text = zlCommFun.NVL(rs("����"))
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                
            End If
                                   
            Call NewDefaultCode(Val(cmd(0).Tag), txtParentCode, txt(0), chk)
                        
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
        On Error Resume Next
        Call mfrmMain.EditRefresh("�����Ϸ���", lngKey)
        On Error GoTo 0
        
        If mlngKey = 0 Then
            
            Call NewDefaultCode(Val(cmd(0).Tag), txtParentCode, txt(0), chk)
            
            txt(1).Text = ""
                            
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
    Case 1, 2
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        If Index = 2 Then zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 2
        zlCommFun.OpenIme False
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

Private Sub txtParentCode_Change()
    Call AdjustCodePostion(txtParentCode, txt(0))
End Sub
