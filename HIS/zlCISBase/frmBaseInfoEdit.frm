VERSION 5.00
Begin VB.Form frmBaseInfoEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14985
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   14985
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraEdit 
      Caption         =   "������Ϣ�༭"
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12165
      Begin VB.ComboBox cbo����1 
         Height          =   300
         Left            =   11880
         TabIndex        =   5
         Text            =   "cbo����1"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   3570
         MaxLength       =   60
         TabIndex        =   1
         Top             =   360
         Width           =   2235
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   0
         Top             =   360
         Width           =   1380
      End
      Begin VB.ComboBox cbo�����Ա� 
         Height          =   300
         Left            =   7035
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2235
      End
      Begin VB.TextBox txt˵�� 
         Height          =   720
         Left            =   1080
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   3285
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   7035
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox chkȱʡ��־ 
         Caption         =   "ȱʡ��־(ע�������־����������)"
         Height          =   255
         Left            =   4980
         TabIndex        =   8
         Top             =   840
         Width           =   3255
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   10320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3000
         TabIndex        =   14
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   13
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ա�"
         Height          =   180
         Left            =   6375
         TabIndex        =   12
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lbl˵�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   360
      End
      Begin VB.Label lbl���� 
         Caption         =   "����"
         Height          =   180
         Left            =   9720
         TabIndex        =   10
         Top             =   420
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmBaseInfoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr����OLD As String          '��ǰ��ʾ����Ŀid
Private mstr����New As String

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------

Public Function zlRefresh(strItemName As String, ByVal str���� As String) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset, rsRec As New ADODB.Recordset
    Dim strTmp As String
    
    mstr����OLD = str����
    
    '����ؼ��ı�
    Me.txt����.Text = "": Me.txt����.Text = "": Me.txt˵��.Text = ""
    Me.txt����.Text = "": Me.txt����.Text = "": Me.cbo�����Ա�.Clear
    Me.chkȱʡ��־.Value = 0: Me.cbo����.Clear

    If str���� = "" Then zlRefresh = True: Exit Function

    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand

    gstrSql = "Select * From " & strItemName & " Where ���� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str����)
    With rsTemp
        If Not .EOF Then
            Me.txt����.Text = Nvl(!����)
            Me.txt����.Text = Nvl(!����)

            Select Case Trim(strItemName)
            Case "���Ƽ���걾"
                Me.txt����.Text = Nvl(!����)
                For i = 0 To cbo�����Ա�.ListCount - 1
                    If cbo�����Ա�.List(i) = Nvl(!�����Ա�) Then
                        cbo�����Ա�.ListIndex = i
                        Exit For
                    End If
                Next
            Case "���Ƽ�������"
                Me.txt����.Text = Nvl(!����)
                Me.txt����.Text = Nvl(!����)
                Me.chkȱʡ��־.Value = Val(Nvl(!ȱʡ��־))
            Case "���鱸ע����", "������������"
                Me.txt����.Text = Nvl(!����)
                Me.txt˵��.Text = Nvl(!˵��)
        
                cbo����.Clear
                gstrSql = "select distinct ���� from ���Ƽ�������"
                Set rsRec = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
                
                With rsRec
                    Do While Not .EOF
                        cbo����.AddItem Nvl(!����)
                        .MoveNext
                    Loop
                End With
                For i = 0 To cbo����.ListCount - 1
                    If cbo����.List(i) = !���� Then
                        cbo����.ListIndex = i
                        Exit For
                    End If
                Next
            Case "������������"
                Me.txt����.Text = Nvl(!����)
                Me.txt˵��.Text = Nvl(!˵��)
            Case "�����������"
                Me.txt˵��.Text = Nvl(!����)
            Case "����걾��̬"
                Me.txt˵��.Text = Nvl(!˵��) '
            Case "����������", "����ϸ�����", "����Ⱦɫ����"
                Me.txt����.Text = Nvl(!����)
                Me.chkȱʡ��־.Value = Val(Nvl(!ȱʡ��־))
            Case "����ϸ������", "�ʿؼ��鷽��", "ϸ����ⷽ��"
                Me.txt����.Text = Nvl(!����)
            Case "�ʿر���ʾ�"
                Me.txt����.Text = Nvl(!����)
                'cbo�����Ա�.Clear
                cbo����1.Clear
                gstrSql = "Select Distinct ���� From �ʿر���ʾ�"
                Set rsRec = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
                
                With rsRec
                    Do While Not .EOF
                        'cbo�����Ա�.AddItem Nvl(!����)
                        cbo����1.AddItem Nvl(!����)
                        .MoveNext
                    Loop
                End With
                For i = 0 To cbo����1.ListCount - 1
                    If cbo����1.List(i) = !���� Then
                        cbo����1.ListIndex = i
                        Exit For
                    End If
                Next
            Case "�ʿ��Լ���Դ"
                Me.txt����.Text = Nvl(!����)
                Me.txt����.Text = Nvl(!QC����)
            Case "����������"
                Me.txt����.Text = Nvl(!����)
                cbo����1.Clear
                gstrSql = "select distinct ���� from ����������"
                Set rsRec = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
                With rsRec
                    Do While Not .EOF
                        cbo����1.AddItem Nvl(!����)
                        .MoveNext
                    Loop
                End With
                    
                For i = 0 To cbo����1.ListCount - 1
                    If cbo����1.List(i) = !���� Then
                        cbo����1.ListIndex = i
                        Exit For
                    End If
                Next
               
            End Select
        End If
    End With
        
    zlRefresh = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlEditStart(blnAdd As Boolean, strItemName As String, str���� As String) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngItemId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim i As Integer
    Dim strTmp As String
    Dim rsTemp As New ADODB.Recordset, rsLength As New ADODB.Recordset
    
    frmBaseInfoList.sbType.Enabled = False
    Err = 0: On Error GoTo ErrHand
    If blnAdd Then
        gstrSql = "Select Nvl(Max(To_Number(����)), 0) As ����, Nvl(Max(Length(����)), 0) As ���� From " & strItemName
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ��һ������ĳ���")
        If rsTemp!���� <> 0 And rsTemp!���� <= Me.txt����.MaxLength Then
            'Me.txt����.MaxLength = rsTemp!����
            Me.txt����.Text = Format(Val(rsTemp!����) + 1, String(rsTemp!����, "0"))
        Else
            gstrSql = " select data_length as ColLength from user_tab_columns where table_name=[1] and column_Name=[2]"
            Set rsLength = zlDatabase.OpenSQLRecord(gstrSql, "���볤��", strItemName, "����")
            Me.txt����.MaxLength = rsLength!Collength
            Me.txt����.Text = Format(Val(rsTemp!����) + 1, String(rsLength!Collength, "0"))
        End If
        
        Me.txt����.Text = "": Me.txt����.Text = "": Me.txt����.Text = ""
        Me.txt˵��.Text = "": Me.chkȱʡ��־.Value = 0
    End If
    
    Select Case Trim(strItemName)
        Case "���Ƽ���걾"
            strTmp = Me.cbo�����Ա�.Text
            Me.cbo�����Ա�.Clear
            Me.cbo�����Ա�.AddItem (""): Me.cbo�����Ա�.AddItem ("��"): Me.cbo�����Ա�.AddItem ("Ů")
            For i = 0 To cbo�����Ա�.ListCount - 1
                If cbo�����Ա�.List(i) = strTmp Then
                    cbo�����Ա�.ListIndex = i
                    Exit For
                End If
            Next
        Case "���鱸ע����", "������������"
            strTmp = cbo����.Text
            cbo����.Clear
            gstrSql = "select distinct ���� from ���Ƽ�������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            
            With rsTemp
                Do While Not .EOF
                    cbo����.AddItem Nvl(!����)
                    .MoveNext
                Loop
            End With
            For i = 0 To cbo����.ListCount - 1
                If cbo����.List(i) = strTmp Then
                    cbo����.ListIndex = i
                    Exit For
                End If
            Next
        Case "����������"
            strTmp = cbo����1.Text
            cbo����1.Clear
            gstrSql = "select distinct ���� from ����������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            With rsTemp
                Do While Not .EOF
                    cbo����1.AddItem Nvl(!����)
                    .MoveNext
                Loop
            End With
            For i = 0 To cbo����1.ListCount - 1
                If cbo����1.List(i) = strTmp Then
                    cbo����1.ListIndex = i
                    Exit For
                End If
            Next
        
        Case "�ʿر���ʾ�"
            strTmp = cbo����1.Text
            cbo����1.Clear
            gstrSql = "Select Distinct ���� From �ʿر���ʾ�"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            
            With rsTemp
                Do While Not .EOF
                    cbo����1.AddItem Nvl(!����)
                    .MoveNext
                Loop
            End With
            For i = 0 To cbo����1.ListCount - 1
                If cbo����1.List(i) = strTmp Then
                    cbo����1.ListIndex = i
                    Exit For
                End If
            Next
    End Select
    
    Me.Tag = IIf(blnAdd, "����", "�޸�")
    Me.Enabled = True: Me.BackColor = RGB(250, 250, 250)
    
    Me.txt����.Enabled = True
    Me.txt����.SetFocus
    
    zlEditStart = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = ""
    Me.Enabled = False: Me.BackColor = &H8000000F
    fraEdit.BackColor = &H8000000F
    frmBaseInfoList.sbType.Enabled = True
    Call Me.zlRefresh(gstrItemName, mstr����OLD)
End Sub

Public Function zlEditSave(strItemName As String) As String
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long, strLists As String
    
    frmBaseInfoList.sbType.Enabled = True
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = "": Exit Function
    End If
    mstr����New = Trim(Me.txt����.Text)
    
    '��������Ƿ��ظ�
    If zlCodeRepeat(mstr����New, strItemName) Then
        txt����.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If strItemName = "�����������" Then
        If Trim(Me.txt˵��.Text) = "" Then
            MsgBox "���������ƣ�", vbInformation, gstrSysName
            Me.txt˵��.SetFocus: zlEditSave = "": Exit Function
        End If
    Else
        If Trim(Me.txt����.Text) = "" Then
            MsgBox "���������ƣ�", vbInformation, gstrSysName
            Me.txt����.SetFocus: zlEditSave = "": Exit Function
        End If
    End If

    '�ַ����Ϸ�����֤
    If zlCommFun.StrIsValid(Trim(txt����.Text), , , "����") = False Then
        txt����.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If zlCommFun.StrIsValid(Trim(IIf(strItemName = "�����������", txt˵��.Text, txt����.Text)), IIf(strItemName = "�����������", _
                txt˵��.MaxLength, txt����.MaxLength), IIf(strItemName = "�����������", txt˵��.hWnd, txt����.hWnd), "����") = False Then
        If strItemName = "�����������" Then
            Me.txt˵��.SetFocus
        Else
            Me.txt����.SetFocus
        End If
        zlEditSave = ""
        Exit Function
    End If
    
    If (strItemName = "���Ƽ���걾" Or strItemName = "���Ƽ�������" Or strItemName = "���鱸ע����" Or strItemName = "������������" _
            Or strItemName = "������������" Or strItemName = "����������" Or strItemName = "����ϸ������" Or strItemName = "����ϸ�����" Or strItemName = "����Ⱦɫ����" Or _
            strItemName = "�ʿر���ʾ�" Or strItemName = "�ʿؼ��鷽��" Or strItemName = "�ʿ��Լ���Դ" Or strItemName = "ϸ����ⷽ��" Or _
            strItemName = "����������" Or strItemName = "ϸ����ҩ����") Then
            
        If zlCommFun.StrIsValid(Trim(txt����.Text), txt����.MaxLength, txt����.hWnd, "����") = False Then Me.txt����.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If (strItemName = "���鱸ע����" Or strItemName = "����걾��̬" Or strItemName = "������������" Or strItemName = "������������") Then
        If zlCommFun.StrIsValid(Trim(txt˵��.Text), txt˵��.MaxLength, Me.txt˵��.hWnd, "˵��") = False Then Me.txt˵��.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If strItemName = "���Ƽ�������" Then
        If zlCommFun.StrIsValid(Trim(txt����.Text), txt����.MaxLength, txt����.hWnd, "����") = False Then Me.txt����.SetFocus: zlEditSave = "": Exit Function
    End If

    If strItemName = "�ʿر���ʾ�" Then
        If zlCommFun.StrIsValid(Trim(cbo����1.Text), 4, cbo����1.hWnd, "����") = False Then Me.cbo����1.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If strItemName = "�ʿ��Լ���Դ" Then
        If zlCommFun.StrIsValid(Trim(txt����.Text), txt����.MaxLength, txt����.hWnd, "QC����") = False Then Me.txt����.SetFocus: zlEditSave = "": Exit Function
    End If
    '���ݱ��������֯
    Select Case Trim(strItemName)
        Case "���Ƽ���걾"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "','" & Trim(txt����.Text) & "','" & Trim(cbo�����Ա�.Text) & "'"
        Case "���Ƽ�������"
            
            If zlCommFun.StrIsValid(Trim(txt����.Text), txt����.MaxLength, txt����.hWnd, "����") = False Then
                Me.txt����.SetFocus: zlEditSave = "": Exit Function
            End If
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "','" & Trim(txt����.Text) & "','" & _
                            chkȱʡ��־.Value & "','" & Trim(txt����.Text) & "'"
        Case "���鱸ע����", "������������"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "','" & Trim(txt����.Text) & "','" & _
                            Trim(txt˵��.Text) & "','" & Trim(cbo����.Text) & "'"
        Case "������������"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "','" & Trim(txt����.Text) & "','" & Trim(txt˵��.Text) & "'"
        Case "����걾��̬"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "','" & Trim(txt˵��.Text) & "'"
        Case "�����������"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(Me.txt˵��.Text) & "'"
        Case "���������;"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "'"
        Case "����������", "����ϸ�����", "����Ⱦɫ����"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & " ','" & Trim(txt����.Text) & "','" & _
            chkȱʡ��־.Value & "'"
        Case "����ϸ������", "�ʿؼ��鷽��", "ϸ����ⷽ��", "ϸ����ҩ����"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "','" & Trim(txt����.Text) & "'"
        Case "�ʿر���ʾ�"
            If Len(Trim(cbo����1.Text)) > 4 Then MsgBox "��ȷ���������Ƴ��Ȳ�����4λ��", vbInformation, gstrSysName: cbo����1.SetFocus: zlEditSave = "": Exit Function
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "','" & Trim(txt����.Text) & "','" & _
                    Trim(cbo����1.Text) & "'"
        Case "�ʿ��Լ���Դ"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "','" & Trim(txt����.Text) & "','" & _
                    Trim(txt����.Text) & "'"
        Case "����������"
            gstrSql = "'" & mstr����New & "','" & mstr����OLD & "','" & Trim(txt����.Text) & "','" & Trim(txt����.Text) & "','" & _
                    Trim(cbo����1.Text) & "'"
    End Select
    
    Err = 0: On Error GoTo ErrHand

    If Me.Tag = "����" Then
        If zlDatabase.OpenSQLRecord("select ���� from " & strItemName & " where ���� ='" & Trim(txt����.Text) & "'", Me.Caption).RecordCount > 0 Then
            MsgBox strItemName & "���Ƴ����ظ���", vbInformation, gstrSysName
            txt����.SetFocus: zlEditSave = "": Exit Function
        End If
        gstrSql = "zl_" & strItemName & "_Edit(1," & gstrSql & ")"
    Else
        gstrSql = "zl_" & strItemName & "_Edit(2," & gstrSql & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Me.Tag = ""
    Me.Enabled = False: Me.BackColor = &H8000000F
    fraEdit.BackColor = &H8000000F
    
    zlEditSave = mstr����New
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlCodeRepeat(strInputCode As String, strItemName As String) As Boolean
    '----------------------------------
    '���ܣ���������Ƿ������б����ظ����ظ��������ʾ
    '��Σ�strInputCode-����ı���
    '���Σ��ظ�����True��������Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    Err = 0: On Error GoTo ErrHand
    'strSQL = "select ����,���� from (select ����,���� from " & strItemName & " where ����<>[1]) where ����=[1]"
    strSql = "select ����,���� from " & strItemName & " where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�жϱ����Ƿ��ظ�", strInputCode)
        
    With rsTmp
        If .RecordCount <> 0 And mstr����OLD <> mstr����New Then
            MsgBox "����Ŀ�롾" & Nvl(!����) & "-" & Nvl(!����) & "�������ظ���", vbExclamation, gstrSysName
            zlCodeRepeat = True
        Else
            zlCodeRepeat = False
        End If
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlCodeRepeat = True
End Function

Private Sub cbo����1_KeyPress(KeyAscii As Integer)
    If Len(Trim(cbo����1.Text)) = 4 And KeyAscii <> 8 Then KeyAscii = 0: Exit Sub
End Sub

'
Private Sub cbo�����Ա�_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkȱʡ��־_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    mstr����OLD = ""
End Sub

Public Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    'Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
            Exit Sub
        End If
    End Select
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
            Exit Sub
        End If
    End Select
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_Change()
    txt����.Text = zlCommFun.SpellCode(txt����.Text)
End Sub

'
Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


