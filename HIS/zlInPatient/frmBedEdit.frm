VERSION 5.00
Begin VB.Form frmBedEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraEdit 
      Height          =   4185
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4830
      Begin VB.ComboBox cbo�ȼ� 
         Height          =   315
         Left            =   975
         TabIndex        =   7
         Text            =   "cbo�ȼ�"
         Top             =   3240
         Width           =   3660
      End
      Begin VB.TextBox txt˳��� 
         Height          =   300
         Left            =   960
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cbo���� 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2730
         Width           =   3660
      End
      Begin VB.CheckBox chkContAdd 
         Caption         =   "��������"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1000
         MaxLength       =   10
         TabIndex        =   0
         Top             =   280
         Width           =   1095
      End
      Begin VB.ComboBox cbo���� 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2220
         Width           =   1455
      End
      Begin VB.TextBox txt����� 
         Height          =   300
         Left            =   975
         MaxLength       =   10
         TabIndex        =   2
         Top             =   735
         Width           =   1455
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1710
         Width           =   1455
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   975
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CboLevel 
         Height          =   315
         Left            =   1905
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3255
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label lbl˳��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˳���"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   1275
         Width           =   540
      End
      Begin VB.Label lbl�ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ȼ�"
         Height          =   180
         Left            =   195
         TabIndex        =   13
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         Height          =   180
         Left            =   195
         TabIndex        =   14
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   375
         TabIndex        =   12
         Top             =   805
         Width           =   540
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ա�"
         Height          =   180
         Left            =   195
         TabIndex        =   11
         Top             =   1785
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   195
         TabIndex        =   10
         Top             =   2805
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   555
         TabIndex        =   9
         Top             =   300
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmBedEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mblnAdd As Boolean '����༭״̬
Public mlngUnit As Long '��ǰ����ID
Public mobjSta As StatusBar
Public mblnChange As Boolean
Public mintCancle As Integer

Private mrs���� As New ADODB.Recordset
Private mrsBedLevel As New ADODB.Recordset
Private mrptRecord As ReportRecord

Private Sub cbo����_Click()
    Dim strTemp As String
    If cbo����.Text = "" Then Exit Sub
    
    If cbo����.ListIndex <> Val(cbo����.Tag) Then
        cbo����.Tag = cbo����.ListIndex
        mblnChange = True
    End If

    strTemp = Split(cbo����.Text, "-")(0)
    
    mrs����.Filter = "����='" & strTemp & "'"
    
    If Not mrs����.EOF Then
        txt����.Text = mrs����!���� & ""
    End If

    If mblnAdd = True Then
        txt����.Text = NextBedNo(mlngUnit, zlCommFun.GetNeedName(cbo����.Text), mrs����!���� & "")
    End If
End Sub

Private Sub cbo�ȼ�_Click()
    If cbo�ȼ�.ListIndex <> Val(cbo�ȼ�.Tag) Then
        cbo�ȼ�.Tag = cbo�ȼ�.ListIndex
        mblnChange = True
    End If
End Sub

Private Sub cbo�ȼ�_GotFocus()
    zlControl.TxtSelAll cbo�ȼ�
End Sub

Private Sub cbo�ȼ�_Validate(Cancel As Boolean)
    If isCheckBedLevelExists(cbo�ȼ�.Text, True, False) = False Then
        cbo�ȼ�.Text = ""
        cbo�ȼ�.ListIndex = -1
    End If
End Sub

Private Sub cbo����_Click()
    If cbo����.ListIndex <> Val(cbo����.Tag) Then
        cbo����.Tag = cbo����.ListIndex
        mblnChange = True
    End If
End Sub

Private Sub cbo�Ա�_Click()
    If cbo�Ա�.ListIndex <> Val(cbo�Ա�.Tag) Then
        cbo�Ա�.Tag = cbo�Ա�.ListIndex
        mblnChange = True
    End If
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo�Ա�.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo�Ա�.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo�Ա�.ListIndex = lngIdx
    ElseIf cbo�Ա�.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    ElseIf cbo����.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    ElseIf cbo����.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo�ȼ�_KeyPress(KeyAscii As Integer)
    '69273:������,2014-01-03,���ٶ�λ��λ�ȼ�
    Dim lngIdx As Long
    Dim i As Long, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim rsTemp As ADODB.Recordset
    
    If KeyAscii <> 13 Then
'        If SendMessage(cbo�ȼ�.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
'        lngIdx = MatchIndex(cbo�ȼ�.hWnd, KeyAscii)
'        If lngIdx <> -2 Then cbo�ȼ�.ListIndex = lngIdx
    Else
        If cbo�ȼ�.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cbo�ȼ�.Text)
        If cbo�ȼ�.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo�ȼ�.List(cbo�ȼ�.ListIndex) Then Call cbo.SetIndex(cbo�ȼ�.hWnd, -1)
        End If
        If strText = "" Then
            cbo�ȼ�.ListIndex = -1
        ElseIf cbo�ȼ�.ListIndex = -1 Then
            strFilter = ""
            '�ȸ��Ƽ�¼��
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrsBedLevel)
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrsBedLevel.Filter = strFilter: iCount = 0
            With mrsBedLevel
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrsBedLevel.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '������������,��Ҫ���:
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                        If Nvl(!����) = strText Then strResult = Nvl(!����): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(Nvl(!����)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!����)
                            iCount = iCount + 1
                        End If
                        
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(Nvl(!����)) Like strText & "*" Then
                            If isCheckBedLevelExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(Nvl(!����)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(Nvl(!����)) Like strCompents Then
                            If isCheckBedLevelExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!����) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!����) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                            If isCheckBedLevelExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                        End If
                    End Select
                    mrsBedLevel.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
            'ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheckBedLevelExists(strResult, True) Then cbo�ȼ�.SetFocus:  zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                rsTemp.Sort = "����,����"
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1130, cbo�ȼ�, rsTemp, True, "", "", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheckBedLevelExists(Nvl(rsReturn!����), True) Then
                                cbo�ȼ�.SetFocus
                                zlCommFun.PressKey vbKeyTab
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    cbo�ȼ�.SetFocus
                    Exit Sub
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: cbo�ȼ�.ListIndex = -1: zlControl.TxtSelAll cbo�ȼ�: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
        End If
        
        If cbo�ȼ�.ListIndex = -1 Then
            cbo�ȼ�.Text = ""
            Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Function isCheckBedLevelExists(ByVal str���� As String, Optional blnLocateItem As Boolean = False, Optional ByVal blnLevel As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڴ�λ�ȼ������б���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If blnLevel = True Then
        For i = 0 To CboLevel.ListCount - 1
            If CboLevel.List(i) = str���� Then
                If blnLocateItem Then cbo�ȼ�.ListIndex = i
                isCheckBedLevelExists = True
                Exit Function
            End If
        Next
    Else
        For i = 0 To cbo�ȼ�.ListCount - 1
            If cbo�ȼ�.List(i) = str���� Then
                If blnLocateItem Then cbo�ȼ�.ListIndex = i
                isCheckBedLevelExists = True
                Exit Function
            End If
        Next
    End If
End Function

Private Sub Form_Activate()
    On Error Resume Next
    Me.SetFocus
    If Err <> 0 Then Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnChange = False
    mintCancle = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mintCancle = Cancel
    If mblnAdd = False And mblnChange And Visible Then
        If MsgBox("���޸��˵�������δ����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            mintCancle = 1: Cancel = 1: Exit Sub
        End If
    End If
    mblnAdd = False
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strTmp As String
    
    '�Ա����
    cbo�Ա�.Clear
    cbo�Ա�.AddItem "1-�д�"
    cbo�Ա�.AddItem "2-Ů��"
    cbo�Ա�.AddItem "3-���޴�"
    If mblnAdd Then cbo�Ա�.ListIndex = 2
    
    'ȷ�������ķ������
    strSQL = "Select ������� From ��������˵�� Where ��������='����' And ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUnit)
    
    cbo����.Clear
    
    If rsTmp!������� = 1 Then
        '����۲������ö�Ӧ�������ٴ�����
        strTmp = "1,3"
    ElseIf rsTmp!������� = 2 Then
        strTmp = "2,3"
    ElseIf rsTmp!������� = 3 Then
        strTmp = "1,2,3"
    End If
    Set rsTmp = GetDeptOrUnit(0, mlngUnit, strTmp)
    
    If Not rsTmp.EOF Then
        cbo����.AddItem "<���ò���>" '���ò���
        For i = 1 To rsTmp.RecordCount
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Next
        If mblnAdd And cbo����.ListIndex = -1 Then cbo����.ListIndex = 1
    Else
        MsgBox "δ��ʼ���ٴ����һ�û�����ò������Ҷ�Ӧ��Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '69273:������,2014-01-03,�ṩ��λ�ǼǵĿ��ٲ���
    '��λ�ȼ�
    strSQL = "Select ID,����,����,zlspellcode(����,20) ���� From �շ���ĿĿ¼ Where ���='J' And (����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL) Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo�ȼ�.Clear
    CboLevel.Clear: CboLevel.Visible = False
    Set mrsBedLevel = rsTmp.Clone
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo�ȼ�.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo�ȼ�.ItemData(i - 1) = rsTmp!ID
            CboLevel.AddItem rsTmp!����
            CboLevel.ItemData(i - 1) = rsTmp!ID
            rsTmp.MoveNext
        Next
        If mblnAdd Then cbo�ȼ�.ListIndex = 0
    Else
        MsgBox "û�г�ʼ����λ�ȼ���Ϣ,���ȵ���λ�ȼ������д���", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��λ����
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ,���� From  ��λ���Ʒ��� Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo����.Clear
    Set mrs���� = rsTmp.Clone
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then cbo����.ListIndex = cbo����.NewIndex
            rsTmp.MoveNext
        Next
    Else
        MsgBox "û�г�ʼ����λ������Ϣ,�뵽�ֵ�����г�ʼ����λ���Ʒ��࣡", vbInformation, gstrSysName
        Exit Function
    End If
    
    InitData = True
End Function

Private Sub txt����_Change()

    txt����.Left = txt����.Left + TextWidth(txt����.Text) + 60
    txt����.width = txt����.Left + txt����.width - txt����.Left - 60
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text = "" Then
            Call Beep: Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub txt�����_Change()
    mblnChange = True
End Sub

Private Sub txt�����_GotFocus()
    zlControl.TxtSelAll txt�����
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Public Function zlEditStart(ByVal blnAdd As Boolean, ByVal lngUnitID As Long, Optional ByVal rptRecord As ReportRecord) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngItemId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    
    Dim i As Integer
    Dim strTmp As String
    Dim rsTemp As New ADODB.Recordset, rsLength As New ADODB.Recordset
    Dim str���� As String
    
    gblnOK = False
    
    mblnAdd = blnAdd
    mlngUnit = lngUnitID
    Set mrptRecord = rptRecord
    
    cbo����.Tag = "-1"
    cbo�Ա�.Tag = "-1"
    cbo�ȼ�.Tag = "-1"
    cbo����.Tag = "-1"
    
    If chkContAdd.Value Then
        txt����.Text = NextBedNo(mlngUnit, zlCommFun.GetNeedName(cbo����.Text), txt����.Text)
    Else
        If Not InitData Then Exit Function
        
        chkContAdd.Visible = blnAdd
        chkContAdd.Value = IIf(blnAdd, 1, 0)
        If blnAdd Then
            Me.Caption = "��������"
            txt����.MaxLength = 10
            If cbo����.Text <> "" Then str���� = Split(cbo����.Text, "-")(1)
            txt����.Text = NextBedNo(mlngUnit, str����, txt����.Text)
        Else
            txt����.Enabled = False
            
            With mrptRecord
    
                
                cbo����.ListIndex = cbo.FindIndex(cbo����, Val(.Item(mCol.����ID).Value))
                If cbo����.ListIndex = -1 Then
                    If .Item(mCol.����ID).Value <> "" Then
                        cbo����.ListIndex = 0
                    End If
                End If
                
                cbo�Ա�.ListIndex = cbo.FindIndex(cbo�Ա�, .Item(mCol.�Ա����).Value, True)
                cbo�ȼ�.ListIndex = cbo.FindIndex(cbo�ȼ�, .Item(mCol.�ȼ�).Value, True)
                If cbo�ȼ�.ListIndex = -1 Then isCheckBedLevelExists .Item(mCol.�ȼ�).Value, True
                cbo����.ListIndex = cbo.FindIndex(cbo����, .Item(mCol.��λ����).Value, True)
                txt����.MaxLength = 10
                txt����.Text = .Item(mCol.����).Value
                txt����.width = TextWidth(txt����.Text)
                txt�����.Text = .Item(mCol.�����).Value
                txt˳���.Text = .Item(mCol.˳���).Value
                txt����.Text = ""
                
                
                '��ΪӰ�촲λ������¼,��ֹ����
                cbo����.Enabled = False
            End With
            Me.Caption = "��������"
        End If
    End If
    
    mblnChange = False
    zlEditStart = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Dim objControl As Control
    For Each objControl In Me.Controls
        If objControl.Enabled = False Then objControl.Enabled = True
    Next
    Me.chkContAdd.Value = 0
    mblnChange = False
End Sub

Public Function zlEditSave() As String
    '���ܣ��������ڽ��еı༭,���������ڱ༭��λ��,����ʧ�ܷ��ؿ�
    Dim strTmp As String, strSQL As String
    Dim objItem As ListItem
    Dim str���� As String, lngDept As Long

    If mblnAdd = True Then
        If Not IsNumeric(txt����.Text) Then
            MsgBox "���ű������룡", vbInformation, gstrSysName
            txt����.SetFocus: Exit Function
        End If
    End If

    If InStr(txt�����.Text, "'") > 0 Then
        MsgBox "������а����Ƿ��ַ�,���飡", vbInformation, gstrSysName
        txt�����.SetFocus: Exit Function
    End If

    If LenB(StrConv(txt�����.Text, vbFromUnicode)) > 10 Then
        MsgBox "����ŵĳ��Ȳ��ܴ���10��", vbInformation, gstrSysName
        txt�����.SetFocus: Exit Function
    End If
    
    If InStr(Trim(txt˳���.Text), ".") <> 0 Then
        If LenB(StrConv(txt˳���.Text, vbFromUnicode)) > 10 Then
            MsgBox "˳��ŵĳ��Ȱ���С�������ڲ��ܴ���10λ��", vbInformation, gstrSysName
            txt˳���.SetFocus: Exit Function
        End If
    Else
        If Len(Trim(txt˳���.Text)) > 9 Then
            MsgBox "˳��ŵĳ��ȳ�ȥС���㲻�ܴ���9λ��", vbInformation, gstrSysName
            txt˳���.SetFocus: Exit Function
        End If
    End If
    
    If InStr(Trim(txt˳���.Text), ".") <> 0 Then
        If Len(Mid(Trim(txt˳���.Text), InStr(Trim(txt˳���.Text), ".") + 1)) > 1 Then
            MsgBox "˳���ֻ����һλС����", vbInformation, gstrSysName
            txt˳���.SetFocus: Exit Function
        End If
    End If
    
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ���ò������ڿ��ң�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Function
    End If
    If cbo�Ա�.ListIndex = -1 Then
        MsgBox "����ȷ���ò������Ա���࣡", vbInformation, gstrSysName
        cbo�Ա�.SetFocus: Exit Function
    End If
    If cbo�ȼ�.ListIndex = -1 Then
        MsgBox "����ȷ���ò����ĵȼ���", vbInformation, gstrSysName
        cbo�ȼ�.SetFocus: Exit Function
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ���ò����ı������ͣ�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Function
    End If

    mblnChange = False

    If mblnAdd = True Then
        str���� = txt���� & txt����.Text
    Else
        str���� = txt����.Text
    End If
    lngDept = cbo����.ItemData(cbo����.ListIndex)

    If mblnAdd Then
        str���� = Trim(str����)
        strTmp = isRepeat(mlngUnit, "'" & str���� & "'")
        If strTmp <> "" Then
            MsgBox "��ǰ����Ĵ����Ѿ����ڣ�", vbInformation, gstrSysName
            txt����.SetFocus: Exit Function
        End If

        gstrSQL = "zl_��λ״����¼_INSERT('" & Trim(str����) & "'," & mlngUnit & "," & _
            IIf(lngDept = 0, "NULL", lngDept) & "," & _
            "'" & txt�����.Text & "'," & _
            IIf(cbo�Ա�.ListIndex = -1, "NULL,", "'" & zlCommFun.GetNeedName(cbo�Ա�.Text) & "',") & _
            IIf(cbo����.ListIndex = -1, "NULL,", "'" & zlCommFun.GetNeedName(cbo����.Text) & "',") & _
            IIf(cbo�ȼ�.ListIndex = -1, "NULL", cbo�ȼ�.ItemData(cbo�ȼ�.ListIndex)) & ",1" & ",'" & txt˳���.Text & "')"
        
    Else
        str���� = Trim(mrptRecord.Item(mCol.����).Value)
        gstrSQL = "zl_��λ״����¼_INSERT('" & Trim(str����) & "'," & mlngUnit & "," & _
             IIf(lngDept = 0, "NULL", lngDept) & "," & _
             "'" & txt�����.Text & "'," & _
             IIf(cbo�Ա�.ListIndex = -1, "NULL,", "'" & zlCommFun.GetNeedName(cbo�Ա�.Text) & "',") & _
             IIf(cbo����.ListIndex = -1, "NULL,", "'" & zlCommFun.GetNeedName(cbo����.Text) & "',") & _
             IIf(cbo�ȼ�.ListIndex = -1, "NULL", cbo�ȼ�.ItemData(cbo�ȼ�.ListIndex)) & ",0" & ",'" & txt˳���.Text & "')"
    End If
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    zlEditSave = str����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt˳���_Change()
    mblnChange = True
End Sub

Private Sub txt˳���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
End Sub

