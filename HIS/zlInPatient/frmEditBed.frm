VERSION 5.00
Begin VB.Form frmEditBed 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmEditBed.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   390
      TabIndex        =   8
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3735
      TabIndex        =   7
      Top             =   2055
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2550
      TabIndex        =   6
      Top             =   2055
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   4830
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1000
         MaxLength       =   5
         TabIndex        =   0
         Top             =   280
         Width           =   1095
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   3345
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   645
         Width           =   1290
      End
      Begin VB.ComboBox cbo�ȼ� 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   3660
      End
      Begin VB.TextBox txt����� 
         Height          =   300
         Left            =   3345
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1290
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   645
         Width           =   1470
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1065
         Width           =   3660
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         Height          =   180
         Left            =   2550
         TabIndex        =   15
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lbl�ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ȼ�"
         Height          =   180
         Left            =   195
         TabIndex        =   14
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   2730
         TabIndex        =   13
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ա�"
         Height          =   180
         Left            =   195
         TabIndex        =   12
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   195
         TabIndex        =   11
         Top             =   1125
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   555
         TabIndex        =   10
         Top             =   300
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmEditBed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mblnModi As Boolean '����༭״̬(ȱʡ����)
Public mlngUnit As Long '��ǰ����ID
Public mlvwBeds As ListView
Public mobjSta As StatusBar
Public mblnChange As Boolean
Private mrs���� As New ADODB.Recordset

Private Sub cbo����_Click()
    Dim strTemp As String
    
    mblnChange = True

    strTemp = Split(cbo����.Text, "-")(0)
    
    mrs����.Filter = "����=" & strTemp
    
    If Not mrs����.EOF Then
        txt����.Text = mrs����!���� & ""
    End If

    If mblnModi = False Then
        txt����.Text = NextBedNo(mlngUnit, NeedName(cbo����.Text), mrs����!���� & "")
    End If
End Sub

Private Sub cbo�ȼ�_Click()
    mblnChange = True
End Sub

Private Sub cbo����_Click()
    mblnChange = True
End Sub

Private Sub cbo�Ա�_Click()
    mblnChange = True
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo�Ա�.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = MatchIndex(cbo�Ա�.hwnd, KeyAscii)
        If lngIdx <> -2 Then cbo�Ա�.ListIndex = lngIdx
    ElseIf cbo�Ա�.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = MatchIndex(cbo����.hwnd, KeyAscii)
        If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    ElseIf cbo����.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = MatchIndex(cbo����.hwnd, KeyAscii)
        If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    ElseIf cbo����.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo�ȼ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo�ȼ�.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = MatchIndex(cbo�ȼ�.hwnd, KeyAscii)
        If lngIdx <> -2 Then cbo�ȼ�.ListIndex = lngIdx
    ElseIf cbo�ȼ�.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String, strSQL As String
    Dim objItem As ListItem
    Dim str���� As String, lngDept As Long
    
    If mblnModi = False Then
        If Not IsNumeric(txt����.Text) Then
            MsgBox "���ű������룡", vbInformation, gstrSysName
            txt����.SetFocus: Exit Sub
        End If
    End If
    
    If InStr(txt�����.Text, "'") > 0 Then
        MsgBox "������а����Ƿ��ַ�,���飡", vbInformation, gstrSysName
        txt�����.SetFocus: Exit Sub
    End If
    
    If LenB(StrConv(txt�����.Text, vbFromUnicode)) > 10 Then
        MsgBox "����ŵĳ��Ȳ��ܴ���10��", vbInformation, gstrSysName
        txt�����.SetFocus: Exit Sub
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ���ò������ڿ��ң�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    If cbo�Ա�.ListIndex = -1 Then
        MsgBox "����ȷ���ò������Ա���࣡", vbInformation, gstrSysName
        cbo�Ա�.SetFocus: Exit Sub
    End If
    If cbo�ȼ�.ListIndex = -1 Then
        MsgBox "����ȷ���ò����ĵȼ���", vbInformation, gstrSysName
        cbo�ȼ�.SetFocus: Exit Sub
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "����ȷ���ò����ı������ͣ�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    
    mblnChange = False
    
    If mblnModi = False Then
        str���� = txt���� & txt����.Text
    Else
        str���� = txt����.Text
    End If
    lngDept = cbo����.ItemData(cbo����.ListIndex)

    If mblnModi Then
        strSQL = "zl_��λ״����¼_INSERT('" & Mid(mlvwBeds.SelectedItem.Key, 2) & "'," & mlngUnit & "," & _
            IIf(lngDept = 0, "NULL", lngDept) & "," & _
            "'" & txt�����.Text & "'," & _
            IIf(cbo�Ա�.ListIndex = -1, "NULL,", "'" & NeedName(cbo�Ա�.Text) & "',") & _
            IIf(cbo����.ListIndex = -1, "NULL,", "'" & NeedName(cbo����.Text) & "',") & _
            IIf(cbo�ȼ�.ListIndex = -1, "NULL", cbo�ȼ�.ItemData(cbo�ȼ�.ListIndex)) & ",0)"
        On Error GoTo errH
            zldatabase.ExecuteProcedure strSQL, Me.Caption
        On Error GoTo 0
        
        Set objItem = mlvwBeds.SelectedItem
        objItem.SubItems(mlvwBeds.ColumnHeaders("_����").Index - 1) = NeedName(cbo����.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_�����").Index - 1) = txt�����.Text
        objItem.SubItems(mlvwBeds.ColumnHeaders("_�Ա����").Index - 1) = NeedName(cbo�Ա�.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_�ȼ�").Index - 1) = NeedName(cbo�ȼ�.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_��λ����").Index - 1) = NeedName(cbo����.Text)
        If cbo�Ա�.ListIndex = 0 Then
            objItem.Icon = "M_Empty"
            objItem.SmallIcon = "M_Empty"
        ElseIf cbo�Ա�.ListIndex = 1 Then
            objItem.Icon = "F_Empty"
            objItem.SmallIcon = "F_Empty"
        Else
            objItem.Icon = "Empty"
            objItem.SmallIcon = "Empty"
        End If
        objItem.Tag = lngDept
        objItem.ListSubItems(1).Tag = ""
        If lngDept = 0 Then objItem.ListSubItems(1).Tag = 1 '���ò���
        
        Call SetBedIcon(mlvwBeds, objItem)
        
        objItem.EnsureVisible
        
        With objItem
            mobjSta.Panels(2) = "����[" & Trim(.Text) & "]" & _
                " ״̬:" & .SubItems(mlvwBeds.ColumnHeaders("_״̬").Index - 1) & _
                " �Ա����:" & .SubItems(mlvwBeds.ColumnHeaders("_�Ա����").Index - 1) & _
                " ����:" & .SubItems(mlvwBeds.ColumnHeaders("_����").Index - 1) & _
                " �ȼ�:" & .SubItems(mlvwBeds.ColumnHeaders("_�ȼ�").Index - 1)
        End With
        gblnOK = True
        Unload Me
    Else
        strTmp = isRepeat(mlngUnit, "'" & str���� & "'")
        If strTmp <> "" Then
            MsgBox "��ǰ����Ĵ����Ѿ����ڣ�", vbInformation, gstrSysName
            txt����.SetFocus: Exit Sub
        End If
        
        strSQL = "zl_��λ״����¼_INSERT('" & str���� & "'," & mlngUnit & "," & _
            IIf(lngDept = 0, "NULL", lngDept) & "," & _
            "'" & txt�����.Text & "'," & _
            IIf(cbo�Ա�.ListIndex = -1, "NULL,", "'" & NeedName(cbo�Ա�.Text) & "',") & _
            IIf(cbo����.ListIndex = -1, "NULL,", "'" & NeedName(cbo����.Text) & "',") & _
            IIf(cbo�ȼ�.ListIndex = -1, "NULL", cbo�ȼ�.ItemData(cbo�ȼ�.ListIndex)) & ",1)"
        
        On Error GoTo errH
        zldatabase.ExecuteProcedure strSQL, Me.Caption
        On Error GoTo 0
        
        If cbo�Ա�.ListIndex = 0 Then
            Set objItem = mlvwBeds.ListItems.Add(, "_" & str����, str����, "M_Empty", "M_Empty")
        ElseIf cbo�Ա�.ListIndex = 1 Then
            Set objItem = mlvwBeds.ListItems.Add(, "_" & str����, str����, "F_Empty", "F_Empty")
        Else
            Set objItem = mlvwBeds.ListItems.Add(, "_" & str����, str����, "Empty", "Empty")
        End If
        
        objItem.SubItems(mlvwBeds.ColumnHeaders("_����").Index - 1) = NeedName(cbo����.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_�����").Index - 1) = txt�����.Text
        objItem.SubItems(mlvwBeds.ColumnHeaders("_״̬").Index - 1) = "�մ�"
        objItem.SubItems(mlvwBeds.ColumnHeaders("_�Ա����").Index - 1) = NeedName(cbo�Ա�.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_�ȼ�").Index - 1) = NeedName(cbo�ȼ�.Text)
        objItem.SubItems(mlvwBeds.ColumnHeaders("_��λ����").Index - 1) = NeedName(cbo����.Text)
        objItem.Tag = lngDept
        If lngDept = 0 Then objItem.ListSubItems(1).Tag = 1 '���ò���
        
        Call SetBedIcon(mlvwBeds, objItem)
        
        objItem.Selected = True
        objItem.EnsureVisible
        With objItem
            mobjSta.Panels(2) = "����[" & Trim(.Text) & "]" & _
                " ״̬:" & .SubItems(mlvwBeds.ColumnHeaders("_״̬").Index - 1) & _
                " �Ա����:" & .SubItems(mlvwBeds.ColumnHeaders("_�Ա����").Index - 1) & _
                " ����:" & .SubItems(mlvwBeds.ColumnHeaders("_����").Index - 1) & _
                " �ȼ�:" & .SubItems(mlvwBeds.ColumnHeaders("_�ȼ�").Index - 1)
        End With
        
        Call frmManageBed.SetBedNOLen
        Call frmManageBed.SetMenuState
        
        txt����.Text = NextBedNo(mlngUnit, NeedName(cbo����.Text), txt����.Text)
        
        gblnOK = True
        
        txt����.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    mblnChange = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    Dim str���� As String
    
    gblnOK = False
    
    If Not InitData Then Unload Me: Exit Sub
    
    If mblnModi Then
        txt����.Enabled = False
        
        With mlvwBeds.SelectedItem

            
            cbo����.ListIndex = FindCboIndex(cbo����, Val(.Tag))
            If cbo����.ListIndex = -1 Then
                If .SubItems(mlvwBeds.ColumnHeaders("_����").Index - 1) <> "" Then
                    cbo����.ListIndex = 0
                End If
            End If
            
            cbo�Ա�.ListIndex = GetCboIndex(cbo�Ա�, .SubItems(mlvwBeds.ColumnHeaders("_�Ա����").Index - 1))
            cbo�ȼ�.ListIndex = GetCboIndex(cbo�ȼ�, .SubItems(mlvwBeds.ColumnHeaders("_�ȼ�").Index - 1))
            cbo����.ListIndex = GetCboIndex(cbo����, .SubItems(mlvwBeds.ColumnHeaders("_��λ����").Index - 1))
            txt����.MaxLength = 10
            txt����.Text = Mid(.Key, 2)
            txt����.Width = TextWidth(txt����.Text)
            txt�����.Text = .SubItems(mlvwBeds.ColumnHeaders("_�����").Index - 1)
            txt����.Text = ""
            
            '��ΪӰ�촲λ������¼,��ֹ����
            cbo����.Enabled = False
        End With
        Me.Caption = "��������"
    Else
        Me.Caption = "��������"
        txt����.MaxLength = 5
        If cbo����.Text <> "" Then str���� = Split(cbo����.Text, "-")(1)
        txt����.Text = NextBedNo(mlngUnit, str����, txt����.Text)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnModi And mblnChange And Visible Then
        If MsgBox("���޸��˵�������δ����,ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    mblnModi = False
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strTmp As String
    
    '�Ա����
    cbo�Ա�.AddItem "1-�д�"
    cbo�Ա�.AddItem "2-Ů��"
    cbo�Ա�.AddItem "3-���޴�"
    If Not mblnModi Then cbo�Ա�.ListIndex = 2
    
    'ȷ�������ķ������
    strSQL = "Select ������� From ��������˵�� Where ��������='����' And ����ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUnit)
    
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
        If Not mblnModi And cbo����.ListIndex = -1 Then cbo����.ListIndex = 1
    Else
        MsgBox "δ��ʼ���ٴ����һ�û�����ò������Ҷ�Ӧ��Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��λ�ȼ�
    strSQL = "Select ID as ���,����,���� From �շ���ĿĿ¼ Where ���='J' And (����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL) Order by ����"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo�ȼ�.Clear
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo�ȼ�.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo�ȼ�.ItemData(i - 1) = rsTmp!���
            rsTmp.MoveNext
        Next
        If Not mblnModi Then cbo�ȼ�.ListIndex = 0
    Else
        MsgBox "û�г�ʼ����λ�ȼ���Ϣ,���ȵ���λ�ȼ������д���", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��λ����
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ,���� From  ��λ���Ʒ��� Order by ����"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
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
    txt����.Width = txt����.Left + txt����.Width - txt����.Left - 60
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
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
    SelAll txt�����
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub SetBedIcon(objLvw As Object, objItem As ListItem)
    If objItem.SubItems(objLvw.ColumnHeaders("_��λ����").Index - 1) = "�Ӵ�" Then
        objItem.Icon = "�Ӵ�_" & objItem.Icon
        objItem.SmallIcon = "�Ӵ�_" & objItem.SmallIcon
    ElseIf objItem.SubItems(objLvw.ColumnHeaders("_��λ����").Index - 1) = "�Ǳ�" Then
        objItem.Icon = "�Ǳ�_" & objItem.Icon
        objItem.SmallIcon = "�Ǳ�_" & objItem.SmallIcon
    End If
    
    If Val(objItem.ListSubItems(1).Tag) <> 0 Then
        objItem.Icon = "����_" & objItem.Icon
        objItem.SmallIcon = "����_" & objItem.SmallIcon
    End If
End Sub

