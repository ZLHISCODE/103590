VERSION 5.00
Begin VB.Form frmModiPatiBaseInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˻�����Ϣ����"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4380
   Icon            =   "frmModiPatiBaseInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboAge 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3135
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1215
      Width           =   705
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   2  'OFF
      Left            =   2115
      TabIndex        =   5
      Top             =   1215
      Width           =   1020
   End
   Begin VB.ComboBox cboSex 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmModiPatiBaseInfo.frx":030A
      Left            =   2115
      List            =   "frmModiPatiBaseInfo.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   690
      Width           =   1725
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2115
      MaxLength       =   64
      TabIndex        =   1
      Top             =   210
      Width           =   1725
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1215
      TabIndex        =   7
      Top             =   1995
      Width           =   1450
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2685
      TabIndex        =   8
      Top             =   1995
      Width           =   1450
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   30
      TabIndex        =   9
      Top             =   1710
      Width           =   5100
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Left            =   495
      Picture         =   "frmModiPatiBaseInfo.frx":030E
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
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
      Left            =   1500
      TabIndex        =   4
      Top             =   1275
      Width           =   480
   End
   Begin VB.Label lblSex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
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
      Left            =   1485
      TabIndex        =   2
      Top             =   750
      Width           =   480
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
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
      Left            =   1530
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmModiPatiBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlng����ID As Long
Private mstr����ID As String
Private mstrģ�� As String
Private mint���� As Integer

Public Function ShowMe(ByVal lng����ID As Long, ByVal str����ID As String, ByVal strģ�� As String, ByVal int���� As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:lng����ID-����ID
    '     str����ID=�����޸�Ϊ�գ����ﲡ��Ϊ�Һ�ID��סԺ����Ϊ��ҳID���������˸���ҵ�����о������磺ҽ��ID����첡��Ϊ���񵥺�
    '     strģ��=���øù��ܵ�ģ����������"����Һ�"��"��鱨��"��
    '     int����=0-����,1-����,2-סԺ,3-��������,4-��첡��
    '����:
    '����:
    '����:������
    '����:2013-10-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng����ID = lng����ID
    mstr����ID = str����ID
    mstrģ�� = strģ��
    mint���� = int����
    
    mblnOK = False
    '��ȡ���˻�����Ϣ
    If Not LoadPatiBaseInfo Then ShowMe = False: Exit Function
    
    Me.Show 1
    ShowMe = mblnOK
End Function

Private Sub InitDicts()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    txtName.Text = ""
    txtName.MaxLength = GetColumnLength("������Ϣ", "����")
    txtAge.Text = ""
    cboAge.Clear
    cboAge.AddItem "��"
    cboAge.AddItem "��"
    cboAge.AddItem "��"
    cboAge.ListIndex = 0
    txtAge.MaxLength = GetColumnLength("������Ϣ", "����")
    
    cboSex.Clear
    
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Order by ����"
    Call gobjComLib.zlDatabase.OpenRecordset(rsTmp, strSQL, "�Ա�")
    Do While Not rsTmp.EOF
        cboSex.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ȱʡ = 1 Then
            cboSex.ListIndex = cboSex.NewIndex
            cboSex.ItemData(cboSex.NewIndex) = 1
        End If
    rsTmp.MoveNext
    Loop
    
    Exit Sub
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Function LoadPatiBaseInfo() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngIndex As Long
    
    On Error GoTo errHand
    
    If mint���� = 1 Then '���ﲡ��
        strSQL = "Select ����,�Ա�,���� from ���˹Һż�¼ where ����ID=[1] And ID=[2]"
    ElseIf mint���� = 2 Then 'סԺ����
        strSQL = " Select Nvl(a.����, b.����) ����, Nvl(a.�Ա�, b.�Ա�) �Ա�, a.����" & vbNewLine & _
                " From ������ҳ a, ������Ϣ b" & vbNewLine & _
                " Where a.����id = b.����id And a.����id = [1] And a.��ҳid = [2]"
    Else
        strSQL = "Select ����,�Ա�,���� From ������Ϣ Where ����ID=[1]"
    End If
    
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˻�����Ϣ", mlng����ID, Val(mstr����ID))
    
    If Not rsTmp.EOF Then
        If mint���� = 0 Then
            Me.Caption = "�������λ�����Ϣ����"
        Else
            Me.Caption = "���˻�����Ϣ����"
        End If
        '������Ϣ��ʼ��
        Call InitDicts
        
        txtName.Text = gobjComLib.zlCommFun.NVL(rsTmp!����)
        txtName.Tag = txtName.Text
        cboSex.Tag = gobjComLib.zlCommFun.NVL(rsTmp!�Ա�)
        lngIndex = GetCboIndex(cboSex, gobjComLib.zlCommFun.NVL(rsTmp!�Ա�))
        If lngIndex <> -1 Then cboSex.ListIndex = lngIndex
        Call LoadOldData("" & rsTmp!����, txtAge, cboAge)
        txtAge.Tag = gobjComLib.zlCommFun.NVL(rsTmp!����)
    Else
        MsgBox "��ȡ���˻�����Ϣʧ��,����ȷ��Ҫ������Ϣ�����Ĳ��ˣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    LoadPatiBaseInfo = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    gobjComLib.SaveErrLog
End Function

Private Sub LoadOldData(ByVal strOld As String, ByRef txtAge As TextBox, ByRef cboAge As ComboBox)
'����:�����ݿ��б�������䰴�淶�ĸ�ʽ���ص�����,���淶��ԭ����ʾ
    Dim strTmp As String, lngIdx As Long
    
    If Trim(strOld) = "" Then Exit Sub
    
    lngIdx = -1
    strTmp = strOld
    If InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 0
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 1
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 2
        End If
    ElseIf IsNumeric(strOld) Then
        lngIdx = 0
    End If
    txtAge.Text = strTmp
    If cboAge.ListCount > 0 Then Call gobjComLib.zlControl.CboSetIndex(cboAge.hwnd, lngIdx)
    If lngIdx = -1 Then
        cboAge.Visible = False
    Else
        If cboAge.Visible = False Then cboAge.Visible = True
    End If
End Sub

Private Function CheckOldData(ByRef txtAge As TextBox, ByRef cboAge As ComboBox) As Boolean
'���ܣ������������ֵ����Ч��
'���أ�
    If Not IsNumeric(txtAge.Text) Then CheckOldData = True: Exit Function
    
    Select Case cboAge.Text
        Case "��"
            If Val(txtAge.Text) > 200 Then
                MsgBox "���䲻�ܴ���200��!", vbInformation, gstrSysName
                If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txtAge.Text) > 2400 Then
                MsgBox "���䲻�ܴ���2400��!", vbInformation, gstrSysName
                If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "��"
            If Val(txtAge.Text) > 73000 Then
                MsgBox "���䲻�ܴ���73000��!", vbInformation, gstrSysName
                If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function

Private Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If gobjComLib.zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Private Function CheckTextLength(strName As String, txtObj As TextBox) As Boolean
'����:��鲢��ʾ�ı������볤���Ƿ���

    CheckTextLength = True
    If gobjComLib.zlCommFun.ActualLen(txtObj.Text) > txtObj.MaxLength Then
        MsgBox strName & "���������ֻ�������� " & txtObj.MaxLength & " ���ַ��� " & txtObj.MaxLength \ 2 & " �����֡�", vbInformation, gstrSysName
        If txtObj.Enabled And txtObj.Visible Then txtObj.SetFocus
        CheckTextLength = False
    End If
End Function

Private Function GetColumnLength(strTable As String, strColumn As String) As Long
'���ܣ���ȡָ������ָ���ֶεĳ���
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(Data_Precision, Data_Length) collen From All_Tab_Columns Where Table_Name = [1] And Column_Name = [2]"
    On Error GoTo errH
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strTable, strColumn)
    GetColumnLength = Val("" & rsTmp!collen)
    
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function GetCboIndex(cbo As ComboBox, strFind As String, _
    Optional blnKeep As Boolean, _
    Optional blnLike As Boolean, Optional strSplit As String = "-") As Long
'���ܣ����ַ�����ComboBox�в�������
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '�Ⱦ�ȷ����
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), strSplit) > 0 Then
            If NeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '���ģ������
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Private Function NeedName(strList As String) As String
    If InStr(strList, Chr(&HA)) > 0 Then
        NeedName = Trim(Mid(strList, InStr(strList, Chr(&HA)) + 1))
    Else
        NeedName = Trim(Mid(strList, InStr(strList, "-") + 1))
    End If
    If InStr(NeedName, Chr(&HD)) > 0 Then
        NeedName = Replace(NeedName, Chr(&HD), "")
    End If
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
'���ܣ��������У��ͱ���
    Dim strSQL As String
    Dim str���� As String
    
    '��һ�������ݺϷ���У��
    If Trim(txtName.Text) = "" Then
        MsgBox "�������벡��������", vbInformation, gstrSysName
        txtName.SetFocus: Exit Sub
    End If
    If cboSex.ListIndex = -1 Then
        MsgBox "����ȷ�������Ա�", vbInformation, gstrSysName
        cboSex.SetFocus: Exit Sub
    End If
    If Trim(txtAge.Text) = "" Then
        MsgBox "�������벡�����䣡", vbInformation, gstrSysName
        txtAge.SetFocus: Exit Sub
    End If
    
    
    If Not CheckTextLength("����", txtName) Then Exit Sub
    If Not CheckTextLength("����", txtAge) Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    str���� = Trim(txtAge.Text)
    If IsNumeric(str����) Then str���� = str���� & cboAge.Text
    
    '�ڶ��������ݱ���
    On Error GoTo errHand
    strSQL = "Zl_������Ϣ_������Ϣ����("
'   ����id_In ������Ϣ�䶯.����id%Type,
    strSQL = strSQL & "" & mlng����ID & ","
'   ����id_In Number := Null,
    strSQL = strSQL & "'" & mstr����ID & "',"
'   ģ��_In   ������Ϣ�䶯.�䶯ģ��%Type,
    strSQL = strSQL & "'" & mstrģ�� & "',"
'   ����_In   ������Ϣ.����%Type,
    strSQL = strSQL & "'" & Trim(txtName.Text) & "',"
'   �Ա�_In   ������Ϣ.�Ա�%Type,
    strSQL = strSQL & "'" & Split(cboSex.Text, "-")(1) & "',"
'   ����_In   ������Ϣ.����%Type
    strSQL = strSQL & "'" & str���� & "',"
'   ����_In   number(1)
    strSQL = strSQL & "" & mint���� & ")"
    
    Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "Zl_������Ϣ_������Ϣ����")
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    gobjComLib.SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        If ActiveControl.Name <> txtName.Name And ActiveControl.Name <> txtAge.Name Then
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtAge_GotFocus()
    Call gobjComLib.zlCommFun.OpenIme
    gobjComLib.zlControl.TxtSelAll txtAge
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboAge.Visible = False And IsNumeric(txtAge.Text) Then
            Call txtAge_Validate(False)
            Call cboAge.SetFocus
        Else
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txtAge.Text) Then Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAge_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAge.Text) And Trim(txtAge.Text) <> "" Then
        cboAge.ListIndex = -1: cboAge.Visible = False
    ElseIf cboAge.Visible = False Then
        cboAge.ListIndex = 0: cboAge.Visible = True
    End If
End Sub

Private Sub txtName_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtName
    Call gobjComLib.zlCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            Call CheckInputLen(txtName, KeyAscii)
        End If
    Else
        If Trim(txtName.Text) = "" Then
            Exit Sub
        Else
            gobjComLib.zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Sub txtName_LostFocus()
    Call gobjComLib.zlCommFun.OpenIme
    txtName.Text = Trim(txtName.Text)
End Sub
