VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "frmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame frame�������� 
      Caption         =   "��������"
      Height          =   2745
      Left            =   150
      TabIndex        =   8
      Top             =   1920
      Width           =   4515
      Begin VB.TextBox txt����Ŀ¼ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   27
         Tag             =   "6"
         Top             =   1890
         Width           =   2235
      End
      Begin VB.CommandButton cmd����Ŀ¼ 
         Caption         =   "��"
         Height          =   300
         Left            =   3945
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1890
         Width           =   285
      End
      Begin VB.CommandButton cmdҽ����ĿĿ¼ 
         Caption         =   "��"
         Height          =   300
         Left            =   3945
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2280
         Width           =   285
      End
      Begin VB.TextBox txtҽ����ĿĿ¼ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   22
         Tag             =   "6"
         Top             =   2280
         Width           =   2235
      End
      Begin VB.CommandButton cmdҽԺ���� 
         Caption         =   "��"
         Height          =   300
         Left            =   3945
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtҽԺ���� 
         Height          =   300
         Left            =   1710
         MaxLength       =   40
         TabIndex        =   10
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox Txt���Ŀ¼ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   13
         Tag             =   "6"
         Top             =   660
         Width           =   2235
      End
      Begin VB.CommandButton Cmd���Ŀ¼ 
         Caption         =   "��"
         Height          =   300
         Left            =   3945
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   660
         Width           =   285
      End
      Begin VB.CommandButton cmd�ϴ�Ŀ¼ 
         Caption         =   "��"
         Height          =   300
         Left            =   3945
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1500
         Width           =   285
      End
      Begin VB.TextBox txt�ϴ�Ŀ¼ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   19
         Tag             =   "6"
         Top             =   1500
         Width           =   2235
      End
      Begin VB.CommandButton cmd����Ŀ¼ 
         Caption         =   "��"
         Height          =   300
         Left            =   3945
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1080
         Width           =   285
      End
      Begin VB.TextBox txt����Ŀ¼ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "6"
         Top             =   1080
         Width           =   2235
      End
      Begin VB.Label lbl����Ŀ¼ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ŀ¼(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   28
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lblҽ����ĿĿ¼ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����ĿĿ¼(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   21
         Top             =   2340
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ҽԺ����(&N)"
         Height          =   180
         Index           =   3
         Left            =   660
         TabIndex        =   9
         Top             =   300
         Width           =   990
      End
      Begin VB.Label Lbl���Ŀ¼ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ŀ¼(&I)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   12
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lbl�ϴ�Ŀ¼ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ�Ŀ¼(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   18
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label lbl����Ŀ¼ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ŀ¼(&O)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   15
         Top             =   1140
         Width           =   990
      End
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1605
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4515
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1095
         Left            =   3330
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4800
      TabIndex        =   25
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4800
      TabIndex        =   24
      Top             =   450
      Width           =   1100
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum�ı�
    textҽ���û� = 0
    Textҽ������ = 1
    Textҽ�������� = 2
    TextҽԺ���� = 3
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����
Dim mblnTest As Boolean
Dim mcnTest As New ADODB.Connection

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag) = False Then
        Exit Sub
    End If
    
    If Not mblnTest Then MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub

Private Sub Cmd���Ŀ¼_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "��ָ�����Ŀ¼��")
    If strPath = "" Then Exit Sub
    Txt���Ŀ¼.Text = strPath
End Sub

Private Sub Cmd����Ŀ¼_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "��ָ������Ŀ¼��")
    If strPath = "" Then Exit Sub
    txt����Ŀ¼.Text = strPath
End Sub

Private Sub cmd�ϴ�Ŀ¼_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "��ָ�����Ŀ¼��")
    If strPath = "" Then Exit Sub
    txt�ϴ�Ŀ¼.Text = strPath
End Sub

Private Sub cmd����Ŀ¼_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "��ָ������Ŀ¼��")
    If strPath = "" Then Exit Sub
    txt����Ŀ¼.Text = strPath
End Sub

Private Sub cmdҽ����ĿĿ¼_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "��ָ��ҽ����ĿĿ¼��")
    If strPath = "" Then Exit Sub
    txtҽ����ĿĿ¼.Text = strPath
End Sub

Private Sub cmdҽԺ����_Click()
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If mcnTest.State = 0 Then
        mblnTest = True
        Call cmdTest_Click
        mblnTest = False
        If mcnTest.State = 0 Then Exit Sub
    End If
    
    gstrSQL = "" & _
        " SELECT A.ҽԺ����,A.ҽԺ����,zlSpellcode(A.ҽԺ����) As ����,B.����||'-'||B.���� AS ҽԺ�ȼ�,C.����||'-'||C.���� AS ҽԺ����" & _
        " FROM ҽԺ�ȼ� A," & _
        "     (SELECT B.����,B.����" & _
        "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
        "     WHERE A.���=B.��� AND A.����='ҽԺ�ȼ�') B," & _
        "     (SELECT B.����,B.����" & _
        "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
        "     WHERE A.���=B.��� AND A.����='ҽԺ����') C" & _
        " WHERE A.ҽԺ�ȼ�=B.����(+) AND A.ҽԺ����=C.����(+) AND A.��Ч����<=SYSDATE"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\���ղ�������", gstrSQL): rsTemp.Open gstrSQL, mcnTest: Call SQLTest
    If rsTemp.RecordCount = 0 Then
        MsgBox "û���ҵ���ҽԺ��Ϣ�������䣡", vbInformation, gstrSysName
        txtҽԺ����.SetFocus
        zlControl.TxtSelAll txtҽԺ����
        Exit Sub
    Else
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_����, rsTemp, "ҽԺ����", "ҽԺ�ȼ�ѡ��", "��ѡ��ҽԺ�ȼ���")
        Else
            blnReturn = True
        End If
    End If
    If blnReturn Then
        txtҽԺ����.Text = rsTemp!ҽԺ����
        txtҽԺ����.Tag = rsTemp!ҽԺ����
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    ElseIf KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'ҽ���û���','" & txtEdit(textҽ���û�).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'ҽ���û�����','" & txtEdit(Textҽ������).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'ҽ��������','" & txtEdit(Textҽ��������).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'ҽԺ����','" & txtҽԺ����.Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'���Ŀ¼','" & Txt���Ŀ¼.Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'����Ŀ¼','" & txt����Ŀ¼.Text & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'�ϴ�Ŀ¼','" & txt�ϴ�Ŀ¼.Text & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'����Ŀ¼','" & txt����Ŀ¼.Text & "',8)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'ҽ����ĿĿ¼','" & txtҽ����ĿĿ¼.Text & "',9)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '����ҽԺ���
    gstrSQL = "Select ����,˵��,�Ƿ��ֹ From ������� Where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����)
    '��������ҽ�������� 204-04-07
    gstrSQL = "zl_�������_Update(" & TYPE_���� & ",'" & rsTemp!���� & "','" & IIf(IsNull(rsTemp!˵��), "", rsTemp!˵��) & "','" & Me.txtҽԺ����.Tag & "'," & IIf(IsNull(rsTemp!�Ƿ��ֹ), 0, rsTemp!�Ƿ��ֹ) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    
    gComInfo_����.ҽԺ���� = txtҽԺ����.Tag
    gComInfo_����.���Ŀ¼ = Txt���Ŀ¼.Text
    gComInfo_����.����Ŀ¼ = txt����Ŀ¼.Text
    
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Textҽ������ Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Textҽ�������� Or Index = Textҽ������ Or Index = textҽ���û� Then
        '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Public Function ��������() As Boolean
'���ܣ������붫�󰢶��ɵ�ҽ���ӿ�
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    
    mblnOK = False
    
    On Error GoTo errHandle
    
    'ȡ���ղ���
    gstrSQL = "select ������,����ֵ from ���ղ��� " & _
              " where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����)
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽ���û���"
                txtEdit(textҽ���û�) = str����ֵ
            Case "ҽ��������"
                txtEdit(Textҽ��������) = str����ֵ
            Case "ҽ���û�����"
                txtEdit(Textҽ������).Text = "        "    '������
                txtEdit(Textҽ������).Tag = str����ֵ
            Case "ҽԺ����"
                txtҽԺ����.Text = str����ֵ
            Case "���Ŀ¼"
                Txt���Ŀ¼.Text = str����ֵ
            Case "����Ŀ¼"
                txt����Ŀ¼.Text = str����ֵ
            Case "�ϴ�Ŀ¼"
                txt�ϴ�Ŀ¼.Text = str����ֵ
            Case "����Ŀ¼"
                txt����Ŀ¼.Text = str����ֵ
            Case "ҽ����ĿĿ¼"
                txtҽ����ĿĿ¼.Text = str����ֵ
        End Select
        rsTemp.MoveNext
    Loop
    
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_����)
    If Not rsTemp.EOF Then txtҽԺ����.Tag = Nvl(rsTemp!ҽԺ����)
    
    mblnChange = False
    frmSet����.Show vbModal, frmҽ�����
    
    �������� = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtҽԺ����_GotFocus()
    Call zlControl.TxtSelAll(txtҽԺ����)
End Sub

Private Sub txtҽԺ����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrInput As String
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    StrInput = UCase(Trim(txtҽԺ����.Text))
    If Trim(StrInput) = "" Then Exit Sub
    
    If mcnTest.State = 0 Then
        mblnTest = True
        Call cmdTest_Click
        mblnTest = False
        If mcnTest.State = 0 Then Exit Sub
    End If
    
    gstrSQL = "SELECT * FROM (" & _
        " SELECT A.ҽԺ����,A.ҽԺ����,zlSpellcode(A.ҽԺ����) As ����,B.����||'-'||B.���� AS ҽԺ�ȼ�,C.����||'-'||C.���� AS ҽԺ����" & _
        " FROM ҽԺ�ȼ� A," & _
        "     (SELECT B.����,B.����" & _
        "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
        "     WHERE A.���=B.��� AND A.����='ҽԺ�ȼ�') B," & _
        "     (SELECT B.����,B.����" & _
        "     FROM ָ������ A,ָ����ϵ���ձ� B" & _
        "     WHERE A.���=B.��� AND A.����='ҽԺ����') C" & _
        " WHERE A.ҽԺ�ȼ�=B.����(+) AND A.ҽԺ����=C.����(+) AND A.��Ч����<=SYSDATE) A" & _
        " WHERE (A.ҽԺ���� Like '" & StrInput & "%' Or A.ҽԺ���� Like '" & StrInput & "%' Or A.���� Like '" & StrInput & "%')"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\���ղ�������", gstrSQL): rsTemp.Open gstrSQL, mcnTest: Call SQLTest
    If rsTemp.RecordCount = 0 Then
        MsgBox "û���ҵ���ҽԺ��Ϣ�������䣡", vbInformation, gstrSysName
        txtҽԺ����.SetFocus
        zlControl.TxtSelAll txtҽԺ����
        Exit Sub
    Else
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_����, rsTemp, "ҽԺ����", "ҽԺ�ȼ�ѡ��", "��ѡ��ҽԺ�ȼ���")
        Else
            blnReturn = True
        End If
    End If
    If blnReturn Then
        txtҽԺ����.Text = rsTemp!ҽԺ����
        txtҽԺ����.Tag = rsTemp!ҽԺ����
    End If
End Sub

Private Function OpenDire(odtvOwner As Form, Optional odtvTitle As String) As String
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = odtvTitle
   With tBrowseInfo
      .hwndOwner = odtvOwner.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      OpenDire = sBuffer
   End If
End Function

