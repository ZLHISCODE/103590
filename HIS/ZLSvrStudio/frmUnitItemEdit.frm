VERSION 5.00
Begin VB.Form frmUnitItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽԺ��Ϣ��Ŀά��"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4845
   Icon            =   "frmUnitItemEdit.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.OptionButton optType 
      Caption         =   "ͼƬ"
      Height          =   180
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.OptionButton optType 
      Caption         =   "�ı�"
      Height          =   180
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Frame fraInfo 
      Height          =   120
      Left            =   0
      TabIndex        =   11
      Top             =   2100
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3525
      TabIndex        =   10
      Top             =   2275
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2340
      TabIndex        =   9
      Top             =   2275
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1290
      Width           =   2625
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   2
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lblNoteNo 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ���:003"
      Height          =   180
      Left            =   2880
      TabIndex        =   3
      Top             =   900
      Width           =   1080
   End
   Begin VB.Label lblMarks 
      BackStyle       =   0  'Transparent
      Caption         =   "�ù�������ҽԺ��Ϣ��Ŀ�Ķ��塢��������Ŀ��������Ŀ���͵ĵ�����Ӱ�����ݵ����á�"
      Height          =   390
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   4590
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   180
      Picture         =   "frmUnitItemEdit.frx":6852
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����"
      Height          =   180
      Left            =   1080
      TabIndex        =   6
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����"
      Height          =   180
      Left            =   1080
      TabIndex        =   4
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����"
      Height          =   180
      Left            =   1080
      TabIndex        =   1
      Top             =   900
      Width           =   720
   End
End
Attribute VB_Name = "frmUnitItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mstrNO As String '��Ŀ����
Private mstrName As String '��Ŀ����
Private mintType As Integer '��Ŀ����
Private mblnChange As Boolean
Private mstrRemarks As String
'ϵͳ�̶����õ���Ŀ
Private Const SYS_ITEMS = "�汾��,������Ŀ¼,�����û�,��������,�ռ�Ŀ¼,�ռ�����,վ����,վ������,վ������,��Ϣ����ƽ̨�ͻ���," & _
                          "�ͻ�����������,�ռ�Ŀ¼S,�����û�S,��������S,�ռ�Ŀ¼F,�����û�F,��������F,���ʶ˿�F,�ռ���ʽ," & _
                          "����Ա�˺�,����Ա����,�ͻ���Ԥ����ʱ���,����Ա,��֤��,ע����,������,��������,��Ȩ֤��,��Ȩ����," & _
                          "��Ȩ�ʴ�,վ����,��Ʒ����,֧���̼���,��Ʒ����,��Ȩ����,��λ����,��Ʒ������,����֧����,֧���̼���," & _
                          "֧����MAIL,֧����URL,֧����BBS,��Ȩվ��,ʹ������,��Ȩ����,Ӱ��DICOM�豸����,Ӱ����Ƶ�豸����," & _
                          "Ӱ��Ƭ��ӡ������,Ӱ���Ƭվ����,������������"
'ϵͳ���õĶ༶��Ŀ
Private Const SYS_ITEMS_EXTEND = "������Ŀ¼[n],�����û�[n],��������[n],FTP������[n],FTP�û�[n],FTP����[n],FTP�˿�[n]"

Private Enum UnitCol
    Col_���� = 0
    Col_��Ŀ = 1
    Col_�Ƿ�ͼƬ = 2
    Col_���� = 3
    Col_Edit = 4
    Col_Del = 5
    Col_�Ƿ�ı� = 6
End Enum

'===========================================================================
'==�����ӿ�
'===========================================================================
Public Function ShowMe(Optional ByRef strNo As String, Optional ByRef strName As String, Optional ByRef intType As Integer) As Boolean
'���ܣ���Ŀ�༭��������
'     strNo=�༭�ı��룬Ϊ�ձ�ʾ����
'���أ��Ƿ�����˱༭
'     strNo=�༭��ı���
'     strName=�༭�������
'     intType=�༭�������
    mblnOk = False
    mstrNO = strNo
    mstrName = strName
    mintType = intType
    Me.Show vbModal
    strNo = mstrNO
    strName = mstrName
    intType = mintType
    ShowMe = mblnOk
End Function
'===========================================================================
'==�¼�
'===========================================================================
Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mstrNO <> "" Then
        '��֤��ݲ��������˵��
        If Not CheckAuditStatus("0312", "������Ŀ", mstrRemarks) Then Exit Sub
    End If
    If ValiData() Then
        If mblnChange Then
            If Not SaveData() Then
                Exit Sub
            End If
            mblnOk = True
        End If
    Else
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Call LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrRemarks = ""
End Sub

Private Sub OptType_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtName_Change()
    mblnChange = True
End Sub

Private Sub txtName_GotFocus()
    SelAll txtName
End Sub

Private Sub txtNO_Change()
    mblnChange = True
End Sub

Private Sub txtNO_GotFocus()
 SelAll txtNO
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

'===========================================================================
'==˽�з���
'===========================================================================
Private Sub LoadData()
'���ܣ����ݼ���
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngMax As Long, strTmp As String
    Dim strNote As String
    
    On Error GoTo errH
    If mstrNO <> "" Then
        txtNO.Enabled = False
        txtNO.BackColor = Me.BackColor
        '������Ŀ
        strSQL = "Select ����, ����, �Ƿ�ͼƬ From Zlunitinfoitem Where ���� = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mstrNO)
        txtNO.Text = rsTmp!����
        txtNO.Tag = rsTmp!����
        txtName.Text = rsTmp!����
        txtName.Tag = rsTmp!����
        optType(Val(rsTmp!�Ƿ�ͼƬ & "")).value = True
        lblType.Tag = Val(rsTmp!�Ƿ�ͼƬ & "")
        lblNoteNo.Visible = False
    Else
        txtNO.Enabled = True
        txtNO.BackColor = &H80000005
        strSQL = "Select Max(Lpad(����, 3, '0')) ������ From Zlunitinfoitem"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mstrNO)
        lngMax = Val(rsTmp!������ & "")
        '�Ѿ��ﵽ�����룬Ѱ�ұ����϶
        If lngMax = 999 Then
            strSQL = "Select b.����" & vbNewLine & _
                    "From (Select Lpad(����, 3, '0') ���� From Zlunitinfoitem) a," & vbNewLine & _
                    "     (Select Lpad(Rownum || '', 3, '0') ���� From Dual Connect By Rownum < [1]) b" & vbNewLine & _
                    "Where a.����(+) = b.���� And a.���� Is Null"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, lngMax)
            If Not rsTmp.EOF Then
                If rsTmp.RecordCount <> 0 Then
                    strTmp = rsTmp!����
                    txtNO.Text = rsTmp!����
                    If rsTmp.RecordCount > 0 Then
                        rsTmp.MoveNext
                        strTmp = strTmp & "," & rsTmp!����
                    End If
                End If
            End If
            lblNoteNo.Visible = True
            If strTmp <> "" Then
                lblNoteNo.Caption = "��ѡ���:" & strTmp
            Else
                lblNoteNo.Caption = "�޷��������,���ֹ�ָ����λ��һλ����"
            End If
        Else
            lblNoteNo.Caption = "��ѡ���:����" & Lpad(lngMax & "", 3, "0")
            txtNO.Text = Lpad((lngMax + 1) & "", 3, "0")
        End If
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function ValiData() As Boolean
'���ܣ���������У��
    Dim intType As Integer, strName As String, strNo As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intTmp As Integer
    
    On Error GoTo errH
    intType = IIf(optType(0).value, 0, 1)
    strName = Trim(txtName.Text)
    strNo = Trim(txtNO.Text)
    If mstrNO = "" Then
        If strNo = "" Then
            MsgBox "��������Ŀ���롣", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        If Not IsNumeric(strNo) Then
            MsgBox "��Ŀ�������Ϊ��ֵ���ͣ����������롣", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        If ActualLen(strNo) > txtNO.MaxLength Then
            MsgBox "��Ŀ���볬��" & txtNO.MaxLength & "λ���ȣ����������롣", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
        
    Else
        '����δ�����ı�
        If intType = Val(lblType.Tag) And txtName.Tag = strName And txtNO.Tag = strNo Then
            mblnChange = False
            ValiData = True
            Exit Function
        End If
    End If
    If Trim(txtName.Text) = "" Then
        MsgBox "��������Ŀ���ơ�", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Function
    End If
    
    If txtName.Tag <> strName Then
        If ActualLen(strName) > txtName.MaxLength Then
            MsgBox "��Ŀ���Ƴ���" & txtName.MaxLength & "λ���ȣ����������롣", vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
    
        '��ϵͳ��Ŀ
        If InStr("," & SYS_ITEMS & ",", "," & strName & ",") > 0 Then
            MsgBox "��������ϵͳ�̶���Ŀ���뻻���������ơ�" & vbNewLine & "ϵͳ��Ŀ��" & SYS_ITEMS, vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
        '���Ƽӱ�ŷ�ʽ���ж����Ƿ���ϵͳ��չ��Ŀ
        If Not IsNumeric(strName) Then
            intTmp = ValEx(strName)
            If strName Like "*" & intTmp Then
                If InStr("," & SYS_ITEMS_EXTEND & ",", "," & Mid(strName, 1, Len(strName) - Len(intTmp & "")) & "[n]" & ",") > 0 Then
                    MsgBox "��������ϵͳ�̶���Ŀ���뻻���������ơ�" & vbNewLine & "ϵͳ��Ŀ��" & SYS_ITEMS_EXTEND, vbInformation, gstrSysName
                    txtName.SetFocus
                    Exit Function
                End If
            End If
        End If
        '�Ƿ�������Ѿ�����
        strSQL = "Select 1 From Zlunitinfoitem Where ���� ='" & strName & "'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            MsgBox "�������Ѿ���ʹ�ã��뻻���������ơ�", vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
        
        '�Ƿ���ϵͳ��ʹ�õ���Ŀ
        strSQL = "Select 1 From Zlreginfo Where ��Ŀ = '" & strName & "' And Not Exists (Select 1 From Zlunitinfoitem Where ���� = '" & strName & "')"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            MsgBox "��������ϵͳ�̶���Ŀ���뻻���������ơ�", vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
    End If
    If mstrNO = "" Then
        '�Ƿ�ñ����Ѿ�����
        strSQL = "Select 1 From Zlunitinfoitem Where ���� ='" & strNo & "'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            MsgBox "�ñ����Ѿ���ʹ�ã��뻻���������ơ�", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Function
        End If
    Else
        '�ж��Ƿ��������
        If Val(lblType.Tag) = 0 Then
            strSQL = "Select 1 From Zlreginfo Where ��Ŀ = '" & txtName.Tag & "'"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        Else
            strSQL = "Select 1 From Zlunitinfoimage Where ��Ŀ = '" & txtName.Tag & "'"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        End If
        If Not rsTmp.EOF Then
            '���͸ı�
            If intType <> Val(lblType.Tag) Then
                If MsgBox("��Ŀ���ͷ����ı䣬��ǰ�����ݻᱻ��ա��Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                    txtNO.SetFocus
                    Exit Function
                End If
            ElseIf txtName.Tag <> strName Then
                If MsgBox("��Ŀ���Ʒ����ı䣬��Ը���Ŀ�����ò���Ӱ�졣�Ƿ������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                    txtNO.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    mblnChange = True
    ValiData = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Function SaveData() As Boolean
'���ܣ��������ݱ���
    Dim intType As Integer, strName As String, strNo As String
    Dim strSQL As String
    
    On Error GoTo errH
    intType = IIf(optType(0).value, 0, 1)
    strName = Trim(txtName.Text)
    strNo = Trim(txtNO.Text)
    strSQL = "Zltools.b_Public.Zlunitinfoitemchange(" & IIf(mstrNO = "", 0, 1) & ",'" & strNo & "','" & strName & "'," & intType & ")"
    Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
    If mstrNO = "" Then
        '������Ҫ������־
        Call SaveAuditLog(1, "������Ŀ", strName)
    Else
        '������Ҫ������־
        Call SaveAuditLog(2, "������Ŀ", "�ɡ�" & mstrName & "������Ϊ��" & strName & "��", mstrRemarks)
    End If
    
    mstrNO = strNo
    mstrName = strName
    mintType = intType
    SaveData = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function



