VERSION 5.00
Begin VB.Form frmIdentifyͭ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentifyͭ��.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   300
      Left            =   3660
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt���� 
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   1635
      MaxLength       =   8
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox txt������ 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1635
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1740
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox txt������ 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1635
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1290
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "���ݲ���(&L)"
      Height          =   315
      Left            =   4350
      TabIndex        =   12
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Timer timRead 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   30
      Top             =   1650
   End
   Begin VB.TextBox txtPwd 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1635
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   825
      Width           =   2355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   405
      Left            =   4530
      TabIndex        =   10
      Top             =   210
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   4530
      TabIndex        =   11
      Top             =   870
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   4170
      TabIndex        =   13
      Top             =   -270
      Width           =   30
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmIdentifyͭ��.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   810
      TabIndex        =   5
      Top             =   2220
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   810
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   555
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   810
      TabIndex        =   15
      Top             =   2790
      Width           =   510
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   810
      TabIndex        =   14
      Top             =   2430
      Width           =   510
   End
   Begin VB.Label lblNote 
      Caption         =   "���ڶ��������̵�����֮���������롣"
      Height          =   540
      Left            =   840
      TabIndex        =   0
      Top             =   165
      Width           =   3180
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   810
      TabIndex        =   1
      Top             =   885
      Width           =   510
   End
   Begin VB.Label lbl���� 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   1  'Fixed Single
      Height          =   885
      Left            =   645
      TabIndex        =   16
      Top             =   2280
      Width           =   3345
   End
End
Attribute VB_Name = "frmIdentifyͭ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'      1����IC��������Ϣ
Private Declare Function ReadICCard Lib "ICREAD.DLL" (iICͭ�� As TICͭ��) As Long
'      2��дIC��������Ϣ
Private Declare Function WriteICCard Lib "ICWRITE.DLL" (iICͭ�� As TICͭ��) As Long

Private mICͭ�� As TICͭ��   '��ʱ���濨��Ϣ

Private mintTimes As Integer
Private mblnOK As Boolean
Private mint���� As Integer
Private mlng����ID As Long
Private mstr���ֱ��� As String
Private mblnԶ����֤ As Boolean, mstrԶ�̵�ַ As String
Private blnUpload As Boolean

Private Sub chk����_Click()
    txtPwd.Text = ""
    lbl����.Caption = "���ţ�"
    lbl����.Caption = "������"
    
    Call SetFace
    txtPwd.SetFocus
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    mintTimes = mintTimes + 1
    '��ʱֹͣ�Զ���ȡ������������������ͻ
    timRead.Enabled = False
    If IsValid = False Then
        '�ָ�
        timRead.Enabled = True
        If mintTimes > 3 Then
            '����������̫��
            Unload Me
        End If
        Exit Sub
    End If
    mlng����ID = Val(txt����.Tag)
    
    If txt������(0).Visible = True Then
        If SavePass() = False Then
            Exit Sub
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Private Function SavePass() As Boolean
'���ܣ��޸��û�IC������
    Dim ic As TICͭ��
    Dim lngReturn  As Long
    
    On Error GoTo errHandle
    ic = mICͭ��
    ic.Password = txt������(0).Text
    MousePointer = vbHourglass
    lngReturn = WriteICCard(ic)
    MousePointer = vbDefault
    If lngReturn <> 0 Then
        '��ȡʧ��
        MsgBox ������Ϣ_ͭ��(lngReturn), vbInformation, gstrSysName
        Exit Function
    End If
    mICͭ�� = ic
    MsgBox "�����뱣��ɹ���", vbInformation, gstrSysName
    SavePass = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MousePointer = vbDefault
End Function

Private Function IsValid() As Boolean
'���ܣ��ж�IC���Ƿ�Ϸ�
    Dim rsTemp As New ADODB.Recordset
    Dim dat��Ч�� As String
    Dim bln����ҽ�� As Boolean
    Dim str����ֵ As String, str������� As String
    Dim lngIndex As Long, lngCount As Long
    
    If ReadICͭ��(True) = False Then
        '����ʧ��
        Exit Function
    End If
    'H�� ����У���Ƿ���ȷ����֤IC����Password�������PasswordΪ9000�������������֤����
    If TruncZero(mICͭ��.Password) <> "9000" Then
        If mblnԶ����֤ = False Then
            If TruncZero(mICͭ��.Password) <> txtPwd.Text Then
                MsgBox "�����������", vbInformation, gstrSysName
                txtPwd.Text = ""
                txtPwd.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '���кϷ�����֤
    If txt������(0).Visible = False Then
        If mint���� = 1 And txt����.Tag = "" Then
            MsgBox "��Ժ����ѡ���֡�", vbInformation, gstrSysName
            Exit Function
        End If
        If mint���� = 0 And txt����.Tag <> "" Then
            '����Ƿ�֧���ñ���
            gstrSQL = "SELECT A.��� FROM ���ղ��� A " & _
                      "  WHERE A.����=81 AND A.����='" & mstr���ֱ��� & "' and A.���>'0'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcnͭ��, adOpenStatic, adLockReadOnly
            If rsTemp.EOF = True Then
                MsgBox "��ҽ�����Ĳ����ҵ��ò��֡�", vbInformation, gstrSysName
                Exit Function
            End If
            str������� = rsTemp("���")
        End If
        
        gstrSQL = "SELECT B.��Ч��,B.�Ƿ����,A.����ģʽ,A.��չ��������,A.��չ�󲡱��� " & _
                   " FROM ��������Ŀ¼ A,�������� B " & _
                   " WHERE A.����=" & TYPE_ͭ�� & " AND A.����='" & mICͭ��.CenterCode & "' AND A.��������=B.���� AND A.����=B.���� "
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnͭ��, adOpenStatic, adLockReadOnly
        If rsTemp.EOF = False Then
            dat��Ч�� = Nvl(rsTemp("��Ч��"), Date)
            bln����ҽ�� = Nvl(rsTemp("�Ƿ����"), 0) And Nvl(rsTemp("����ģʽ"), 0)
        Else
            MsgBox "���ȴ�ҽ�������������ݺ���ʹ�ñ����ܡ�", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        
        If str������� >= "1" And str������� <= "5" Then
            If Nvl(rsTemp("��չ��������"), 0) <> 1 Then
                MsgBox "��Ժ�ڲ�����������δ��չ����������", vbInformation, gstrSysName
                Exit Function
            End If
        ElseIf str������� >= "6" And str������� <= "9" Then
            If Nvl(rsTemp("��չ�󲡱���"), 0) <> 1 Then
                MsgBox "��Ժ�ڲ�����������δ��չ�󲡱�����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        'B�� �Ƿ����Ч���жϣ����ж�Center����CenterCode����IC���е�CenterCode�ļ�¼��UseExpired�ֶ���Ϣ�Ƿ�С�ڵ�ǰ���ڡ�
        If dat��Ч�� < CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")) Then
            MsgBox "��������ҽ�������Ѿ�������Ч�ڡ�", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        
        'C�� �жϸ����˻��Ƿ���ָ��������ж�IC����InPerAcc-OutPerAcc�Ƿ�Ϊ������
        If mICͭ��.InPerAcc - mICͭ��.OutPerAcc < 0 Then
            MsgBox "���˸����˻��Ѿ����ָ�����", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        
        'D�� �ж��Ƿ񶨵�ҽ�ƻ��������ж�Center����CenterCode����IC���е�CenterCode�ļ�¼��IsAppoint�ֶ���Ϣ��
        '    ���IsAppoint=1���ǣ�IsAppoint=0���
        If bln����ҽ�� = False Then
            MsgBox "��Ժ�����ڸò��˵Ķ���ҽ�ƻ�����", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        
        'E�� �ж��Ƿ����ڳ�פ���ְ�����ж�IC����DomainCode�Ƿ����1
        If mICͭ��.DomainCode = 1 Then
            MsgBox "�ò������ڳ�פ���ְ����", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        
        'F�� �ж��Ƿ�������ذ���ְ�����ж�IC����DomainCode�Ƿ����2��
        If mICͭ��.DomainCode = 2 Then
            MsgBox "�ò���������ذ���ְ����", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        
        'G�� �ж�ְ���Ƿ���סԺ���ж�IC����InpatientFlag����סԺ���㲻���д��жϣ�
'        If mbln�ж���Ժ = True Then
            If mICͭ��.InpatientFlag = "1" Then
                MsgBox "�ò�����Ȼ��Ժ��", vbInformation, gstrSysName
                mintTimes = 10 'ֱ���˳���ǰ����
                Exit Function
            End If
            
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��Ѿ���Ժ", TYPE_ͭ��, CStr(TrimStr(mICͭ��.Cardno)))
            If Not rsTemp.EOF Then
                If rsTemp!״̬ = 1 Then
                    MsgBox "�ò�����Ȼ��Ժ��", vbInformation, gstrSysName
                    mintTimes = 10 'ֱ���˳���ǰ����
                    Exit Function
                End If
            End If
'        End If
    Else
        '�����޸�
        If txt������(0).Text <> txt������(1).Text Then
            txt������(0).Text = ""
            txt������(1).Text = ""
            txt������(0).SetFocus
            MsgBox "��������ȷ�����벻��ͬ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        For lngIndex = 0 To 1
            If Len(txt������(lngIndex).Text) <> txt������(lngIndex).MaxLength Then
                txt������(lngIndex).Text = ""
                txt������(lngIndex).SetFocus
                MsgBox "���볤�Ȳ�����", vbInformation, gstrSysName
                Exit Function
            End If
            
            For lngCount = 1 To Len(txt������(lngIndex).Text)
                If InStr("0123456789", Mid(txt������(lngIndex).Text, lngCount, 1)) = 0 Then
                    txt������(lngIndex).Text = ""
                    txt������(lngIndex).SetFocus
                    MsgBox "����ֻ����������ɡ�", vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        Next
    End If
    IsValid = True
End Function

Private Sub cmd����_Click()
    Dim rs���� As ADODB.Recordset
    
    'סԺҪѡ����ͨ����
    '����ѡ�����ز�
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=[1] And A.��� IN (" & IIf(mint���� = 0, "1,2)", "0)")
    Set rs���� = New ADODB.Recordset
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", TYPE_ͭ��)
    If rs����.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_ͭ��, rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�") = True Then
            txt����.Text = rs����("����")
            txt����.Tag = rs����("ID")
            mstr���ֱ��� = rs����("����")
            txt����.ForeColor = txtPwd.ForeColor
        End If
    End If
End Sub

Private Sub Form_Load()
    Shell "cmd /c route delete 0.0.0.0", vbNormal
    Shell "cmd /c route add 0.0.0.0 mask 0.0.0.0 192.168.0.1", vbNormal
End Sub

Private Sub timRead_Timer()
    If mblnԶ����֤ = False Then
        Call ReadICͭ��
    End If
End Sub

Private Sub txtPwd_GotFocus()
    zlControl.TxtSelAll txtPwd
End Sub

Public Function GetPatient(ByVal int���� As Integer, ByVal bln�޸����� As Boolean, ����ID As Long) As Boolean
    Dim lngIndex As Long
    Dim blnԶ����֤ As Boolean, strԶ�̵�ַ As String
    
    If Get���ղ���_ͭ��(blnԶ����֤, strԶ�̵�ַ) = False Then
        Exit Function
    End If
    If bln�޸����� = True And blnԶ����֤ = True Then
        MsgBox "���ڲ���Զ�������֤�����ܽ��������޸ġ�", vbInformation, gstrSysName
        Exit Function
    End If
    mblnԶ����֤ = blnԶ����֤
    mstrԶ�̵�ַ = strԶ�̵�ַ
    
    mintTimes = 0
    mblnOK = False
    timRead.Enabled = True
    txtPwd.MaxLength = Len(mICͭ��.Password)
    txt������(0).MaxLength = txtPwd.MaxLength
    txt������(1).MaxLength = txtPwd.MaxLength
    
    mint���� = int����
    If int���� = 0 Then
        '���ȼ���Ƿ����ʹ������
        
    End If
    
    '��Ԥ��һ��
    blnUpload = False
    Call ReadICͭ��
    
    '�����Ƿ��޸����룬�ı���ʾ״̬
    If bln�޸����� = False Then
'        If int���� = 0 Or int���� = 1 Then
            '��������Ժ��Ҫ�����벡��
            lbl����.Top = lbl������(0).Top
            txt����.Top = txt������(0).Top
            cmd����.Top = txt����.Top + 30
            lbl����.Visible = True
            txt����.Visible = True
            cmd����.Visible = True
'        End If
        lbl������(1).Caption = "����"
        lbl������(1).Left = lblPwd.Left
        Call SetFace
    Else
        Me.Caption = "IC�������޸�"
        cmdOK.Default = False
        chk����.Visible = False
        
        For lngIndex = 0 To 1
            lbl������(lngIndex).Visible = True
            txt������(lngIndex).Visible = True
        Next
    End If
    
    frmIdentifyͭ��.Show vbModal
    DoEvents
    '����ֵ
    If mblnOK = True Then
        ����ID = mlng����ID
        gICͭ�� = mICͭ��
    End If
    GetPatient = mblnOK
    
End Function

Private Function ReadICͭ��(Optional ByVal blnMessage As Boolean = False) As Boolean
'���ܣ���IC���ϵ���Ϣ
    Dim lngReturn As Long
    
    If chk����.Value = 0 Then
        If mblnԶ����֤ = False Then
            lngReturn = ReadICCard(mICͭ��)
        Else
            'Զ������
            If Trim(txtPwd.Text) = "" Then
                If blnMessage = True Then MsgBox "���������֤���롣", vbInformation, gstrSysName
                Exit Function
            End If
            If blnUpload = False Then
                If frmSockͭ��.CommIC(mstrԶ�̵�ַ, True, IIf(mint���� = 1, 1, 0), txtPwd.Text & "|" & txt������(1).Text) = False Then
                    Exit Function
                End If
                blnUpload = True
                mICͭ�� = gICͭ��Temp
            End If
        End If
    Else
        '�������嵥�ж�ȡ�������������IC���ṹ��
        If Get���ݲ���_ͭ��(Trim(txtPwd.Text), mICͭ��, False) = False Then
            If blnMessage = True Then MsgBox "δ�ҵ����֤��Ϊ " & txtPwd.Text & " �����ݲ��ˡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If lngReturn = 0 Then
        '��ȡ�ɹ�
        lbl����.Caption = "���ţ�" & TruncZero(mICͭ��.Cardno)
        lbl����.Caption = "������" & TruncZero(mICͭ��.Name)
        
        ReadICͭ�� = True
    Else
        '��ȡʧ��
        If blnMessage = True Then
            MsgBox ������Ϣ_ͭ��(lngReturn), vbInformation, gstrSysName
        End If
        lbl����.Caption = "���ţ�"
        lbl����.Caption = "������"
    End If
End Function

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        blnUpload = False
        '�������֤�����س�ʱ,ǿ���ϴ�����
        If ReadICͭ�� = True Then
            If txt������(0).Visible = False Then
                txt����.SetFocus
            Else
                txt������(0).SetFocus
            End If
        Else
            zlControl.TxtSelAll txtPwd
        End If
    End If
End Sub

Private Sub txt����_Change()
    txt����.Tag = ""
    txt����.ForeColor = &HC0&
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt����.Text = "" Or txt����.Tag <> "" Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strText = txt����.Text
    gstrSQL = "Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ⲡ','��ͨ��') ��� " & _
             "   FROM ���ղ��� A WHERE A.����=[1] And A.��� IN ([2]) And (" & _
             " A.���� like [3] || '%' or A.����  like [3] || '%' or  A.����  like [3] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_ͭ��, IIf(mint���� = 0, "1,2", "0"), strText)
    
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_ͭ��, rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '�϶����м�¼����
        txt����.Text = rsTemp("����")
        txt����.Tag = rsTemp("ID")
        mstr���ֱ��� = rsTemp("����")
        txt����.ForeColor = txtPwd.ForeColor
        SendKeys "{TAB}"
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt������_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt������(Index)
End Sub

Private Sub SetFace()
'���ܣ�����״̬���ý�����ʽ
    If chk����.Value = 1 Or mblnԶ����֤ = True Then
        '���ݲ���
        txtPwd.MaxLength = 18
        txtPwd.PasswordChar = ""
        lblPwd.Caption = "���֤"
        lblNote.Caption = "�����벡�˵����֤��"
        timRead.Enabled = False
        If chk����.Value = 1 Then
            '���ݲ���Ҫ����
            lbl������(1).Visible = False
            txt������(1).Visible = False
        Else
            'Զ����֤
            lbl������(1).Visible = True
            txt������(1).Visible = True
        End If
    Else
        'ֱ�Ӷ�IC��
        txtPwd.MaxLength = Len(mICͭ��.Password)
        txtPwd.PasswordChar = "*"
        lblPwd.Caption = "����"
        lblNote.Caption = "���ڶ��������̵�����֮���������롣"
        timRead.Enabled = True
    End If
End Sub

Private Sub txt������_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub


