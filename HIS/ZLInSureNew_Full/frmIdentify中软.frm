VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   3255
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
   Icon            =   "frmIdentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txt������ 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1635
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   1290
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "���ݲ���(&L)"
      Height          =   315
      Left            =   4350
      TabIndex        =   5
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
      TabIndex        =   0
      Top             =   825
      Width           =   2355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   4530
      TabIndex        =   3
      Top             =   210
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   4530
      TabIndex        =   4
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
      TabIndex        =   7
      Top             =   -270
      Width           =   30
   End
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
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
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
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
      TabIndex        =   12
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   2430
      Width           =   510
   End
   Begin VB.Label lblNote 
      Caption         =   "���ڶ��������̵�����֮���������롣"
      Height          =   540
      Left            =   840
      TabIndex        =   6
      Top             =   165
      Width           =   3180
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
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
      TabIndex        =   8
      Top             =   885
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmIdentify����.frx":030A
      Top             =   405
      Width           =   480
   End
   Begin VB.Label lbl���� 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   1  'Fixed Single
      Height          =   885
      Left            =   645
      TabIndex        =   11
      Top             =   2250
      Width           =   3345
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'      1����IC��������Ϣ
Private Declare Function ReadICCard Lib "ICREAD.DLL" (iIC���� As TIC����) As Long
'      2��дIC��������Ϣ
Private Declare Function WriteICCard Lib "ICWRITE.DLL" (iIC���� As TIC����) As Long

Private mIC���� As TIC����   '��ʱ���濨��Ϣ
Private mintTimes As Integer
Private mblnOK As Boolean
Private mbln�ж���Ժ As Boolean      '�Ƿ���Ҫ�Ըò�����Ժ�������ж�

Private Sub chk����_Click()
    txtPwd.Text = ""
    lbl����.Caption = "���ţ�"
    lbl����.Caption = "������"
    If chk����.Value = 0 Then
        '��ͨ����
        txtPwd.MaxLength = Len(mIC����.Password)
        txtPwd.PasswordChar = "*"
        lblPwd.Caption = "����"
        lblNote.Caption = "���ڶ��������̵�����֮���������롣"
        timRead.Enabled = True
        cmdOK.Default = True
    Else
        '���ݲ���
        txtPwd.MaxLength = 18
        txtPwd.PasswordChar = ""
        lblPwd.Caption = "���֤"
        lblNote.Caption = "���������ݲ��˵����֤��"
        timRead.Enabled = False
        cmdOK.Default = False
    End If
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
    Dim ic As TIC����
    Dim lngReturn  As Long
    
    On Error GoTo errHandle
    ic = mIC����
    ic.Password = txt������(0).Text
    MousePointer = vbHourglass
    lngReturn = WriteICCard(ic)
    MousePointer = vbDefault
    If lngReturn <> 0 Then
        '��ȡʧ��
        MsgBox ������Ϣ_����(lngReturn), vbInformation, gstrSysName
        Exit Function
    End If
    mIC���� = ic
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
    Dim str��Ч�� As String
    Dim bln����ҽ�� As Boolean
    Dim str����ֵ As String
    Dim lngIndex As Long, lngCount As Long
    
    If ReadIC����(True) = False Then
        '����ʧ��
        Exit Function
    End If
    'H�� ����У���Ƿ���ȷ����֤IC����Password�������PasswordΪ9000�������������֤����
    If TruncZero(mIC����.Password) <> "9000" Then
        If TruncZero(mIC����.Password) <> txtPwd.Text Then
            MsgBox "�����������", vbInformation, gstrSysName
            txtPwd.Text = ""
            txtPwd.SetFocus
            Exit Function
        End If
    End If
    
    '���кϷ�����֤
    If txt������(0).Visible = False Then
        str��Ч�� = Get���ղ���_����(mIC����.CenterCode, "��Ч��", True)
        bln����ҽ�� = (Get���ղ���_����(mIC����.CenterCode, "����ҽ�ƻ���", False) = "1")
        
        'B�� �Ƿ����Ч���жϣ����ж�Center����CenterCode����IC���е�CenterCode�ļ�¼��UseExpired�ֶ���Ϣ�Ƿ�С�ڵ�ǰ���ڡ�
        If IsDate(str��Ч��) = False Then
            MsgBox "���ȴ�ҽ�������������ݺ���ʹ�ñ����ܡ�", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        If CDate(str��Ч��) < CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")) Then
            MsgBox "��������ҽ�������Ѿ�������Ч�ڡ�", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        
        'C�� �жϸ����˻��Ƿ���ָ��������ж�IC����InPerAcc-OutPerAcc�Ƿ�Ϊ������
        If mIC����.InPerAcc - mIC����.OutPerAcc < 0 Then
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
        If mIC����.DomainCode = 1 Then
            MsgBox "�ò������ڳ�פ���ְ����", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        
        'F�� �ж��Ƿ�������ذ���ְ�����ж�IC����DomainCode�Ƿ����2��
        If mIC����.DomainCode = 2 Then
            MsgBox "�ò���������ذ���ְ����", vbInformation, gstrSysName
            mintTimes = 10 'ֱ���˳���ǰ����
            Exit Function
        End If
        
        'G�� �ж�ְ���Ƿ���סԺ���ж�IC����InpatientFlag����סԺ���㲻���д��жϣ�
        If mbln�ж���Ժ = True Then
            If mIC����.InpatientFlag = "1" Then
                MsgBox "�ò�����Ȼ��Ժ��", vbInformation, gstrSysName
                mintTimes = 10 'ֱ���˳���ǰ����
                Exit Function
            End If
        End If
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If txt������(0).Visible = True Then Exit Sub
    
    If InStr("0123456789X", Chr(KeyAscii)) > 0 Then
        If Not ActiveControl Is txtPwd Then
            'ֱ�ӽ������������
            txtPwd.SetFocus
            DoEvents
            txtPwd.Text = Chr(KeyAscii)
            txtPwd.SelStart = Len(txtPwd.Text)
            txtPwd.SelLength = 0
        End If
    End If
End Sub

Private Sub timRead_Timer()
    Call ReadIC����
End Sub

Private Sub txtPwd_GotFocus()
    zlControl.TxtSelAll txtPwd
End Sub

Public Function GetPatient(ByVal bln�ж���Ժ As Boolean, Optional ByVal bln�޸����� As Boolean = False) As Boolean
    Dim lngIndex As Long
    
    mintTimes = 0
    mblnOK = False
    timRead.Enabled = True
    txtPwd.MaxLength = Len(mIC����.Password)
    txt������(0).MaxLength = txtPwd.MaxLength
    txt������(1).MaxLength = txtPwd.MaxLength
    
    mbln�ж���Ժ = bln�ж���Ժ
    '��Ԥ��һ��
    Call ReadIC����
    
    '�����Ƿ��޸����룬�ı���ʾ״̬
    If bln�޸����� = False Then
        lbl����.Top = lbl����.Top - 900
        lbl����.Top = lbl����.Top - 900
        lbl����.Top = lbl����.Top - 900
        
        Me.Height = Me.Height - 900
    Else
        Me.Caption = "IC�������޸�"
        cmdOK.Default = False
        chk����.Visible = False
        
        For lngIndex = 0 To 1
            lbl������(lngIndex).Visible = True
            txt������(lngIndex).Visible = True
        Next
    End If
    
    frmIdentify����.Show vbModal
    DoEvents
    '����ֵ
    If mblnOK = True Then
        gIC���� = mIC����
    End If
    GetPatient = mblnOK
    
End Function

Private Function ReadIC����(Optional ByVal blnMessage As Boolean = False) As Boolean
'���ܣ���IC���ϵ���Ϣ
    Dim lngReturn As Long
    
    If chk����.Value = 0 Then
        lngReturn = ReadICCard(mIC����)
    Else
        '�������嵥�ж�ȡ�������������IC���ṹ��
        If Get���ݲ���_����(Trim(txtPwd.Text), mIC����, False) = False Then
            Exit Function
        End If
    End If
    If lngReturn = 0 Then
        '��ȡ�ɹ�
        lbl����.Caption = "���ţ�" & TruncZero(mIC����.Cardno)
        lbl����.Caption = "������" & TruncZero(mIC����.Name)
        
        ReadIC���� = True
    Else
        '��ȡʧ��
        If blnMessage = True Then
            MsgBox ������Ϣ_����(lngReturn), vbInformation, gstrSysName
        End If
        lbl����.Caption = "���ţ�"
        lbl����.Caption = "������"
    End If
End Function

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If ReadIC���� = True Then
            If txt������(0).Visible = False Then
                cmdOK.SetFocus
            Else
                txt������(0).SetFocus
            End If
        End If
    End If
End Sub

Private Sub txt������_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt������(Index)
End Sub

Private Sub txt������_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
