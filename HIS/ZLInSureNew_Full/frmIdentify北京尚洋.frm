VERSION 5.00
Begin VB.Form frmIdentify�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   1560
      Left            =   75
      TabIndex        =   5
      Top             =   75
      Width           =   3075
      Begin VB.TextBox txtNO 
         Height          =   300
         Left            =   990
         TabIndex        =   0
         Top             =   270
         Width           =   1890
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   990
         TabIndex        =   1
         Text            =   "ҽ������"
         Top             =   660
         Width           =   1890
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1065
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   735
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "���˱��"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   345
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   3
         Left            =   540
         TabIndex        =   6
         Top             =   1140
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   945
      TabIndex        =   3
      Top             =   1755
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   2055
      TabIndex        =   4
      Top             =   1755
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long, mstrIdentify As String, mbytType As Byte

Public Function Identify(ByVal bytType As Byte, Optional lng����ID As Long) As String
    mbytType = bytType
    mlng����ID = lng����ID
    Me.Show 1
    Identify = mstrIdentify
End Function

Private Sub cmdCancel_Click()
    mstrIdentify = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim strIdentify As String
    Dim strAddition As String
    Dim lngSequence As String
    On Error GoTo errHandle
    '�ж��Ƿ�����ҽ����������
    If Trim(txtNO.Text) = "" Then
        MsgBox "������ҽ�����˵ĸ��˱�ţ�", vbInformation, gstrSysName
        txtNO.SetFocus
        Exit Sub
    End If
'    If Trim(Text1.Text) = "" Then
'        If Text1.Enabled = True Then
'            MsgBox "������ҽ����������", vbInformation, gstrSysName
'            Text1.SetFocus
'            Exit Sub
'        End If
'    End If
    WriteInfo Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MI:SS") & "��ʼ���ɲ��������Ϣ"
    lngSequence = UCase(txtNO.Text)
'      strInfo='0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
'      8����;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(1,2,3);15����֤��;16�����;17�Ҷȼ�
'      18�ʻ������ۼ�;19�ʻ�֧���ۼ�;20����ͳ���ۼ�;21ͳ�ﱨ���ۼ�;22סԺ�����ۼ�;23�������
'      24��������;25�����ۼ�;26����ͳ���޶�
    
    strIdentify = lngSequence & ";"                             '0����
    strIdentify = strIdentify & lngSequence & ";"               '1ҽ���ţ����˱�ţ�
    strIdentify = strIdentify & ";"                             '2����
    strIdentify = strIdentify & Trim(Text1.Text) & ";"          '3����
    strIdentify = strIdentify & Combo1.Text & ";"               '4�Ա�
    strIdentify = strIdentify & txtNO.Tag & ";"                 '5��������
    strIdentify = strIdentify & ";"                             '6���֤
    strIdentify = strIdentify & Text1.Tag & ";"                 '7.��λ����(����)
    strAddition = "0;"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & Combo1.Tag & ";"                '10��Ա���
    '��Ϊ����ȡ�ø��������Խ����������㹻���ֵ������֧�������ҽ��ȷ��
    strAddition = strAddition & "1000000;"                      '11�ʻ����
    strAddition = strAddition & "0;"                            '12��ǰ״̬
    strAddition = strAddition & ";"                             '13����ID
    strAddition = strAddition & "1;"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & "1000000;"                      '18�ʻ������ۼ�
    strAddition = strAddition & ";"                             '19�ʻ�֧���ۼ�
    strAddition = strAddition & "0;"                            '20����ͳ���ۼ�
    strAddition = strAddition & "0;"                            '21ͳ�ﱨ���ۼ�
    strAddition = strAddition & "0;"                            '22סԺ�����ۼ�
    strAddition = strAddition & ";"                             '23��������
    WriteInfo Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MI:SS") & "��ʼ�������˵��������ݣ� strIdentify & strAddition"
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_��������)
    WriteInfo Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MI:SS") & "��ɵ�������"
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrIdentify = strIdentify & mlng����ID & ";" & strAddition
    End If
    Me.Hide
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Private Sub Form_Load()
    Combo1.AddItem "��"
    Combo1.AddItem "Ů"
    Combo1.ListIndex = 0
End Sub

Private Sub Text1_GotFocus()
    Call zlControl.TxtSelAll(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Combo1.SetFocus
End Sub

Private Sub txtNO_GotFocus()
    Call zlControl.TxtSelAll(txtNO)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    Dim StrInput As String, strSQL As String, rsTmp As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txtNO <> "" Then
        txtNO.Text = UCase(txtNO.Text)
        If mbytType = 1 Then
            Set rsTmp = gcn����.Execute("Select * From SICK_VISIT_INFO Where PERSONAL_NUMBER='" & txtNO.Text & "' And HOSPITAL_NUMBER='" & gstrҽԺ���� & "'")
        Else
            Set rsTmp = gcn����.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where PERSONAL_NUMBER='" & txtNO.Text & "' And HOSPITAL_NUMBER='" & gstrҽԺ���� & "'")
        End If
        If rsTmp.EOF Then
            Text1.Enabled = True
            Combo1.Enabled = True
            txtNO.Tag = ""
            Text1.Tag = ""
            Combo1.Tag = ""
            Text1.SetFocus
        Else
            Text1.Enabled = False
            Combo1.Enabled = False
            Text1.Text = rsTmp!Name                                    '����
            Text1.Tag = rsTmp!UNIT_NUMBER                              '��λ����
            Combo1.ListIndex = IIf(Trim(Nvl(rsTmp!Sex, "0")) = "1", 0, 1)
            If mbytType = 1 Then
                Call DebugTool("��������:" & rsTmp!BIRTHMONTH)
                txtNO.Tag = Mid(rsTmp!BIRTHMONTH, 1, 4) & "-" & Mid(rsTmp!BIRTHMONTH, 5, 2) & "-01"     '��������
            Else
                txtNO.Tag = Format(rsTmp!BIRTH_DATE, "yyyy-MM-dd")         '��������
            End If
            Combo1.Tag = rsTmp!PERSONAL_TYPE                           '��Ա���
            cmdOK.SetFocus
        End If
    End If
End Sub

