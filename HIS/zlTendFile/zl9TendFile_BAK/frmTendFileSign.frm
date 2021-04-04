VERSION 5.00
Begin VB.Form frmTendFileSign 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ǩ��"
   ClientHeight    =   2835
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5295
   Icon            =   "frmTendFileSign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   2505
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2670
      TabIndex        =   4
      Top             =   2370
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3930
      TabIndex        =   5
      Top             =   2370
      Width           =   1095
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -375
      TabIndex        =   6
      Top             =   2250
      Width           =   5670
   End
   Begin VB.CheckBox chkEsign 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����ǩ��(&E)"
      Height          =   195
      Left            =   3930
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1860
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ǩ��"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   750
      TabIndex        =   9
      Top             =   990
      Width           =   540
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ƽǩ��"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   750
      TabIndex        =   8
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������ѡ���ݵ������ǩ�˵���߼��𣬳����Զ�ѡ������Ӧ�ĸ��߼���"
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   3960
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   210
      Picture         =   "frmTendFileSign.frx":000C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���Լ��޸Ĺ������ݽ���ǩ��������ȱʡѡ����߼��𣻶�������ǩ���������޸ĺ�ǩ���������Զ�ѡ����ͬ����"
      ForeColor       =   &H00FF0000&
      Height          =   540
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   420
      Width           =   3960
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ǩ������(&L)"
      Height          =   180
      Left            =   255
      TabIndex        =   1
      Top             =   1860
      Width           =   990
   End
End
Attribute VB_Name = "frmTendFileSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '������
Private mblnOK As Boolean
Private Sign As cEPRSign                    'ǩ������

Private objESign As Object                  '����ǩ���ӿڲ���
Private lngCertID As Long                   '֤��ID
Private lngPassType As Long                 '������֤����ϵͳ������ 0-���룻1�����֣�2�����߽Կ�
Private mbln��ǩ As Boolean                 '�Ƿ���ǩ
Private mlngCur As Long, mlngLast As Long   '��ǰ��Ա������ǩ�˼���

Private mlng�ļ�ID As Long
Private mstrSource As String                 '����ǩ����Դ�ַ���
Private mstr״̬ As String
Private mstrPrivs As String

Private Enum SignLevel
    ���� = 1
    ���� = 2
    �м� = 3
    ʦ�� = 4
    Աʿ = 5
    δ���� = 9
End Enum

'######################################################################################################################
'���˻�������.����ˣ��������һ��ǩ�������һ��ǩ���ˣ���ʽΪ����ǩ/ǩ��
'��¼���� = 1 And ��ֹ�汾 Is NULLΪԭʼ��¼
'���˻�������.�����ΪNULL��δǩ��������/��ʾ��ǩ������/��ʾ����ǩ
'δ��ǩ֮ǰ��ͬ�������໥�޸ģ�ǩ����һ����ǩ�󣬾�ֻ�ܼ�����ǩ��
'ȡ����ǩʱ���Զ�ɾ���޸ĺۼ�
'��ǩ�󣬲������Ӽ�¼����=5��ǩ����¼
'������ǩ��¼ʱ���������޸ģ�Ҫô���ϵ���ǩ��Ҫôһֱ���˵���ͨǩ����¼״̬
'�����µ���ǩ��¼����ǩ����ʱ��������ֶ�Ҫ����
'######################################################################################################################

'����ǩ��ʹ�ó��ϣ�
'26  ����ǩ��ʹ�ó���(4λ�ַ�) �Բ�ͬ�����Ƿ�ʹ�õ���ǩ�����п���,����λ���ֱ�Ϊ:����,סԺ,ҽ��,���� 0-������,1-����
Public Function ShowMe(ByVal objParent As Object, ByVal strPrivs As String, ByVal lng�ļ�ID As Long, ByVal intLevel As Integer, _
    ByVal sSource As String, ByVal bln��ǩ As Boolean, Optional str״̬ As String, Optional str���� As String) As cEPRSign
    '******************************************************************************************************************
    '���ܣ� ��ʾǩ������
    '������ edtThis     :IN     �༭���ؼ�
    '       fParent     :IN     ������
    '       mstrSource   :IN     ����ǩ����Դ�ַ��������ı�����ȡ��ȥ��ǩ����٣�
    '       str״̬     :IN     ��������ǩ��ʱ���룬����Ƶ������ǩ������
    '       str��ǩ��   :IN     ��ǩʱ�����ϴ���ǩ���������Ա��ʵ��ǩȨ��
    '******************************************************************************************************************
    
    Set Sign = New cEPRSign
    Set frmParent = objParent
    mstrSource = sSource
    mstr״̬ = str״̬
    mbln��ǩ = bln��ǩ
    mlngLast = intLevel
    mlng�ļ�ID = lng�ļ�ID
    mstrPrivs = strPrivs
    
    '�����û���ǩ����������ʼ����ǩ������
    Call GetUserLevel(glngUserId)           '��ȡ�û�ǩ������
    
    '��ǩ�������ϴμ���ߵ�;ƽǩ��ֻ�����ϴ���ͬ�������
    If bln��ǩ Or mlngLast = δ���� Then
        If Not (mlngCur < mlngLast) Then
            str���� = "��Ҫ��������¼��ǩ���߻��ϴ���ǩ�ߵļ��������ǩ��"
            Unload Me
            Exit Function
        End If
        If mlngCur <= ���� And ���� < mlngLast Then cmbLevel.AddItem "5-���λ�ʦ"
        If mlngCur <= ���� And ���� < mlngLast Then cmbLevel.AddItem "4-�����λ�ʦ"
        If mlngCur <= �м� And �м� < mlngLast Then cmbLevel.AddItem "3-���ܻ�ʦ"
        If mlngCur <= ʦ�� And ʦ�� < mlngLast Then cmbLevel.AddItem "2-��ʦ"
        If mlngCur <= Աʿ And Աʿ < mlngLast Then cmbLevel.AddItem "1-��ʿ"
        If mlngCur > Աʿ Then cmbLevel.AddItem "0-δ����"
    Else
        If Not (mlngCur <= mlngLast) Then
            str���� = "������Ҫ�ﵽ�ϴ�ǩ���ߵļ������ǩ����"
            Unload Me
            Exit Function
        End If
        Select Case mlngCur
        Case ����
            cmbLevel.AddItem "5-���λ�ʦ"
        Case ����
            cmbLevel.AddItem "4-�����λ�ʦ"
        Case �м�
            cmbLevel.AddItem "3-���ܻ�ʦ"
        Case ʦ��
            cmbLevel.AddItem "2-��ʦ"
        Case Աʿ
            cmbLevel.AddItem "1-��ʿ"
        End Select
    End If
    cmbLevel.ListIndex = 0
    
    '��ȡ��ǰǩ����ʽ��ϵͳ����26��
    lngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), 4, 1))     '����,סԺ,ҽ��,���� (1111),Ϊ��Ĭ�ϲ�������ģʽ
    chkEsign.Value = Val(zlDatabase.GetPara("��������ǩ��", glngSys, 1255, "0"))
    
    Call RefControls
    Call RestoreState
    
    If mstr״̬ <> "" Then
        '����ǩ��ʱ
        Call cmdOK_Click
    Else
        Me.Show vbModal, frmParent
    End If
    
    If mblnOK Then
        str״̬ = mstr״̬
        Set ShowMe = Sign
    Else
        Set ShowMe = Nothing
    End If
End Function

Public Sub GetUserLevel(ByVal lngUserID As Long)
    Dim strǩ���� As String, str��ǩ�� As String
    Dim rs As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    mlngCur = δ����
    '�����Ƿ��ŵģ�1����������ԣ��ж�ֵ����С����ǩ�˵ļ��𣬷�������ǩ��
    
    'ȡ��ǰ����Ա�ļ���
    gstrSQL = "select /*+ RULE */ Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        mlngCur = NVL(rs("Ƹ�μ���ְ��"), δ����)
    End If
errHand:
    Exit Sub
End Sub

Private Sub RestoreState()
    Dim arrData
    
    If mstr״̬ <> "" Then
        arrData = Split(mstr״̬, "|")
        cmbLevel.ListIndex = arrData(0)
        chkEsign.Value = arrData(1)
    End If
End Sub

Private Function Validation() As Boolean
    '******************************************************************************************************************
    '
    '���ܣ�  ����ǩ�����ڲ�ǩ���鲢ˢ����ʾ����֤�����������ǩ����
    '
    '******************************************************************************************************************
    On Error GoTo errHand
    Dim intLevel As Integer '0-����,ԭ�����1,Ϊ�˼���ǩ������Ķ���
    Dim strUserName As String, lngUserID As Long, strSign As String, strʱ��� As String
    
    If chkEsign.Value = vbChecked Then
        '����ǩ��
        Err.Clear
        On Error Resume Next
        If objESign Is Nothing Then
            Set objESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err = 0: strSign = ""
        End If
        If Not objESign Is Nothing Then
            Call objESign.Initialize(gcnOracle, glngSys)
        End If
        lngCertID = 0
        strSign = objESign.signature(mstrSource, UCase(gcnOracle.Properties(23)), lngCertID, strʱ���) '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
        If strSign = "" Then
            MsgBox "ǩ��ʧ�ܣ�", vbInformation + vbOKOnly, "ǩ��"
            Exit Function
        End If
    End If
    strUserName = gstrUserName
    lngUserID = glngUserId
    
    '�´ζ�ȡ��+1
    Select Case Mid(cmbLevel.Text, 1, 1)
    Case 5
        intLevel = 0    '1
    Case 4
        intLevel = 1    '2
    Case 3
        intLevel = 2    '3
    Case 2
        intLevel = 3    '4
    Case 1
        intLevel = 4    '5
    End Select
    
    '------------------------------------------------------------------------------------------------------------------
    Sign.���� = strUserName
    Sign.ǩ������ = intLevel                    '-1��Ϊ�˼���ǩ��������
    Sign.ǩ����Ϣ = strSign
    Sign.ǩ����ʽ = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.ǩ������ = 1
    Sign.֤��ID = IIf(Sign.ǩ����ʽ = 2, lngCertID, 0)
    Sign.ʱ��� = strʱ���
    
    Validation = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��֤�û��������Ƿ���ȷ
'################################################################################################################
Private Function OraDataOpen(ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    Dim strSQL As String
    Dim strError As String
    Dim Cn As New ADODB.Connection
    
    On Error Resume Next
    Err = 0
    With Cn
        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
        .Open gcnOracle.ConnectionString, strUserName, strUserPwd
        If Err <> 0 Then
            OraDataOpen = False
            Exit Function
        End If
        .Close
    End With
    Set Cn = Nothing
    OraDataOpen = True
    Exit Function
errHand:
    Set Cn = Nothing
    OraDataOpen = False
    Err = 0
End Function

'################################################################################################################
'## ���ܣ�  ˢ�¿ؼ�
'################################################################################################################
Private Sub RefControls()
    Select Case lngPassType
    Case 0
        '����ǩ��
        chkEsign.Value = vbUnchecked
        chkEsign.Visible = False
    Case 1
        '1������
        chkEsign.Value = vbChecked
        chkEsign.Visible = True
        chkEsign.Enabled = False
    Case 2
        '2�����߽Կ�
    End Select
End Sub

Private Sub cmbLevel_Click()
    cmdOK.Enabled = (Mid(Me.cmbLevel.Text, 1, 1) > 0)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Validation Then
        mstr״̬ = cmbLevel.ListIndex & "|" & chkEsign.Value
        
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lngPassType = 2 Then
        Call zlDatabase.SetPara("��������ǩ��", chkEsign.Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    End If
End Sub
