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
      ItemData        =   "frmTendFileSign.frx":000C
      Left            =   1365
      List            =   "frmTendFileSign.frx":000E
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
   Begin VB.Label lblsinName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ǩ���ˣ�������"
      Height          =   180
      Left            =   255
      TabIndex        =   10
      Top             =   2455
      Width           =   1260
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
      Picture         =   "frmTendFileSign.frx":0010
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
Private Sign As cTendSign                    'ǩ������

Private lngCertID As Long                   '֤��ID
Private mlngPassType As Long                '�����Ƿ����õ���ǩ����ϵͳ������ 0-�����ƣ�1������
Private mbln��ǩ As Boolean                 '�Ƿ���ǩ
Private mlngCur As Long, mlngLast As Long   '��ǰ��Ա������ǩ�˼���

Private mlng�ļ�ID As Long
Private mstrSource As String                 '����ǩ����Դ�ַ���
Private mstr״̬ As String
Private mstrPrivs As String
Private mlngUnitID As Long                   '��ǰ����ID

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
Public Function ShowMe(ByVal objParent As Object, ByVal strPrivs As String, ByVal lng�ļ�ID As Long, ByVal lngUnitId As Long, ByVal intLevel As Integer, _
    ByVal sSource As String, ByVal bln��ǩ As Boolean, Optional str״̬ As String, Optional str���� As String, _
    Optional ByVal intSignMode As Integer = 0, Optional ByVal blnExchange As Boolean = False) As cTendSign
    '******************************************************************************************************************
    '���ܣ� ��ʾǩ������
    '������ edtThis     :IN     �༭���ؼ�
    '       objParent     :IN     ������
    '       lng�ļ�ID   :IN      �ļ�ID
    '       lngUnitId   :IN      ����ID
    '       mstrSource   :IN     ����ǩ����Դ�ַ��������ı�����ȡ��ȥ��ǩ����٣�
    '       str״̬     :IN     ��������ǩ��ʱ���룬����Ƶ������ǩ������
    '       str��ǩ��   :IN     ��ǩʱ�����ϴ���ǩ���������Ա��ʵ��ǩȨ��
    '******************************************************************************************************************
    Dim strLastInfo As String
    
    Set Sign = New cTendSign
    Set frmParent = objParent
    mstrSource = sSource
    mstr״̬ = str״̬
    mbln��ǩ = bln��ǩ
    mlngLast = intLevel
    mlng�ļ�ID = lng�ļ�ID
    mstrPrivs = strPrivs
    mlngUnitID = lngUnitId
    '76700:LPF:ǩ���ɹ�����ǩʱ�����ǩ������رհ�ť���ͻᵼ��ǩ����Ϊ�յļ�¼��
    mblnOK = False
    
    '�����û���ǩ����������ʼ����ǩ������
    Call GetUserLevel(glngUserId)           '��ȡ�û�ǩ������
    strLastInfo = ""
    If Not mlngLast = δ���� Then
        Select Case mlngLast
        Case ����
            strLastInfo = "5-���λ�ʦ"
        Case ����
            strLastInfo = "4-�����λ�ʦ"
        Case �м�
            strLastInfo = "3-���ܻ�ʦ"
        Case ʦ��
            strLastInfo = "2-��ʦ"
        Case Աʿ
            strLastInfo = "1-��ʿ"
        End Select
    End If
    '��ǩ�������ϴμ���ߵ�;ƽǩ��ֻ�����ϴ���ͬ�������
    If bln��ǩ Or mlngLast = δ���� Then
        '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
        If mlngCur = δ���� Then
            If bln��ǩ = True Then
                str���� = "��Ҫ��������¼��ǩ���߻��ϴ���ǩ�ߵļ��������ǩ��" & vbCrLf & _
                    "������ǰ��δ����Ƹ�μ���ְ��������Ա���������ã�"
            Else
                str���� = "����ǰ��δ����Ƹ�μ���ְ��������Ա���������ã�"
            End If
            Unload Me
            Exit Function
        End If

        If IIf(bln��ǩ = True And intSignMode = 1, False, (Not (mlngCur < mlngLast))) Then
            str���� = "��Ҫ��������¼��ǩ���߻��ϴ���ǩ�ߵļ��������ǩ��" & IIf(strLastInfo = "", "", "�ϴ�ǩ���˼���" & strLastInfo & "��")
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
            str���� = "������Ҫ�ﵽ�ϴ�ǩ���ߵļ������ǩ����" & IIf(strLastInfo = "", "", "�ϴ�ǩ���˼���" & strLastInfo & "��")
            Unload Me
            Exit Function
        End If
        '51589:������,2013-03-01,��ӽ���ǩ��
        If bln��ǩ = False And blnExchange = True Then
            Select Case mlngLast
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
        Else
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
    End If
    If bln��ǩ = True And intSignMode = 1 Then
        cmbLevel.ListIndex = cmbLevel.ListCount - 1
    Else
        cmbLevel.ListIndex = 0
    End If
    
    lblsinName.Caption = "ǩ���ˣ�" & gstrUserName
    
    If RefControls = False Then
        Unload Me
        Exit Function
    End If
    
    '43588:������,2012-09-13,��Ӽ�¼����ǩģʽ
    '51589:������,2013-03-01,��ӽ���ǩ��
    If mstr״̬ <> "" Or (bln��ǩ = True And intSignMode = 1) Or (bln��ǩ = False And blnExchange = True) Then
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
    
    Err = 0: On Error GoTo ErrHand
    mlngCur = δ����
    '�����Ƿ��ŵģ�1����������ԣ��ж�ֵ����С����ǩ�˵ļ��𣬷�������ǩ��
    
    'ȡ��ǰ����Ա�ļ���
    gstrSQL = "select  Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        mlngCur = NVL(rs("Ƹ�μ���ְ��"), δ����)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Validation() As Boolean
    '******************************************************************************************************************
    '
    '���ܣ�  ����ǩ�����ڲ�ǩ���鲢ˢ����ʾ����֤�����������ǩ����
    '
    '******************************************************************************************************************
    On Error GoTo ErrHand
    Dim intLevel As Integer '0-����,ԭ�����1,Ϊ�˼���ǩ������Ķ���
    Dim strUserName As String, lngUserID As Long, strSign As String, strʱ��� As String, strʱ�����Ϣ As String
    
    If chkEsign.Value = vbChecked Then
        '����ǩ��
        If InitESign = False Then
            MsgBox "����ǩ������δ����ȷ��װ��ǩ���������ܼ�����", vbInformation, gstrSysName
            Exit Function
        End If
        lngCertID = 0
        strSign = gobjESign.signature(mstrSource, gstrDBUser, lngCertID, strʱ���, , strʱ�����Ϣ)  '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
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
    Sign.ʱ�����Ϣ = strʱ�����Ϣ
    
    Validation = True
    Exit Function
ErrHand:
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
ErrHand:
    Set Cn = Nothing
    OraDataOpen = False
    Err = 0
End Function

'################################################################################################################
'## ���ܣ�  ˢ�¿ؼ�
'################################################################################################################
Private Function RefControls() As Boolean
    Dim arrData
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '63955:������,2013-09-16,���õ��ǩ�������ҵ�ǰǩ���Ĳ��������õĵ���ǩ�����ò����в���ʹ�õ���ǩ��
    '˵�������û�����õ���ǩ����Ҫ���õĲ���,��˵�����õ���ǩ���Ĳ���Ϊ���в���
    If mstr״̬ <> "" And InStr(1, mstr״̬, "|") <> 0 Then
        arrData = Split(mstr״̬, "|")
        mlngPassType = Val(arrData(1))
        cmbLevel.ListIndex = Val(arrData(0))
    Else
        gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) ����ǩ�� From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ǩ�����ò���", 4, mlngUnitID)
        If rsTemp.RecordCount > 0 Then
            mlngPassType = Val(NVL(rsTemp!����ǩ��, 0))
        Else
            mlngPassType = 0
        End If
        '123565,����ǩ������
        If mlngPassType = 1 Then
            If InitESign = True Then
                If gobjESign.CheckCertificate(gstrDBUser) = True Then ''֤���Ѿ�ע�ᣬ��֤��û��ͣ�ã��Ҳ�����key���������������ǩ��������ǩ����ֹ
                    If gobjESign.CertificateStoped(gstrUserName) = True Then mlngPassType = 0 '���ǩ���˵�֤���Ƿ�ͣ�ã�ͣ�õĻ�����ʹ�õ���ǩ����ʹ������ǩ��
                Else
                    '��ֹǩ������
                    Exit Function
                End If
            Else
                mlngPassType = 0  'ǩ����������ʧ�ܣ���������ǩ��
            End If
        End If
    End If
    
    Select Case mlngPassType
    Case 1
        '1�����õ���ǩ��
        chkEsign.Value = vbChecked
        chkEsign.Visible = True
        chkEsign.Enabled = False
    Case Else
        '�����õ���ǩ��
        chkEsign.Value = vbUnchecked
        chkEsign.Visible = False
    End Select
    
    RefControls = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitESign() As Boolean
'���ܣ�����ǩ����ʼ��
    If gobjESign Is Nothing Then
        On Error Resume Next
        Err.Clear

        Set gobjESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        On Error GoTo 0
        If Not gobjESign Is Nothing Then
            Call gobjESign.Initialize(gcnOracle, glngSys)
        End If
    End If
    InitESign = Not gobjESign Is Nothing
End Function

Private Sub cmbLevel_Click()
    cmdOK.Enabled = (Mid(Me.cmbLevel.Text, 1, 1) > 0)
End Sub

Private Sub cmdCanCel_Click()
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
