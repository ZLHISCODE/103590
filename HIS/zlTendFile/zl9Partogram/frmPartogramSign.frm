VERSION 5.00
Begin VB.Form frmPartogramSign 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ǩ��"
   ClientHeight    =   3345
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5295
   Icon            =   "frmPartogramSign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      Picture         =   "frmPartogramSign.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   2500
      Width           =   240
   End
   Begin VB.ComboBox cboMan 
      Height          =   300
      Left            =   1365
      TabIndex        =   2
      Text            =   "cboMan"
      Top             =   1800
      Width           =   2505
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Width           =   2505
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2640
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3960
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -375
      TabIndex        =   8
      Top             =   2760
      Width           =   5670
   End
   Begin VB.CheckBox chkEsign 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����ǩ��(&E)"
      Height          =   195
      Left            =   3930
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1305
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʾ��"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   360
      TabIndex        =   12
      Top             =   2530
      Width           =   540
   End
   Begin VB.Label lblǩ���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ǩ����(&P)"
      Height          =   180
      Left            =   435
      TabIndex        =   1
      Top             =   1860
      Width           =   810
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   1200
      Width           =   3960
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   210
      Picture         =   "frmPartogramSign.frx":034E
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
      TabIndex        =   3
      Top             =   2220
      Width           =   990
   End
End
Attribute VB_Name = "frmPartogramSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '������
Private mblnOK As Boolean
Private Sign As cPartogramSign                    'ǩ������

Private lngCertID As Long                   '֤��ID
Private mlngPassType As Long                 '������֤����ϵͳ������ 0-���룻1�����֣�2�����߽Կ�
Private mbln��ǩ As Boolean                 '�Ƿ���ǩ
Private mlngCur As Long, mlngLast As Long   '��ǰ��Ա������ǩ�˼���

Private mlng�ļ�ID As Long
Private mlngDeptID As Long
Private mstrSource As String                 '����ǩ����Դ�ַ���
Private mstr״̬ As String
Private mstrUserInfo As String
Private mstrPrivs As String
Private mblnDrop As Boolean
'--��Ա��Ϣ
Private mlngUserID As Long
Private mstrUserName As String               '��ǰ�û�����
Private mstrUserAbbr As String               '��ǰ�û�����
Private mrsǩ���� As New ADODB.Recordset
Private gstrLike As String

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
Public Function ShowMe(ByVal objParent As Object, ByVal strPrivs As String, ByVal lng�ļ�ID As Long, ByVal lngDeptID As Long, ByVal intLevel As Integer, _
    ByVal sSource As String, ByVal bln��ǩ As Boolean, Optional str״̬ As String, Optional strUserInfo As String) As cPartogramSign
    '******************************************************************************************************************
    '���ܣ� ��ʾǩ������
    '������ edtThis     :IN     �༭���ؼ�
    '       fParent     :IN     ������
    '       mstrSource   :IN     ����ǩ����Դ�ַ��������ı�����ȡ��ȥ��ǩ����٣�
    '       str״̬     :IN     ��������ǩ��ʱ���룬����Ƶ������ǩ������
    '       str��ǩ��   :IN     ��ǩʱ�����ϴ���ǩ���������Ա��ʵ��ǩȨ��
    '       strUseringo :IN   ��ԱID'��Ա����'��Ա����
    '******************************************************************************************************************
    
    Dim arrUser
    
    Set Sign = New cPartogramSign
    Set frmParent = objParent
    mstrSource = sSource
    mstr״̬ = str״̬
    mbln��ǩ = bln��ǩ
    mlngLast = intLevel
    mlng�ļ�ID = lng�ļ�ID
    mlngDeptID = lngDeptID
    mstrPrivs = strPrivs
    mblnOK = False
    '��һ�ε���
    If mstr״̬ = "" Then
        mlngUserID = glngUserId
        mstrUserName = Replace(gstrUserName, "-", "")
        mstrUserAbbr = Replace(gstrUserAbbr, "-", "")
    Else
        arrUser = Split(strUserInfo, "'")
        mlngUserID = Val(arrUser(0))
        mstrUserName = CStr(arrUser(1))
        mstrUserAbbr = CStr(arrUser(2))
    End If
    
    Call GetUser(lngDeptID)
    
    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")

    Call RefControls
    
    If mstr״̬ <> "" Then
        '����ǩ��ʱ
        Call cmdOK_Click
    Else
        Me.Show vbModal, frmParent
    End If
    
    If mblnOK Then
        str״̬ = mstr״̬
        strUserInfo = mstrUserInfo
        Set ShowMe = Sign
    Else
        Set ShowMe = Nothing
    End If
End Function

Public Sub GetUser(ByVal lngWorkID As Long)
    Dim rs As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
    '��ȡ������������Ա��Ϣ
    gstrSQL = " Select Distinct a.Id, b.����id, a.���, a.����, Upper(a.����) As ����, c.��Ա����, Nvl(a.Ƹ�μ���ְ��, 0) As ְ��, b.ȱʡ" & vbNewLine & _
            " From ��Ա�� A, ������Ա B, ��Ա����˵�� C, ��������˵�� D" & vbNewLine & _
            " Where a.Id = b.��Աid And a.Id = c.��Աid And b.����id = d.����id And c.��Ա���� In ('ҽ��', '��ʿ') And" & vbNewLine & _
            "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And b.����id = [1]" & vbNewLine & _
            " Order By ����, ȱʡ Desc"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngWorkID)
    Set mrsǩ���� = rs
    With cboMan
        .Clear
        If Not rs.EOF Then
            Do While Not rs.EOF
            .AddItem Replace(NVL(rs!����), "-", "") & "-" & Replace(NVL(rs!����), "-", "")
            .ItemData(.NewIndex) = Val(rs!ID)
            rs.MoveNext
            Loop
        Else
            .AddItem mstrUserAbbr & "-" & mstrUserName
            .ItemData(.NewIndex) = mlngUserID
        End If
    End With
    
    '��λ����ǰ����Ա
    Call isCheckǩ����Exists(mstrUserName, True)
    If cboMan.ListIndex = -1 Then cboMan.ListIndex = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub GetUserLevel(ByVal lngUserID As Long)
    Dim strǩ���� As String, str��ǩ�� As String
    Dim rs As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    mlngCur = δ����
    '�����Ƿ��ŵģ�1����������ԣ��ж�ֵ����С����ǩ�˵ļ��𣬷�������ǩ��
    
    'ȡ��ǰ����Ա�ļ���
    gstrSQL = "select  Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        mlngCur = NVL(rs("Ƹ�μ���ְ��"), δ����)
    End If
    
    Exit Sub
errHand:
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
    On Error GoTo errHand
    Dim intLevel As Integer '0-����,ԭ�����1,Ϊ�˼���ǩ������Ķ���
    Dim strUserName As String, lngUserID As Long, strSign As String, strʱ��� As String, strʱ�����Ϣ As String
    
    '���ǩ�����Ƿ����
    If Not isCheckǩ����Exists(Mid(cboMan.Text, InStr(cboMan.Text, "-") + 1)) Then
        lblInfo.Caption = "��ʾ��ǩ������Ϣ��������,���飡"
        If cboMan.Enabled = True And cboMan.Visible = True Then cboMan.SetFocus
        Exit Function
    End If
    
    mstrUserName = Mid(cboMan.Text, InStr(cboMan.Text, "-") + 1)
    mstrUserAbbr = Mid(cboMan.Text, 1, InStr(cboMan.Text, "-") - 1)
    mlngUserID = Val(cboMan.ItemData(cboMan.ListIndex))
    If chkEsign.Value = vbChecked Then
        '����ǩ��
        Err.Clear
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err.Clear: strSign = ""
            On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If gobjESign Is Nothing Then
            MsgBox "����ǩ������δ����ȷ��װ��ǩ���������ܼ�����", vbInformation, gstrSysName
            Exit Function
        End If
        lngCertID = 0
        strSign = gobjESign.signature(mstrSource, UCase(gcnOracle.Properties(23)), lngCertID, strʱ���, , strʱ�����Ϣ) '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
        If strSign = "" Then
            MsgBox "ǩ��ʧ�ܣ�", vbInformation + vbOKOnly, "ǩ��"
            Exit Function
        End If
    End If
    strUserName = mstrUserName
    lngUserID = mlngUserID
    
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
    Sign.���� = mstrUserName
    Sign.ǩ������ = intLevel                    '-1��Ϊ�˼���ǩ��������
    Sign.ǩ����Ϣ = strSign
    Sign.ǩ����ʽ = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.ǩ������ = 1
    Sign.֤��ID = IIf(Sign.ǩ����ʽ = 2, lngCertID, 0)
    Sign.ʱ��� = strʱ���
    Sign.ʱ�����Ϣ = strʱ�����Ϣ
    
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
    Dim rsTemp As New ADODB.Recordset
    Dim arrData
    On Error GoTo errHand
    
    '63955:������,2013-09-16,���õ��ǩ�������ҵ�ǰǩ���Ĳ��������õĵ���ǩ�����ò����в���ʹ�õ���ǩ��
    '˵�������û�����õ���ǩ����Ҫ���õĲ���,��˵�����õ���ǩ���Ĳ���Ϊ���в���
    If mstr״̬ <> "" And InStr(1, mstr״̬, "|") <> 0 Then
        arrData = Split(mstr״̬, "|")
        mlngPassType = Val(arrData(1))
        cmbLevel.ListIndex = Val(arrData(0))
    Else
        gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) ����ǩ�� From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ǩ�����ò���", 4, mlngDeptID)
        If rsTemp.RecordCount > 0 Then
            mlngPassType = Val(NVL(rsTemp!����ǩ��, 0))
        Else
            mlngPassType = 0
        End If
        If mlngPassType = 1 Then
            If CertificateStoped(gstrUserName) = True Then mlngPassType = 0
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
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CertificateStoped(ByVal strName As String) As Boolean
'���ܣ����ǩ���˵�֤���Ƿ�ͣ�ã�ͣ�õĻ�����ʹ�õ���ǩ��
    On Error Resume Next
    CertificateStoped = True
    Err.Clear
    If gobjESign Is Nothing Then
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        If Not gobjESign Is Nothing Then Call gobjESign.Initialize(gcnOracle, glngSys)
    End If
    If gobjESign Is Nothing Then Exit Function
    CertificateStoped = gobjESign.CertificateStoped(strName)
    If Err <> 0 Then Err.Clear
End Function

Private Sub cboMan_Click()
    cmbLevel.Clear
    lblInfo.Caption = "��ʾ��"
     '�����û���ǩ����������ʼ����ǩ������
    Call GetUserLevel(Val(cboMan.ItemData(cboMan.ListIndex)))            '��ȡ�û�ǩ������
    
    '��ǩ�������ϴμ���ߵ�;ƽǩ��ֻ�����ϴ���ͬ�������
    If mbln��ǩ Or mlngLast = δ���� Then
        If Not (mlngCur < mlngLast) Then
            If mbln��ǩ = True Then
                lblInfo.Caption = "��ʾ����ǰǩ����Ҫ��������¼��ǩ���߻��ϴ���ǩ�ߵļ��������ǩ��"
            Else
                '˵����¼��û�н��й�ǩ��
                lblInfo.Caption = "��ʾ����ǩ���˻�δ����Ƹ��ְ�񼶱�������Ա�����н������ã�"
            End If
            cmdOK.Enabled = False
            Exit Sub
        End If
        If mlngCur <= ���� And ���� < mlngLast Then cmbLevel.AddItem "5-���λ�ʦ"
        If mlngCur <= ���� And ���� < mlngLast Then cmbLevel.AddItem "4-�����λ�ʦ"
        If mlngCur <= �м� And �м� < mlngLast Then cmbLevel.AddItem "3-���ܻ�ʦ"
        If mlngCur <= ʦ�� And ʦ�� < mlngLast Then cmbLevel.AddItem "2-��ʦ"
        If mlngCur <= Աʿ And Աʿ < mlngLast Then cmbLevel.AddItem "1-��ʿ"
        If mlngCur > Աʿ Then cmbLevel.AddItem "0-δ����"
    Else
        If Not (mlngCur <= mlngLast) Then
            lblInfo.Caption = "��ʾ����ǰǩ��������Ҫ�ﵽ�ϴ�ǩ���ߵļ������ǩ����"
            cmdOK.Enabled = False
            Exit Sub
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
End Sub

Private Sub cboMan_KeyDown(KeyCode As Integer, Shift As Integer)
    If cboMan.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cboMan.Hwnd, &H157, 0, 0) = 1
End Sub

Private Sub cboMan_KeyPress(KeyAscii As Integer)
    'Call zlControl.CboMatchIndex(cboMan.Hwnd, KeyAscii)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset
    If KeyAscii = 13 Then
        If cboMan.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cboMan.Text)
        If cboMan.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cboMan.List(cboMan.ListIndex) Then Call zlControl.CboSetIndex(cboMan.Hwnd, -1)
        End If
        If strText = "" Then
            cboMan.ListIndex = -1
        ElseIf cboMan.ListIndex = -1 Then
            intIdx = -1
            strFilter = ""
            '�ȸ��Ƽ�¼��
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrsǩ����)
            Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
            Dim strCompents As String 'ƥ�䴮
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrsǩ����.Filter = strFilter: iCount = 0
            With mrsǩ����
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrsǩ����.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '������������,��Ҫ���:
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        
                        
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                        If NVL(!���) = strText Then strResult = NVL(!����): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(NVL(!���)) = Val(strText) Then
                            If iCount = 0 Then strResult = NVL(!����)
                            iCount = iCount + 1
                        End If
                        
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(NVL(!���)) Like strText & "*" Then
                            If isCheckǩ����Exists(NVL(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsǩ����, rsTemp)
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(NVL(!����)) = strText Then
                            If iCount = 0 Then strResult = NVL(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(NVL(!����)) Like strCompents Then
                            If isCheckǩ����Exists(NVL(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsǩ����, rsTemp)
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                            If iCount = 0 Then strResult = NVL(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!���) Like strText & "*" Or Trim(NVL(!����)) Like strCompents Or Trim(NVL(!����)) Like strCompents Then
                            If isCheckǩ����Exists(NVL(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsǩ����, rsTemp)
                        End If
                    End Select
                    mrsǩ����.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = NVL(rsTemp!����)
            'ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheckǩ����Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                Select Case intInputType
                Case 0 '����ȫ����
                    rsTemp.Sort = "���"
                Case 1 '����ȫƴ��
                    rsTemp.Sort = "����"
                Case Else
                    '����ѡ������
                    rsTemp.Sort = "����"
                End Select
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cboMan, rsTemp, True, "", "ȱʡ,ְ��,���ȼ���", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheckǩ����Exists(NVL(rsReturn!����), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cboMan: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
             
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cboMan_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cboMan.ListIndex = -1 Then
            cboMan.Text = ""
            Exit Sub
        Else
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cboMan_Click
            ElseIf intIdx <> cboMan.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cboMan.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cboMan_Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Function isCheckǩ����Exists(ByVal str���� As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cboMan.ListCount - 1
        If NeedName(cboMan.List(i)) = str���� Then
            If blnLocateItem Then cboMan.ListIndex = i
            isCheckǩ����Exists = True
            Exit Function
        End If
    Next
End Function

Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Private Sub cboMan_Validate(Cancel As Boolean)
    If cboMan.Text <> "" Then
        If GetCboIndex(cboMan, NeedName(cboMan.Text)) = -1 Then cboMan.ListIndex = -1: cboMan.Text = ""
    End If
    If cboMan.Text = "" Then '˵��¼�����Ϣ���������б���
        lblInfo.Caption = "��ʾ��ǩ������Ϣ��������,���飡"
        cmdOK.Enabled = False
        Cancel = True
    End If
End Sub

Private Sub cmbLevel_Click()
    cmdOK.Enabled = (Mid(Me.cmbLevel.Text, 1, 1) > 0)
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cmbLevel.Hwnd, KeyAscii)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Validation Then
        mstr״̬ = cmbLevel.ListIndex & "|" & chkEsign.Value
        mstrUserInfo = mlngUserID & "'" & mstrUserName & "'" & mstrUserAbbr
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

Private Function GetCboIndex(cbo As ComboBox, strFind As String, Optional blnKeep As Boolean, Optional blnLike As Boolean) As Long
'���ܣ����ַ�����ComboBox�в�������
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '�Ⱦ�ȷ����
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), "-") > 0 Then
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

Private Sub PicInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(PicInfo.Hwnd, lblInfo.Caption)
End Sub
