VERSION 5.00
Begin VB.Form frmChange 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ת��"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frmChange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   9
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3810
      TabIndex        =   8
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2610
      TabIndex        =   7
      Top             =   2040
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   135
      TabIndex        =   10
      Top             =   45
      Width           =   5055
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1005
         Width           =   3795
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   615
         Width           =   1605
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   615
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4095
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Width           =   690
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   675
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   990
         TabIndex        =   6
         Text            =   "cbo����"
         Top             =   1395
         Width           =   3810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   180
         TabIndex        =   17
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ��λ"
         Height          =   180
         Left            =   2370
         TabIndex        =   16
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3690
         TabIndex        =   14
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2370
         TabIndex        =   13
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   540
         TabIndex        =   12
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ת�����"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   360
         TabIndex        =   15
         Top             =   675
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mlng����ID As Long
Public mlng��ҳID As Long
Public mlngUnit As Long
Public mstrPrivs As String
Private mstr������� As String
Private mintFlag As Integer
Private mstrDeptName As String
Private mrsPatiInfo As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cbo����_GotFocus()
    '����27370 by lesfeng 2010-02-03
    With cbo����
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    '����27370 by lesfeng 2010-02-03
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSql As String, intIdx As Long, i As Long
    Dim lngUnit As Long
    
    If KeyAscii = 13 Then
        mintFlag = 0
        strInput = UCase(cbo����.Text)
  
        Set rsTmp = InputDept(Me, Frame1, cbo����, "�ٴ�", mstr�������, strInput, blnCancel, -1, 0)
        If Not rsTmp Is Nothing Then
            intIdx = cbo.FindIndex(cbo����, rsTmp!ID)
            If intIdx <> -1 Then
                cbo����.ListIndex = intIdx
            End If
        Else
            If Not blnCancel Then
                MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
                cbo����.Text = mstrDeptName
                mintFlag = 1
            End If
        End If
    Else
        mintFlag = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '����27370 by lesfeng 2010-02-03
    If KeyCode = 13 And mintFlag = 0 Then cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    If isValid = False Then Exit Sub
    
    strSql = "zl_���˱䶯��¼_Change(" & mlng����ID & "," & mlng��ҳID & "," & _
        cbo����.ItemData(cbo����.ListIndex) & ",'" & UserInfo.��� & "'," & "'" & UserInfo.���� & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    '����96847
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng����ID, mlng��ҳID) <> 1 Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    gblnOK = True
    
    On Error Resume Next
    'ת�Ƴɹ��󴥷���Ϣ
    If mclsMipModule.IsConnect = True Then
        mclsXML.ClearXmlText '��������е�XML
        '--������Ϣ��װ
        '������Ϣ
        mclsXML.AppendNode "in_patient"
        'patient_id      ����id  1   N
        mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
        'page_id     ��ҳid  1   N
        mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
        'patient_name        ����    1   S
        mclsXML.appendData "patient_name", txt����.Text, xsString '����
        'patient_sex     �Ա�    0..1    S
        mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString '�Ա�
        'in_number       סԺ��  1   S
        mclsXML.appendData "in_number", Nvl(mrsPatiInfo!סԺ��), xsString 'סԺ��
        mclsXML.AppendNode "in_patient", True
        
        'ת����Ϣ
        'current_state       ת����Ϣ    1
        mclsXML.AppendNode "current_state"
        'current_area_id     ת������id  0..1    N
        mclsXML.appendData "current_area_id", Val(Nvl(mrsPatiInfo!��ǰ����ID)), xsNumber
        'current_area_title      ת������    0..1    S
        mclsXML.appendData "current_area_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        'current_dept_id     ת������id  1   N
        mclsXML.appendData "current_dept_id", Val(Nvl(mrsPatiInfo!��Ժ����id, 0)), xsNumber
        'current_dept_title      ת������    1   S
        mclsXML.appendData "current_dept_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        'current_room        ת������    0..1    S
        mclsXML.appendData "current_room", txt����.Tag, xsString
        'current_bed     ת������    1   S
        mclsXML.appendData "current_bed", Nvl(mrsPatiInfo!��Ҫ����), xsString
        mclsXML.AppendNode "current_state", True
        
        strSql = " Select ID �䶯ID,sysdate �䶯ʱ�� From ���˱䶯��¼  Where ����ID=[1] And ��ҳId=[2] And ��ʼԭ��=[3] and ��ʼʱ�� IS NULL"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���˱䶯��¼", mlng����ID, mlng��ҳID, 3)
        'ת����Ϣ
        'change_state        ת����Ϣ    1
        mclsXML.AppendNode "change_state"
        'change_id       ת�Ʊ��id  1   N
        mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
        'change_date     ���ʱ��    1   S
        mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
        'change_area_id      ת�벡��id  0..1    N
        'mclsXML.appendData "change_area_id", Val(Nvl(mrsPatiInfo!��ǰ����id)), xsNumber
        'change_area_title       ת�벡��    0..1    S
        'mclsXML.appendData "change_area_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        'change_dept_id      ת�����id  0..1    N
        mclsXML.appendData "change_dept_id", Val(cbo����.ItemData(cbo����.ListIndex)), xsNumber
        'change_dept_title       ת�����    0..1    S
        mclsXML.appendData "change_dept_title", zlCommFun.GetNeedName(cbo����.Text), xsString
        mclsXML.AppendNode "change_state", True

        mclsMipModule.CommitMessage "ZLHIS_PATIENT_003", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    
    '������ҽӿ�
    If CreatePlugInOK(glngModul) Then
        On Error Resume Next
        Call gobjPlugIn.InPatiCheckInBranchAfter(mlng����ID, mlng��ҳID)
        Call zlPlugInErrH(Err, "InPatiCheckInBranchAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function isValid() As Boolean
    
    Dim strSql As String
    Dim strInfo As String
    Dim lng����ID As Long, lngת�����ID As Long
    Dim blnSameUnit As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    lng����ID = mrsPatiInfo!��ס����id
    lngת�����ID = cbo����.ItemData(cbo����.ListIndex)
    
    If gbytת��ʱ���δִ�� <> 0 Then
        'ͬһ����֮��ת��,��ʾ������ֹ
        strSql = "Select Distinct (A.����id) ����id " & _
                 "From �������Ҷ�Ӧ A, �������Ҷ�Ӧ B " & _
                 "Where A.����id = B.����id And A.����id = [1] And B.����id = [2]"
        
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lngת�����ID)
        If rsTemp.RecordCount > 0 Then blnSameUnit = True
            
        strInfo = ExistWaitExe(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbytת��ʱ���δִ�� = 1 Or blnSameUnit = True Then
                If MsgBox("�ò��˴�����δִ����ɵ����ݣ�" & _
                    vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫת����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "�ò��˴�����δִ����ɵ����ݣ�" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "������ת��.", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '����30208 by lesfeng 2010-08-02 ���ֲ���22��32 ����154��155
    If gbytת��ʱ���ҩƷδִ�� <> 0 Then
        'ͬһ����֮��ת��,��ʾ������ֹ
        strSql = "Select Distinct (A.����id) ����id " & _
                 "From �������Ҷ�Ӧ A, �������Ҷ�Ӧ B " & _
                 "Where A.����id = B.����id And A.����id = [1] And B.����id = [2]"
        
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lngת�����ID)
        If rsTemp.RecordCount > 0 Then blnSameUnit = True
        
        strInfo = ExistWaitDrug(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbytת��ʱ���ҩƷδִ�� = 1 Or blnSameUnit = True Then
                If MsgBox("�ò���" & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫת����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "�ò���" & strInfo & vbCrLf & vbCrLf & "������ת�ơ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    
    '61429:������,2013-11-11,ת��ʱ����δ��˵��ݼ��
    If gbytת��ʱδ������ʵ��ݼ�� <> 0 Then
        'ͬһ����֮��ת��,��ʾ������ֹ
        blnSameUnit = False
        strSql = "Select Distinct (A.����id) ����id " & _
                 "From �������Ҷ�Ӧ A, �������Ҷ�Ӧ B " & _
                 "Where A.����id = B.����id And A.����id = [1] And B.����id = [2]"
        
        On Error GoTo errHandle
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lngת�����ID)
        If rsTemp.RecordCount > 0 Then blnSameUnit = True
        
        strInfo = ExistWaitQuittance(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbytת��ʱδ������ʵ��ݼ�� = 1 Or blnSameUnit = True Then
                If MsgBox("�ò���" & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫת����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "�ò���" & strInfo & vbCrLf & vbCrLf & "������ת�ơ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    isValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    isValid = False
End Function

Private Function LoadBed() As Boolean
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Integer, lng����ID As Long
    Dim byt�������� As Byte
    Dim strTmp As String, str���� As String, str����� As String

    '����27370 by lesfeng 2010-02-03
    mintFlag = 0
    
    On Error GoTo errH
    
    gblnOK = False
    Set mrsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID)
    'סԺ����
    With mrsPatiInfo
        txt����.Text = !����
        txt�Ա�.Text = "" & !�Ա�
        txt����.Text = "" & !����
        txtסԺ��.Text = "" & !סԺ��
        txt����.Text = !��ǰ����
        
        lng����ID = !��ס����id
        byt�������� = Val("" & !��������)
    End With
    
    Set rsTmp = GetPatiBeds(mlng����ID)
    If rsTmp.RecordCount = 0 Then
        str���� = "��ͥ����"
        str����� = ""
    Else
        Do While Not rsTmp.EOF
            str���� = str���� & "," & rsTmp!����
            If Nvl(rsTmp!����) = Nvl(mrsPatiInfo!��Ҫ����) And Nvl(rsTmp!����ID) = Nvl(mrsPatiInfo!��ס����id) Then
                str����� = Nvl(rsTmp!�����)
            End If
            rsTmp.MoveNext
        Loop
        str���� = Mid(str����, 2)
    End If
    txt����.Text = str����
    txt����.Tag = str�����
        
    'ȷ�������ķ������
    strSql = "Select ������� From ��������˵�� Where ��������='����' And ����ID=[1]"
     Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit)
    
    If rsTmp!������� = 1 Then
        strTmp = "1,3"
    ElseIf rsTmp!������� = 2 Then
        strTmp = "2,3"
    ElseIf rsTmp!������� = 3 Then
        If byt�������� = 1 Then
            strTmp = "1,3"
        Else
            strTmp = "2,3"
        End If
    End If
    '����27370 by lesfeng 2010-02-03
    mstr������� = strTmp
    
    '��ѡ����Ϊ�ٴ�����,û�д�λ��Ҳ�г�,��Ϊ����ʹ�ò����Ĺ��ô�
    Set rsTmp = GetDepts("�ٴ�", strTmp)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!ID <> lng����ID Then
                cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        cbo����.ListIndex = 0
    Else
        MsgBox "û���ҵ��뵱ǰ�������������ͬ���ٴ�����,�뵽���Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    LoadBed = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowMe(frmMain As Object, ByVal lngUnit As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
    
    
    Set mfrmParent = frmMain
    mlngUnit = lngUnit
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstrPrivs = strPrivs
    If LoadBed = False Then Exit Function
    Me.Show 1, frmMain
    
    ShowMe = gblnOK
End Function

Private Sub Form_Load()
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub
