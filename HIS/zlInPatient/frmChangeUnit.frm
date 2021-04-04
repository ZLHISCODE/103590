VERSION 5.00
Begin VB.Form frmChangeUnit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ת����"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmChangeUnit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1395
         Width           =   3810
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   225
         Width           =   675
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4095
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   225
         Width           =   690
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   615
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   615
         Width           =   1605
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1005
         Width           =   3795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   675
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ת�벡��"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   540
         TabIndex        =   15
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2370
         TabIndex        =   14
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3690
         TabIndex        =   13
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ��λ"
         Height          =   180
         Left            =   2370
         TabIndex        =   12
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1065
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2715
      TabIndex        =   1
      Top             =   2115
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3915
      TabIndex        =   2
      Top             =   2115
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   270
      TabIndex        =   3
      Top             =   2115
      Width           =   1100
   End
End
Attribute VB_Name = "frmChangeUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrPrivs As String
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlngUnit As Long
Private mrsPatiInfo As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Function InitData() As Boolean
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset
    Dim str���� As String, str����� As String
    
    On Error GoTo errHandle
    
    Set mrsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID)
    
    With mrsPatiInfo
        txt����.Text = !����
        txt�Ա�.Text = "" & !�Ա�
        txt����.Text = "" & !����
        txtסԺ��.Text = "" & !סԺ��
    End With
    
    str����� = ""
    Set rsTmp = GetPatiBeds(mlng����ID)
    If rsTmp.RecordCount = 0 Then
        str���� = "��ͥ����"
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
    
    'Ŀǰ��������۲���
    
    gstrSQL = "Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,�������Ҷ�Ӧ B,��������˵�� C " & _
            " Where B.����ID=A.ID And B.����ID=[1] " & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " And C.����ID=A.ID And Instr(',' || [2]|| ',',',' || C.������� || ',')>0 " & _
            " And C.��������='����' " & _
            " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlInPatient", Val("" & mrsPatiInfo!��Ժ����id), "1,2,3")
    'Set rsTmp = GetDeptOrUnit(1, mrsPatiInfo!��Ժ����ID, "1,2,3")
    If Not rsTmp.EOF Then
        cboUnit.Clear
        For i = 1 To rsTmp.RecordCount
            
            If rsTmp!ID = mlngUnit Then
                txt����.Text = rsTmp!����
            Else
                
                cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
                cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            End If
            rsTmp.MoveNext
        Next
        If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0  '������InitBed���ô�λ
    End If
    
    If cboUnit.ListCount = 0 Then
        MsgBox "�ò������ڿ���û������������Ӧ�Ĳ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function isValid() As Boolean
    
    Dim strSql As String
    Dim strInfo As String
    Dim rsTemp As New ADODB.Recordset

    If gbytת��ʱ���δִ�� <> 0 Then

        strInfo = ExistWaitExe(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbytת��ʱ���δִ�� = 1 Then
                If MsgBox("�ò��˴�����δִ����ɵ����ݣ�" & _
                    vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫת������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "�ò��˴�����δִ����ɵ����ݣ�" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "������ת������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If

    If gbytת��ʱ���ҩƷδִ�� <> 0 Then
        strInfo = ExistWaitDrug(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbytת��ʱ���ҩƷδִ�� = 1 Then
                If MsgBox("�ò���" & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫת������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "�ò���" & strInfo & vbCrLf & vbCrLf & "������ת������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '61429:������,2013-11-11,ת��ʱ����δ��˵��ݼ��
    If gbytת��ʱδ������ʵ��ݼ�� <> 0 Then
        strInfo = ""
        strInfo = ExistWaitQuittance(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbytת��ʱδ������ʵ��ݼ�� = 1 Then
                If MsgBox("�ò���" & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫת������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "�ò���" & strInfo & vbCrLf & vbCrLf & "������ת������", vbInformation, gstrSysName
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

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
        ByVal strPrivs As String) As Boolean
'#########################################################################################################
'### ������
'### ���أ�Ŀ�괲��
'#########################################################################################################
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstrPrivs = strPrivs
    
    If InitData = False Then Exit Function
    
    Me.Show 1, frmParent
    
    ShowMe = gblnOK
End Function


Private Sub cmdOK_Click()
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQLtmp As String, rsPati As Recordset
    
    On Error GoTo errH
    
    If isValid = False Then Exit Sub
    
    'ת�������ü��
    If CreatePublicExpenseBillOperation() And gblnת����ת���� Then
        strSQLtmp = "Select ID, ����id" & vbNewLine & _
                    "From ���˱䶯��¼" & vbNewLine & _
                    "Where ����id = [1] And ��ҳid = [2] And ��ʼʱ�� Is Not Null And ��ֹʱ�� Is Null And NVL(���Ӵ�λ,0) = 0"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQLtmp, Me.Caption, mlng����ID, mlng��ҳID)
        If rsPati.RecordCount > 0 Then
            If gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(Me, 2, mlng����ID, mlng��ҳID, Val(rsPati!ID & ""), Val(rsPati!����ID & ""), cboUnit.ItemData(cboUnit.ListIndex)) = False Then Exit Sub
        End If
    End If
    
    strSql = "zl_���˱䶯��¼_ChangeUnit(" & mlng����ID & "," & mlng��ҳID & "," & _
        cboUnit.ItemData(cboUnit.ListIndex) & ",'" & UserInfo.��� & "'," & "'" & UserInfo.���� & "')"
        
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    '����96847��118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng����ID, mlng��ҳID) <> 1 Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    gblnOK = True
    
    On Error Resume Next
    'ת�����ɹ��󴥷���Ϣ
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���˱䶯��¼", mlng����ID, mlng��ҳID, 15)
        'ת����Ϣ
        'change_state        ת����Ϣ    1
        mclsXML.AppendNode "change_state"
        'change_id       ת�Ʊ��id  1   N
        mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
        'change_date     ���ʱ��    1   S
        mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
        
        'change_area_id      ת�벡��id  0..1    N
        mclsXML.appendData "change_area_id", Val(cboUnit.ItemData(cboUnit.ListIndex)), xsNumber
        'change_area_title       ת�벡��    0..1    S
        mclsXML.appendData "change_area_title", zlCommFun.GetNeedName(cboUnit.Text), xsString
        'change_dept_id      ת�����id  0..1    N
        mclsXML.appendData "change_dept_id", Val(Nvl(mrsPatiInfo!��Ժ����id, 0)), xsNumber
        'change_dept_title       ת�����    0..1    S
        mclsXML.appendData "change_dept_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        mclsXML.AppendNode "change_state", True

        mclsMipModule.CommitMessage "ZLHIS_PATIENT_003", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

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
