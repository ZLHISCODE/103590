VERSION 5.00
Begin VB.Form frmClinicPlanOfficeAndUnitRegModify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ҵ���"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9375
   Icon            =   "frmClinicPlanOfficeAndUnitRegModify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8130
      TabIndex        =   32
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8130
      TabIndex        =   33
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   8130
      TabIndex        =   34
      Top             =   6210
      Width           =   1100
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "������Ϣ"
      Height          =   1065
      Left            =   30
      TabIndex        =   14
      Top             =   1560
      Width           =   7965
      Begin VB.TextBox txt��Լ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6720
         TabIndex        =   21
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txt�޺��� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4800
         TabIndex        =   20
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txt�ϰ�ʱ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   18
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txt�������� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   16
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txt����ҽ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   23
         Top             =   675
         Width           =   1065
      End
      Begin VB.CheckBox chk��ſ��� 
         Caption         =   "������ſ���"
         Enabled         =   0   'False
         Height          =   225
         Left            =   4800
         TabIndex        =   26
         Top             =   713
         Width           =   1395
      End
      Begin VB.CheckBox chkʱ�� 
         Caption         =   "����ʱ��"
         Enabled         =   0   'False
         Height          =   225
         Left            =   6690
         TabIndex        =   27
         Top             =   713
         Width           =   1035
      End
      Begin VB.TextBox txtԤԼ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   25
         Top             =   675
         Width           =   1605
      End
      Begin VB.Label lbl��Լ�� 
         AutoSize        =   -1  'True
         Caption         =   "��Լ��"
         Height          =   180
         Left            =   6120
         TabIndex        =   28
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lbl�޺��� 
         AutoSize        =   -1  'True
         Caption         =   "�޺���"
         Height          =   180
         Left            =   4230
         TabIndex        =   19
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl����ҽ�� 
         AutoSize        =   -1  'True
         Caption         =   "����ҽ��"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   735
         Width           =   720
      End
      Begin VB.Label lbl�ϰ�ʱ�� 
         AutoSize        =   -1  'True
         Caption         =   "�ϰ�ʱ��"
         Height          =   180
         Left            =   2130
         TabIndex        =   17
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblԤԼ���� 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ����"
         Height          =   180
         Left            =   2130
         TabIndex        =   24
         Top             =   735
         Width           =   720
      End
   End
   Begin VB.Frame fraӦ������ 
      Caption         =   "Ӧ������"
      Height          =   4125
      Left            =   30
      TabIndex        =   29
      Top             =   2730
      Width           =   7965
      Begin zl9RegEvent.ClinicPlanOffice cpoRoom 
         Height          =   3855
         Left            =   60
         TabIndex        =   30
         Top             =   210
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   6800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin zl9RegEvent.ClinicPlanUnit cpuUnit 
         Height          =   3855
         Left            =   60
         TabIndex        =   31
         Top             =   210
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   6800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fra��Դ��Ϣ 
      Caption         =   "��Դ������Ϣ"
      Height          =   1395
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7965
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3210
         TabIndex        =   4
         Top             =   285
         Width           =   1275
      End
      Begin VB.TextBox txtSignalNO 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   2
         Top             =   285
         Width           =   1035
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�Һ�ʱ���뽨��"
         Enabled         =   0   'False
         Height          =   180
         Left            =   5160
         TabIndex        =   13
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txt���տ��� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   12
         Top             =   1020
         Width           =   1935
      End
      Begin VB.TextBox txtDoctor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5160
         TabIndex        =   10
         Top             =   652
         Width           =   2625
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5160
         TabIndex        =   6
         Top             =   285
         Width           =   2625
      End
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   300
         Left            =   870
         TabIndex        =   8
         Top             =   652
         Width           =   3615
      End
      Begin VB.Label lblSignalNO 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   2820
         TabIndex        =   3
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ"
         Height          =   180
         Left            =   480
         TabIndex        =   7
         Top             =   712
         Width           =   360
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   4770
         TabIndex        =   5
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         Height          =   180
         Left            =   4770
         TabIndex        =   9
         Top             =   705
         Width           =   360
      End
      Begin VB.Label lbl���տ��� 
         AutoSize        =   -1  'True
         Caption         =   "���տ���"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmClinicPlanOfficeAndUnitRegModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As Byte '1-������������,2-������λ�Һſ���
Private mobj��Դ As �����Դ, mobj�����¼ As �����¼
Private mblnRecord As Boolean '�Ƿ��������¼
Private mblnFirst As Boolean

Private mblnOk As Boolean

Public Function ShowMe(frmParent As Form, ByVal bytFun As Byte, _
    ByVal obj��Դ As �����Դ, ByVal obj�����¼ As �����¼, Optional ByVal blnRecord As Boolean) As Boolean
    '�������
    '������
    '   bytFun 1-������������,2-������λ�Һſ���
    If obj��Դ Is Nothing Then Exit Function
    If obj�����¼ Is Nothing Then Exit Function
    
    mbytFun = bytFun
    Set mobj��Դ = obj��Դ: Set mobj�����¼ = obj�����¼
    mblnRecord = blnRecord
    
    On Error Resume Next
    If CheckDepend() = False Then Exit Function
    mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Function CheckDepend() As Boolean
    '����:�������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    '���ܶ���ʷ�İ��Ž��в���
    If DateDiff("s", mobj�����¼.��ֹʱ��, zlDatabase.Currentdate) >= 0 Then
        MsgBox "��ǰϵͳʱ���Ѵ����˰���ʱ�ε���ֹʱ�䣬���ܽ���" & IIf(mbytFun = 1, "��������", "������λ�Һſ���") & "����������", vbInformation, gstrSysName
        Exit Function
    End If
    '�Ѿ�ͣ���δ���ﰲ�ŵģ����������
    strSQL = "Select 1 from �ٴ������¼ Where ID=[1] and �ϰ�ʱ��=[2] And ͣ�￪ʼʱ�� Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������¼", mobj�����¼.��¼ID, mobj�����¼.ʱ���)
    If rsTemp.EOF Then
        MsgBox "��ǰ����ʱ�β����ڻ���ͣ����ܽ���" & IIf(mbytFun = 1, "��������", "������λ�Һſ���") & "����������", vbInformation, gstrSysName
        Exit Function
    End If
    CheckDepend = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Err = 0: On Error GoTo errHandler
    If mbytFun = 1 Then '��������
        If cpoRoom.IsValied() = False Then Exit Sub
    ElseIf mbytFun = 2 Then '������λ����
        If cpuUnit.IsValied() = False Then Exit Sub
    End If
    
    If SaveData() = False Then Exit Sub
    mblnOk = True
    Unload Me
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function InitData() As Boolean
    Dim i As Integer
'    Dim obj���з������� As �������Ҽ�
    Dim obj���к�����Ϣ�� As ������Ϣ��, obj���к�����λ As ������λ���Ƽ�
    
    Err = 0: On Error GoTo errHandler
    cpoRoom.Visible = False
    cpuUnit.Visible = False
    If mbytFun = 1 Then '��������
        cpoRoom.Visible = True
        Me.Caption = "�������ҵ���"
        fraӦ������.Caption = "��������"
    Else '������λ����
        cpuUnit.Visible = True
        Me.Caption = "������λ�Һſ��Ƶ���"
        fraӦ������.Caption = "������λ�Һſ���"
    End If
    
    '��Դ��Ϣ
    txtSignalNO.Text = mobj��Դ.����
    txt����.Text = mobj��Դ.����
    txtDept.Text = mobj��Դ.��������
    txtItem.Text = mobj��Դ.��Ŀ����
    txtDoctor.Text = mobj��Դ.ҽ������
    txt���տ���.Text = Decode(mobj��Դ.���տ���״̬, 1, "����ԤԼ", 2, "��ֹԤԼ", 3, "�ܽڼ������ÿ���", "���ϰ�")
    chk����.Value = IIf(mobj��Դ.�Ƿ񽨲���, vbChecked, vbUnchecked)
    If IsDate(mobj�����¼.��������) Then
        txt��������.Text = Format(mobj�����¼.��������, "yyyy-mm-dd")
    Else
        txt��������.Text = mobj�����¼.��������
    End If
    
    txt�ϰ�ʱ��.Text = mobj�����¼.ʱ���
    txt����ҽ��.Text = mobj�����¼.����ҽ��
    txtԤԼ����.Text = Choose(mobj�����¼.ԤԼ���� + 1, "����ԤԼ", "��ֹԤԼ", "����ֹ��������ԤԼ")
    chk��ſ���.Value = IIf(mobj�����¼.�Ƿ���ſ���, vbChecked, vbUnchecked)
    chkʱ��.Value = IIf(mobj�����¼.�Ƿ��ʱ��, vbChecked, vbUnchecked)
    txt�޺���.Text = IIf(mobj�����¼.�޺��� = 0, "", mobj�����¼.�޺���)
    txt��Լ��.Text = IIf(mobj�����¼.��Լ�� = 0, "", mobj�����¼.��Լ��)
    
    If mbytFun = 1 Then '��������
'        Set obj���з������� = GetVisitRoomsObjects(GetDoctorRooms(mobj�����¼.����ID))
'        obj���з�������.���﷽ʽ = mobj�����¼.���﷽ʽ
'        cpoRoom.LoadData mobj�����¼.�����������Ҽ�, obj���з�������
    Else '������λ����
        Set obj���к�����Ϣ�� = GetTimeIntervalObjects(GetTimeInterval(mobj�����¼.��¼ID, True))
        With obj���к�����Ϣ��
            .�Ƿ��ʱ�� = mobj�����¼.�Ƿ��ʱ��
            .�Ƿ���ſ��� = mobj�����¼.�Ƿ���ſ���
            .�޺��� = mobj�����¼.�޺���
            .��Լ�� = mobj�����¼.��Լ��
            .ԤԼ���� = mobj�����¼.ԤԼ����
        End With
        Set obj���к�����λ = GetUnitsObjects(GetUnitAll())
        cpuUnit.LoadData mobj�����¼.������λ���Ƽ�, obj���к�����Ϣ��, obj���к�����λ
    End If
    InitData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Activate()
    Dim obj���з������� As �������Ҽ�
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = True
    
    If mbytFun = 1 Then '��������
        '�������������ΪListView�ؼ���ԭ�������Load�¼��м������ݣ��ᵼ��������ʾ��ȫ�������ʡ�Ժ�
        Set obj���з������� = GetVisitRoomsObjects(GetDoctorRooms(mobj�����¼.����ID))
        obj���з�������.���﷽ʽ = mobj�����¼.���﷽ʽ
        cpoRoom.LoadData mobj�����¼.�����������Ҽ�, obj���з�������
    End If
    If cpoRoom.Visible And cpoRoom.EditMode = ED_RegistPlan_Edit Then cpoRoom.SetFocus
    If cpuUnit.Visible And cpuUnit.EditMode = ED_RegistPlan_Edit Then cpuUnit.SetFocus
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    
    mblnFirst = True
    If mbytFun = 2 Then
        If mobj�����¼.ԤԼ���� = 1 Then
            '��ֹԤԼ
            MsgBox "��ǰ����Ϊ��ֹԤԼ�����ܵ���������λ���ƣ�", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    Call InitData
    Call SetEnabledBackColor(Me.Controls)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, cllPro As New Collection, i As Integer
    Dim byt���﷽ʽ As Byte, str���� As String, obj���� As ��������
    Dim obj������λ As ������λ����, obj���� As ������Ϣ
    Dim cll���� As Collection, str���� As String, strTemp As String
    Dim blnTrans As Boolean
    Dim lng�䶯ID As Long
    
    Err = 0: On Error GoTo errHandler
    If mbytFun = 1 Then '��������
        Set mobj�����¼.�����������Ҽ� = cpoRoom.Get�����������Ҽ�
        '��������
        byt���﷽ʽ = mobj�����¼.�����������Ҽ�.���﷽ʽ
        str���� = ""
        For Each obj���� In mobj�����¼.�����������Ҽ�
            '����_In:����1,����2,...
            str���� = str���� & "," & obj����.����ID
        Next
        If str���� <> "" Then str���� = Mid(str����, 2)
        
        'Zl_�ٴ���������_Update(
        strSQL = "Zl_�ٴ���������_Update("
        'Id_In       �ٴ���������.Id%Type,
        strSQL = strSQL & "" & mobj�����¼.��¼ID & ","
        '���﷽ʽ_In �ٴ���������.���﷽ʽ%Type := Null,
        strSQL = strSQL & "" & byt���﷽ʽ & ","
        '����_In     Varchar2 := Null,
        strSQL = strSQL & "'" & str���� & "',"
        '�����¼_In Number:=0--�Ƿ��ǶԳ����¼����ɾ��
        strSQL = strSQL & "" & IIf(mblnRecord, 1, 0) & ")"
        cllPro.Add strSQL
    Else '������λ����
        Set mobj�����¼.������λ���Ƽ� = cpuUnit.Get������λ������Ϣ��
        If mblnRecord Then
            lng�䶯ID = zlDatabase.GetNextId("�ٴ�����䶯��¼")
            'Zl_�ٴ�����ԤԼ���Ʊ䶯(
            strSQL = "Zl_�ٴ�����ԤԼ���Ʊ䶯("
            '�䶯����_In   �ٴ�����䶯��ϸ.�䶯����%Type,
            strSQL = strSQL & "" & 1 & ","
            'Id_In         �ٴ�����䶯��¼.Id%Type,
            strSQL = strSQL & "" & lng�䶯ID & ","
            '��¼id_In     �ٴ�����䶯��¼.��¼id%Type := Null,
            strSQL = strSQL & "" & mobj�����¼.��¼ID & ","
            '��ԤԼ����_In �ٴ�����䶯��¼.��ԤԼ����%Type := Null
            strSQL = strSQL & "" & "NULL" & ")"
            cllPro.Add strSQL
        End If
        For Each obj������λ In mobj�����¼.������λ���Ƽ�
            'ԤԼ����:0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
            '����:1-��������;2-ԤԼ��ʽ
            Set cll���� = New Collection
            str���� = ""
            For Each obj���� In obj������λ.������Ϣ��
                strTemp = obj����.��� & "," & obj����.����
                If zlCommFun.ActualLen(str���� & "|" & strTemp) > 2000 Then
                    '���ſ���_in:���1,����|���2,����|...
                    str���� = Mid(str����, 2)
                    cll����.Add str����
                    str���� = ""
                End If
                str���� = str���� & "|" & strTemp
            Next
            If str���� <> "" Then
                str���� = Mid(str����, 2)
                cll����.Add str����
            End If
            For i = 1 To IIf(cll����.Count = 0, 1, cll����.Count)
                If mblnRecord Then
                    'Zl_�ٴ�����Һſ��Ƽ�¼_Insert(
                    strSQL = "Zl_�ٴ�����Һſ��Ƽ�¼_Insert("
                    '��¼id_In   �ٴ�����Һſ��Ƽ�¼.��¼id%Type,
                    strSQL = strSQL & "" & mobj�����¼.��¼ID & ","
                    '����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
                    strSQL = strSQL & "" & obj������λ.���� & ","
                    '����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
                    strSQL = strSQL & "" & 1 & ","
                    '����_In     �ٴ�����Һſ��Ƽ�¼.����%Type,
                    strSQL = strSQL & "'" & obj������λ.������λ���� & "',"
                    '���Ʒ�ʽ_In �ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ%Type,
                    strSQL = strSQL & "" & obj������λ.ԤԼ���Ʒ�ʽ & ","
                    '�Ƿ��ռ_In �ٴ������¼.�Ƿ��ռ%Type,
                    strSQL = strSQL & "" & IIf(mobj�����¼.������λ���Ƽ�.�Ƿ��ռ, 1, 0) & ","
                    '���ſ���_In Varchar2,
                    str���� = ""
                    If cll����.Count > 0 Then str���� = cll����(i)
                    strSQL = strSQL & "'" & str���� & "',"
                    'ɾ��_In Number:=0
                    strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                    cllPro.Add strSQL
                Else
                    'Zl_�ٴ�����Һſ���_Insert(
                    strSQL = "Zl_�ٴ�����Һſ���_Insert("
                    '����id_In   �ٴ�����Һſ���.����id%Type,
                    strSQL = strSQL & "" & mobj�����¼.��¼ID & ","
                    '����_In     �ٴ�����Һſ���.����%Type,
                    strSQL = strSQL & "" & obj������λ.���� & ","
                    '����_In     �ٴ�����Һſ���.����%Type,
                    strSQL = strSQL & "" & 1 & ","
                    '����_In     �ٴ�����Һſ���.����%Type,
                    strSQL = strSQL & "'" & obj������λ.������λ���� & "',"
                    '���Ʒ�ʽ_In �ٴ�����Һſ���.���Ʒ�ʽ%Type,
                    strSQL = strSQL & "" & obj������λ.ԤԼ���Ʒ�ʽ & ","
                    '�Ƿ��ռ_In �ٴ���������.�Ƿ��ռ%Type,
                    strSQL = strSQL & "" & IIf(mobj�����¼.������λ���Ƽ�.�Ƿ��ռ, 1, 0) & ","
                    '���ſ���_In Varchar2,
                    str���� = ""
                    If cll����.Count > 0 Then str���� = cll����(i)
                    strSQL = strSQL & "'" & str���� & "',"
                    'ɾ��_In Number:=0
                    strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
                    cllPro.Add strSQL
                End If
            Next
        Next
        
        'Zl_�ٴ�����ԤԼ���Ʊ䶯(
        strSQL = "Zl_�ٴ�����ԤԼ���Ʊ䶯("
        '�䶯����_In   �ٴ�����䶯��ϸ.�䶯����%Type,
        strSQL = strSQL & "" & 2 & ","
        'Id_In         �ٴ�����䶯��¼.Id%Type,
        strSQL = strSQL & "" & lng�䶯ID & ","
        '��¼id_In     �ٴ�����䶯��¼.��¼id%Type := Null,
        strSQL = strSQL & "" & mobj�����¼.��¼ID & ")"
        '��ԤԼ����_In �ٴ�����䶯��¼.��ԤԼ����%Type := Null
        cllPro.Add strSQL
    End If
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    blnTrans = False
    SaveData = True
    Exit Function
errHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
