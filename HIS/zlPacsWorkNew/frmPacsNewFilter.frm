VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPacsNewFilter 
   Caption         =   "���ݲ�ѯ"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPacsNewFilter.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   7710
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.Slider sldDays 
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   529
      _Version        =   393216
      Max             =   180
      TickFrequency   =   7
   End
   Begin VB.ComboBox cboAgeType 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":000C
      Left            =   6720
      List            =   "frmPacsNewFilter.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox cboAgeWhere 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":0030
      Left            =   5280
      List            =   "frmPacsNewFilter.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtEndAge 
      Height          =   300
      Left            =   6120
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.ComboBox cboQueryTime 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":0074
      Left            =   1320
      List            =   "frmPacsNewFilter.frx":0081
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox cboPatientFrom 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":00A3
      Left            =   5400
      List            =   "frmPacsNewFilter.frx":00B6
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame fraControl 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   52
      Top             =   6240
      Width           =   7215
      Begin VB.CommandButton cmdCustomQuery 
         Caption         =   "�Զ����ѯ(&C)"
         Height          =   375
         Left            =   2760
         TabIndex        =   36
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "ȷ  ��(&O)"
         Height          =   375
         Left            =   4440
         TabIndex        =   34
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "�� ��(&Q)"
         Height          =   375
         Left            =   5760
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkInputReportInf 
         Caption         =   "��������"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveSchema 
         Caption         =   "���淽��(&S)"
         Height          =   375
         Left            =   5400
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelSchema 
         Caption         =   "ɾ������(&D)"
         Height          =   375
         Left            =   3720
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboSchemaName 
         Height          =   330
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label labSchema 
         Caption         =   "��ѯ������"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   1095
      End
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
      Height          =   3255
      Left            =   240
      TabIndex        =   42
      Top             =   2880
      Width           =   7215
      Begin VB.ComboBox cboYangXingLv 
         Height          =   330
         ItemData        =   "frmPacsNewFilter.frx":00D4
         Left            =   4800
         List            =   "frmPacsNewFilter.frx":00E2
         TabIndex        =   21
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cboStudyDoctor 
         Height          =   330
         Left            =   4800
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboBodyPart 
         Height          =   330
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboImageType 
         Height          =   330
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboDevice 
         Height          =   330
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboPatientRoom 
         Height          =   330
         Left            =   1320
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboProcedure 
         Height          =   330
         ItemData        =   "frmPacsNewFilter.frx":00FC
         Left            =   4800
         List            =   "frmPacsNewFilter.frx":011E
         TabIndex        =   19
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   4800
         TabIndex        =   29
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   1320
         TabIndex        =   28
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txt������� 
         Height          =   300
         Left            =   4800
         TabIndex        =   27
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtReportContext 
         Height          =   300
         Left            =   1320
         TabIndex        =   26
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   4800
         TabIndex        =   25
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtIllnessRes 
         Height          =   300
         Left            =   1320
         TabIndex        =   24
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox cboQuality 
         Height          =   330
         ItemData        =   "frmPacsNewFilter.frx":016C
         Left            =   1320
         List            =   "frmPacsNewFilter.frx":0179
         TabIndex        =   20
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cboDiagnoseDoctor 
         Height          =   330
         Left            =   4800
         TabIndex        =   23
         Top             =   1680
         Width           =   2175
      End
      Begin VB.ComboBox cboAuditingDoctor 
         Height          =   330
         Left            =   1320
         TabIndex        =   22
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label labFilter 
         Caption         =   "�� �� �ԣ�"
         Height          =   255
         Index           =   22
         Left            =   3720
         TabIndex        =   62
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "��鲿λ��"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "��鼼ʦ��"
         Height          =   255
         Index           =   7
         Left            =   3720
         TabIndex        =   60
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "Ӱ�����"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   59
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "����豸��"
         Height          =   255
         Index           =   9
         Left            =   3720
         TabIndex        =   58
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "���˿��ң�"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   57
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "�����̣�"
         Height          =   255
         Index           =   11
         Left            =   3720
         TabIndex        =   56
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "��    �飺"
         Height          =   255
         Index           =   20
         Left            =   3720
         TabIndex        =   51
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "��������"
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   50
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "Ӱ��������"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   49
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "���ҽ����"
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   48
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "���ҽ����"
         Height          =   255
         Index           =   16
         Left            =   3720
         TabIndex        =   47
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "������ϣ�"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   46
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "��    �ã�"
         Height          =   255
         Index           =   14
         Left            =   3720
         TabIndex        =   45
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "�������ݣ�"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   44
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label labFilter 
         Caption         =   "���������"
         Height          =   255
         Index           =   12
         Left            =   3720
         TabIndex        =   43
         Top             =   2520
         Width           =   1095
      End
   End
   Begin VB.ComboBox cboNumType 
      Height          =   330
      ItemData        =   "frmPacsNewFilter.frx":0187
      Left            =   1320
      List            =   "frmPacsNewFilter.frx":01A3
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpBirthDay 
      Height          =   420
      Left            =   4680
      TabIndex        =   9
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   63438849
      CurrentDate     =   40372
   End
   Begin VB.TextBox txtStartAge 
      Height          =   300
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.ComboBox cboSex 
      Height          =   330
      Left            =   1320
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtNum 
      Height          =   300
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   330
      Left            =   5400
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   63438851
      CurrentDate     =   38082
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   330
      Left            =   3000
      TabIndex        =   11
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   63438851
      CurrentDate     =   40372
   End
   Begin VB.Label labDays 
      Alignment       =   2  'Center
      Caption         =   "����"
      Height          =   255
      Left            =   840
      TabIndex        =   66
      Top             =   2640
      Width           =   5895
   End
   Begin VB.Label Label 
      Caption         =   "������"
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   65
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "����"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   64
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   63
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label labFilter 
      Caption         =   "��ѯʱ�䣺"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   55
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "������Դ��"
      Height          =   255
      Index           =   21
      Left            =   4320
      TabIndex        =   54
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "�������ڣ�"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   41
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "��    �䣺"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   40
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "��    ��"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   39
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "��    ����"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   38
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label labFilter 
      Caption         =   "��ѯ���룺"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   37
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmPacsNewFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDepartmentId As Long   '���浱ǰ�����Ĳ���ID
Private marrParValue(25) As String '������������õĲ���ֵ
Private mblnOk As Boolean


Public index_��ѯ���� As Integer
Public index_������Դ As Integer
Public index_�������� As Integer
Public index_��ʼ���� As Integer
Public index_�������� As Integer
Public index_�����Ա� As Integer
Public index_�������� As Integer
Public index_��ѯ��ʼʱ�� As Integer
Public index_��ѯ����ʱ�� As Integer
Public index_��鲿λ As Integer
Public index_��鼼ʦ As Integer
Public index_Ӱ����� As Integer
Public index_����豸 As Integer
Public index_���˿��� As Integer
Public index_������ As Integer
Public index_Ӱ������ As Integer
Public index_������ As Integer
Public index_���ҽ�� As Integer
Public index_���ҽ�� As Integer
Public index_������� As Integer
Public index_��� As Integer
Public index_�������� As Integer
Public index_������� As Integer
Public index_������ As Integer
Public index_���� As Integer



Public Function ShowFilter(ByVal lngDepartmentId, arrParameter() As String, owner As Form) As String
    mblnOk = False
    mDepartmentId = lngDepartmentId
    
    Call SetParIndex
    
    Me.Show 1, owner
    
    If mblnOk Then
        Call SetParValue(arrParameter)
        ShowFilter = GetQueryFilter()
    End If
End Function

Private Sub SetParIndex()
'********************************************
'
'��ʼ����������ȡֵ
'
'********************************************
    On Error GoTo errHandle
        index_��ѯ���� = 1
        index_������Դ = 2
        index_�������� = 3
        index_��ʼ���� = 4
        index_�������� = 5
        index_�����Ա� = 6
        index_�������� = 7
        index_��ѯ��ʼʱ�� = 8
        index_��ѯ����ʱ�� = 9
        index_��鲿λ = 10
        index_��鼼ʦ = 11
        index_Ӱ����� = 12
        index_����豸 = 13
        index_���˿��� = 14
        index_������ = 15
        index_Ӱ������ = 16
        index_������ = 17
        index_���ҽ�� = 18
        index_���ҽ�� = 19
        index_������� = 20
        index_��� = 21
        index_�������� = 22
        index_������� = 23
        index_������ = 24
        index_���� = 25
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub SetParValue(arrParValue() As String)
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrParValue(index_��ѯ����) = txtNum.Text
    
    'ȡ�ò�����Դ
    'arrParValue(index_������Դ) = cboPatientFrom.Text
    Select Case cboPatientFrom.Text
        Case "����"
            arrParValue(index_������Դ) = 1
        Case "סԺ"
            arrParValue(index_������Դ) = 2
        Case "����"
            arrParValue(index_������Դ) = 3
        Case "���"
            arrParValue(index_������Դ) = 4
    End Select
    
    arrParValue(index_��������) = txtName.Text
    
    '��Ҫ����������ת����ָ��������
    'arrParValue(index_��ʼ����) = txtStartAge.Text & cboAgeType.Text
    'arrParValue(index_��������) = txtEndAge.Text & cboAgeType.Text
    Select Case cboAgeType.Text
        Case "��"
            arrParValue(index_��ʼ����) = Val(txtStartAge.Text) * 365
            arrParValue(index_��������) = Val(txtEndAge.Text) * 365
        Case "��"
            arrParValue(index_��ʼ����) = Val(txtStartAge.Text) * 30
            arrParValue(index_��������) = Val(txtEndAge.Text) * 30
        Case "��"
            arrParValue(index_��ʼ����) = Val(txtStartAge.Text) * 7
            arrParValue(index_��������) = Val(txtEndAge.Text) * 7
        Case "��"
            arrParValue(index_��ʼ����) = Val(txtStartAge.Text) * 1
            arrParValue(index_��������) = Val(txtEndAge.Text) * 1
    End Select
    
    
    arrParValue(index_�����Ա�) = cboSex.Text
    
    If Not dtpBirthDay Is Nothing Then
        arrParValue(index_��������) = dtpBirthDay.value
    End If
    
    arrParValue(index_��ѯ��ʼʱ��) = dtpBegin.value
    arrParValue(index_��ѯ����ʱ��) = dtpEnd.value
    arrParValue(index_��鲿λ) = cboBodyPart.Text
    arrParValue(index_��鼼ʦ) = cboStudyDoctor.Text
    arrParValue(index_Ӱ�����) = cboImageType.Text
    
    If Trim(cboDevice.Text <> "") Then arrParValue(index_����豸) = cboDevice.ItemData(cboDevice.ListIndex) '�������豸��
    If Trim(cboPatientRoom.Text <> "") Then arrParValue(index_���˿���) = cboPatientRoom.ItemData(cboPatientRoom.ListIndex) '���没�˵Ŀ���ID
    
    arrParValue(index_������) = cboProcedure.Text
    arrParValue(index_Ӱ������) = cboQuality.Text
    
    '�ڴ���������ʱ��0��ʾ���ԣ�1��ʾ����
    If cboYangXingLv.Text = "�������" Then
        arrParValue(index_������) = 1
    Else
        arrParValue(index_������) = 0
    End If
    
    arrParValue(index_���ҽ��) = cboAuditingDoctor.Text
    arrParValue(index_���ҽ��) = cboDiagnoseDoctor.Text
    arrParValue(index_�������) = txtIllnessRes.Text
    arrParValue(index_���) = txt���.Text
    arrParValue(index_��������) = txtReportContext.Text
    arrParValue(index_�������) = txt�������.Text
    arrParValue(index_������) = txt������.Text
    arrParValue(index_����) = txt����.Text
           
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetQueryFilter() As String
    Dim strFilter As String
    Dim strSubFilter As String
    Dim strQueryField As String
    
    On Error GoTo errHandle
    
    '��ѯ����
    If Trim(txtNum.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strQueryField = GetQueryNumField(cboNumType.Text)
        
        If strQueryField <> "Ӱ��걾����ȡ��.�����" Then
            strFilter = strFilter & GetQueryNumField(cboNumType.Text) & "=[" & index_��ѯ���� & "]"
        Else
            strFilter = strFilter & "����ҽ����¼.������ĿID IN (select a.������Ŀ from Ӱ��걾����ȡ�� a where a.�����=[" & index_��ѯ���� & "])"
        End If
    End If
    
    '������Դ
    If Trim(cboPatientFrom.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "����ҽ����¼.������Դ=[" & index_������Դ & "]"
    End If
    
    '��������
    If Trim(txtName.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "������Ϣ.����=[" & index_�������� & "]"
    End If
    
    '��������-��ʼ����(ֻ�е�����ʹ�á����������ڶ�������֮��ʱ����ʹ�ÿ�ʼ����)
    If Trim(txtStartAge.Text) <> "" Then
        If cboAgeWhere.Text = "��" Then
            If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "AgeToDays(������Ϣ.����)>=[" & index_��ʼ���� & "]"
        End If
    End If
    
    '��������-��������
    If Trim(txtEndAge.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        If cboAgeWhere.Text = "��" Then
            strFilter = strFilter & "AgeToDays(������Ϣ.����)<=[" & index_�������� & "]"
        Else
            strFilter = strFilter & "AgeToDays(������Ϣ.����)" & GetQueryAgeWhere(cboAgeWhere.Text) & "[" & index_�������� & "]"
        End If
    End If
    
    '�����Ա�
    If Trim(cboSex.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "������Ϣ.�Ա�=[" & index_�����Ա� & "]"
    End If
    
    '��������
    If Trim(dtpBirthDay.value) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "������Ϣ.��������=[" & index_�������� & "]"
    End If
    
    '�������-��ʼ����(�������Ǳ�ѡ������)
    If Trim(dtpBegin.value) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & GetQueryTimeField(cboQueryTime.Text) & ">=[" & index_��ѯ��ʼʱ�� & "]"
    End If
    
    '�������-��������(�������Ǳ�ѡ������)
    If Trim(dtpEnd.value) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & GetQueryTimeField(cboQueryTime.Text) & "<=[" & index_��ѯ����ʱ�� & "]"
    End If
    
    '��鲿λ
    If Trim(cboBodyPart.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "Instr(����ҽ����¼.ҽ������, [" & index_��鲿λ & "]) > 0"
    End If
    
    '��鼼ʦ
    If Trim(cboStudyDoctor.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "Ӱ�����¼.��鼼ʦ=[" & index_��鼼ʦ & "]"
    End If
    
    'Ӱ�����
    If Trim(cboImageType.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "Ӱ�����¼.Ӱ�����=[" & index_Ӱ����� & "]"
    End If
    
    '����豸
    If Trim(cboDevice.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "Ӱ�����¼.����豸=[" & index_����豸 & "]"
    End If
    
    '���˿��� "+0"��ʾ������������Щ�ط�ʹ��������ѯ��Ч������ԱȽϵ�Ч
    If Trim(cboPatientRoom.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "����ҽ����¼.���˿���ID+0=[" & index_���˿��� & "]"
    End If
    
    '������
    If Trim(cboProcedure.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "

        Select Case cboProcedure.Text
            Case "�ѵǼ�"
                strFilter = strFilter & " ( ����ҽ������.ִ�й���=0 or ����ҽ������.ִ�й��� = 1 or ����ҽ������.ִ�й��� IS NULL) "
            Case "�ѱ���"
                strFilter = strFilter & " ( ����ҽ������.ִ�й���=2 and Ӱ�����¼.������ IS NULL)"
            Case "�Ѽ��"
                strFilter = strFilter & " ( ����ҽ������.ִ�й���=3 and Ӱ�����¼.������ IS NULL)"
            Case "������"
                strFilter = strFilter & " ( not Ӱ�����¼.������� IS NULL)"
            Case "������"
                strFilter = strFilter & " (( ����ҽ������.ִ�й��� =2 or ����ҽ������.ִ�й���=3) and not Ӱ�����¼.������ is null and Ӱ�����¼.������� is null) "
            Case "�ѱ���"
                strFilter = strFilter & " (����ҽ������.ִ�й���=4 and Ӱ�����¼.������ is null) "
            Case "�����"
                strFilter = strFilter & " (����ҽ������.ִ�й���=4 and not Ӱ�����¼.������ is null) "
            Case "�����"
                strFilter = strFilter & " ����ҽ������.ִ�й���=5 "
            Case "�����"
                strFilter = strFilter & " ����ҽ������.ִ�й���=6 "
        End Select
    End If
    
    'Ӱ������
    If Trim(cboQuality.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "Ӱ�����¼.Ӱ������=[" & index_Ӱ������ & "]"
    End If
    
    '������
    If Trim(cboYangXingLv.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " Nvl(����ҽ������.�������, 0)=[" & index_������ & "]"
    End If
    
    '���ҽ��
    If Trim(cboAuditingDoctor.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "Ӱ�����¼.������=[" & index_���ҽ�� & "]"
    End If
    
    '���ҽ��
    If Trim(cboDiagnoseDoctor.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "Ӱ�����¼.������=[" & index_���ҽ�� & "]"
    End If
    

    '���
    If Trim(txt���.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "Ӱ�����¼.�������=[" & index_��� & "]"
    End If
    
    '������� - ��Ҫ����������й�����ѯ
    If Trim(txtIllnessRes.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "����ҽ����¼.ID IN (Select t.ҽ��id From ����ҽ������ t Where t.����id IN " & _
                                                                        " (Select Distinct a.ID  " & _
                                                                        " From ���Ӳ�����¼ a,���Ӳ������� b " & _
                                                                        " Where a.����ʱ��>[" & index_��ѯ��ʼʱ�� & "] AND a.Id=b.�ļ�ID  " & _
                                                                        " And b.��������=7 And instr(b.��������,'52;')>0 And instr(b.�����ı�,[" & index_������� & "])>0))"
    End If
    
    '��������
    If Trim(txtReportContext.Text) <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & " And ����ҽ����¼.ID IN (Select t.ҽ��id From ����ҽ������ t Where t.����id IN " & _
                                                                " (Select Distinct a.ID " & _
                                                                " From ���Ӳ�����¼ a,���Ӳ������� b " & _
                                                                " Where a.����ʱ��>[" & index_��ѯ��ʼʱ�� & "] AND A.Id=b.�ļ�ID " & _
                                                                " And b.��������=2 And instr(b.�����ı�,[" & index_�������� & "])>0 And b.��ֹ�� = 0)) "
    End If
    
    '�������
    If Trim(txt�������.Text) <> "" Then
        If Trim(strSubFilter) <> "" Then strSubFilter = strSubFilter & " or "
        strSubFilter = strSubFilter & " (b.�����ı� ='�������' And Instr(c.�����ı�, [" & index_������� & "]) > 0)"
    End If
    
    '������
    If Trim(txt������.Text) <> "" Then
        If Trim(strSubFilter) <> "" Then strSubFilter = strSubFilter & " or "
        strSubFilter = strSubFilter & " (b.�����ı� ='������' And Instr(c.�����ı�, [" & index_������ & "]) > 0)"
    End If
    
    '����
    If Trim(txt����.Text) <> "" Then
        If Trim(strSubFilter) <> "" Then strSubFilter = strSubFilter & " or "
        strSubFilter = strSubFilter & " (b.�����ı� ='����' And Instr(c.�����ı�, [" & index_���� & "]) > 0)"
    End If
    
    If strSubFilter <> "" Then
        If Trim(strFilter) <> "" Then strFilter = strFilter & " and "
        
        strSubFilter = " (" & strSubFilter & ")"
        
        strFilter = strFilter & " ����ҽ����¼.ID IN ( Select t.ҽ��id From ����ҽ������ t Where t.����id IN " & _
            " (Select Distinct a.ID From ���Ӳ�����¼ a, ���Ӳ������� b,���Ӳ������� c " _
            & " Where a.����ʱ�� > [" & index_��ѯ��ʼʱ�� & "] And a.Id = b.�ļ�id And b.Id = C.��ID And b.�������� = 3 And c.�������� = 2 And c.��ֹ�� = 0 and " _
            & strSubFilter & "))"
    End If
    
    GetQueryFilter = strFilter
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetQueryTimeField(ByVal strTimeType As String) As String
'********************************************
'
'ȡ�����ڲ�ѯ������ֶ�
'
'********************************************
    Dim strField As String
    
    On Error GoTo errHandle
    
    Select Case strTimeType
        Case "����ʱ��"
            strField = "����ҽ������.����ʱ��"
        Case "����ʱ��"
            strField = "����ҽ������.�״�ʱ��"
        Case "��ͼʱ��"
            strField = "Ӱ�����¼.��������"
    End Select
    
    GetQueryTimeField = strField
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog

End Function


Private Function GetQueryAgeWhere(ByVal strAgeWhere) As String
'********************************************
'
'��ȡ��ʹ��������в�ѯʱ�����õĲ�ѯ��������
'
'********************************************
    Dim strWhere As String
    
    On Error GoTo errHandle
    
    Select Case strAgeWhere
        Case "����"
            strWhere = ">"
        Case "���ڵ���"
            strWhere = ">="
        Case "С��"
            strWhere = "<"
        Case "С�ڵ���"
            strWhere = "<="
        Case "����"
            strWhere = "="
        Case "��"
            strWhere = ""
    End Select
    
    GetQueryAgeWhere = strWhere
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    
End Function


Private Function GetQueryNumField(ByVal strNumType As String) As String
'********************************************
'
'ȡ�ú����ѯ����Ҫ���ֶ�
'
'********************************************
    Dim strField As String
    
    On Error GoTo errHandle
    
    Select Case strNumType
        Case "�����"
            strField = "������Ϣ.�����"
        Case "סԺ��"
            strField = "������Ϣ.סԺ��"
        Case "���￨��"
            strField = "������Ϣ.���￨��"
        Case "���ݺ�"
            strField = "����ҽ������.No"
        Case "IC������"
            strField = "������Ϣ.IC����"
        Case "����"
            strField = "Ӱ�����¼.����"
        Case "�����"
            strField = "Ӱ��걾����ȡ��.�����"
        Case "���֤"
            strField = "������Ϣ.���֤��"
    End Select
    
    GetQueryNumField = strField
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub InitFaceData()
'********************************************
'
'��ʼ����ѯ��������
'
'********************************************

    '�����ڿؼ��������޸�Ϊ��������
    dtpBirthDay.value = Now
    dtpBegin.value = CDate(Now - time)
    dtpEnd.value = Now
    
    dtpBirthDay.value = ""
            
            
    '�����鲿λ
    Call LoadStudyPart
    '����ͼ�����
    Call LoadImageType
    '�������ҽ��
    Call LoadDoctor
    '���벡���Ա�
    Call LoadSex
    '���벡�˿���
    Call LoadPatientRoom
    '��ȡ����豸
    Call LoadStudyDevice
    
    
    cboNumType.ListIndex = 2 'Ĭ�ϰ��վ��￨�Ų�ѯ
    cboAgeWhere.ListIndex = 4 'Ĭ������Ĳ�ѯ����Ϊ����
    cboAgeType.ListIndex = 0 'Ĭ�ϵ����䵥λΪ��
    cboQueryTime.ListIndex = 0 'Ĭ�ϵĲ�ѯʱ��Ϊ����ʱ��
    cboPatientFrom.ListIndex = 0 'Ĭ�ϲ��Բ�����Դ�����ж�


    'Call txtName.SetFocus
End Sub


Private Sub LoadStudyDevice()
'********************************************
'
'��ȡ����豸
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    strSql = "Select �豸��, �豸�� From Ӱ���豸Ŀ¼ where ����=4 and ״̬=1"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����豸")
    cboDevice.Clear
    cboDevice.AddItem ""
    cboDevice.ItemData(0) = -1
        
    With Me.cboDevice
        Do While Not rsTmp.EOF
            .AddItem rsTmp!�豸�� & "-" & Nvl(rsTmp!�豸��)
            .ItemData(cboDevice.NewIndex) = rsTmp!�豸��
            
            rsTmp.MoveNext
        Loop
    End With
    
    cboDevice.ListIndex = 0
 
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadPatientRoom()
'********************************************
'
'��ȡ���˿���
'
'********************************************
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Long
    
    strSql = "Select Distinct A.ID,A.����,A.����,B.�������" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.�������� IN('�ٴ�','����')" & _
        " And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.����"
        
    On Error GoTo errHandle
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    cboPatientRoom.Clear
    cboPatientRoom.AddItem ""
    cboPatientRoom.ItemData(0) = -1
    
    For i = 1 To rsTmp.RecordCount
        cboPatientRoom.AddItem rsTmp!���� & "-" & rsTmp!����
        cboPatientRoom.ItemData(cboPatientRoom.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next
    
    cboPatientRoom.ListIndex = 0

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadSex()
'********************************************
'
'��ȡ�����Ա�
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    strSql = "Select ���� From �Ա�"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�����Ա�")
    cboSex.Clear
    cboSex.AddItem ""
        
    With Me.cboSex
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("����"))) & "-" & Nvl(rsTmp("����"))
            rsTmp.MoveNext
        Loop
    End With
    
    cboSex.ListIndex = 0
 
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadDoctor()
'********************************************
'
'��ȡ���ҽ�������ҽ�������ҽ��
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    
    cboDiagnoseDoctor.Clear
    cboAuditingDoctor.Clear
    cboStudyDoctor.Clear
    
    cboDiagnoseDoctor.AddItem ""
    cboAuditingDoctor.AddItem ""
    cboStudyDoctor.AddItem ""
        
    
    strSql = "select distinct A.����,A.���� from ��Ա�� A,������Ա B where B.����ID=[1] AND A.ID=B.��ԱID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ҽ��", mDepartmentId)
    
    If rsTmp Is Nothing Then Exit Sub
    
    Do While Not rsTmp.EOF
        cboDiagnoseDoctor.AddItem rsTmp!���� & "-" & rsTmp!����
        cboAuditingDoctor.AddItem rsTmp!���� & "-" & rsTmp!����
        cboStudyDoctor.AddItem rsTmp!���� & "-" & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    cboDiagnoseDoctor.ListIndex = 0
    cboAuditingDoctor.ListIndex = 0
    cboStudyDoctor.ListIndex = 0
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub LoadStudyPart()
'********************************************
'
'��ȡ��鲿λ
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    strSql = "Select Distinct ���� From ���Ƽ�鲿λ"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��λ")
    
    cboBodyPart.Clear
    cboBodyPart.AddItem ""
        
    With Me.cboBodyPart
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("����"))) & "-" & Nvl(rsTmp("����"))
            rsTmp.MoveNext
        Loop
    End With
    
    cboBodyPart.ListIndex = 0
 
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadImageType()
'********************************************
'
'��ȡͼ�����
'
'********************************************
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "select ����,���� from Ӱ�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Ӱ�������")
    
    cboImageType.Clear
    cboImageType.AddItem ""
    
    Do While Not rsTemp.EOF
        cboImageType.AddItem rsTemp!���� & "-" & rsTemp!����
        rsTemp.MoveNext
    Loop
    
    cboImageType.ListIndex = 0
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetDaysHint(ByVal intDays As Integer) As String
'********************************************
'
'ȡ�õ�ǰ������Ӧ������˵��
'
'intDays����Ҫת��������˵���ľ�������
'
'********************************************
    Dim strReturn As String
    
    On Error GoTo errHandle
    
    If intDays = 0 Then strReturn = "����"
    If intDays >= 1 And intDays < 7 Then strReturn = intDays & "��(��" & intDays & "��)"
    If intDays >= 7 And intDays < 30 Then strReturn = intDays & "��(��" & Int(intDays / 7) & "��)"
    If intDays >= 30 And intDays < 180 Then strReturn = intDays & "��(��" & Int(intDays / 30) & "��)"
    If intDays >= 180 And intDays < 365 Then strReturn = intDays & "��(������)"
    
    GetDaysHint = strReturn
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function



Private Sub cmdQuit_Click()
    mblnOk = False
    Me.Hide
End Sub

Private Sub cmdSure_Click()
    If dtpEnd.value < dtpBegin.value Then
        MsgBox "��ѯʱ��Ŀ�ʼʱ�䲻�ܴ��ڽ�ֹʱ�䣬���飡", vbInformation, gstrSysName
        dtpEnd.SetFocus
        Exit Sub
    End If
    
    mblnOk = True
    Me.Hide
End Sub

Private Sub Form_Load()
    mblnOk = False
    
    Call InitFaceData
End Sub

Private Sub sldDays_Change()
    '���ò�ѯʱ�䷶Χ
    dtpBegin.value = CDate(dtpEnd.value - sldDays.value)
End Sub


Private Sub sldDays_Scroll()
    labDays.Caption = GetDaysHint(sldDays.value)
End Sub

