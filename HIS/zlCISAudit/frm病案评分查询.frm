VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�������ֲ�ѯ 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������ּ���"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   Icon            =   "frm�������ֲ�ѯ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm�������ֲ�ѯ.frx":000C
   ScaleHeight     =   7875
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo 
      Height          =   300
      Left            =   1620
      TabIndex        =   38
      Top             =   4455
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   10
      Left            =   1620
      TabIndex        =   19
      Tag             =   "��Ժ����"
      Top             =   4050
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   9
      Left            =   1620
      TabIndex        =   17
      Tag             =   "�����"
      Top             =   3615
      Width           =   3525
   End
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   1140
      Top             =   7095
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "ȱʡ(&W)"
      Height          =   350
      Left            =   360
      TabIndex        =   36
      Top             =   7410
      Width           =   1100
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   8
      Left            =   1620
      TabIndex        =   26
      Tag             =   "��Ժ����"
      Top             =   5640
      Width           =   3525
   End
   Begin MSComCtl2.DTPicker dt��Ժ���� 
      Height          =   300
      Index           =   0
      Left            =   2655
      TabIndex        =   22
      Top             =   4845
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   136314880
      CurrentDate     =   38373
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   7
      Left            =   1620
      TabIndex        =   15
      Tag             =   "������"
      Top             =   3180
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   6
      Left            =   1620
      TabIndex        =   13
      Tag             =   "���λ�ʿ"
      Top             =   2790
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   5
      Left            =   1620
      TabIndex        =   11
      Tag             =   "����ҽʦ"
      Top             =   2415
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   4
      Left            =   1620
      TabIndex        =   9
      Tag             =   "סԺҽʦ"
      Top             =   2025
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   3
      Left            =   1620
      TabIndex        =   7
      Tag             =   "����"
      Top             =   1650
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   2
      Left            =   1620
      TabIndex        =   5
      Tag             =   "��ҳID"
      Top             =   1260
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   1
      Left            =   1620
      TabIndex        =   3
      Tag             =   "סԺ��"
      Top             =   885
      Width           =   3525
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   0
      Left            =   1620
      TabIndex        =   1
      Tag             =   "����ID"
      Top             =   495
      Width           =   3525
   End
   Begin VB.OptionButton optOr 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "������һ����(&V)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3450
      TabIndex        =   33
      Top             =   6855
      Width           =   1665
   End
   Begin VB.OptionButton optAnd 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "������������(&U)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1620
      TabIndex        =   32
      Top             =   6855
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2922
      TabIndex        =   34
      Top             =   7410
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4272
      TabIndex        =   35
      Top             =   7410
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dt��Ժ���� 
      Height          =   300
      Index           =   1
      Left            =   2655
      TabIndex        =   24
      Top             =   5265
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   136314880
      CurrentDate     =   38373
   End
   Begin MSComCtl2.DTPicker dt��Ժ���� 
      Height          =   300
      Index           =   0
      Left            =   2655
      TabIndex        =   29
      Top             =   6030
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   136314880
      CurrentDate     =   38373
   End
   Begin MSComCtl2.DTPicker dt��Ժ���� 
      Height          =   300
      Index           =   1
      Left            =   2655
      TabIndex        =   31
      Top             =   6420
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   136314880
      CurrentDate     =   38373
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&S)"
      Height          =   180
      Left            =   525
      TabIndex        =   39
      Top             =   4500
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ժ����(&L)"
      Height          =   180
      Index           =   16
      Left            =   525
      TabIndex        =   18
      Top             =   4110
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� ��(&K)"
      Height          =   180
      Index           =   15
      Left            =   525
      TabIndex        =   16
      Top             =   3675
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� ID(&A)"
      Height          =   180
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Top             =   555
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ס Ժ ��(&B)"
      Height          =   180
      Index           =   1
      Left            =   525
      TabIndex        =   2
      Top             =   945
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "סԺ����(&D)"
      Height          =   180
      Index           =   2
      Left            =   525
      TabIndex        =   4
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&E)"
      Height          =   180
      Index           =   3
      Left            =   525
      TabIndex        =   6
      Top             =   1710
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ҽʦ(&F)"
      Height          =   180
      Index           =   4
      Left            =   525
      TabIndex        =   8
      Top             =   2085
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ҽʦ(&G)"
      Height          =   180
      Index           =   5
      Left            =   525
      TabIndex        =   10
      Top             =   2475
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���λ�ʿ(&I)"
      Height          =   180
      Index           =   6
      Left            =   525
      TabIndex        =   12
      Top             =   2850
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� ��(&J)"
      Height          =   180
      Index           =   7
      Left            =   525
      TabIndex        =   14
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ժ����(&M)"
      Height          =   180
      Index           =   8
      Left            =   525
      TabIndex        =   20
      Top             =   4935
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ժ����(&Q)"
      Height          =   180
      Index           =   9
      Left            =   525
      TabIndex        =   25
      Top             =   5700
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ժ����(&R)"
      Height          =   180
      Index           =   10
      Left            =   525
      TabIndex        =   27
      Top             =   6090
      Width           =   990
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   180
      TabIndex        =   37
      Top             =   90
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   5940
      Y1              =   7245
      Y2              =   7245
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   -105
      X2              =   5940
      Y1              =   7260
      Y2              =   7260
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&T)"
      Height          =   180
      Index           =   14
      Left            =   1620
      TabIndex        =   30
      Top             =   6480
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����(&S)"
      Height          =   180
      Index           =   13
      Left            =   1620
      TabIndex        =   28
      Top             =   6090
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&P)"
      Height          =   180
      Index           =   12
      Left            =   1620
      TabIndex        =   23
      Top             =   5325
      Width           =   990
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����(&N)"
      Height          =   180
      Index           =   11
      Left            =   1620
      TabIndex        =   21
      Top             =   4935
      Width           =   990
   End
End
Attribute VB_Name = "frm�������ֲ�ѯ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnForce                As Boolean
Public mstrReturn               As String
Public mblnOK                   As Boolean
Public mbln��Ŀ������           As Boolean
Public mblnCancel               As Boolean
Private intPara                 As Integer

Private mlngSickID              As Long             '����ID
Private mlngHospitalID          As Long             'סԺ��
Private mlngHospitalTimes       As Long             'סԺ����
Private mstrSickName            As String           '��������
Private mstrMainDoctor          As String           '����ҽʦ
Private mstrOutpatientDoctor    As String           '����ҽʦ
Private mstrNurses              As String           '���λ�ʿ
Private mstrRatingMan           As String           '������
Private mstrAuditMan            As String           '�����
Private mstrOutDept             As String           '��Ժ����
Private mstrInDept              As String           '��Ժ����
Private mdatStarOutDate         As Date             '��Ժ��ʼ����
Private mdatEndOutDate          As Date             '��Ժ��ʼ����
Private mdatStarInDate          As Date             '��Ժ��ʼ����
Private mdatEndInDate           As Date             '��Ժ��ʼ����
Private mstrSickType            As String           '��������

'����ID
Public Property Get lngSickID() As Long
    lngSickID = mlngSickID
End Property

'סԺ��
Public Property Get lngHospitalID() As Long
    lngHospitalID = mlngHospitalID
End Property

'סԺ����
Public Property Get lngHospitalTimes() As Long
    lngHospitalTimes = mlngHospitalTimes
End Property

'��������
Public Property Get strSickName() As String
    strSickName = mstrSickName
End Property

'����ҽʦ
Public Property Get strMainDoctor() As String
    strMainDoctor = mstrMainDoctor
End Property

'����ҽʦ
Public Property Get strOutpatientDoctor() As String
    strOutpatientDoctor = mstrOutpatientDoctor
End Property

'���λ�ʿ
Public Property Get strNurses() As String
    strNurses = mstrNurses
End Property

'������
Public Property Get strRatingMan() As String
    strRatingMan = mstrRatingMan
End Property

'�����
Public Property Get strAuditMan() As String
    strAuditMan = mstrAuditMan
End Property

'��Ժ����
Public Property Get strOutDept() As String
    strOutDept = mstrOutDept
End Property

'��Ժ����
Public Property Get strInDept() As String
    strInDept = mstrInDept
End Property

'��Ժ��ʼ����
Public Property Get datStarOutDate() As Date
    datStarOutDate = mdatStarOutDate
End Property

'��Ժ��������
Public Property Get datEndOutDate() As Date
    datEndOutDate = mdatEndOutDate
End Property

'��Ժ��ʼ����
Public Property Get datStarInDate() As Date
    datStarInDate = mdatStarInDate
End Property

'��Ժ��������
Public Property Get datEndInDate() As Date
    datEndInDate = mdatEndInDate
End Property

'��������
Public Property Get strSickType() As String
    strSickType = mstrSickType
End Property

'==============================================================================
'=���ܣ�ȡ���˳�
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo errH
    txtInfo(0).SetFocus
    mblnCancel = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����ȱʡֵ
'==============================================================================
Private Sub cmdDefault_Click()
    Dim i           As Long
    
    On Error GoTo errH
    
    For i = 0 To 10
        txtInfo(i).Text = ""
    Next
    
    dt��Ժ����(0) = DateAdd("M", -1, Now)
    dt��Ժ����(1) = Now
    dt��Ժ����(0) = ""
    dt��Ժ����(1) = ""
    optAnd.Value = True
    txtInfo(0).SetFocus
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ȷ����ѯ�˳�
'==============================================================================
Private Sub CmdOK_Click()
    Dim i           As Long
    
    On Error GoTo errH
    
    intPara = 1
        
    mstrReturn = " and (1=1 "
    For i = 0 To 10
        If Trim(txtInfo(i)) <> "" Then
            If txtInfo(i).Tag = "��Ժ����" Then '��Ժ����
                If InStrRev(txtInfo(i), ",") > 0 Then
                    mstrReturn = mstrReturn & IIf(optAnd.Value = True, " And ", " Or ") & txtInfo(i).Tag & "  In (" & Get��������(UserInfo.ID, 1) & ")"
                Else
                    mstrReturn = mstrReturn & IIf(optAnd.Value = True, " And ", " Or ") & txtInfo(i).Tag & " = [" & intPara & "] "
                End If
                
            Else
                mstrReturn = mstrReturn & IIf(optAnd.Value = True, " And ", " Or ") & txtInfo(i).Tag & " = [" & intPara & "] "
            End If
        End If
        intPara = intPara + 1
    Next
    
    If cbo.Text <> "" Then
        mstrReturn = mstrReturn & IIf(optAnd.Value = True, " And ", " Or ") & "��������=[16] "
    End If
    
    If Not IsNull(dt��Ժ����(0).Value) Then
        If IsNull(dt��Ժ����(1)) Then
            mstrReturn = mstrReturn & IIf(optAnd.Value = True, " and ", " or ") & _
                "��Ժ���� >= [" & intPara & "]"
        Else
            mstrReturn = mstrReturn & IIf(optAnd.Value = True, " and ", " or ") & _
                "��Ժ����>= [" & intPara & "] and " & _
                "��Ժ����<= [" & intPara + 1 & "]"
        End If
    End If
    
    intPara = intPara + 2
    
    If Not IsNull(dt��Ժ����(0).Value) Then
        If IsNull(dt��Ժ����(1)) Then
            mstrReturn = mstrReturn & IIf(optAnd.Value = True, " and ", " or ") & _
                "��Ժ���� >= [" & intPara & "]"
        Else
            mstrReturn = mstrReturn & IIf(optAnd.Value = True, " and ", " or ") & _
                "��Ժ����>= [" & intPara & "] and " & _
                "��Ժ����<= [" & intPara + 1 & "]"
        End If
    End If
    
    '�������
    mlngSickID = Val(txtInfo(0).Text)               '����ID
    mlngHospitalID = Val(txtInfo(1).Text)           'סԺ��
    mlngHospitalTimes = Val(txtInfo(2).Text)        'סԺ����
    mstrSickName = txtInfo(3).Text                  '��������
    mstrMainDoctor = txtInfo(4).Text                '����ҽʦ
    mstrOutpatientDoctor = txtInfo(5).Text          '����ҽʦ
    mstrNurses = txtInfo(6).Text                    '���λ�ʿ
    mstrRatingMan = txtInfo(7).Text                 '������
    mstrInDept = txtInfo(8).Text                    '��Ժ����
    mstrAuditMan = txtInfo(9).Text                  '�����
    mstrOutDept = txtInfo(10).Text                  '��Ժ����
    mstrSickType = cbo.Text                         '��������
    
    If Not IsNull(dt��Ժ����(0).Value) Then
        mdatStarOutDate = Format(dt��Ժ����(0).Value, "yyyy-mm-dd 00:00:00")      '��Ժ��ʼ����
    End If
    If Not IsNull(dt��Ժ����(1).Value) Then
        mdatEndOutDate = Format(dt��Ժ����(1).Value, "yyyy-mm-dd 23:59:59")       '��Ժ��ʼ����
    End If
    If Not IsNull(dt��Ժ����(0).Value) Then
        mdatStarInDate = Format(dt��Ժ����(0).Value, "yyyy-mm-dd 00:00:00")       '��Ժ��ʼ����
    End If
    If Not IsNull(dt��Ժ����(1).Value) Then
        mdatEndInDate = Format(dt��Ժ����(1).Value, "yyyy-mm-dd 23:59:59")        '��Ժ��ʼ����
    End If
    If mstrReturn = " and (1=1 " Then
        MsgBox "���������������", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    mblnOK = True
    
    If mbln��Ŀ������ Then
        mstrReturn = mstrReturn & ") and ��Ŀ���� is not null "
    Else
        mstrReturn = mstrReturn & ") "
    End If
    mblnCancel = False
    Unload Me
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���Ժ���������ж�
'==============================================================================
Private Sub dt��Ժ����_Change(Index As Integer)
    On Error GoTo errH
    If dt��Ժ����(0).Value > dt��Ժ����(1).Value Then MsgBox "��ʼ����Ӧ�ñȽ��������磡", vbExclamation, gstrSysName: dt��Ժ����(0).SetFocus: Exit Sub
    If Index = 0 Then
        If IsNull(dt��Ժ����(0).Value) Then dt��Ժ����(1).Value = Null
    Else
        If IsNull(dt��Ժ����(0).Value) And Not IsNull(dt��Ժ����(1).Value) Then dt��Ժ����(1).Value = Null
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���Ժ�������ݼ��
'==============================================================================
Private Sub dt��Ժ����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���Ժ���������ж�
'==============================================================================
Private Sub dt��Ժ����_Change(Index As Integer)
    On Error GoTo errH
    If dt��Ժ����(0).Value > dt��Ժ����(1).Value Then MsgBox "��ʼ����Ӧ�ñȽ��������磡", vbExclamation, gstrSysName: dt��Ժ����(0).SetFocus: Exit Sub
    If Index = 0 Then
        If IsNull(dt��Ժ����(0).Value) Then dt��Ժ����(1).Value = Null
    Else
        If IsNull(dt��Ժ����(0).Value) And Not IsNull(dt��Ժ����(1).Value) Then dt��Ժ����(1).Value = Null
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���Ժ�������ݼ��
'==============================================================================
Private Sub dt��Ժ����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ����ڳ�ʼ��'��ȡϵͳ�������Ƿ��Ŀ���������
'==============================================================================
Private Sub Form_Load()
    On Error GoTo errH
    mblnCancel = True
    mbln��Ŀ������ = Val(zlDatabase.GetPara(91, glngSys, 0)) = 1
    mblnForce = False
    
    Call Fill��������
    
    dt��Ժ����(0) = DateAdd("M", -1, Now)
    dt��Ժ����(1) = Now
    dt��Ժ����(0) = DateAdd("M", -1, Now)
    dt��Ժ����(1) = Now
    dt��Ժ����(0) = ""
    dt��Ժ����(1) = ""
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ҳ��ر����ݴ���
'==============================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errH
    If mblnForce Then
        '�ڲ������������ǿ�ƹر�
        If IsCompiled = True Then
            Call SetWindowLong(Me.hWnd, GWL_WNDPROC, OldWindowProc)
        End If
    Else
        '�û�ռ���˹رհ�ť�������رյ�����
        Cancel = 1
        Me.Hide
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ������������ɲ�ѯ���
'==============================================================================
Public Function GetFilter(ByVal strPrivs As String, ByVal txtDept As String) As String
    On Error GoTo errH
    If Not IsPrivs(strPrivs, "���п���") Then
        If txtDept <> UserInfo.�������� Then
            txtInfo(10).Text = txtDept
        Else
            txtInfo(10).Text = UserInfo.��������
        End If
        txtInfo(10).Locked = True
        txtInfo(10).BackColor = &H80000000
    
    End If
    mblnOK = False
    Me.Show vbModal
    If mblnOK Then
        GetFilter = mstrReturn
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ������� ѡ��
'==============================================================================
Private Sub optAnd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ������� ѡ��
'==============================================================================
Private Sub optOr_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����¼�����
'==============================================================================
Private Sub txtInfo_Change(Index As Integer)
    On Error GoTo errH
    If InStr(txtInfo(Index), "'") <> 0 Then txtInfo(Index) = Replace(txtInfo(Index), "'", "")
    If InStr(txtInfo(Index), "|") <> 0 Then txtInfo(Index) = Replace(txtInfo(Index), "|", "")
    txtInfo(Index).SelStart = Len(txtInfo(Index))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��������뷨����
'==============================================================================
Private Sub txtInfo_GotFocus(Index As Integer)
    On Error GoTo errH
    zlControl.TxtSelAll txtInfo(Index)
    Select Case Index
        Case 0, 1, 2
            Call zlCommFun.OpenIme(False)
        Case Else
            Call zlCommFun.OpenIme(True)
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�������������
'==============================================================================
Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo errH
    If InStr("'|", Chr(KeyAscii)) <> 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    Select Case Index
        Case 0, 1, 2
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If InStr("1234567890." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�Shift�س� ȷ����ѯ
'==============================================================================
Private Sub txtInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = vbKeyReturn And Shift = 2 Then
        Call CmdOK_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��ؼ�ͨ��Tips��ʾ
'==============================================================================
Private Sub ShowTips(ctl As Control, str���� As String, Optional str���� As String = "��ʾ��Ϣ", Optional lngʱ�� As Long = 2500)
    Dim X           As Single
    Dim Y           As Single
    
    On Error GoTo errH
    
    X = (ctl.Left + ctl.Width / 2) / Screen.TwipsPerPixelX
    Y = (ctl.Top + ctl.Height) / Screen.TwipsPerPixelY
    If Len(str����) > 0 Then
        tipPopup1.Hide
        tipPopup1.StandardIcon = IDI_INFORMATION
        tipPopup1.ShowCloseButton = True
        tipPopup1.TimeOut = lngʱ��
        tipPopup1.Title = str����
        tipPopup1.Text = str����
        tipPopup1.Show Me.hWnd, X, Y
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Fill��������()
    Dim rs As New ADODB.Recordset
    On Error GoTo errH
    gstrSQL = "" & _
        "Select ����,����,����,ȱʡ��־ From ��������"
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    With cbo
        .Clear
        .AddItem ""
        .ItemData(.NewIndex) = 1
        
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                .AddItem zlCommFun.NVL(rs!����)
                 .ItemData(.NewIndex) = .NewIndex + 1

                rs.MoveNext
            Loop
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
        
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function Get��������(ByVal lng��ԱId As Long, ByVal lngMode As Long) As String
    Dim strSQL As String
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    ' lngMode =0 ��ʾ��ʽ lngMode =1 ���ڲ�ѯ��ʽ
    
    strSQL = "SELECT  distinct C.���� AS ����" & vbNewLine & _
                "      FROM ��Ա�� A,��Ա����˵�� B,���ű� C,������Ա D" & vbNewLine & _
                "      WHERE A.ID=B.��Աid AND C.ID=D.����id AND D.��Աid=A.ID And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
                "      AND A.id =[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ԱId)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        Do Until rsTemp.EOF
            If lngMode = 0 Then
                If Len(strTmp) = 0 Then
                    strTmp = NVL(rsTemp!����)
                Else
                    strTmp = strTmp & "," & NVL(rsTemp!����)
                End If
            Else
                If Len(strTmp) = 0 Then
                    strTmp = "'" & NVL(rsTemp!����) & "'"
                Else
                    strTmp = strTmp & ",'" & NVL(rsTemp!����) & "'"
                End If
            End If
            
            rsTemp.MoveNext
        Loop
        
        Get�������� = strTmp
    Else
        Get�������� = UserInfo.��������
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Function
End Function

