VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlidkind.ocx"
Begin VB.Form frmTriageRoomRegNum 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdNewPati 
      Height          =   345
      Left            =   3945
      Picture         =   "frmTriageRoomRegNum.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "��������(F4)"
      Top             =   90
      Width           =   375
   End
   Begin VB.CommandButton cmdGetNum 
      Caption         =   "ȡ��(&O)"
      Height          =   405
      Left            =   9990
      TabIndex        =   13
      Top             =   465
      Width           =   1065
   End
   Begin zlIDKind.CommandEx cmdExRoom 
      Height          =   285
      Left            =   9570
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   525
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin zlIDKind.CommandEx cmdExDoctor 
      Height          =   285
      Left            =   6690
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   510
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin zlIDKind.TextEx txtExDoctor 
      Height          =   360
      Left            =   4530
      TabIndex        =   8
      Top             =   495
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483645
      Appearance      =   0
      Text            =   ""
   End
   Begin zlIDKind.CommandEx cmdExDept 
      Height          =   285
      Left            =   3660
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   525
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin zlIDKind.PatiIdentify PatiIdentify 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmTriageRoomRegNum.frx":058A
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindAppearance=   0
      InputAppearance =   0
      ShowSortName    =   -1  'True
      DefaultCardType =   "0"
      IDKindWidth     =   555
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAutoCommCard=   -1  'True
      AllowAutoICCard =   -1  'True
      AllowAutoIDCard =   -1  'True
      NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
   End
   Begin zlIDKind.TextEx txtExDept 
      Height          =   360
      Left            =   705
      TabIndex        =   5
      Top             =   495
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483645
      Appearance      =   0
      Text            =   ""
   End
   Begin zlIDKind.TextEx txtExRoom 
      Height          =   360
      Left            =   7530
      TabIndex        =   11
      Top             =   495
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483645
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.Label lblBookingNO 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ԤԼ��:A0001"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   9750
      TabIndex        =   3
      Top             =   135
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H80000003&
      X1              =   11220
      X2              =   -30
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�  ���䣺   ����ţ� �ѱ�"
      Height          =   180
      Left            =   4380
      TabIndex        =   2
      Top             =   180
      Width           =   2880
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   585
      Width           =   360
   End
   Begin VB.Label lblDoctor 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��"
      Height          =   180
      Left            =   4125
      TabIndex        =   7
      Top             =   585
      Width           =   360
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   7080
      TabIndex        =   10
      Top             =   585
      Width           =   360
   End
   Begin XtremeSuiteControls.ShortcutCaption srtcBack 
      Height          =   990
      Left            =   15
      TabIndex        =   14
      Top             =   0
      Width           =   11100
      _Version        =   589884
      _ExtentX        =   19579
      _ExtentY        =   1746
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmTriageRoomRegNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************************************************************************************************
'����:ȡ�Ž���(��ҪӦ�ó��������ڹҺŴ��ڣ�ֱ���ڷ���ȡ�ſ���)
'����ӿ�:zlInitVar-��ʼ����ر�����Ϣ,����������ȵ���
'         LockedScreen-�����¼�(��Ҫ�����洦���������)
'         GetNumSucces-ȡ�ųɹ��¼�
'�ڲ��ӿ�:
'     CreateDeptStructure-�������ż��ṹ:ID,���룬����,����
'����:���˺�
'����:2018-01-03 11:29:56
'���ݴ������˵��:
'   1.�����ɹҺż�¼����Ϊ0,����ҽ������ʱ�����ɻ��۷��ã�Ȼ���շ�
'   2.�������ò���Ϊ����Һ�ģʽ
'**********************************************************************************************************************************************
Private mlngModule As Long, mstrPrivs As String
Private mstrNo As String '���ݺ�
Private mbytMode As Byte '0-ȡ��;1-ԤԼ����ȡ��;2-����ȡ��

Private mfrmMain As Object
Private mobjPati As PatiInfor
Private mstr������� As String
Private mobjCardSqure As Object
Private mrsRegData As ADODB.Recordset
Private mrsBookData As ADODB.Recordset

Private mobjRegister As clsRegist
Private mbytRegMode As Byte '1-�����ģʽ;0-��ͳģʽ
Private mlngPreDeptID As Long  '�ϴ�ѡ��Ŀ���ID
Private mlngPreItemID As Long  '�ϴ�ѡ�����ĿID
Private mstrPreDoctorName As String '�ϴ�ѡ���ҽ��
Private mstrPreRoomName As String '�ϴ�ѡ�������
Private mrsDept As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mrsRooms As ADODB.Recordset '��ȡ���Ҽ�
Private mobjSetFocus As Object '��ǰ����ƶ��Ŀؼ�
Private mblnFirst  As Boolean
Public Event LockedScreen(ByVal blnLocked As Boolean, blnCancel As Boolean)   '������������Ҫ���ڱ�������ʱ����Ҫ��ֹ��������
Public Event GetNumSucces(ByVal strNO As String)    '����ɹ���ˢ������
Private Type ty_Para
    blnBusy  As Boolean ' ����æʱ�������
    intԤԼʧЧ���� As Integer  'ԤԼʧԼ����
    intԤԼ��Чʱ�� As Integer  'ԤԼ��Чʱ��
End Type
Private mblnNotChange As Boolean
Private mPara As ty_Para
Private Sub LoadPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-15 15:46:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mPara
        .blnBusy = Val(zlDatabase.GetPara("����æʱ�������", glngSys, mlngModule, 0)) = 1
        .intԤԼʧЧ���� = Val(zlDatabase.GetPara("ԤԼʧԼ����", glngSys, 1111, 0))
        .intԤԼ��Чʱ�� = Val(zlDatabase.GetPara("ԤԼ��Чʱ��", glngSys, 1111, 0))
    End With
End Sub

Public Function zlInitVar(ByVal frmMain As Object, ByVal str������� As String, ByVal lngModule As Long, ByVal strPrivs As String, Optional ByVal objCardSqure As Object, Optional objRegister As clsRegist) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر�����
    '���:str�������-�������(IDs)
    '     objCardSqure-����������д��ڸö�����Ҫ���룬������Զ��ٴ���
    '     objRegister-�ҺŶ���
    '���� :������سɹ�������true,���򷵻�False(����Flaseʱ����������Ҫ�����ر�)
    '����:���˺�
    '����:2018-01-03 11:27:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstr������� = str�������: mlngModule = lngModule: mbytRegMode = 0: mbytMode = 0
    mstrPrivs = strPrivs: Set mfrmMain = frmMain
    On Error GoTo errHandle
    
    Call LoadPara   '���ز���
   
    Set mobjCardSqure = objCardSqure
    Call PatiIdentify.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, mobjCardSqure, , gstrProductName)
    Set mobjRegister = objRegister
    If mobjRegister Is Nothing Then
         If CreateRegisterObject = False Then Exit Function
         Set mobjRegister = gobjRegist
    End If
    
    cmdNewPati.ToolTipText = "��������(F4)"
    cmdNewPati.Visible = InStr(mstrPrivs, ";�����޸�;") > 0
    lblPati.Left = cmdNewPati.Left + IIf(InStr(mstrPrivs, ";�����޸�;") > 0, cmdNewPati.Width + 50, 0)
    zlInitVar = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdExDept_Click()
    If SelectDept("") = False Then
        DoEvents
        If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        zlControl.TxtSelAll txtExDept
        Exit Sub
    End If
    DoEvents
    If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
End Sub
 
Private Sub cmdExDoctor_Click()
    Dim lngItemID As Long, lngDeptID As Long, varTemp As Variant
    varTemp = Split(txtExDept.Tag & ":", ":")
    
    lngDeptID = Val(varTemp(0))
    lngItemID = Val(varTemp(1))
    
    If SelectDoctor(lngDeptID, lngItemID, "") = False Then
        DoEvents
        If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
        zlControl.TxtSelAll txtExDoctor
        Exit Sub
    End If
    
    If mbytMode = 0 Then Call LoadRoomsData   '������������
    DoEvents
    If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
End Sub
 

Private Sub cmdExRoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If SelectRooms("") = False Then
        DoEvents
        If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
        zlControl.TxtSelAll txtExRoom
        Exit Sub
    End If
    DoEvents
    If cmdGetNum.Enabled And cmdGetNum.Visible Then cmdGetNum.SetFocus
End Sub

Private Sub cmdGetNum_Click()
    Dim lng����ID As Long, strNO As String
    '����
    If LockedScreen(True) = False Then Exit Sub
    Select Case mbytMode
    Case 0  '����ȡ��
        If CheckDataValied(lng����ID) = False Then Exit Sub
        If SaveData(lng����ID, strNO) = False Then Call LockedScreen(False): Exit Sub
        RaiseEvent GetNumSucces(strNO)
        Call LockedScreen(False)
    Case 1 'ԤԼȡ��
        If SaveBooking(strNO) = False Then Call LockedScreen(False): Exit Sub
    Case 2  '����ȡ��
        If SaveHzGetNum(strNO) = False Then Call LockedScreen(False): Exit Sub
    End Select
    
    
    '��ӡƾ��
    Call PrintBill(strNO)
    '���������Ϣ
    txtExDept.Text = ""
    txtExDoctor.Text = ""
    txtExRoom.Text = ""
    PatiIdentify.Text = ""
    cmdExDept.Tag = ""
    cmdExDoctor.Tag = ""
    cmdExRoom.Tag = ""
    cmdNewPati.ToolTipText = "��������(F4)"
    lblPati.Caption = "�Ա�  ���䣺   ����ţ� �ѱ�"
    mstrNo = "": mbytMode = 0
    Set mobjPati = Nothing
    lblBookingNO.Visible = False
    Call LockedScreen(False)
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
End Sub


Private Function LockedScreen(ByVal blnLocked As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������
    '���:blnLocked-true-��ʾ����;False-���ǽ���
    '����:�����ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2018-01-09 16:05:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean
    On Error GoTo errHandle
    
    txtExDept.Enabled = Not blnLocked And cmdExDept.Tag <> "F"
    txtExDoctor.Enabled = Not blnLocked And cmdExDoctor.Tag <> "F"
    txtExRoom.Enabled = Not blnLocked And cmdExRoom.Tag <> "F"
    cmdExDept.Enabled = Not blnLocked And cmdExDept.Tag <> "F"
    cmdExDoctor.Enabled = Not blnLocked And cmdExDoctor.Tag <> "F"
    cmdExRoom.Enabled = Not blnLocked And cmdExRoom.Tag <> "F"
    cmdGetNum.Enabled = Not blnLocked
    PatiIdentify.Enabled = Not blnLocked
    
    
    blnCancel = False
    RaiseEvent LockedScreen(blnLocked, blnCancel)
    LockedScreen = Not blnCancel
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdNewPati_Click()
    Dim lng����ID As Long, lng����ID_Out As Long
    If mobjPati Is Nothing Then
        lng����ID = 0
    Else
        lng����ID = mobjPati.����ID
    End If
    If mobjRegister.zlPatiEdit(mfrmMain, lng����ID, lng����ID_Out) = False Then
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Sub
    End If
    
    PatiIdentify.Text = "-" & lng����ID_Out
    If GetPatient(PatiIdentify.GetCurCard, PatiIdentify.Text, False, mobjPati) = False Then
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Sub
    End If
    
    cmdNewPati.ToolTipText = "�޸Ĳ���(F4)"
    If SelectBooking(mobjPati.����ID, "") = False Then
        Call ReadRegData    '��ȡ�ҺŰ�������
    End If
    
    If txtExDept.Enabled And txtExDept.Visible Then
        txtExDept.SetFocus
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

 
Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If cmdGetNum.Enabled And cmdGetNum.Visible Then Call cmdGetNum_Click
        Case vbKeyF3
            If PatiIdentify.Visible = True And PatiIdentify.Enabled Then
                Call PatiIdentify.SetFocus
            End If
        Case vbKeyF4
            If cmdNewPati.Enabled And cmdNewPati.Visible Then Call cmdNewPati_Click
        Case Else
            PatiIdentify.ActiveFastKey
    End Select
End Sub
Private Sub Form_Load()
    Set mobjPati = Nothing
    mblnFirst = True
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With srtcBack
        lnTop.X1 = ScaleWidth
        lnTop.X2 = 0
        lnTop.Y1 = 0
        lnTop.Y2 = 0
        
        .Left = ScaleLeft
        .Top = ScaleTop + 15
        .Height = ScaleHeight - .Top
        .Width = ScaleWidth
        
        lblBookingNO.Left = ScaleHeight - lblBookingNO.Width - 50
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    Set mobjPati = Nothing
End Sub

Private Sub PatiIdentify_Change()
    If mblnNotChange Then Exit Sub
    Set mobjPati = Nothing
    cmdNewPati.ToolTipText = "��������(F4)"
    PatiIdentify.Tag = ""
    cmdExDept.Tag = ""
    cmdExDoctor.Tag = ""
    cmdExRoom.Tag = ""
    Call LockedScreen(False)
End Sub
 
 
Private Function CheckPatiCheck(ByVal lng����ID As Long) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������
    '���:lng����ID-����ID
    '����:�Ϸ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-08 16:26:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    If CreatePlugInOK(mlngModule) = False Then CheckPatiCheck = True: Exit Function
    On Error Resume Next
    'PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, _
    '    ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, _
    '    ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
    ''���ܣ���鵱ǰ�����Ƿ���ָ�������ⲡ��
    ''���أ�trueʱ�������������Falseʱ���������
    ''������
    ''      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
    ''      lngType �������ͣ�1������Һţ�2��סԺ��Ժ��3�������շѣ�4��סԺ���ʡ�
    ''      lngPatiID-����ID: �½����ģ�Ϊ0,�����뽨������ID
    ''      lngPageID-��ҳID: �½����ģ�Ϊ0,�����뽨����ҳID(סԺ������ҳID) ����˵������ lngType=4 ʱ�Ŵ��� lngPageID����������0
    ''      strPatiInforXML-������Ϣ:���δ�������˴��룬"�������Ա����䣬�������ڣ�ҽ���ţ����֤��"���������� ��ʽ:2016-11-11 12:12:12
    ''                      �̶���ʽ��<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH>
    ''      strReserve=��������,������չʹ��
    Dim blnChecked As Boolean
    blnChecked = gobjPlugIn.PatiValiedCheck(glngSys, mlngModule, 1, lng����ID, 0, "<YSXM>" & txtExDoctor.Text & "</YSXM>")
    
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "PatiValiedCheck"): Err.Clear
        On Error GoTo 0
        CheckPatiCheck = True: Exit Function
    End If
    CheckPatiCheck = blnChecked
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strNO As String
    
    If PatiIdentify.Tag <> "" Then blnFindPatied = True: Exit Sub
    
    blnFindPatied = False
    If GetPatient(objCard, strShowText, blnCancel, objCardData, strNO) = False Then
        DoEvents
        If Me.Enabled Then Me.SetFocus
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        blnCancel = True
        Exit Sub
    End If
    cmdNewPati.ToolTipText = "�޸Ĳ���(F4)"
    strShowText = objCardData.����
    PatiIdentify.Tag = objCardData.����ID
    Set mobjPati = objCardData
    
    blnFindPatied = True
    If strNO = "" Then
        If SelectBooking(objCardData.����ID, strNO) = False Then
            Call ReadRegData    '��ȡ�ҺŰ�������
        End If
    End If
    
    DoEvents
    If Me.Enabled Then Me.SetFocus
    If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
End Sub


Private Function GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, objPati As zlIDKind.PatiInfor, Optional ByRef strBookNo_out As String) As Boolean
                        
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:objCard-��ǰ������
    '     strInput-��ǰ���봮
    '     blnCard-��ǰ�Ƿ�ˢ��
    '����:strBookNo_out-��ԤԼ������ʱ������ԤԼ���ݺ�
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-10 13:50:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strOtherName As String, strOtherValue As String, blnCancel As Boolean
    Dim lng����ID As Long, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim vRect As RECT, rsTmp As ADODB.Recordset
    
    Set objPati = Nothing
    
    On Error GoTo errHandle
    
    strBookNo_out = ""
    If objCard Is Nothing Then Exit Function
    strOtherName = "": strOtherValue = "": lng����ID = 0
    If blnCard And (objCard.���� Like "����*" Or objCard.�Ƿ�ģ������) And InStr("-+*.", Left(strInput, 1)) = 0 Then     'ˢ��
        
        If PatiIdentify.Cards.��ȱʡ������ And Not PatiIdentify.GetfaultCard Is Nothing Then
            lng�����ID = PatiIdentify.GetfaultCard.�ӿ����
        ElseIf PatiIdentify.GetCurCard.�ӿ���� > 0 Then
            lng�����ID = PatiIdentify.GetCurCard.�ӿ����
        Else
            If lng�����ID = 0 Then lng�����ID = -1
        End If
        
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        
        If PatiIdentify.IsMobileNO(strInput) And lng����ID = 0 Then
            If gobjSquare.objSquareCard.zlGetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        End If
        If lng����ID <= 0 Then GoTo NotFoundPati:
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then        '�����
        strOtherName = "�����": strOtherValue = Val(Mid(strInput, 2))
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then         '����ID
        lng����ID = Val(Mid(strInput, 2))
    ElseIf Left(strInput, 1) = "." Then
        strBookNo_out = Mid(strInput, 2)
        strBookNo_out = GetFullNO(strBookNo_out, 12)
        PatiIdentify.Text = strBookNo_out
        If ReadBooking(strBookNo_out, True, objPati) = False Then Exit Function
        GetPatient = True: Exit Function
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                 If zlCommFun.ActualLen(strInput) <= 2 Then 'С��һ������ʱ�������й���
                    MsgBox "���������̫�򵥣�������2�������Ͻ��в��Ҳ��ˡ�", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                 End If
                 
                 strSQL = _
                     " Select /*+Rule */distinct A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ " & _
                     " From ������Ϣ A " & _
                     " Where Rownum <101 And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & _
                     " Order by  ����"
                 vRect = zlControl.GetControlRect(PatiIdentify.Hwnd)
                 Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, PatiIdentify.Height, blnCancel, False, True, strInput & "%")
                 
                 If blnCancel = True Then Exit Function
                 
                 If rsTmp Is Nothing Then GoTo NotFoundPati:
                 If rsTmp.EOF Then GoTo NotFoundPati:
                 lng����ID = Val(Nvl(rsTmp!����ID))
            Case "ҽ����"
                strInput = UCase(strInput)
                strOtherName = "ҽ����": strOtherValue = strInput
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                If objCard.�ӿ���� <> 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.�ӿ����, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                
                If lng����ID = 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                If lng����ID = 0 Then GoTo NotFoundPati:
                 
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If objCard.�ӿ���� <> 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.�ӿ����, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                If lng����ID = 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                If lng����ID = 0 Then GoTo NotFoundPati:
            Case "ԤԼ��"
                strBookNo_out = UCase(strInput)
                strBookNo_out = GetFullNO(strBookNo_out, 12)
                PatiIdentify.Text = strBookNo_out
                
                If ReadBooking(strBookNo_out, True, objPati) = False Then Exit Function
                GetPatient = True: Exit Function
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strOtherName = "�����": strOtherValue = Val(strInput)
             Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
        End Select
    End If
    
    If PatiIdentify.zlGetPatiObjectFromPatiID(lng����ID, objPati, strErrMsg, strOtherName, strOtherValue) = False Then GoTo NotFoundPati:
    
    If objPati Is Nothing Then GoTo NotFoundPati:
    If objPati.����ID = 0 Then GoTo NotFoundPati:
    If CheckPatiCheck(objPati.����ID) = False Then Exit Function
    Call SetPatiColor(PatiIdentify, objPati.��������, IIf(objPati.���� = 0, lblPati.ForeColor, vbRed))
    mblnNotChange = True
    PatiIdentify.Text = objPati.����
    mblnNotChange = False
    
    lblPati.Caption = "�Ա�:" & objPati.�Ա� & Space(4)
    lblPati.Caption = lblPati.Caption & "����:" & objPati.���� & Space(4)
    lblPati.Caption = lblPati.Caption & "�����:" & objPati.����� & Space(4)
    lblPati.Caption = lblPati.Caption & "�ѱ�:" & objPati.�ѱ� & Space(4)
    lblPati.Caption = lblPati.Caption & "��������:" & objPati.�������� & Space(4)
    lblPati.Caption = lblPati.Caption & "���ʽ:" & objPati.ҽ�Ƹ��ʽ & Space(4)
    lblPati.Caption = lblPati.Caption & "���֤��:" & objPati.���֤�� & Space(4)
    lblPati.Caption = lblPati.Caption & IIf(objPati.�������� = "", "", "��������:" & objPati.�������� & Space(4))
  
    GetPatient = True
    Exit Function
NotFoundPati:
    mblnNotChange = False
    MsgBox "δ�ҵ����������Ĳ���", vbInformation + vbOKOnly, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotChange = False
    Call SaveErrLog
End Function



Private Sub PatiIdentify_GotFocus()
    Call zlControl.TxtSelAll(PatiIdentify.objTxtInput)
End Sub
Private Function ReadRegData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�Һ�����
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-04 09:55:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strҽ������ As String, lngҽ��ID As Long, lngDeptID As Long, lngItemID As Long
    Dim str����IDs As String, strTemp As String
    Dim lngID As Long
    
    On Error GoTo errHandle
    txtExDept.Text = ""
    txtExDoctor.Text = ""
    txtExRoom.Text = ""
    
    Call CreateDeptStructure
    If mobjRegister Is Nothing Then Exit Function
    If mobjRegister.zlGetRegisterData(mrsRegData, mstr�������, , False, mbytRegMode) = False Then Exit Function
     
    If mrsRegData Is Nothing Then Exit Function
    If mrsRegData.RecordCount = 0 Then Exit Function
    lngID = 1
    txtExDept.Tag = "": txtExDoctor.Tag = ""
    With mrsRegData
        .MoveFirst
        Do While Not .EOF
            lngDeptID = Val(Nvl(mrsRegData!����ID))
            lngItemID = Val(Nvl(mrsRegData!��ĿID))
            
            strTemp = lngDeptID & ":" & lngItemID
            
            If strTemp = mlngPreDeptID & ":" & mlngPreItemID Then
                If txtExDept.Text = "" Then
                    txtExDept.Text = Nvl(mrsRegData!���ұ���) & "-" & Nvl(mrsRegData!��������) & "��" & mrsRegData!��Ŀ���� & "��"
                    txtExDept.Tag = strTemp
                End If
                
                strҽ������ = Nvl(mrsRegData!ҽ������): lngҽ��ID = Val(Nvl(mrsRegData!ҽ��ID))
                If Nvl(!ҽ������) = mstrPreDoctorName Then
                     txtExDoctor.Text = Nvl(mrsRegData!ҽ������)
                     txtExDoctor.Tag = Val(Nvl(mrsRegData!ҽ��ID)) & ":" & Nvl(mrsRegData!ҽ������)
                End If
            End If
            
            If InStr(str����IDs & ",", "," & strTemp & ",") = 0 Then
                mrsDept.AddNew
                mrsDept!ID = lngID
                mrsDept!����ID = lngDeptID
                mrsDept!���� = CStr(Nvl(mrsRegData!���ұ���))
                mrsDept!���� = CStr(Nvl(mrsRegData!��������))
                mrsDept!���� = CStr(Nvl(mrsRegData!���Ҽ���))
                mrsDept!��ĿID = lngItemID
                mrsDept!��Ŀ���� = CStr(Nvl(mrsRegData!��Ŀ����))
                mrsDept!��Ŀ���� = CStr(Nvl(mrsRegData!��Ŀ����))
                mrsDept!�Ƿ�ԭ���� = 0
                mrsDept.Update
                str����IDs = str����IDs & "," & strTemp
                lngID = lngID + 1
            End If
            .MoveNext
        Loop
    End With
    
    mrsRegData.MoveFirst
    If txtExDept.Tag = "" Then
        mlngPreDeptID = Val(Nvl(mrsRegData!����ID))
        mlngPreItemID = Val(Nvl(mrsRegData!��ĿID))
        strTemp = mlngPreDeptID & ":" & mlngPreItemID
        
        txtExDept.Text = Nvl(mrsRegData!���ұ���) & "-" & Nvl(mrsRegData!��������) & "��" & mrsRegData!��Ŀ���� & "��"
        txtExDept.Tag = strTemp
    End If
    If txtExDoctor.Tag = "" Then
        strҽ������ = Nvl(mrsRegData!ҽ������): lngҽ��ID = Val(Nvl(mrsRegData!ҽ��ID))
        txtExDoctor.Text = strҽ������
        txtExDoctor.Tag = lngҽ��ID & ":" & strҽ������
        mstrPreDoctorName = strҽ������
    End If
    ReadRegData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CreateDeptStructure() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ṹ
    '����:��ʼ���ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-04 18:19:09
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    
    Set mrsDept = New ADODB.Recordset
    With mrsDept
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "ID", adBigInt, , adFldIsNullable
            .Append "����ID", adBigInt, , adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 50, adFldIsNullable
            .Append "��ĿID", adVarChar, 50, adFldIsNullable
            .Append "��Ŀ����", adVarChar, 50, adFldIsNullable
            .Append "��Ŀ����", adVarChar, 200, adFldIsNullable
            .Append "�Ƿ�ԭ����", adBigInt, , adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    CreateDeptStructure = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SelectDept(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ�����
    '���:strInput-Ϊ��ʱ����ʾ��ѯ���е�
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-05 16:35:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset, intCount As Integer
    Dim lngDeptID As Long, lngItemID As Long, str���ұ��� As String, str�������� As String, str��Ŀ���� As String, str��Ŀ���� As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    On Error GoTo errHandle
    
    Set rsTemp = Nothing
    
    If Trim(strInput) = "" Then GoTo GoSel:
    strInput = UCase(strInput)
    
    strCompents = Replace(gstrLike, "%", "*") & strInput & "*"
    
    If mrsDept Is Nothing Then Exit Function
    If mrsDept.RecordCount = 0 Then Exit Function
    If mrsDept.RecordCount = 1 Then
        txtExDept.Text = mrsDept!���� & "-" & mrsDept!���� & "��" & mrsDept!��Ŀ���� & "��"
        lngDeptID = Val(Nvl(mrsDept!����ID))
        lngItemID = Val(Nvl(mrsDept!��ĿID))
        mlngPreDeptID = lngDeptID: mlngPreItemID = lngItemID
        
        txtExDept.Tag = lngDeptID & ":" & lngItemID
        Set mrsDoctor = LoadDoctorData(lngDeptID, lngItemID, mbytMode, Nvl(mrsDept!��Ŀ����), Nvl(mrsDept!��Ŀ����))
        Call LoadDefaultDoctor(lngDeptID)
        SelectDept = True
        Exit Function
    End If
     
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsDept)
    '��Ҫ����Ƿ��ж������������ļ�¼
    If IsNumeric(strInput) Then     '�������ȫ����
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strInput) Then     '�������ȫ��ĸ
        intInputType = 1
    Else
        intInputType = 2   ' 2-����
    End If
    
    lngDeptID = 0
    With mrsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strInput Then
                    txtExDept.Text = Nvl(!����) & "-" & Nvl(!����) & "��" & !��Ŀ���� & "��"
                    lngDeptID = Val(Nvl(!����ID)): lngItemID = Val(Nvl(mrsDept!��ĿID))
                    
                    txtExDept.Tag = lngDeptID & ":" & lngItemID
                    
                    Call LoadDoctorData(lngDeptID, lngItemID, , Nvl(!��Ŀ����), Nvl(!��Ŀ����))
                    Call LoadDefaultDoctor(lngDeptID)
                    SelectDept = True
                    Exit Function
                End If
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strInput) Then
                    If intCount = 0 Then
                        str���ұ��� = Nvl(!����): lngDeptID = Val(Nvl(!����ID)): lngItemID = Val(Nvl(!��ĿID))
                        str�������� = Nvl(!����): str��Ŀ���� = Nvl(!��Ŀ����): str��Ŀ���� = Nvl(!��Ŀ����)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Val(Nvl(!����)) Like strInput & "*" Then
                        Call zlDatabase.zlInsertCurrRowData(mrsDept, rsTemp)
                 End If
                 
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = UCase(strInput) Then
                    If intCount = 0 Then
                         str���ұ��� = Nvl(!����): lngDeptID = Val(Nvl(!����ID)): lngItemID = Val(Nvl(!��ĿID))
                        str�������� = Nvl(!����):: str��Ŀ���� = Nvl(!��Ŀ����): str��Ŀ���� = Nvl(!��Ŀ����)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsDept, rsTemp)
                    intCount = intCount + 1
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strInput Or Trim(!����) = strInput Or UCase(Trim(Nvl(!����))) = strInput Then
                    If intCount = 0 Then
                         
                        str���ұ��� = Nvl(!����): lngDeptID = Val(Nvl(!����ID)): lngItemID = Val(Nvl(!��ĿID))
                        str�������� = Nvl(!����):  str��Ŀ���� = Nvl(!��Ŀ����): str��Ŀ���� = Nvl(!��Ŀ����)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If Trim(Nvl(!����)) Like strInput & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsDept, rsTemp)
                End If
            End Select
            .MoveNext
        Loop
    End With
    
    If intCount > 1 Then lngDeptID = 0
GoSel:
    If Trim(strInput) = "" Then Set rsTemp = mrsDept
    If rsTemp Is Nothing Then
        If PatiIdentify.Text = "" Then
            MsgBox "����ѡ����Ҫȡ�ŵĲ���", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        Else
            MsgBox "δ�ҵ����������Ŀ���", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
    End If
    If lngDeptID = 0 And rsTemp.RecordCount = 1 Then
        rsTemp.MoveFirst
        str���ұ��� = Nvl(rsTemp!����): lngDeptID = Val(Nvl(rsTemp!����ID)): lngItemID = Val(Nvl(rsTemp!��ĿID))
        str�������� = Nvl(rsTemp!����):: str��Ŀ���� = Nvl(rsTemp!��Ŀ����): str��Ŀ���� = Nvl(rsTemp!��Ŀ����)
    End If
    
    'ֱ�Ӷ�λ
    If lngDeptID <> 0 Then
        If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
        txtExDept.Text = str���ұ��� & "-" & str�������� & "��" & str��Ŀ���� & "��"
        txtExDept.Tag = lngDeptID & ":" & lngItemID
        mlngPreDeptID = lngDeptID: mlngPreItemID = lngItemID
        
        Call LoadDoctorData(lngDeptID, lngItemID, , str��Ŀ����, str��Ŀ����)
        Call LoadDefaultDoctor(lngDeptID)
        SelectDept = True
        Exit Function
    End If


    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = "����"
    Case Else
        rsTemp.Sort = "���"
    End Select
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "δ�ҵ����������Ŀ���", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    
    '����ѡ����
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtExDept, rsTemp, True, "", "ID,����ID,��ĿID,�Ƿ�ԭ����", rsReturn) = False Then Exit Function
    If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
    
    If rsReturn Is Nothing Then Exit Function
    If rsReturn.RecordCount = 0 Then Exit Function
    
    lngDeptID = Val(Nvl(rsReturn!����ID)): lngItemID = Val(Nvl(rsReturn!��ĿID))
    txtExDept.Text = Nvl(rsReturn!����) & "-" & Nvl(rsReturn!����) & "��" & rsReturn!��Ŀ���� & "��"
    txtExDept.Tag = lngDeptID & ":" & lngItemID
    mlngPreDeptID = lngDeptID: mlngPreItemID = lngItemID
    
    Set mrsDoctor = LoadDoctorData(lngDeptID, lngItemID, , Nvl(rsReturn!��Ŀ����), Nvl(rsReturn!��Ŀ����))
    Call LoadDefaultDoctor(lngDeptID)
    
    rsReturn.Close: Set rsReturn = Nothing
    SelectDept = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SelectDoctor(ByVal lngDeptID As Long, ByVal lngItemID As Long, strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ҽ��
    '���:lngDeptID-����ID
    '     lngItemID-��ĿID
    '     strInput-������ҵ�ֵ
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-04 18:48:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset, intCount As Integer
    Dim lngҽ��ID As Long, strҽ������ As String, lng����ID As Long, str���ұ��� As String, str�������� As String
    Dim lng��Ŀid As Long, str��Ŀ���� As String, str��Ŀ���� As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    
    
    On Error GoTo errHandle
    
    strInput = UCase(strInput)  'ȫ����д����
    strCompents = Replace(gstrLike, "%", "*") & strInput & "*"
    
    If mrsDoctor Is Nothing Then
        Set mrsDoctor = LoadDoctorData(lngDeptID, lngItemID, mbytMode)
        If mrsDoctor Is Nothing Then Exit Function
    End If
    If mrsDoctor.State <> 1 Then
         Set mrsDoctor = LoadDoctorData(lngDeptID, lngItemID, mbytMode)
         If mrsDoctor Is Nothing Then Exit Function
    End If
    
    If mrsDoctor.RecordCount = 0 Then Exit Function
    
    If Trim(strInput) = "" Then GoTo GoSel:
    
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsDoctor)
           
            
    '��Ҫ����Ƿ��ж������������ļ�¼
    If IsNumeric(strInput) Then     '�������ȫ����
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strInput) Then     '�������ȫ��ĸ
        intInputType = 1
    Else
        intInputType = 2   ' 2-����
    End If
    
    With mrsDoctor
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not mrsDoctor.EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!���) = strInput Then
                    txtExDoctor.Text = Nvl(!����)
                    txtExDoctor.Tag = Val(Nvl(!ҽ��ID)) & ":" & Nvl(!����)
                    mstrPreDoctorName = Nvl(!����)
                    If txtExDept.Tag = "" Then
                        txtExDept.Text = Nvl(!���ұ���) & "-" & Nvl(!��������) & "��" & !��Ŀ���� & "��"
                        mlngPreDeptID = Val(Nvl(!����ID)): mlngPreItemID = Val(Nvl(!��ĿID))
                        txtExDept.Tag = mlngPreDeptID & ":" & mlngPreItemID
                    End If
                    SelectDoctor = True
                    Exit Function
                End If
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!���)) = Val(strInput) Then
                    If intCount = 0 Then
                        strҽ������ = Nvl(!����): lngҽ��ID = Val(Nvl(!ҽ��ID))
                        lng����ID = Val(Nvl(!����ID)): lng��Ŀid = Val(Nvl(!��ĿID))
                         str���ұ��� = Nvl(!���ұ���): str�������� = Nvl(!��������)
                         str��Ŀ���� = Nvl(!��Ŀ����): str��Ŀ���� = Nvl(!��Ŀ����)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Val(Nvl(!���)) Like strInput & "*" Then
                        Call zlDatabase.zlInsertCurrRowData(mrsDoctor, rsTemp)
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If UCase(Trim(Nvl(!����))) = strInput Then
                    If intCount = 0 Then
                        strҽ������ = Nvl(!����): lngҽ��ID = Val(Nvl(!ҽ��ID))
                        lng����ID = Val(Nvl(!����ID)): lng��Ŀid = Val(Nvl(!��ĿID))
                        str���ұ��� = Nvl(!���ұ���): str�������� = Nvl(!��������)
                        str��Ŀ���� = Nvl(!��Ŀ����): str��Ŀ���� = Nvl(!��Ŀ����)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.���ݲ�����ƥ����ͬ����
                If UCase(Trim(Nvl(!����))) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsDoctor, rsTemp)
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!���) = strInput Or UCase(Trim(!����)) = strInput Or UCase(Trim(!����)) = strInput Then
                    If intCount = 0 Then
                        strҽ������ = Nvl(!����): lngҽ��ID = Val(Nvl(!ҽ��ID))
                        lng����ID = Val(Nvl(!����ID)): lng��Ŀid = Val(Nvl(!��ĿID))
                        str���ұ��� = Nvl(!���ұ���): str�������� = Nvl(!��������)
                        str��Ŀ���� = Nvl(!��Ŀ����): str��Ŀ���� = Nvl(!��Ŀ����)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If Trim(!���) Like strInput & "*" Or UCase(Trim(Nvl(!����))) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsDoctor, rsTemp)
                End If
            End Select
            .MoveNext
        Loop
    End With
    
    If intCount > 1 Then strҽ������ = ""

GoSel:
    If Trim(strInput) = "" Then Set rsTemp = mrsDoctor
    If strҽ������ = "" And rsTemp.RecordCount = 1 Then
        rsTemp.MoveFirst
        strҽ������ = Nvl(rsTemp!����): lngҽ��ID = Val(Nvl(rsTemp!ҽ��ID))
        lng����ID = Val(Nvl(rsTemp!����ID)): lng��Ŀid = Val(Nvl(rsTemp!��ĿID))
        str���ұ��� = Nvl(rsTemp!���ұ���): str�������� = Nvl(rsTemp!��������)
        str��Ŀ���� = Nvl(rsTemp!��Ŀ����): str��Ŀ���� = Nvl(rsTemp!��Ŀ����)
    End If
    
    'ֱ�Ӷ�λ
    If strҽ������ <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        txtExDoctor.Text = strҽ������
        txtExDoctor.Tag = lngҽ��ID & ":" & strҽ������
        mstrPreDoctorName = strҽ������
        
        If txtExDept.Tag = "" And str�������� <> "" Then
            txtExDept.Text = str���ұ��� & "-" & str�������� & "��" & str��Ŀ���� & "��"
            txtExDept.Tag = lng����ID & ":" & lng��Ŀid
            mlngPreDeptID = lng����ID: mlngPreItemID = lng��Ŀid
        End If
        SelectDoctor = True
        Exit Function
    End If
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = "���"
    Case 1 '����ȫƴ��
        rsTemp.Sort = "����"
    Case Else
        rsTemp.Sort = "���"
    End Select
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "δ�ҵ�����������ҽ��", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtExDoctor, rsTemp, True, "", "ID,����ID,ҽ��ID", rsReturn) = False Then Exit Function
     If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
    
    If rsReturn Is Nothing Then Exit Function
    If rsReturn.State <> 1 Then Exit Function
    
    If rsReturn.RecordCount = 0 Then Exit Function
 
    
    txtExDoctor.Text = Nvl(rsReturn!����)
    txtExDoctor.Tag = Val(Nvl(rsReturn!ҽ��ID)) & ":" & Nvl(rsReturn!����)
    mstrPreDoctorName = Nvl(rsReturn!����)
    If txtExDept.Tag = "" And Nvl(rsReturn!��������) <> "" Then
        txtExDept.Text = Nvl(rsReturn!���ұ���) & "-" & Nvl(rsReturn!��������) & "��" & rsReturn!��Ŀ���� & "��"
        txtExDept.Tag = Val(Nvl(rsReturn!����ID)) & ":" & Val(Nvl(rsReturn!��ĿID))
        mlngPreDeptID = Val(Nvl(rsReturn!����ID)):  mlngPreItemID = Val(Nvl(rsReturn!��ĿID))
    End If
    
    rsReturn.Close: Set rsReturn = Nothing
    SelectDoctor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDefaultDoctor(ByVal lngDeptID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡҽ��
    '���:lngDeptID-����ID
    '����:ȱʡ�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-08 09:39:49
    '˵��:
    '   ȱʡ��ʽ:1.ֻ��һ��,ȱʡ���ҽ��;2.ȱʡ��һ��ѡ��ҽ��;3.ȱʡ��һ��ҽ��(�����)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngҽ��ID As Long, strҽ������ As String
    On Error GoTo errHandle
    txtExDoctor.Text = ""
    txtExDoctor.Tag = ""
    If mrsDoctor Is Nothing Then Exit Function
    If mrsDoctor Is Nothing Then Exit Function
    If mrsDoctor.RecordCount = 0 Then Exit Function
    If mrsDoctor.RecordCount = 1 Then
        txtExDoctor.Text = Nvl(mrsDoctor!����)
        txtExDoctor.Tag = Val(Nvl(mrsDoctor!ҽ��ID)) & ":" & Nvl(mrsDoctor!����)
        mstrPreDoctorName = Nvl(mrsDoctor!����)
        LoadDefaultDoctor = True: Exit Function
    End If
    If mstrPreDoctorName <> "" Then
        mrsDoctor.Filter = "����='" & mstrPreDoctorName & "'"
        If mrsDoctor.RecordCount <> 0 Then
            txtExDoctor.Text = Nvl(mrsDoctor!����)
            txtExDoctor.Tag = Val(Nvl(mrsDoctor!ҽ��ID)) & ":" & Nvl(mrsDoctor!����)
            mstrPreDoctorName = Nvl(mrsDoctor!����)
            mrsDoctor.Filter = 0
            LoadDefaultDoctor = True: Exit Function
        End If
    End If
    
    mrsDoctor.Filter = 0: mrsDoctor.Sort = "���": mrsDoctor.MoveFirst 'ȱʡ��һ��
    txtExDoctor.Text = Nvl(mrsDoctor!����)
    txtExDoctor.Tag = Val(Nvl(mrsDoctor!ҽ��ID)) & ":" & Nvl(mrsDoctor!����)
    mstrPreDoctorName = Nvl(mrsDoctor!����)
    
    LoadDefaultDoctor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDoctorData(ByVal lngDeptID As Long, ByVal lngItemID As Long, _
    Optional bytMode As Byte, Optional str��Ŀ���� As String, Optional str��Ŀ���� As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ�����ݼ�
    '���:lngDeptID-����ID,0ʱ����ʾ�����а���ҽ��
    '     lngItemID-��Ŀ
    '     bytMode-0-��ͨ;1-ԤԼȡ��;2-����ȡ��
    '����:����ҽ����
    '����:���˺�
    '����:2018-01-04 18:19:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsDeptDoctor As ADODB.Recordset, strDoctors As String, strTemp As String, i As Long
    Dim blnLoadDoctorFromDept As Boolean '�Ƿ����ȱʡ�Ŀ���ҽ��,�԰������Ƿ���ֻ���ŵ����ŵģ��У��ͼ��أ��޾���ҽ��Ϊ׼
    
    
    On Error GoTo errHandle
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "ID", adBigInt, , adFldIsNullable
            .Append "ҽ��ID", adBigInt, , adFldIsNullable
            .Append "���", adVarChar, 20, adFldIsNullable
            .Append "����", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 50, adFldIsNullable
            .Append "����ID", adBigInt, , adFldIsNullable
            .Append "���ұ���", adVarChar, 50, adFldIsNullable
            .Append "��������", adVarChar, 100, adFldIsNullable
            .Append "��ĿID", adBigInt, , adFldIsNullable
            .Append "��Ŀ����", adVarChar, 50, adFldIsNullable
            .Append "��Ŀ����", adVarChar, 200, adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    If bytMode = 1 Then blnLoadDoctorFromDept = True: GoTo gotoDoctor: 'ԤԼģʽ��ֻ��ȡ��������Ӧ��ҽ��
    
    
    If mrsRegData Is Nothing Then
        If ReadRegData = False Then Exit Function
    End If
    
    blnLoadDoctorFromDept = False
    mrsRegData.Filter = IIf(lngDeptID = 0, "", "����ID=" & lngDeptID & " And ��ĿID=" & lngItemID)
    strDoctors = "": i = 1
    
    With mrsRegData
        Do While Not .EOF
           
            strTemp = Val(Nvl(mrsRegData!ҽ��ID)) & ":" & Nvl(!ҽ������)
            If lngDeptID <> 0 And Val(Nvl(!ҽ��ID)) = 0 And Nvl(!ҽ������) = "" Then blnLoadDoctorFromDept = True ' ����ֻ���ŵ����ҵĺ�,���������Ҫ���ÿ��ҵĵ�ҽ��ȫ����ʾ��������ѡ��
            
            If InStr(strDoctors & ",", "," & strTemp & ",") = 0 And Nvl(mrsRegData!ҽ������) <> "" Then
                rsTemp.AddNew
                rsTemp!ID = i
                rsTemp!ҽ��ID = Val(Nvl(mrsRegData!ҽ��ID))
                rsTemp!��� = CStr(Nvl(mrsRegData!ҽ�����))
                rsTemp!���� = CStr(Nvl(mrsRegData!ҽ������))
                If Val(Nvl(mrsRegData!ҽ��ID)) = 0 Then
                    rsTemp!���� = zlCommFun.SpellCode(CStr(Nvl(mrsRegData!ҽ������)))
                Else
                    rsTemp!���� = CStr(Nvl(mrsRegData!ҽ������))
                End If
                rsTemp!����ID = Val(Nvl(mrsRegData!����ID))
                rsTemp!���ұ��� = CStr(Nvl(mrsRegData!���ұ���))
                rsTemp!�������� = CStr(Nvl(mrsRegData!��������))
                rsTemp!��ĿID = Val(Nvl(mrsRegData!��ĿID))
                rsTemp!��Ŀ���� = CStr(Nvl(mrsRegData!��Ŀ����))
                rsTemp!��Ŀ���� = CStr(Nvl(mrsRegData!��Ŀ����))
                rsTemp.Update
                i = i + 1
                strDoctors = strDoctors & "," & strTemp
            End If
            
            .MoveNext
        Loop
    End With
     mrsRegData.Filter = 0
gotoDoctor:
    
    If blnLoadDoctorFromDept Or bytMode = 2 Then
       '��ҽ��
       rsTemp.AddNew
       rsTemp!ID = i
       rsTemp!ҽ��ID = 0
       rsTemp!���� = ""
       rsTemp.Update
       i = i + 1
       If mobjRegister.zlGetDoctorFromDeptID(lngDeptID, rsDeptDoctor) Then  '���ݲ���ID����ȡ���漰��ҽ����
            With rsDeptDoctor
                Do While Not .EOF
                    strTemp = Val(Nvl(!ID)) & ":" & Nvl(!����)
                    If InStr(strDoctors & ",", "," & strTemp & ",") = 0 And Nvl(!����) <> "" Then
                        rsTemp.AddNew
                        rsTemp!ID = i
                        rsTemp!ҽ��ID = Val(Nvl(!ID))
                        rsTemp!��� = CStr(Nvl(!���))
                        rsTemp!���� = CStr(Nvl(!����))
                        If Val(Nvl(!ID)) = 0 Then
                            rsTemp!���� = zlCommFun.SpellCode(CStr(Nvl(!����)))
                        Else
                            rsTemp!���� = CStr(Nvl(!����))
                        End If
                        rsTemp!����ID = Val(Nvl(!����ID))
                        rsTemp!���ұ��� = CStr(Nvl(!���ұ���))
                        rsTemp!�������� = CStr(Nvl(!��������))
                        rsTemp!��ĿID = lngItemID
                        rsTemp!��Ŀ���� = str��Ŀ����
                        rsTemp!��Ŀ���� = str��Ŀ����
                        i = i + 1
                        rsTemp.Update
                        strDoctors = strDoctors & "," & strTemp
                    End If
                    .MoveNext
                Loop
            End With
       End If
    End If
    Set LoadDoctorData = rsTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Set LoadDoctorData = rsTemp
End Function

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    If txtExDept.Enabled And txtExDept.Visible Then
        txtExDept.SetFocus
    Else
        
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtExDept_Change()
    txtExDept.Tag = ""
    txtExDoctor.Text = ""
    Set mrsRooms = Nothing
    
End Sub
Private Sub txtExDept_GotFocus()
    zlControl.TxtSelAll txtExDept
End Sub

Private Sub txtExDept_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtExDept, KeyAscii, m�ı�ʽ)
    If KeyAscii <> 13 Then Exit Sub
    If txtExDept.Tag = "" Then
        If SelectDept(Trim(txtExDept.Text)) = False Then
            DoEvents
            If Me.Enabled Then Me.SetFocus
            If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
            zlControl.TxtSelAll txtExDept
            Exit Sub
        End If
    End If
    DoEvents
    If Me.Enabled Then Me.SetFocus
    If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
End Sub
Private Sub txtExDept_LostFocus()
    zlCommFun.OpenIme False
End Sub
Private Sub txtExDoctor_Change()
    txtExDoctor.Tag = ""
    Set mrsRooms = Nothing
    txtExRoom.Text = ""
End Sub

Private Sub txtExDoctor_GotFocus()
    zlControl.TxtSelAll txtExDoctor
End Sub

Private Sub txtExDoctor_KeyPress(KeyAscii As Integer)
    Dim lngItemID As Long, lngDeptID As Long, varTemp As Variant
    
    Call zlControl.TxtCheckKeyPress(txtExDoctor, KeyAscii, m�ı�ʽ)
    If KeyAscii <> 13 Then Exit Sub
    
    If txtExDoctor.Tag <> "" Then
        If txtExRoom.Enabled And txtExRoom.Visible Then
            txtExRoom.SetFocus: Exit Sub
        Else
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    End If
    varTemp = Split(txtExDept.Tag & ":", ":")
    lngDeptID = Val(varTemp(0))
    lngItemID = Val(varTemp(1))
        
    If SelectDoctor(lngDeptID, lngItemID, Trim(txtExDoctor.Text)) = False Then
        DoEvents
        If Me.Enabled Then Me.SetFocus
        If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
        zlControl.TxtSelAll txtExDoctor
        Exit Sub
    End If
    
    If mbytMode = 0 Then Call LoadRoomsData  '������������
    DoEvents
    If Me.Enabled Then Me.SetFocus
    If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
End Sub

Private Sub txtExDoctor_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txtExRoom_Change()
    txtExRoom.Tag = ""
End Sub

Private Sub txtExRoom_GotFocus()
    zlControl.TxtSelAll txtExRoom
End Sub

Private Sub txtExRoom_KeyPress(KeyAscii As Integer)

    Call zlControl.TxtCheckKeyPress(txtExRoom, KeyAscii, m�ı�ʽ)
    If KeyAscii <> 13 Then Exit Sub
    If txtExRoom.Tag = "" And Trim(txtExRoom.Text) <> "" Then
        If SelectRooms(Trim(txtExRoom.Text)) = False Then
            DoEvents
            If Me.Enabled Then Me.SetFocus
            Set mobjSetFocus = txtExRoom
            If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
            zlControl.TxtSelAll txtExRoom
            Exit Sub
        End If
    End If
    DoEvents
    If Me.Enabled Then Me.SetFocus
    Set mobjSetFocus = cmdGetNum
    If cmdGetNum.Enabled And cmdGetNum.Visible Then cmdGetNum.SetFocus
End Sub

Private Sub txtExRoom_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Function GetRegisterPlanID(ByRef lng����Id_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ��Ĳ��ţ�ҽ������ȡ���尲�ŵ�ID
    '���:
    '����:lng����ID_Out-����ID(�°汾Ϊ��¼ID)
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-08 14:18:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngDoctorID As Long, strҽ������ As String, varData As Variant
    Dim lngItemID As Long, lngDeptID As Long, varTemp As Variant
    Dim lng����ID As Long, lng�ƻ�ID As Long
    
    On Error GoTo errHandle
    If mbytMode = 1 Then
      If mrsBookData Is Nothing Then Exit Function
      If mrsBookData.State <> 1 Then Exit Function
      If mrsBookData.RecordCount = 0 Then Exit Function
      lng����Id_Out = Val(Nvl(mrsBookData!�����¼ID))
      If lng����Id_Out <> 0 Then GetRegisterPlanID = True: Exit Function
      
       GetRegisterPlanID = mobjRegister.zlGetRegisterPlanID_Tradition(Nvl(mrsBookData!�ű�), lng����Id_Out, lng�ƻ�ID)
       Exit Function
    End If
        
     
     
    If mrsRegData Is Nothing Then Exit Function
    
    varTemp = Split(txtExDept.Tag & ":", ":")
    
    lngDeptID = Val(varTemp(0))
    lngItemID = Val(varTemp(1))
     
    If txtExDept.Tag = "" Then Exit Function
    
    varData = Split(txtExDoctor.Tag & ":", ":")
    lngDoctorID = Val(varData(0))
    strҽ������ = varData(1)
    mrsRegData.Filter = "����ID=" & lngDeptID & " And ��ĿID=" & lngItemID
    
    With mrsRegData
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Val(Nvl(!ҽ��ID)) = lngDoctorID And Nvl(!ҽ������) = strҽ������ Then
                lng����Id_Out = Val(Nvl(mrsRegData!ID)): Exit Do
            End If
            If Val(Nvl(!ҽ��ID)) = 0 And Nvl(!ҽ������) = "" Then lng����ID = Val(Nvl(mrsRegData!ID))
            .MoveNext
        Loop
    End With
    
    mrsRegData.Filter = 0
    If lng����Id_Out = 0 Then lng����Id_Out = lng����ID
    If lng����Id_Out = 0 Then Exit Function
    GetRegisterPlanID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadRoomsData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������Ϣ��
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-08 14:11:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim lngItemID As Long, lngDeptID As Long, varTemp As Variant
    varTemp = Split(txtExDept.Tag & ":", ":")
    
    lngDeptID = Val(varTemp(0))
    lngItemID = Val(varTemp(1))
    
    On Error GoTo errHandle
    If GetRegisterPlanID(lng����ID) = False Then
         If mobjRegister.zlGetRegRoomsFromDeptid(lngDeptID, lngItemID, Trim(txtExDoctor.Text), mrsRooms) = False Then Set mrsRooms = Nothing: Exit Function
    Else
        If mobjRegister.zlGetRegRoomsFromPlanID(lng����ID, mrsRooms, mPara.blnBusy) = False Then Set mrsRooms = Nothing: Exit Function
    End If
    LoadRoomsData = True
    If mrsRooms Is Nothing Then Exit Function
    If mrsRooms.RecordCount = 0 Then Exit Function
    
    '����ȱʡֵ
    txtExRoom.Text = mrsRooms!����
    txtExRoom.Tag = mrsRooms!����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SelectRooms(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ������
    '���:strInput-Ϊ��ʱ����ʾ��ѯ���е�
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-05 16:35:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngRoomID As Long, str���� As String, str���� As String, intCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    On Error GoTo errHandle
    
    If mrsRooms Is Nothing Then Call LoadRoomsData  '�������Ҽ�
    
    strInput = UCase(Trim(strInput))
    If Trim(strInput) = "" Then GoTo GoSel:
    
    strCompents = Replace(gstrLike, "%", "*") & strInput & "*"
    
    If mrsRooms Is Nothing Then
        Call LoadRoomsData
        If mrsRooms Is Nothing Then
            MsgBox "δ�ҵ���������������", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    If mrsRooms.State <> 1 Then
         Call LoadRoomsData
        If mrsRooms Is Nothing Then
            MsgBox "δ�ҵ���������������", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If mrsRooms.RecordCount = 0 Then
        MsgBox "δ�ҵ���������������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'b.ID, b.����, b.����,b.����, b.λ��
    If mrsRooms.RecordCount = 1 Then
        lngRoomID = Val(Nvl(mrsRooms!ID))
        txtExRoom.Text = mrsRooms!����
        txtExRoom.Tag = mrsRooms!����
        SelectRooms = True
        Exit Function
    End If
     
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsRooms)
    '��Ҫ����Ƿ��ж������������ļ�¼
    If IsNumeric(strInput) Then     '�������ȫ����
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strInput) Then     '�������ȫ��ĸ
        intInputType = 1
    Else
        intInputType = 2   ' 2-����
    End If
    
    lngRoomID = 0
    With mrsRooms
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strInput Then
                    lngRoomID = Val(Nvl(!ID))
                    txtExRoom.Text = Nvl(!����)
                    txtExRoom.Tag = Nvl(!����)
                    SelectRooms = True
                    Exit Function
                End If
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strInput) Then
                    If intCount = 0 Then
                        str���� = Nvl(!����): lngRoomID = Val(Nvl(!ID))
                        str���� = Nvl(!����)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Val(Nvl(!����)) Like strInput & "*" Then
                        Call zlDatabase.zlInsertCurrRowData(mrsRooms, rsTemp)
                 End If
                 
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strInput Then
                    If intCount = 0 Then
                         str���� = Nvl(!����): lngRoomID = Val(Nvl(!ID))
                        str���� = Nvl(!����)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.���ݲ�����ƥ����ͬ����
                If UCase(Trim(Nvl(!����))) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsRooms, rsTemp)
                    intCount = intCount + 1
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strInput Or UCase(Trim(!����)) = strInput Or UCase(Trim(!����)) = strInput Then
                    If intCount = 0 Then
                        str���� = Nvl(!����): lngRoomID = Val(Nvl(!ID))
                        str���� = Nvl(!����)
                    End If
                    intCount = intCount + 1
                End If
                
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If Trim(Nvl(!����)) Like strInput & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                    Call zlDatabase.zlInsertCurrRowData(mrsRooms, rsTemp)
                    intCount = intCount + 1
                End If
            End Select
            .MoveNext
        Loop
    End With
    
    If intCount > 1 Then lngRoomID = 0
GoSel:
    If Trim(strInput) = "" Then Set rsTemp = mrsRooms
    If rsTemp Is Nothing Then
        MsgBox "δ�ҵ���������������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.State <> 1 Then
         MsgBox "δ�ҵ���������������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If lngRoomID = 0 And rsTemp.RecordCount = 1 Then
        rsTemp.MoveFirst
        str���� = Nvl(rsTemp!����): lngRoomID = Val(Nvl(rsTemp!ID))
        str���� = Nvl(rsTemp!����)
    End If
    
    'ֱ�Ӷ�λ
    If lngRoomID <> 0 Then
        If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
        txtExRoom.Text = str����
        txtExRoom.Tag = str����
        SelectRooms = True
        Exit Function
    End If
    

    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = "����"
    Case Else
        rsTemp.Sort = "����"
    End Select
    If rsTemp.RecordCount = 0 Then
        MsgBox "δ�ҵ���������������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtExRoom, rsTemp, True, "", "ID", rsReturn) = False Then Exit Function
    If Trim(strInput) <> "" Then rsTemp.Close: Set rsTemp = Nothing
    
    If rsReturn Is Nothing Then Exit Function
    If rsReturn.RecordCount = 0 Then Exit Function
    
    lngRoomID = Val(Nvl(rsReturn!ID))
    txtExRoom.Text = Nvl(rsReturn!����)
    txtExRoom.Tag = Nvl(rsReturn!����)
    rsReturn.Close: Set rsReturn = Nothing
    SelectRooms = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckDataValied(ByRef lng����Id_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ݵĺϷ���
    '���:
    '����:lng����Id_Out-���ص�ǰ����ID
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-08 14:54:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnExist As Boolean
    Dim strErrMsg As String
    
    lng����Id_Out = 0
    
    On Error GoTo errHandle
    If mobjPati Is Nothing Then
       MsgBox "δѡ���ˣ�����ȡ��!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Function
    End If
    If mobjPati.����ID = 0 Or PatiIdentify.Text = "" Then
        MsgBox "δѡ���ˣ�����ȡ��!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Function
    End If
    
    If txtExDept.Tag = "" Then
        MsgBox "δѡ����Ҫȡ�ŵĿ��ң�����ȡ��!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        Exit Function
    End If
    If txtExDoctor.Tag = "" And txtExDoctor.Text <> "" Then
        MsgBox "ҽ��ѡ�����,��ѡ����ȷ��ҽ��!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
        Exit Function
    End If
    If txtExRoom.Tag = "" And txtExRoom.Text <> "" Then
        MsgBox "����ѡ�����,��ѡ����ȷ������!", vbInformation + vbOKOnly, gstrSysName
        Call LockedScreen(False)
        If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
        Exit Function
    End If
    
    If Not mrsRooms Is Nothing Then
        If mrsRooms.RecordCount <> 0 And txtExDept.Tag = "" Then
            MsgBox "�㻹δѡ������,������ȡ�ţ�!", vbInformation + vbOKOnly, gstrSysName
            Call LockedScreen(False)
            If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
            Exit Function
        End If
    End If

    '���������
    If CheckPatiCheck(mobjPati.����ID) = False Then
        Call LockedScreen(False)
        If PatiIdentify.Enabled And PatiIdentify.Visible Then PatiIdentify.SetFocus
        Exit Function
    End If
    
    If mbytMode = 0 Then
        '����Ƿ���ڰ���
        blnExist = GetRegisterPlanID(lng����Id_Out)
    End If
    
    
    If blnExist Then blnExist = lng����Id_Out <> 0
    If Not blnExist Then
        If txtExDoctor.Tag <> "" Then
            MsgBox "δ�ҵ�����Ϊ" & txtExDept.Text & "��ҽ��Ϊ" & txtExDoctor.Text & " �İ��ţ�����ȡ��!", vbInformation + vbOKOnly, gstrSysName
            Call LockedScreen(False)
            If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        Else
            MsgBox "δ�ҵ�����Ϊ" & txtExDept.Text & "�İ��ţ�����ȡ��!", vbInformation + vbOKOnly, gstrSysName
             Call LockedScreen(False)
            If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        End If
        Exit Function
    End If
    
    '����Ƿ񳬺�
    If mobjRegister.zlRegisterCheckValied(mobjPati.����ID, lng����Id_Out, strErrMsg) = False Then
        If strErrMsg <> "" Then
            ShowMsgbox txtExDept.Text & " " & strErrMsg & ",��ѡ���������Ҿ���!"
            Call LockedScreen(False)
            DoEvents
            If Me.Enabled Then Me.SetFocus
            If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
            Exit Function
        End If
    End If
    
    CheckDataValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call LockedScreen(False)
End Function

Public Function SaveHzGetNum(ByRef strNo_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ȡ��
    '����:ȡ�ųɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-16 16:17:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�Һ�ID As Long, lng����ID As Long, lng��Ŀid As Long, str���� As String, strҽ�� As String, lngҽ��ID As Long
    Dim blnYes As Boolean, strSQL As String, varTemp As Variant
        
    On Error GoTo errHandle
    
    If mrsBookData Is Nothing Then Exit Function
    If mrsBookData.State <> 1 Then Exit Function
    If mrsBookData.RecordCount = 0 Then Exit Function
    mrsBookData.MoveFirst
    lng�Һ�ID = Val(Nvl(mrsBookData!�Һ�ID))
    strNo_Out = Nvl(mrsBookData!NO)
    
    If Val(Nvl(mrsBookData!��¼��־)) <> 2 Then
        MsgBox "�Һŵ�Ϊ" & mrsBookData!NO & "���ǻ��ﵥ�ݣ�����ȡ��!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    varTemp = Split(txtExDept.Tag & ":", ":")
    lng����ID = Val(varTemp(0))
    lng��Ŀid = Val(varTemp(1))
     
    str���� = txtExRoom.Text
    strҽ�� = txtExDoctor.Text
    lngҽ��ID = Val(Split(txtExDoctor.Tag & ":", ":")(0))
    
    If lng����ID = -1 Then
        MsgBox "��ȷ��Ҫ����Ŀ��ҡ�", vbInformation, gstrSysName
        If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        Exit Function
    End If
    If lng����ID <> mrsBookData!����ID Then
        If MsgBox("ע��:" & vbCrLf & "  ��ѡ��Ŀ��������Ŀ��Ҳ�һ��,���Ƿ�Ҫ�������˻������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If txtExDept.Enabled And txtExDept.Visible Then txtExDept.SetFocus
        Exit Function
        End If
        blnYes = True
    End If
    If str���� <> Nvl(mrsBookData!����) And blnYes = False Then
        If MsgBox("ע��:" & vbCrLf & "  ��ѡ����������������Ҳ�һ��,���Ƿ�Ҫ�������˵Ļ�������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtExRoom.Enabled And txtExRoom.Visible Then txtExRoom.SetFocus
            Exit Function
        End If
        blnYes = True
    End If
    
    If strҽ�� <> Nvl(mrsBookData!ҽ������) And blnYes = False Then
        If MsgBox("ע��:" & vbCrLf & "  ��ѡ���ҽ��������ҽ����һ��,���Ƿ�Ҫ�������˵Ļ���ҽ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtExDoctor.Enabled And txtExDoctor.Visible Then txtExDoctor.SetFocus
            Exit Function
        End If
        blnYes = True
    End If
    
    'Zl_���˹Һż�¼_����
    strSQL = "Zl_���˹Һż�¼_����("
    '  Id_In         ���˹Һż�¼.ID%Type,
    strSQL = strSQL & "" & lng�Һ�ID & ","
    '  ��ִ�п���_In ���˹Һż�¼.ִ�в���id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ������_In     ���˹Һż�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ��ҽ��_In     ���˹Һż�¼.ִ����%Type,
    strSQL = strSQL & "'" & strҽ�� & "',"
    '  �����_In Integer:=0
    strSQL = strSQL & "0)"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveHzGetNum = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
 End Function

Private Function SaveBooking(ByRef strNo_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ԤԼ����ȡ��
    '���:
    '����:strNo_Out-����ȡ�ŵĵ��ݺ�
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-15 17:31:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�Һ�ID As Long
    On Error GoTo errHandle
    If mrsBookData Is Nothing Then Exit Function
    If mrsBookData.State <> 1 Then Exit Function
    If mrsBookData.RecordCount = 0 Then Exit Function
    
    mrsBookData.MoveFirst
    strNo_Out = Nvl(mrsBookData!NO)
    If strNo_Out = "" Then Exit Function
    If mrsBookData!��¼״̬ = 1 Then
        '������ȡ�ţ�ֱ�Ӳ���ǩ����ʽ
        lng�Һ�ID = Val(Nvl(mrsBookData!�Һ�ID))
 
        ' Zl_���˹Һż�¼_ǩ��
        strSQL = "Zl_���˹Һż�¼_ǩ��("
        '  Id_In       ���˹Һż�¼.Id%Type,
        strSQL = strSQL & "" & lng�Һ�ID & ","
        '  ��������_In Integer := 0,
        strSQL = strSQL & "" & 0 & ","
        '  ԤԼ��ʽ_In ԤԼ��ʽ.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In     ���˹Һż�¼.����%Type := Null,
        strSQL = strSQL & "'" & txtExRoom.Text & "',"
        '  ҽ��_In     ���˹Һż�¼.ִ����%Type := Null
        strSQL = strSQL & "'" & txtExDoctor.Text & "')"
 
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        SaveBooking = True: Exit Function
    End If
    '    Zl_����ԤԼ����_ȡ��
    strSQL = "Zl_����ԤԼ����_ȡ��("
    '  No_In         ������ü�¼.No%Type
    strSQL = strSQL & "'" & strNo_Out & "',"
    '  ����_In       ������ü�¼.��ҩ����%Type,
    strSQL = strSQL & "'" & txtExRoom.Text & "',"
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mobjPati.����ID & ","
    '  ҽ������_In   ������ü�¼.ִ���� %Type,
    strSQL = strSQL & "'" & txtExDoctor.Text & "',"
    '  ����Ա���_In ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
    strSQL = strSQL & "sysdate,"
    '  ժҪ_In       ���˹Һż�¼.ժҪ%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����_In       ���˹Һż�¼.����%Type := Null
    strSQL = strSQL & "NULL)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveBooking = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function SaveData(ByVal lng����ID As Long, ByRef strNo_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���(ֻ�Ǵ滮�Һŷ�0�ļ�¼)
    '���:lng����ID-��ǰ�İ���ID
    '����:strNo_out-����ɹ��󣬷��صĵ��ݺ�
    '����:����ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-08 16:38:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNO As String
    
    
    strNO = zlDatabase.GetNextNo(12)
    
    On Error GoTo errHandle
    
    'Zl_�������ȡ��_Insert
    strSQL = "Zl_�������ȡ��_Insert("
    '  ����id_In     ������Ϣ.����id%Type,
    strSQL = strSQL & "" & mobjPati.����ID & ","
    '  ��¼id_In     �ٴ������¼.Id%Type,
    strSQL = strSQL & "" & IIf(mbytRegMode = 1, lng����ID, "NULL") & ","
    '  ����id_In     �ҺŰ���.Id%Type,
    strSQL = strSQL & "" & IIf(mbytRegMode <> 1, lng����ID, "NULL") & ","
    '  ���ݺ�_In     ���˹Һż�¼.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  ����_In       ��������.����%Type,
    strSQL = strSQL & "" & IIf(Trim(txtExRoom.Text) = "", "NULL", "'" & Trim(txtExRoom.Text) & "'") & ","
    '  ҽ������_In   �ҺŰ���.ҽ������%Type,
    strSQL = strSQL & "" & IIf(Trim(txtExDoctor.Text) = "", "NULL", "'" & Trim(txtExDoctor.Text) & "'") & ","
    '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
    strSQL = strSQL & "" & IIf(Trim(txtExDoctor.Tag) = "", "NULL", "'" & Val(Split(txtExDoctor.Tag & ":", ":")(0)) & "'") & ","
    '  ��������id_In ������ü�¼.��������id%Type,
    strSQL = strSQL & "" & UserInfo.����ID & ","
    '  ����Ա���_In ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �˺�����_In   Integer := 0,
    strSQL = strSQL & "0,"
    '  վ��_In Varchar2:=Null
    strSQL = strSQL & "" & IIf(Trim(gstrNodeNo) = "", "NULL", "'" & Trim(gstrNodeNo) & "'") & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    strNo_Out = strNO
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function PrintBill(ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ����
    '���:strNo-�Һŵ���
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-09 15:39:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    Select Case Val(zlDatabase.GetPara("�Һ�ƾ����ӡ��ʽ", glngSys, 9000, "0"))
        Case 0    '����ӡ
           Exit Function
        Case 1    '�Զ���ӡ
        Case 2    'ѡ���ӡ
            If MsgBox("Ҫ��ӡȡ��ƾ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    End Select
    strSQL = "select ID From ���˹Һż�¼ where NO =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then
        MsgBox "δ�ҵ����ݺ�Ϊ" & strNO & "��ȡ�ż�¼,����"
    End If
    
    '�ݶ�Ϊ�����ظ���ӡ
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1113", Me, "�Һ�ID=" & Val(Nvl(rsTemp!ID)), "��Ʊ��=��", 2)
    PrintBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SelectBooking(ByVal lng����ID As Long, ByRef strNo_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ԤԼ����
    '���:lng����ID-ָ������ID
    '����:strNO_out-����ԤԼ���嵥�ݺ�
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-15 16:29:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmNew As frmSelRegist, rsTemp As ADODB.Recordset, strNO As String
    Dim lng����ID1 As Long
    
    On Error GoTo errHandle
    Set mrsBookData = Nothing
    
    lblBookingNO.Visible = False    '����ԤԼ��
    
    'bytType-0-������ԤԼ����;1-�������Ѿ�֧����û��ǩ����ԤԼ��;2-���������ﲡ��;3-����(0,1,2)
    If mobjRegister.zlGetRegisterBookData(lng����ID, rsTemp, , mstr�������, 3) = False Then Exit Function
    
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.EOF Then Exit Function
    
    strNo_Out = Nvl(rsTemp!NO)
    lng����ID1 = Val(Nvl(rsTemp!����ID))
    
    If rsTemp.RecordCount > 1 Then
        '���һ��ﲡ��
        rsTemp.Filter = "��¼��־=2"
        If rsTemp.EOF = False Then strNo_Out = Nvl(rsTemp!NO): GoTo LoadPati:
        rsTemp.Filter = "��¼״̬<>0 "  '��ȡ�Ѿ����ѵ�
        If rsTemp.EOF = False Then strNo_Out = Nvl(rsTemp!NO): GoTo LoadPati:
        
         '���ò����Ƿ���ԤԼ����
        rsTemp.Filter = 0
        Set frmNew = New frmSelRegist
        If frmNew.ShowRegist(Me, mstrPrivs, False, mPara.intԤԼʧЧ����, strNo_Out, rsTemp, lng����ID, 1, mstr�������) = False Then
            If Not frmNew Is Nothing Then Unload frmNew
            Set frmNew = Nothing
            Exit Function
        End If
        If Not frmNew Is Nothing Then Unload frmNew
        Set frmNew = Nothing
        lng����ID1 = Val(Nvl(rsTemp!����ID))
    End If
    
LoadPati:
    If lng����ID <> lng����ID1 Then
        '���ز�����Ϣ
        If GetPatient(PatiIdentify.GetCurCard, "-" & Val(Nvl(rsTemp!����ID)), False, mobjPati) = False Then Exit Function
        cmdNewPati.ToolTipText = "�޸Ĳ���(F4)"
    End If
    
    '����ԤԼ��
    If ReadBooking(strNo_Out) = False Then Exit Function

    SelectBooking = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function ReadBooking(ByVal strNO As String, Optional blnReadPati As Boolean, Optional objPati As PatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤԼ����
    '���: strNo-ԤԼ����,Ϊ��ʱ����ʾ���ݲ���ID�����ԤԼ����,�������ԤԼ��������ԤԼ����
    '     blnReadPati-�Ƿ���Ҫ���ݹҺŵ��еĲ�����Ϣ���¶�ȡ����Ϣ
    '����:objPati-���ز�����Ϣ,blnReadPatiΪtrueʱ������,����ΪNothing
    '����:��ȡ�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-15 15:39:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean, dtDate As Date
    On Error GoTo errHandle
    
    
    If mobjRegister.zlGetRegisterBookData(0, mrsBookData, strNO, , 3) = False Then Exit Function
    
    If mrsBookData Is Nothing Then
        If strNO <> "" Then MsgBox "ԤԼ��:" & strNO & "������,���ܱ������˽��ջ���Ч��ԤԼ����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBookData.EOF Then
        If strNO <> "" Then MsgBox "ԤԼ��:" & strNO & "������,���ܱ������˽��ջ���Ч��ԤԼ����!", vbInformation + vbOKOnly, gstrSysName
        Set mrsBookData = Nothing
        Exit Function
    End If
    
    
    mbytMode = IIf(Val(Nvl(mrsBookData!��¼��־)) = 2, 2, 1)
       
    If InStr("," & mstr������� & ",", "," & Val(Nvl(mrsBookData!����ID)) & ",") = 0 And mstr������� <> "" Then
        MsgBox strNO & "��" & IIf(mbytMode <> 2, "ԤԼ", "����") & "���ݲ����ڱ�����̨ȡ�ţ����飡", vbInformation + vbOKOnly, gstrSysName
        Set mrsBookData = Nothing
        mbytMode = 0
        Exit Function
    End If
    
    If mPara.intԤԼ��Чʱ�� <> 0 And mbytMode <> 2 Then
        dtDate = DateAdd("n", 1 * mPara.intԤԼ��Чʱ��, zlDatabase.Currentdate)
        If Format(dtDate, "yyyy-MM-dd hh:mm:ss") > Format(mrsBookData!ԤԼʱ��, "yyyy-MM-dd hh:mm:ss") Then
           dtDate = DateAdd("n", -1 * mPara.intԤԼ��Чʱ��, CDate(Format(mrsBookData!ԤԼʱ��, "yyyy-MM-dd hh:mm:ss")))
           MsgBox "��ԤԼ���ѹ�ԤԼ������ʱ�� " & Format(dtDate, "yyyy-MM-dd hh:mm:00") & ",���ܽ���", vbInformation, gstrSysName
           Set mrsBookData = Nothing: mbytMode = 0
           Exit Function
        End If
    End If
    
    If blnReadPati Then
        If GetPatient(PatiIdentify.GetCurCard, "-" & Val(Nvl(mrsBookData!����ID)), False, objPati) = False Then Exit Function
    End If
    
    
    '�ݲ����ڽ���ʱ���ŵĴ���
    'If ReadRegData = False Then
        Call CreateDeptStructure    '�����ȡʧ�ܣ�ֻ��ԤԼ�Һŵ��Ŀ���
    'End If
    If mrsDept Is Nothing Then Exit Function
    If mrsDept.State <> 1 Then Exit Function
     
    mbytMode = IIf(Val(Nvl(mrsBookData!��¼��־)) = 2, 2, 1)
    If mbytMode <> 2 Then
        lblBookingNO.Caption = "ԤԼ��:" & strNO
    Else
        lblBookingNO.Caption = "����(����:" & strNO & ")"
    End If
    lblBookingNO.Left = lblPati.Left + lblPati.Width + 200
    If mbytMode = 2 Then
        '�������ѡ�����
        Call ReadRegData
    End If
    
    mrsDept.Filter = "����ID=" & Val(Nvl(mrsBookData!����ID)) & " And ��ĿID=" & Val(Nvl(mrsBookData!��ĿID))
    If mrsDept.EOF Then '���ϱ����ҵ����ѡ��
        mrsDept.Filter = 0
        mrsDept.AddNew
        mrsDept!ID = mrsDept.RecordCount + 1
        mrsDept!����ID = Val(Nvl(mrsBookData!����ID))
        mrsDept!���� = CStr(Nvl(mrsBookData!���ұ���))
        mrsDept!���� = CStr(Nvl(mrsBookData!��������))
        mrsDept!���� = CStr(Nvl(mrsBookData!���Ҽ���))
        
        mrsDept!��ĿID = CStr(Nvl(mrsBookData!��ĿID))
        mrsDept!��Ŀ���� = CStr(Nvl(mrsBookData!��Ŀ����))
        mrsDept!��Ŀ���� = CStr(Nvl(mrsBookData!�Һ���Ŀ))
        
        mrsDept!�Ƿ�ԭ���� = 1
        mrsDept.Update
    End If

    '����ȱʡ����
    txtExDept.Text = mrsDept!���� & "-" & mrsDept!���� & "��" & mrsDept!��Ŀ���� & "��"
    txtExDept.Tag = Val(Nvl(mrsDept!����ID)) & ":" & Val(Nvl(mrsDept!��ĿID))
    
    '����ȱʡҽ��
    Set mrsDoctor = LoadDoctorData(Val(Nvl(mrsDept!����ID)), Val(Nvl(mrsDept!��ĿID)), mbytMode)
    mrsDept.Filter = 0
    
    txtExDoctor.Text = Trim(Nvl(mrsBookData!ҽ������))
    txtExDoctor.Tag = Val(Nvl(mrsBookData!ҽ��ID)) & ":" & Trim(Nvl(mrsBookData!ҽ������))
    
    Call LoadRoomsData  '��������
    '��������
    txtExRoom.Text = Nvl(mrsBookData!����)
    txtExRoom.Tag = Nvl(mrsBookData!����)
    lblBookingNO.Visible = True
    
    
    txtExDept.Enabled = mbytMode = 2: cmdExDept.Enabled = mbytMode = 2: cmdExDept.Tag = IIf(mbytMode = 2, "", "F")
    
    
    blnEnabled = txtExDoctor.Text = "" Or mbytMode = 2
    If mbytMode <> 2 Then
        If blnEnabled Then
            If mrsDoctor Is Nothing Then
                blnEnabled = False
            ElseIf mrsDoctor.State <> 1 Then
                blnEnabled = False
            ElseIf mrsDoctor.RecordCount = 0 Then
                 blnEnabled = False
            End If
        End If
    End If
    txtExDoctor.Enabled = blnEnabled: cmdExDoctor.Enabled = blnEnabled: cmdExDoctor.Tag = IIf(blnEnabled, "", "F")
    
    
    blnEnabled = txtExRoom.Text <> "" Or mbytMode = 2
    If mbytMode <> 2 Then
        If blnEnabled Then
            If mrsRooms Is Nothing Then
                blnEnabled = False
            ElseIf mrsRooms.State <> 1 Then
                blnEnabled = False
            ElseIf mrsRooms.RecordCount = 0 Then
                 blnEnabled = False
            End If
        End If
    End If
    txtExRoom.Enabled = blnEnabled: cmdExRoom.Enabled = blnEnabled: cmdExRoom.Tag = IIf(blnEnabled, "", "F")
    

    ReadBooking = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function





