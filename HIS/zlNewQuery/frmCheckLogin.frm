VERSION 5.00
Begin VB.Form frmCheckLogin 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "������֤"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6420
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCheckLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer tmrError 
      Interval        =   6000
      Left            =   0
      Top             =   0
   End
   Begin zl9NewQuery.ctlButton ctlCancel 
      Height          =   720
      Left            =   4500
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1270
      Caption         =   "ȡ��"
      AutoSize        =   0   'False
      ButtonHeight    =   600
   End
   Begin VB.Timer Time 
      Interval        =   4000
      Left            =   1170
      Top             =   2700
   End
   Begin VB.TextBox TxtCardID 
      Height          =   435
      Left            =   2595
      TabIndex        =   0
      Top             =   1830
      Width           =   3135
   End
   Begin VB.TextBox Txtpwd 
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2595
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2430
      Width           =   3135
   End
   Begin zl9NewQuery.ctlButton ctlOK 
      Height          =   720
      Left            =   2745
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1270
      Caption         =   "ȷ��"
      AutoSize        =   0   'False
      ButtonHeight    =   600
   End
   Begin zl9NewQuery.ctlButton ctlReset 
      Height          =   720
      Left            =   960
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1270
      Caption         =   "����"
      AutoSize        =   0   'False
      ButtonHeight    =   600
   End
   Begin VB.Label Lblreg 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCheckLogin.frx":1E26
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   2025
      TabIndex        =   7
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Lblinfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   450
      Left            =   105
      TabIndex        =   6
      Top             =   405
      Width           =   5685
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ��ĹҺ���ĿΪ��"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1290
      TabIndex        =   5
      Top             =   30
      Width           =   3105
   End
   Begin VB.Label LBLErr 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "�������,����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2010
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label lblCardID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����  "
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1845
      TabIndex        =   3
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label Lblpwd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1845
      TabIndex        =   2
      Top             =   2475
      Width           =   660
   End
   Begin VB.Image Imgbak 
      Height          =   2130
      Left            =   180
      Picture         =   "frmCheckLogin.frx":1E64
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1605
   End
End
Attribute VB_Name = "frmCheckLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private Type PARAM_IN
    RegisterMode As String                  '��ǰ���õĹҺŷ�ʽ
    Depart As String
    RegisterItem As String
    DoctorName As String
    DoctorID As Long
    ����ID As Long
    BillNo As String
    RegisterPrice As Double
    DetailID As Long
    DepartID As Long
    �ű� As String
End Type
Private mParamIn As PARAM_IN

Private mCurPayNeed As Currency          '������Ҫ�ķ���
Private mCurLeft As Currency
Private mlngTime As Long

Private Type PATIENTINFO
    PatientID As String          '���˵�ID
    Name As String
    Sex As String                '��¼���˵��������Ա�
    DoorPost As String
    Age As String                '��¼���˵�����ź�����
    FareClass As String
    strIDCard As String           '���֤��
    str�������� As String
    str������ַ As String
    str���� As String
End Type
Private mPatient As PATIENTINFO
Private mBrushIDCardPatiInfor As PATIENTINFO    'ˢ��ʱ�Ĳ�����Ϣ
Private mlng�����ID As Long 'ҽ�ƿ����ID ˢҽ�ƿ�ʱ����
Private mblnCanCommit As Boolean         '�Ƿ��ܹ��ύ����
Private mblnCharge As Boolean
Private mblnNoChange As Boolean
Private mblnBrushCard As Boolean  'ˢ��

Private mobjICCard As Object 'IC������
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1

'######################################################################################################################

Public Function ShowLogin(ByVal frmMain As Object, ByVal strRegisterMode As String, _
                            ByVal StrDepart As String, ByVal strRegisterItem As String, ByVal StrDoctorName As String, ByVal lng����ID As Long, ByVal strBillNo As String, _
                            ByVal lngDoctorID As Long, ByVal dbRegisterPrice As Double, ByVal lngDetailID As Long, ByVal lngDepartID As Long, ByVal str�ű� As String, _
                            Optional ByVal lng�����ID As Long) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    mParamIn.RegisterMode = strRegisterMode
    mParamIn.Depart = StrDepart
    mParamIn.RegisterItem = strRegisterItem
    mParamIn.DoctorName = StrDoctorName
    mParamIn.����ID = lng����ID
    mParamIn.BillNo = strBillNo
    mParamIn.DoctorID = lngDoctorID
    mParamIn.RegisterPrice = dbRegisterPrice
    mParamIn.DetailID = lngDetailID
    mParamIn.DepartID = lngDepartID
    mParamIn.�ű� = str�ű�
    mlng�����ID = lng�����ID
    
    Me.Show 1, frmMain
    ShowLogin = True
    
End Function

Private Function ShowErrorInfo(ByVal strError As String) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim strErrorInfo As String
    
    mblnNoChange = True

    Select Case strError
    Case "����"
        strErrorInfo = "ע�⣺������㣬���Ƚɷ�"
    Case "������Ժ"
        strErrorInfo = "ע�⣺���Ѿ���Ժ,���ܹҺ�"
    Case "�����֤"
        
        Select Case mParamIn.RegisterMode
        Case "���֤�Һ�"
            strErrorInfo = "������֤������������ԣ�"
        Case "���￨�Һ�"
            strErrorInfo = "��Ŀ��Ż�������������ԣ�"
        Case "�ɣÿ��Һ�"
            strErrorInfo = "������֤������������ԣ�"
        End Select
    Case "���￨"
            strErrorInfo = "������Ϣ������"
    End Select
    
    TxtCardID.Text = ""
    Txtpwd.Text = ""
    LBLErr.Caption = strErrorInfo
    LBLErr.Visible = True
    Lblreg.Visible = False
    If TxtCardID.Enabled And TxtCardID.Visible Then TxtCardID.SetFocus
            
    mblnNoChange = False
    
End Function

Private Function ShowPatientInfo(ByVal lng����ID As Long) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
        
    '������������ʼ��
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    mlngTime = Val(GetPara("������֤����ͣ��ʱ��"))
    mPatient.Name = "null"
    mPatient.Sex = "nul"
    mPatient.DoorPost = "null"
    mPatient.Age = "null"
    mPatient.FareClass = "null"
    mPatient.PatientID = lng����ID
    
    '�����˵ļ���Ϣ���
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select ����,�Ա�,�����,����,�ѱ�,Trunc(��������) as �������� from ������Ϣ where ����ID=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mPatient.PatientID))
    If rs.BOF = False Then
        mPatient.Name = zlCommFun.Nvl(rs("����").Value)
        mPatient.DoorPost = zlCommFun.Nvl(rs("�����").Value, 0)
        mPatient.Sex = zlCommFun.Nvl(rs("�Ա�").Value)
        mPatient.Age = zlCommFun.Nvl(rs("����").Value)
        mPatient.FareClass = zlCommFun.Nvl(rs("�ѱ�").Value)
    End If

    '��һЩ�ؼ��Ŀɼ����Ըı�
    '------------------------------------------------------------------------------------------------------------------
    ctlOK.Visible = True
    '68550,������,2014-01-08,������ť��ʾ���ݴ��������
    ctlOK.Caption = "ȷ��"
    ctlReset.Visible = False
    TxtCardID.Visible = False
    Lblpwd.Visible = False
    Txtpwd.Visible = False
    Lblreg.Caption = Chr(10) + Chr(13) + "��������ѡ���밴��ȡ������" + Chr(10) + Chr(13) + "�����йҺţ��밴��ȷ�ϡ�"
    If mPatient.Age <> "" And mPatient.Age <> "0" Then strTmp = "/" + mPatient.Age Else mPatient.Age = ""
    lblCardID.Caption = "�����Ϣ:" + mPatient.Name + "/" + mPatient.Sex + strTmp
    
    ShowPatientInfo = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData() As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    
    '���Һŵ���Ϣ������ʾ

    lblInfo.Caption = mParamIn.Depart + "/" + mParamIn.RegisterItem + "/" + mParamIn.DoctorName
    
    ctlOK.Visible = False
    ctlReset.Visible = True
    mblnCanCommit = False
    Txtpwd.Text = ""
    TxtCardID.Text = ""
    
    Select Case mParamIn.RegisterMode
    Case "���֤�Һ�"
        Txtpwd.Visible = False
        Lblpwd.Visible = False
        ctlReset.Visible = False
        ctlOK.Visible = False
        Lblreg.Caption = "��������ѡ���밴��ȡ��������������ȷ�������֤��"
    Case "�ɣÿ��Һ�"
        Txtpwd.Visible = False
        Lblpwd.Visible = False
        ctlReset.Visible = False
        ctlOK.Visible = True
        ctlOK.Caption = "����"
        Lblreg.Caption = "��������ѡ���밴��ȡ��������������ȷ���ãɣÿ�������������"
    Case Else
        Txtpwd.Visible = True
        Lblpwd.Visible = True
        ctlReset.Visible = True
        ctlOK.Visible = True
        ctlOK.Caption = "ȷ��"
        Lblreg.Caption = "��������ѡ���밴��ȡ��������ȷ������ˢ��������������"
        If mParamIn.RegisterMode = "���￨�Һ�" Then Me.Txtpwd.MaxLength = 0
    End Select
    
    If GetPara("������ʾ����") = "1" Then
        TxtCardID.PasswordChar = "*"
    End If

End Function

Private Function CheckIdentify(ByVal strMode As String, ByRef lng����ID As Long, Optional ByVal strUser As String, Optional ByVal strPsw As String) As Boolean
    '******************************************************************************************************************
    '����:�����֤
    '����:
    '����:
    '******************************************************************************************************************
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strDentify As String
    Dim varAry As Variant
    
    On Error GoTo errHand
    
    Select Case strMode
    '------------------------------------------------------------------------------------------------------------------
    Case "ҽ�����Һ�"
        
        strDentify = gclsInsure.Identify2(UCase(strUser), strPsw, 3, , gintInsure)
        If strDentify = "" Then Exit Function

        varAry = Split(strDentify, ";")
        If UBound(varAry) >= 8 Then
            lng����ID = Val(varAry(8))
        Else
            Exit Function
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case "���￨�Һ�"
        
        
        If strPsw = "" And mlng�����ID = 0 Then
            strSQL = "Select ����ID From ������Ϣ Where ���￨�� = [1] And ����֤�� Is Null"
        ElseIf strPsw <> "" And mlng�����ID = 0 Then
            strSQL = "Select ����ID From ������Ϣ Where ���￨�� = [1] And ����֤�� = [2]"
        ElseIf strPsw = "" And mlng�����ID <> 0 Then
            strSQL = "Select ����ID From ����ҽ�ƿ���Ϣ Where ���� = [1] And ���� Is Null And �����id= " & mlng�����ID
        Else
            strSQL = "Select ����ID From ����ҽ�ƿ���Ϣ Where ���� = [1] And ���� = [2] And �����id= " & mlng�����ID
        End If
                
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strUser), strPsw)
        If rs.BOF Then Exit Function
        lng����ID = rs("����id").Value
        
    '------------------------------------------------------------------------------------------------------------------
    Case "���֤�Һ�"
        
        strSQL = "Select  ����ID From ������Ϣ Where ���֤�� = [1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strUser))
        If rs.BOF Then Exit Function
        lng����ID = rs("����id").Value
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ɣÿ��Һ�"
    
        strSQL = "Select  ����ID From ������Ϣ Where IC���� = [1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strUser))
        If rs.BOF Then Exit Function
        lng����ID = rs("����id").Value
        
    End Select
    
    CheckIdentify = True
    
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function CheckIsHosptial(ByVal lng����ID As Long) As Boolean
    '******************************************************************************************************************
    '����:�жϵ�ǰ�����Ƿ���Ժ
    '����:
    '����:
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    CheckIsHosptial = True
    
    strSQL = "select ��ǰ����ID,��ǰ����ID from ������Ϣ where ����ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rs.BOF = False Then
        If ((Not IsNull(rs("��ǰ����ID").Value)) Or (Not IsNull(rs("��ǰ����ID").Value))) Then Exit Function
    End If
    
    CheckIsHosptial = False
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckMoney(ByVal lng����ID As Long) As Boolean
    '******************************************************************************************************************
    '����:ͨ���ѱ������Ҫ�ķ���
    '����:
    '����:
    '******************************************************************************************************************
    Dim aryItem As Variant
    Dim strSQL As String
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    strSQL = "Select Nvl(�ѱ�,'ȫ��') As �ѱ�,nvl(C.Ԥ�����,0)-nvl(C.�������,0) as ��� From ������Ϣ A,������� C Where A.����ID=[1] and A.����ID=C.����ID(+)  And C.����(+)=1 And C.����(+)=1 "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rs.BOF Then Exit Function
    
    If rs("�ѱ�").Value = "ȫ��" Then
        mCurPayNeed = CCur(mParamIn.RegisterPrice)
    Else
        aryItem = GetRegistPrice(CLng(mParamIn.DetailID))
        mCurPayNeed = 0
        For intLoop = 0 To UBound(aryItem)
            mCurPayNeed = mCurPayNeed + ActualMoney(CStr(rs("�ѱ�").Value), aryItem(intLoop, 1), aryItem(intLoop, 0))
        Next
    End If

    If Val(Nvl(rs!���)) < mCurPayNeed Then
        CheckMoney = False
    Else
        CheckMoney = True
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ctlCancel_CommandClick()
    Unload Me
End Sub

Private Sub ctlOK_CommandClick()
    
    Dim lng����ID As Long
    
    If mblnCanCommit = False Then
        '�����֤
        Select Case mParamIn.RegisterMode
        '--------------------------------------------------------------------------------------------------------------
        Case "���֤�Һ�", "���￨�Һ�", "�ɣÿ��Һ�", "ҽ�����Һ�"
            
            If mParamIn.RegisterMode = "�ɣÿ��Һ�" Then
                TxtCardID.Text = ""
                If Not (mobjICCard Is Nothing) Then
                    TxtCardID.Text = mobjICCard.Read_Card(Me)
                End If
                If TxtCardID.Text = "" Then Exit Sub
            End If
            
            '�������ַ�ȥ��
            If mParamIn.RegisterMode = "���￨�Һ�" Then
                TxtCardID.Text = Replace(TxtCardID.Text, ":", "")
                TxtCardID.Text = Replace(TxtCardID.Text, "��", "")
                TxtCardID.Text = Replace(TxtCardID.Text, ";", "")
                TxtCardID.Text = Replace(TxtCardID.Text, "��", "")
                TxtCardID.Text = Replace(TxtCardID.Text, "?", "")
                TxtCardID.Text = Replace(TxtCardID.Text, "��", "")
            End If
        
            If CheckIdentify(mParamIn.RegisterMode, lng����ID, TxtCardID.Text, Txtpwd.Text) = False Then
                Call ShowErrorInfo("�����֤")
                Exit Sub
            End If
            
            If CheckIsHosptial(lng����ID) = True Then
                Call ShowErrorInfo("������Ժ")
                Exit Sub
            End If
            
            If mblnCharge = False Then
                If CheckMoney(lng����ID) = False Then
                    Call ShowErrorInfo("����")
                    Exit Sub
                End If
            End If
            
            Call ShowPatientInfo(lng����ID)
            
            mblnCanCommit = True
            
        End Select
                
    Else
        '�ύ
        Call CommitData
    End If

End Sub

Private Sub ctlReset_CommandClick()
    Call InitData
    If TxtCardID.Enabled Then TxtCardID.SetFocus
End Sub

Private Sub Form_Activate()
'    If mBlnUse = True Then Unload Me
End Sub

Private Sub Form_Load()

    '���Һŵ���Ϣ������ʾ
    mblnCharge = (Val(GetPara("�Һ�ʱ���ɻ��۵�", "1")) = 1)
    mlngTime = Val(GetPara("������֤����ͣ��ʱ��")) / 2
    If Dir(App.Path & "\ͼ��\�Һ�ȷ�ϴ������汳��.pic") <> "" Then
        Imgbak.Picture = LoadPicture(App.Path & "\ͼ��\�Һ�ȷ�ϴ������汳��.pic")
    End If
    
    ctlReset.Picture = frmselectinfo.ilsImage.ListImages("reset")
    ctlOK.Picture = frmselectinfo.ilsImage.ListImages("ok")
    ctlCancel.Picture = frmselectinfo.ilsImage.ListImages("close")
    
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    On Error Resume Next
    Set mobjICCard = CreateObject("zlICCard.clsICCard")
    On Error GoTo 0
    ctlOK.Width = ctlReset.Width
    
    Call InitData

End Sub

Private Sub Form_Paint()
    Call DrawColorToColor(Me, Me.BackColor, &HFFC0C0, , True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mobjIDCard Is Nothing) Then
        On Error Resume Next
        Call mobjIDCard.SetEnabled(False)
        On Error GoTo 0
    End If
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
   Dim lngPreIDKind As Long
   
    If Not TxtCardID.Locked And TxtCardID.Text = "" And Me.ActiveControl Is TxtCardID Then
        With mBrushIDCardPatiInfor
            .strIDCard = strID
            .Name = strName
            .Sex = strSex
            .str�������� = Format(datBirthDay, "yyyy-mm-dd")
            .str������ַ = strAddress
            .str���� = strNation
        End With
        TxtCardID.Text = strID
        Call zlSave������Ϣ
        mblnCanCommit = False
        Call ctlOK_CommandClick
    Else
        mBrushIDCardPatiInfor.strIDCard = ""
    End If
        
End Sub

Private Sub Time_Timer()
    On Error Resume Next
   
    Time.Tag = Val(Time.Tag) - 1
    If Val(Time.Tag) = 0 Then Unload Me
   
End Sub

Private Sub ResetTime()
    Time.Tag = mlngTime
    If LBLErr.Visible = True Then
        LBLErr.Visible = False
        Lblreg.Visible = True
    End If
    
End Sub

Private Sub tmrError_Timer()
    If LBLErr.Visible = True Then
        LBLErr.Visible = False
        Lblreg.Visible = True
    End If
End Sub

Private Sub TxtCardID_Change()
    Dim strTmp As String
    Dim intLen As Integer
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
    
    If mblnNoChange Then Exit Sub
    
    Call ResetTime
    
    Select Case mParamIn.RegisterMode
    Case "���￨�Һ�"

'        strTmp = zlDatabase.GetPara(20, glngSys, , "")
'        If UBound(Split(strTmp, "|")) >= 4 Then intLen = Val(Split(strTmp, "|")(4))
'
'        If Len(TxtCardID.Text) = intLen Then
'
'            '��������Ƿ�Ϊ��
'            gstrSQL = "Select 1 From ������Ϣ Where ���￨�� = [1] And ����֤�� Is Null"
'            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(TxtCardID.Text))
'            If rs.BOF = False Then
'                mblnCanCommit = False
'                Call ctlOK_CommandClick
'            Else
'               Txtpwd.SetFocus
'            End If
'        End If
    Case "���֤�Һ�"
    
        If Me.ActiveControl Is TxtCardID Then
            If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(TxtCardID.Text = "")
        End If
    End Select

    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub TxtCardID_GotFocus()
    
    If mParamIn.RegisterMode = "���֤�Һ�" Then
        If Not (mobjIDCard Is Nothing) Then
            On Error Resume Next
            Call mobjIDCard.SetEnabled(True)
            On Error GoTo 0
        End If
    Else
        If Not (mobjIDCard Is Nothing) Then
            On Error Resume Next
            Call mobjIDCard.SetEnabled(False)
            On Error GoTo 0
        End If
    End If
    
End Sub

Private Sub TxtCardID_KeyPress(KeyAscii As Integer)
    Dim lng����ID As Long, blnCard As Boolean
    
    If KeyAscii = 13 Then
    
        Select Case mParamIn.RegisterMode
        '--------------------------------------------------------------------------------------------------------------
        Case "ҽ�����Һ�", "���￨�Һ�"
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        Case "���֤�Һ�", "�ɣÿ��Һ�"
            mblnCanCommit = False
            Call ctlOK_CommandClick
        End Select
    Else
        If mParamIn.RegisterMode = "���￨�Һ�" Then
            Select Case Chr(KeyAscii)
            Case ":", "��", ";", "��", "?", "��"
                KeyAscii = 0
            Case Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
        End If
    End If
    
End Sub

Private Sub TxtCardID_LostFocus()
    Dim rs As ADODB.Recordset
    Dim strPwd As String
    If Not (mobjIDCard Is Nothing) Then
        On Error Resume Next
        Call mobjIDCard.SetEnabled(True)
        On Error GoTo 0
    End If
     If Me.ActiveControl Is Me.ctlCancel Or Me.ActiveControl Is Me.ctlReset Then Exit Sub
     If Trim(TxtCardID.Text) = "" Then Exit Sub
     Select Case mParamIn.RegisterMode
     Case "���￨�Һ�"
          '��������Ƿ�Ϊ��
            Me.Txtpwd.SetFocus
            If mlng�����ID = 0 Then
                 gstrSQL = "Select 1 From ������Ϣ Where ���￨�� = [1] And ����֤�� Is Null"
            Else
                 gstrSQL = "Select 1 From ����ҽ�ƿ���Ϣ Where ���� = [1] And ���� Is Null" & IIf(mlng�����ID = 0, "", " And �����ID=[2] ")
            End If
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(TxtCardID.Text), mlng�����ID)
            If rs.BOF = False Then
                mblnCanCommit = False
                Call ctlOK_CommandClick
            Else
                If mlng�����ID = 0 Then
                     gstrSQL = "Select ����֤�� From ������Ϣ Where ���￨�� = [1]"
                Else
                     gstrSQL = "Select ���� as  ����֤��  From ����ҽ�ƿ���Ϣ Where ���� = [1] " & IIf(mlng�����ID = 0, "", " And �����ID=[2] ")
                End If
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(TxtCardID.Text), mlng�����ID)
                If rs.EOF Then ShowErrorInfo "���￨": Exit Sub
                If frmCardPass.ShowCardPass(Nvl(rs!����֤��)) Then
                    Txtpwd.Text = Nvl(rs!����֤��)
                    Call ctlOK_CommandClick
                Else
                    Call ctlReset_CommandClick
                End If
            End If
      Case "ҽ�����Һ�"
        If frmCardPass.GetCardPass(strPwd) Then
            Txtpwd.Text = Nvl(strPwd):    Call ctlOK_CommandClick: Exit Sub
        Else
            Call ctlReset_CommandClick
        End If
     End Select
End Sub

 

Private Sub Txtpwd_Change()
    If mblnNoChange Then Exit Sub
     Call ResetTime
End Sub

Private Sub Txtpwd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Txtpwd.Visible = False Then Exit Sub
        mblnCanCommit = False
        Call ctlOK_CommandClick
    End If
End Sub

Private Sub CommitData()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset, rsPati As New ADODB.Recordset, rs As New ADODB.Recordset
    Dim aryItem As Variant, Str����ID As String, strNo As String, i As Integer, str�վݷ�Ŀ As String, str���㷽ʽ As String
    Dim CurԤ��֧�� As Currency, Curҽ��֧�� As Currency, curʵ�� As Currency, curӦ�� As Currency
    Dim Arrҽ�� As Variant, str���մ���ID As String, str������Ŀ�Ƿ� As String, strͳ���� As String
    Dim StrRoom As String, strBed As String, str�ѱ� As String, strTmp As String, strNow As String, str����NO As String
    Dim cllProBefor As Collection, cllPro As Collection, cllproAfter As Collection, strSQL As String
    If mblnCanCommit = False Then Exit Sub
    
    strNow = "To_Date('" & CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd")) & "','yyyy-mm-dd')"
    '�����ǰ��Ͽ���
    '------------------------------------------------------------------------------------------------------------------
    StrRoom = GetRoom(mParamIn.�ű�)
    If StrRoom = "" Then
        StrRoom = "null"
    Else
        StrRoom = "'" + StrRoom + "'"
    End If
        
    Set cllProBefor = New Collection: Set cllPro = New Collection: Set cllproAfter = New Collection
    
    '������㷽ʽ
    '------------------------------------------------------------------------------------------------------------------
    str���մ���ID = "Null": str������Ŀ�Ƿ� = "Null": strͳ���� = "Null"
    
    If mParamIn.RegisterMode = "ҽ�����Һ�" Then
        str���㷽ʽ = "ҽ������": Curҽ��֧�� = CCur(mCurPayNeed)
    Else
        str���㷽ʽ = "�����ʻ�": CurԤ��֧�� = CCur(mCurPayNeed)
    End If
    
    '������к�
    '------------------------------------------------------------------------------------------------------------------
    strNo = zlDatabase.GetNextNo(12)

    '�������ID
    '------------------------------------------------------------------------------------------------------------------
    Str����ID = CStr(zlDatabase.GetNextId("���˽��ʼ�¼"))
    aryItem = GetRegistPrice(mParamIn.DetailID)
    
    gstrSQL = "Select C.���� as ������" & _
                " From ������Ϣ A,ҽ�Ƹ��ʽ C" & _
                " Where A.����ID=[1] " & _
                " And A.ҽ�Ƹ��ʽ=C.����(+)"
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(mPatient.PatientID))
    If rsPati.BOF = False Then strBed = zlCommFun.Nvl(rsPati("������").Value)
    
                
    On Error GoTo ErrHandle
            
    '------------------------------------------------------------------------------------------------------------------
    For i = 0 To UBound(aryItem)
    
        '��ȡͨ��ҽ���õ�������
        If mParamIn.RegisterMode = "ҽ�����Һ�" Then
            str���㷽ʽ = "ҽ������"
            If gintInsure > 0 Then Arrҽ�� = Split(gclsInsure.GetItemInsure(CLng(mPatient.PatientID), Val(aryItem(i, 5)), ActualMoney(mPatient.FareClass, aryItem(i, 1), aryItem(i, 4) * aryItem(i, 0)), True, gintInsure), ";")
            str������Ŀ�Ƿ� = CStr(Arrҽ��(0))
            If CStr(Arrҽ��(1)) <> "0" Then str���մ���ID = CStr(Arrҽ��(1))
            strͳ���� = CStr(CCur(Arrҽ��(2)))
        Else
            str���㷽ʽ = "�����ʻ�"
        End If
        
        gstrSQL = "Select �վݷ�Ŀ From ������Ŀ where ID =[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(aryItem(i, 1)))
        If Not rsTmp.BOF And Not IsNull(rsTmp("�վݷ�Ŀ")) Then str�վݷ�Ŀ = CStr(rsTmp("�վݷ�Ŀ"))
        
        
        curӦ�� = CCur(Val(Format(aryItem(i, 0) * aryItem(i, 4), "0.00")))
'        curӦ�� = IIf(mblnCharge, "0.00", CCur(Val(Format(aryItem(i, 0), "0.00"))))
        curʵ�� = curӦ��
        str�ѱ� = mPatient.FareClass
        
        '��ȡ�ѱ�ʵ�ս��
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) As ��� From Dual"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, str�ѱ�, CLng(aryItem(i, 5)), CLng(aryItem(i, 1)), curӦ��, 0, 0)
        If rs.BOF = False Then
            strTmp = Trim(zlCommFun.Nvl(rs("���").Value))
            If strTmp <> "" Then
                If InStr(strTmp, ":") > 0 Then
                    curʵ�� = Format(Val(Mid(strTmp, InStr(strTmp, ":") + 1)), "0.00")
                    str�ѱ� = Trim(Mid(strTmp, 1, InStr(strTmp, ":") - 1))
                End If
            End If
        End If
            

        If i > 0 Then
            Curҽ��֧�� = 0
            CurԤ��֧�� = 0
        End If
        

        gstrSQL = VB_���˹Һż�¼_Insert(mPatient.PatientID, mPatient.DoorPost, mPatient.Name, mPatient.Sex, mPatient.Age, Val(strBed), str�ѱ�, strNo, mParamIn.BillNo, _
                                       i + 1, Val(aryItem(i, 4)), CLng(aryItem(i, 5)), Format(ActualMoney(mPatient.FareClass, aryItem(i, 1), aryItem(i, 0)), "0.00"), CLng(aryItem(i, 1)), str�վݷ�Ŀ, _
                                    str���㷽ʽ, IIf(mblnCharge, 0, curӦ��), IIf(mblnCharge, 0, curʵ��), _
                                    mParamIn.DepartID, mParamIn.DepartID, strNow, strNow, mParamIn.DoctorName, mParamIn.DoctorID, mParamIn.�ű�, StrRoom, Val(Str����ID), mParamIn.����ID, CurԤ��֧��, _
                                    Curҽ��֧��, str���մ���ID, str������Ŀ�Ƿ�, strͳ����, Val(aryItem(i, 6)), Val(aryItem(i, 7)))
        zlAddArray cllPro, gstrSQL
        
        '����:31187:��Ҫ�ǽ��ҺŻ��ܵ�������
        If mParamIn.�ű� <> "" And i + 1 = 1 Then
           strSQL = "zl_���˹ҺŻ���_Update("
           '  ҽ������_In   �ҺŰ���.ҽ������%Type,
           strSQL = strSQL & "'" & mParamIn.DoctorName & "',"
           '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
           strSQL = strSQL & "" & IIf(mParamIn.DoctorID = 0, "NULL", mParamIn.DoctorID) & ","
           '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
           strSQL = strSQL & "" & mParamIn.DetailID & ","
           '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
           strSQL = strSQL & "" & mParamIn.DepartID & ","
           '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
           strSQL = strSQL & "" & strNow & ","
           '  ԤԼ��־_In   Number := 0  --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����
           strSQL = strSQL & "" & 0 & ","
            '  ����_In       �ҺŰ���.����%Type := Null
            strSQL = strSQL & "'" & mParamIn.�ű� & "')"
           
           Call zlAddArray(cllproAfter, strSQL)
        End If
        '----------------------------------------------------------------------------------------------------------
        If mblnCharge Then
            If str����NO = "" Then str����NO = zlDatabase.GetNextNo(13)
                gstrSQL = "zl_���ﻮ�ۼ�¼_Insert("
                '    No_In         ������ü�¼.NO%Type,
                gstrSQL = gstrSQL & "'" & str����NO & "',"
                '    ���_In       ������ü�¼.���%Type,
                gstrSQL = gstrSQL & "" & i + 1 & ","
                '    ����id_In     ������ü�¼.����id%Type,
                gstrSQL = gstrSQL & "" & mPatient.PatientID & ","
                '    ��ҳid_In     סԺ���ü�¼.��ҳid%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    ��ʶ��_In     ������ü�¼.��ʶ��%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    ���ʽ_In   ������ü�¼.���ʽ%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    ����_In       ������ü�¼.����%Type,
                gstrSQL = gstrSQL & "'" & mPatient.Name & "',"
                '    �Ա�_In       ������ü�¼.�Ա�%Type,
                gstrSQL = gstrSQL & "'" & mPatient.Sex & "',"
                '    ����_In       ������ü�¼.����%Type,
                gstrSQL = gstrSQL & "'" & mPatient.Age & "',"
                '    �ѱ�_In       ������ü�¼.�ѱ�%Type,
                gstrSQL = gstrSQL & "'" & str�ѱ� & "',"
                '    �Ӱ��־_In   ������ü�¼.�Ӱ��־%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    ���˿���id_In ������ü�¼.���˿���id%Type,
                gstrSQL = gstrSQL & "" & mParamIn.DepartID & ","
                '    ��������id_In ������ü�¼.��������id%Type,
                gstrSQL = gstrSQL & "" & UserInfo.����ID & ","
                '    ������_In     ������ü�¼.������%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
                '    ��������_In   ������ü�¼.��������%Type,
                gstrSQL = gstrSQL & IIf(Val(aryItem(i, 7)) = 0, "NULL", Val(aryItem(i, 7))) & ","   '57045
                '    �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
                gstrSQL = gstrSQL & "" & CLng(aryItem(i, 5)) & ","
                '    �շ����_In   ������ü�¼.�շ����%Type,
                gstrSQL = gstrSQL & "'1',"
                '    ���㵥λ_In   ������ü�¼.���㵥λ%Type,
                gstrSQL = gstrSQL & "'" & CStr(aryItem(i, 3)) & "',"
                '    ��ҩ����_In   ������ü�¼.��ҩ����%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    ����_In       ������ü�¼.����%Type,
                gstrSQL = gstrSQL & "1,"
                '    ����_In       ������ü�¼.����%Type,
                gstrSQL = gstrSQL & "" & CLng(aryItem(i, 4)) & ","
                '    ���ӱ�־_In   ������ü�¼.���ӱ�־%Type,
                gstrSQL = gstrSQL & "0,"
                '    ִ�в���id_In ������ü�¼.ִ�в���id%Type,
                gstrSQL = gstrSQL & "" & mParamIn.DepartID & ","
                '    �۸񸸺�_In   ������ü�¼.�۸񸸺�%Type,
                gstrSQL = gstrSQL & IIf(Val(aryItem(i, 6)) = 0, "NULL", Val(aryItem(i, 6))) & ","
                '    ������Ŀid_In ������ü�¼.������Ŀid%Type,
                gstrSQL = gstrSQL & "" & CLng(aryItem(i, 1)) & ","
                '    �վݷ�Ŀ_In   ������ü�¼.�վݷ�Ŀ%Type,
                gstrSQL = gstrSQL & "'" & str�վݷ�Ŀ & "',"
                '    ��׼����_In   ������ü�¼.��׼����%Type,
                gstrSQL = gstrSQL & "" & CLng(aryItem(i, 0)) & ","
                '    Ӧ�ս��_In   ������ü�¼.Ӧ�ս��%Type,
                gstrSQL = gstrSQL & "" & curӦ�� & ","
                '    ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
                gstrSQL = gstrSQL & "" & curʵ�� & ","
                '    ����ʱ��_In   ������ü�¼.����ʱ��%Type,
                gstrSQL = gstrSQL & "" & strNow & ","
                '    �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type,
                gstrSQL = gstrSQL & "" & strNow & ","
                '    ҩƷժҪ_In   ҩƷ�շ���¼.ժҪ%Type,
                gstrSQL = gstrSQL & "NULL,"
                '    ����Ա����_In ������ü�¼.����Ա����%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
                '    ���id_In     ҩƷ��������.���id%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    ����ժҪ_In   ������ü�¼.ժҪ%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    ҽ�����_In   ������ü�¼.ҽ�����%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    Ƶ��_In       ҩƷ�շ���¼.Ƶ��%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    ����_In       ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '    �÷�_In       ҩƷ�շ���¼.�÷�%Type := Null, --�÷�[|�巨]
                gstrSQL = gstrSQL & "NULL,"
                '    ��Ч_In       ҩƷ�շ���¼.����%Type := Null,
                 gstrSQL = gstrSQL & "1,"
                '    �Ƽ�����_In   ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "0,"
                '    ������Դ_In   Number := 1,
                gstrSQL = gstrSQL & "4)"
                '    ���ձ���_In   ������ü�¼.���ձ���%Type := Null,
                '    ��������_In   ������ü�¼.��������%Type := Null,
                '    ������Ŀ��_In ������ü�¼.������Ŀ��%Type := Null,
                '    ���մ���id_In ������ü�¼.���մ���id%Type := Null,
                '    ��ҩ��̬_In       ������ü�¼.����%Type := Null
            zlAddArray cllPro, gstrSQL
        End If
    Next
    
    '�޸�ҽ���Ľӿ�����
    '------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrFirst:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    Err = 0: On Error GoTo ErrHandle:
    If mParamIn.RegisterMode = "ҽ�����Һ�" And gintInsure > 0 Then
         If gclsInsure.RegistSwap(CLng(Str����ID), Curҽ��֧��, gintInsure) = False Then
            gcnOracle.RollbackTrans
            Unload Me
            Exit Sub
         End If
    End If
    zlExecuteProcedureArrAy cllproAfter, Me.Caption, False, True
    
    
    Err = 0: On Error GoTo ErrEnd:
    '��ӡ����
    '------------------------------------------------------------------------------------------------------------------
    'If mblnCharge = False Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1111", Me, "NO=" & strNo, 2)
    'End If
 
    Call frmClose.ShowForm(Me, mPatient.Name, strNo, str����NO)
        
    Unload Me
    
    Exit Sub
    '-----------------------------------------------------------------------------------------------------------------
ErrFirst:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
ErrEnd:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
ErrHandle:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
    
End Sub

Private Function VB_���˹Һż�¼_Insert(ByVal lng����ID As String, ByVal lng����� As String, ByVal str���� As String, ByVal str�Ա� As String, ByVal str���� As String, _
    ByVal str���� As String, ByVal str�ѱ� As String, ByVal str���ݺ� As String, ByVal strƱ�ݺ� As String, ByVal int��� As String, ByVal lng���� As Long, ByVal lng�շ�ϸĿid As String, _
    ByVal db��׼���� As String, ByVal lng������Ŀid As String, ByVal str�վݷ�Ŀ As String, ByVal str���㷽ʽ As String, ByVal dbӦ�ս�� As String, ByVal dbʵ�ս�� As String, _
    ByVal lng���˿���id As String, ByVal lngִ�в���id As String, ByVal str����ʱ�� As String, ByVal str�Ǽ�ʱ�� As String, ByVal strҽ������ As String, ByVal lngҽ��id As String, _
    ByVal str�ű� As String, ByVal str��ҩ���� As String, ByVal lng����id As String, ByVal lng����ID As String, _
    ByVal CurԤ��֧�� As String, ByVal Curҽ��֧�� As String, ByVal str���մ���ID As String, ByVal str������Ŀ�Ƿ� As String, _
    ByVal strͳ���� As String, ByVal int�۸񸸺� As Integer, ByVal int�������� As Integer) As String
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    
    Dim strSQL As String, bln���ɶ��� As Boolean
    bln���ɶ��� = Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, 1113)) <> 0
 
    'Zl_���˹Һż�¼_Insert
    strSQL = "zl_���˹Һż�¼_Insert("
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �����_In     ������ü�¼.��ʶ��%Type,
    strSQL = strSQL & "" & lng����� & ","
    '  ����_In       ������ü�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  �Ա�_In       ������ü�¼.�Ա�%Type,
    strSQL = strSQL & "'" & str�Ա� & "',"
    '  ����_In       ������ü�¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ���ʽ_In   ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
    strSQL = strSQL & "'" & Val(str����) & "',"
    '  �ѱ�_In       ������ü�¼.�ѱ�%Type,
    strSQL = strSQL & "'" & str�ѱ� & "',"
    '  ���ݺ�_In     ������ü�¼.NO%Type,
    strSQL = strSQL & "'" & str���ݺ� & "',"
    '  Ʊ�ݺ�_In     ������ü�¼.ʵ��Ʊ��%Type,
    strSQL = strSQL & "'" & strƱ�ݺ� & "',"
    '  ���_In       ������ü�¼.���%Type,
    strSQL = strSQL & "" & int��� & ","
    '  �۸񸸺�_In   ������ü�¼.�۸񸸺�%Type,
    strSQL = strSQL & "" & IIf(int�۸񸸺� = 0, "NULL", int�۸񸸺�) & ","
    '  ��������_In   ������ü�¼.��������%Type,
    strSQL = strSQL & "" & IIf(int�������� = 0, "NULL", int��������) & ","
    '  �շ����_In   ������ü�¼.�շ����%Type,
    strSQL = strSQL & "'1',"
    '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
    strSQL = strSQL & "" & lng�շ�ϸĿid & ","
    '  ����_In       ������ü�¼.����%Type,
    strSQL = strSQL & "" & lng���� & ","
    '  ��׼����_In   ������ü�¼.��׼����%Type,
    strSQL = strSQL & "" & db��׼���� & ","
    '  ������Ŀid_In ������ü�¼.������Ŀid%Type,
    strSQL = strSQL & "" & lng������Ŀid & ","
    '  �վݷ�Ŀ_In   ������ü�¼.�վݷ�Ŀ%Type,
    strSQL = strSQL & "'" & str�վݷ�Ŀ & "',"
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  Ӧ�ս��_In   ������ü�¼.Ӧ�ս��%Type,
    strSQL = strSQL & "" & dbӦ�ս�� & ","
    '  ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & dbʵ�ս�� & ","
    '  ���˿���id_In ������ü�¼.���˿���id%Type,
    strSQL = strSQL & "" & lng���˿���id & ","
    '  ��������id_In ������ü�¼.��������id%Type,
    strSQL = strSQL & "" & UserInfo.����ID & ","
    '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
    strSQL = strSQL & "" & lngִ�в���id & ","
    '  ����Ա���_In ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
    strSQL = strSQL & "" & str����ʱ�� & ","
    '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "" & str�Ǽ�ʱ�� & ","
    '  ҽ������_In   �ҺŰ���.ҽ������%Type,
    strSQL = strSQL & "'" & strҽ������ & "',"
    '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
    strSQL = strSQL & "" & IIf(lngҽ��id = 0, "NULL", lngҽ��id) & ","
    '  ������_In Number, --������¼�Ƿ���������
    strSQL = strSQL & "0,"
    '  ����_In       Number,
    strSQL = strSQL & "0,"
    '  �ű�_In       �ҺŰ���.����%Type,
    strSQL = strSQL & "'" & str�ű� & "',"
    '����:48508
    '  ����_In       ���˷��ü�¼.��ҩ����%Type,
    strSQL = strSQL & "'" & Replace(str��ҩ����, "'", "") & "',"
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����id & ","
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(lng����ID = 0, "NULL", lng����ID) & ","
    '  Ԥ��֧��_In   ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
    strSQL = strSQL & "" & Round(Val(CurԤ��֧��), 2) & ","
    '  �ֽ�֧��_In   ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
    strSQL = strSQL & "" & 0 & ","
    '  ����֧��_In   ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
    strSQL = strSQL & "" & Round(Val(Curҽ��֧��), 2) & ","
    '  ���մ���id_In ������ü�¼.���մ���id%Type,
    strSQL = strSQL & "" & IIf(Val(str���մ���ID) = 0, "NULL", str���մ���ID) & ","
    '  ������Ŀ��_In ������ü�¼.������Ŀ��%Type,
    strSQL = strSQL & "" & Val(str������Ŀ�Ƿ�) & ","
    '  ͳ����_In   ������ü�¼.ͳ����%Type,
    strSQL = strSQL & "" & Val(strͳ����) & ","
    '  ժҪ_In       ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
    strSQL = strSQL & "NULL,"
    '  ԤԼ�Һ�_In   Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
    strSQL = strSQL & "0,"
    '  �շ�Ʊ��_In   Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
    strSQL = strSQL & "0,"
    '  ���ձ���_In   ������ü�¼.���ձ���%Type,
    strSQL = strSQL & "NULL,"
    '  ����_In       ���˹Һż�¼.����%Type := 0,
    strSQL = strSQL & "0,"
    '  ����_In       �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
    strSQL = strSQL & "NULL,"
    '  ����_In       ���˹Һż�¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ԤԼ����_In   Number := 0,
    strSQL = strSQL & "0,"
    '  ԤԼ��ʽ_In   ԤԼ��ʽ.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ���ɶ���_In Number:=0
    strSQL = strSQL & IIf(bln���ɶ���, 1, 0) & ")"
    VB_���˹Һż�¼_Insert = strSQL
    
End Function

Private Function zlSave������Ϣ() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����֤�Һ�ʱ,�������֤ˢ��ʱ,�϶�Ҫ�Ƚ���
    '���:
    '����:
    '����:�ɹ�,����true,����False
    '����:���˺�
    '����:2009-12-04 14:05:53
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strSQL As String, lng����ID As Long, intType As Integer
    Dim rsTemp As New ADODB.Recordset, str�������� As String, str���� As String
    
    'δˢ��,������
    If mBrushIDCardPatiInfor.strIDCard = "" Then Exit Function
    
    
    strSQL = "Select ����ID,����,�Ա�,�����,����,�ѱ�,Trunc(��������) as ��������,��ͥ��ַ,���� From ������Ϣ  where ���֤��=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mBrushIDCardPatiInfor.strIDCard)
    If rsTemp.EOF = False Then
       '���ڲ�����Ϣ
        mPatient.Name = zlCommFun.Nvl(rsTemp!����)
        mPatient.DoorPost = zlCommFun.Nvl(rsTemp!�����, 0)
        mPatient.Sex = zlCommFun.Nvl(rsTemp!�Ա�)
        mPatient.Age = zlCommFun.Nvl(rsTemp!����)
        mPatient.FareClass = zlCommFun.Nvl(rsTemp!�ѱ�)
        mPatient.strIDCard = mBrushIDCardPatiInfor.strIDCard
        mPatient.PatientID = zlCommFun.Nvl(rsTemp!����id)
        mPatient.str�������� = Format(rsTemp!��������, "yyyy-mm-dd")
        mPatient.str������ַ = Nvl(rsTemp!��ͥ��ַ)
        mPatient.str���� = Nvl(rsTemp!����)
        '���ڵĻ����Ͳ�������
        zlSave������Ϣ = True
        Exit Function
    End If
 
    
    '�²���,�Ƚ���
    lng����ID = zlDatabase.GetNextNo(1): mPatient.PatientID = lng����ID
    If IsDate(mBrushIDCardPatiInfor.str��������) Then
        strSQL = "Select (Sysdate-to_date([1],'yyyy-mm-dd'))/365 As �� From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mBrushIDCardPatiInfor.str��������)
        str���� = Format(Val(Nvl(rsTemp!��)), "###0.00") & "��"
    Else
        str���� = ""
    End If
    
    '  --�������ͣ�
    '  --             1=�½�������Ϣ�����ﲡ��(�����¹ҺŲ���)
    '  --             2=�޸Ĳ�����Ϣ���½����ﲡ��(�����޲����Ĳ���)
    '  --             3=�޸Ĳ�����Ϣ�����������ﲡ��(�����в����Ĳ���,�������޸��˲����������)
    '  --����ҩ��ָ���ʽ��"ID~����~~ID~����...",�������޸Ĳ�����Ϣʱ�á�
    
    'Zl_�ҺŲ��˲���_Insert
    strSQL = "Zl_�ҺŲ��˲���_Insert("
    '  ��������_In     Number,
    strSQL = strSQL & "" & 1 & ","
    '  ����id_In       ������Ϣ.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �����_In       ������Ϣ.�����%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ���￨��_In     ������Ϣ.���￨��%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����֤��_In     ������Ϣ.����֤��%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����_In         ������Ϣ.����%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.Name & "',"
    '  �Ա�_In         ������Ϣ.�Ա�%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.Sex & "',"
    '  ����_In         ������Ϣ.����%Type,
    strSQL = strSQL & "" & IIf(str���� = "", "NULL", "'" & str���� & "'") & ","
    '  �ѱ�_In         ������Ϣ.�ѱ�%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����_In         ������Ϣ.����%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����_In         ������Ϣ.����%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.str���� & "',"
    '  ����_In         ������Ϣ.����״��%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ְҵ_In         ������Ϣ.ְҵ%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ���֤��_In     ������Ϣ.���֤��%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.strIDCard & "',"
    '  ������λ_In     ������Ϣ.������λ%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��ͬ��λid_In   ������Ϣ.��ͬ��λid%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��λ�绰_In     ������Ϣ.��λ�绰%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��λ�ʱ�_In     ������Ϣ.��λ�ʱ�%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    strSQL = strSQL & "'" & mBrushIDCardPatiInfor.str������ַ & "',"
    '  ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  �����ʱ�_In     ������Ϣ.�����ʱ�%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  �Ǽ�ʱ��_In     ������Ϣ.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����ҩ��_In     Varchar2,
    strSQL = strSQL & "" & "NULL" & ","
    '  �Һŵ�_In       ���˹Һż�¼.NO%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ��������_In     ������Ϣ.��������%Type := Null,
    If IsDate(mBrushIDCardPatiInfor.str��������) Then
        strSQL = strSQL & "to_date('" & mBrushIDCardPatiInfor.str�������� & "','yyyy-mm-dd'),"
    Else
        strSQL = strSQL & "" & "NULL" & ","
    End If
    '  ҽ����_In       ������Ϣ.ҽ����%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  Ic����_In       ������Ϣ.Ic����%Type := Null
    strSQL = strSQL & "" & "NULL" & ")"
    
    Err = 0: On Error GoTo errHand:
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    With mPatient
        .Name = mBrushIDCardPatiInfor.Name
        .PatientID = lng����ID
        .Sex = mBrushIDCardPatiInfor.Sex
        .Age = str����
        .str�������� = mBrushIDCardPatiInfor.str��������
        .strIDCard = mBrushIDCardPatiInfor.strIDCard
        .str������ַ = mBrushIDCardPatiInfor.str������ַ
        .str���� = mBrushIDCardPatiInfor.str����
    End With
    zlSave������Ϣ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
