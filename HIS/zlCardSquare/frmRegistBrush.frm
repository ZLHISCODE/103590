VERSION 5.00
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#3.4#0"; "zlIDKind.ocx"
Begin VB.Form frmRegistBrush 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�Һ�ˢ����֤"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   375
      Left            =   800
      TabIndex        =   9
      Top             =   1515
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   661
      Appearance      =   2
      IDKindStr       =   $"frmRegistBrush.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "����"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   5805
      TabIndex        =   5
      Top             =   0
      Width           =   5805
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   6000
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʣ���1000.00�����ν�1000.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   465
         Width           =   4320
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4845
         Picture         =   "frmRegistBrush.frx":009D
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ˣ����������У�30��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   255
         TabIndex        =   6
         Top             =   105
         Width           =   2640
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   2700
      Width           =   6900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4485
      TabIndex        =   3
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3270
      TabIndex        =   2
      Top             =   2865
      Width           =   1100
   End
   Begin VB.TextBox txtCard 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1455
      TabIndex        =   1
      Top             =   1500
      Width           =   3015
   End
   Begin VB.CommandButton cmdReadIC 
      Caption         =   "����"
      Height          =   405
      Left            =   4500
      TabIndex        =   0
      Top             =   1500
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   1590
      Width           =   570
   End
End
Attribute VB_Name = "frmRegistBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mintCount As Integer
'Private mobjICCard As Object 'IC������
Private mlng����ID As String
Private mlngSys As Long
Private mblnTest As Boolean
Private mblnPreCard As Boolean
Private mblnUnload As Boolean

'--------------------------------------------------
'�����:
Private mobjKeyboard As Object
Private mblnPassInputCardNo As Boolean  '�Ƿ��������뿨��
Private mobjSquareCard As Object
Private mlngҽ�ƿ����� As Long
Private mlngModul As Long
Private mstrPassWord As String
Private mlngDefaultCardTypeID As Long 'ȱʡ��ˢ�����ID
Private mblnBrushCard As Boolean
Private mlngCardTypeID As Long  '�����ID
Private mstrCardNo As String    '����
Private mobjPatiCardObject As clsCardObject
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'--------------------------------------------------
Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByVal lng����ID As Long, _
    ByVal cur��� As Currency, Optional lngModul As Long = 0, _
    Optional lngCardTypeID As Long, Optional ByRef strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤�������
    '���:frmParent-���õ�������
    '       lngSys-ϵͳ��
    '       lng����ID-ָ���Ĳ���ID
    '       lngModul-ģ���
    '       lngCardTypeID-ȱʡ�����ID
    '����:strCardNo-���ؿ���
    '       lngCardTypeID-�����ID
    '����:��֤�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-10 16:35:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, rsTemp As ADODB.Recordset
     
    Dim strSQL As String, intMouse As Integer
    mlngSys = lngSys: mlngModul = lngModul: mlngDefaultCardTypeID = lngCardTypeID
    mblnOK = False: mintCount = 3: mlng����ID = lng����ID
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0
    mblnTest = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard", "TestCardNO", 0)) = 1
    mblnTest = IsDesinMode Or mblnTest
 
    '��ȡ���￨��Ϣ
    On Error GoTo errH
    strSQL = "" & _
    "   Select A.����,A.�Ա�,A.����,A.���￨��,A.����֤��, " & _
    "              nvl(B.���,0) as ���" & _
    "   From ������Ϣ A, " & _
    "       (   Select ����ID,nvl(Sum(Ԥ�����),0)-nvl(sum(�������),0) as ��� " & _
    "           From  ������� " & _
    "           Where ����ID=[1] and ����=1 and decode([2],0,0,����)=[2]  Group by ����ID) B " & _
    "   Where A.����ID=[1] And A.����ID=B.����ID(+) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng����ID, 1)
    If rsTmp.EOF Then
        MsgBox "������Ϣ������,����!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If IIf(IsNull(rsTmp!���￨��), "", rsTmp!���￨�� & "") = "" Then
        '����:43449
        strSQL = "Select Count(Distinct �����ID) as ����� From ����ҽ�ƿ���Ϣ Where  ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng����ID)
        If IIf(IsNull(rsTemp!�����), 0, Val(rsTemp!����� & "")) = 0 Then
            '--δ����,ֱ�ӷ���true,�����鿨
            ShowMe = False: Exit Function
        End If
    End If
    
    Me.lblPati.Caption = "���ˣ�" & zlCommFun.Nvl(rsTmp!����) & _
        IIf(Not IsNull(rsTmp!�Ա�), "��" & rsTmp!�Ա�, "") & _
        IIf(Not IsNull(rsTmp!����), "��" & rsTmp!����, "")
    Me.lblMoney.Caption = "ʣ���" & Format(rsTmp!���, "0.00") & "�����ν�" & Format(cur���, "0.00")
    Me.txtCard.Tag = zlCommFun.Nvl(rsTmp!���￨��)
    mstrCardNo = "": lngCardTypeID = 0
    On Error GoTo 0
    'IC������
    On Error Resume Next
    'Set mobjICCard = CreateObject("zlICCard.clsICCard")
    On Error GoTo 0
    Me.Show 1, frmParent
    ShowMe = mblnOK
    lngCardTypeID = mlngCardTypeID
    strCardNo = mstrCardNo
    Screen.MousePointer = intMouse
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub CmdCancel_Click()
    Unload Me
End Sub
Private Sub CmdOK_Click()
    Dim strPassWord As String
    If UCase(txtCard.Text) <> UCase(txtCard.Tag) Then
        MsgBox "��ǰ�����벡�˵Ŀ��Ų������", vbExclamation, gstrSysName
        Unload Me: Exit Sub '���Ų�ƥ�䣬��׼����
    End If
    If Val(cmdReadIC.Tag) <> mlng����ID Or Val(cmdReadIC.Tag) = 0 Then
        MsgBox "��ǰ�����벡�˵Ŀ��Ų������", vbExclamation, gstrSysName
        Unload Me: Exit Sub '���Ų�ƥ�䣬��׼����
    End If
    mstrCardNo = txtCard.Text
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdReadIC_Click()
    Call IDKind_Click(IDKind.GetCurCard)
End Sub

Private Sub Form_Activate()
    If IDKind.ListCount = 0 Then Unload Me: Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
       Select Case KeyCode
        Case vbKeyF4
            If IDKind.Enabled Then
                If Shift = vbShiftMask Then
                    IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDKindStr, ";")), IDKind.IDKind - 1)
                Else
                    IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDKindStr, ";")), 0, IDKind.IDKind + 1)
                End If
            End If
        End Select
End Sub
Private Sub Form_Load()
    Call CreateObjectKeyboard
    Call zlInitData
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Set mobjICCard = Nothing
    Set mobjKeyboard = Nothing
End Sub

 

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String
     If objCard Is Nothing Then Exit Sub
    If IsCardType(IDKind, "IC����") Then
        Exit Sub
    End If
    lng�����ID = objCard.�ӿ����
    mlngCardTypeID = lng�����ID
    Call CreatePayObject
    If lng�����ID = 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    strExpand = lng�����ID
    If mobjSquareCard.zlReadCard(Me, mlngModul, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txtCard.Text = strOutCardNo
    '�����:42948
    If txtCard.Text <> "" Then
        If GetPatient(Trim(txtCard.Text)) = False Then
                txtCard.Text = ""
                If txtCard.Enabled Then txtCard.SetFocus
                zlControl.TxtSelAll txtCard
                Exit Sub
        End If
     End If
     If txtCard.Text <> "" Then
        Call CmdOK_Click
     Else
         txtCard.SetFocus
     End If
End Sub

'��ȡidkind��Ĭ��kindֵ
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

'�ؼ������Ƿ�ƥ��
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
     End Select
End Function
                

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    mlngҽ�ƿ����� = objCard.���ų���
    '��7λ��,��ֻ��������,��Ȼȡ������
    mblnPassInputCardNo = IDKind.ShowPassText
    txtCard.MaxLength = mlngҽ�ƿ�����
    txtCard.PasswordChar = IIf(mblnPassInputCardNo, "*", "")
    '85565,���ϴ�,2015/7/19:��������
'    mblnBrushCard = Mid(objCard.��������, 1, 1) = 0 And Mid(objCard.��������, 2, 1) = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtCard.Text <> "" Then txtCard.Text = ""
    txtCard.Locked = Not (objCard.�Ƿ�ˢ�� Or objCard.�Ƿ�ɨ��)
    cmdReadIC.Visible = objCard.�Ƿ�Ӵ�ʽ����
    If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtCard.Locked Or txtCard.Text <> "" Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
     

    intIndex = IDKind.GetKindIndex(objCard.����)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex

    txtCard.Text = objPatiInfor.����
    Call txtCard_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
   
End Sub

'
'Private Sub txtPass_LostFocus()
'    ClosePassKeyboard txtPass
'End Sub
Private Sub txtCard_Change()
    txtCard.Tag = "": cmdReadIC.Tag = ""
    'lblPass.Tag = "":
    'txtPass.Enabled = txtCard.Text <> ""
    'If Not txtPass.Enabled Then txtPass.Text = ""
End Sub

Private Sub txtCard_GotFocus()
    Call zlControl.TxtSelAll(txtCard)
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    mblnPreCard = False

    '�Ƿ�ˢ�����
    blnCard = KeyAscii <> 8 And Len(txtCard.Text) = mlngҽ�ƿ����� - 1 And txtCard.SelLength <> Len(txtCard.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txtCard.Text = txtCard.Text & Chr(KeyAscii)
            txtCard.SelStart = Len(txtCard.Text)
        End If
        KeyAscii = 0
        If GetPatient(Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            zlControl.TxtSelAll txtCard
            Exit Sub
       End If
       mblnPreCard = blnCard
       If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
       If blnCard Then Call CmdOK_Click
       Exit Sub
    Else
        If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        
        If mblnTest Then Exit Sub
        '��ȫˢ�����
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = timer
            If txtCard.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(txtCard.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                txtCard.Text = Chr(KeyAscii)
                txtCard.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub

Private Sub txtCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtCard.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtCard.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtCard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtCard.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
'Private Sub txtPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button <> 2 Or mblnTest Then Exit Sub
'    glngTXTProc = GetWindowLong(txtPass.hWnd, GWL_WNDPROC)
'    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, AddressOf WndMessage)
'End Sub
'
'Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'     If Button <> 2 Or mblnTest Then Exit Sub
'    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, glngTXTProc)
'End Sub
'
'Private Sub txtPass_GotFocus()
'    Call zlControl.TxtSelAll(txtPass)
'    OpenPassKeyboard txtPass
'End Sub
'
'Private Sub txtPass_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        If mblnPreCard Then
'             If (GetAsyncKeyState(VK_RETURN) And &H1) <> 0 Then
'                txtPass.Text = ""
'                Exit Sub
'             End If
'        End If
'        mblnPreCard = False
'        Call cmdOK_Click
'    ElseIf KeyAscii = 22 Then
'        KeyAscii = 0 '������ճ��
'    Else
'        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
'                KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
'        End If
'    End If
'
'End Sub

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function OpenPassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Sub CreatePayObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����֧������ӿ�
    '����:���˺�
    '����:2011-06-22 13:15:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�����ID As Long, bln���ѿ� As Boolean, int�Զ���ȡ As Integer
    Dim strKey As String
    Dim i As Long
    Set mobjSquareCard = Nothing:
    Err = 0: On Error Resume Next
    If zlGetCardObj(Me, mlngCardTypeID, False, mobjPatiCardObject) = False Then
        Set mobjPatiCardObject = Nothing
        Set mobjSquareCard = Nothing
        Exit Sub
    End If
    Set mobjSquareCard = mobjPatiCardObject.CardObject
    If Err <> 0 Then
        MsgBox "δ�ҵ�" & IDKind.GetCurCard.���� & "����Ӧ�Ĳ���,����", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If mobjSquareCard Is Nothing Then Exit Sub
End Sub
Private Sub zlInitData()
    Dim strExpend As String, i As Integer
    Dim strKey As String
    Dim lngCardID As Long
    strKey = GetIDKindStr("", True)
    If strKey = "" Then
        mblnUnload = True
        Exit Sub
    End If
    IDKind.IDKindStr = strKey
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, "", txtCard)
    IDKind.ShowPropertySet = InStr(";" & gstrPrivs & ";", "��������") > 0
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
End Sub
Private Function GetPatient(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    On Error GoTo errH
    mstrPassWord = ""
    lng�����ID = IDKind.GetCurCard.�ӿ����
    If lng�����ID = 0 Then
      If mobjSquareCard.zlGetPatiID(IDKind.GetCurCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
    Else
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If GetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
    End If
    If lng����ID <= 0 Then GoTo NotFoundPati:
    If mlng����ID <> lng����ID Then
       MsgBox "��ǰ�����벡�������еĿ��Ų����,���飡", vbExclamation, gstrSysName
       txtCard.Text = ""
       Exit Function '���Ų�ƥ�䣬��׼����
    End If
    txtCard.Tag = strInput
    cmdReadIC.Tag = lng����ID
    mstrPassWord = strPassWord
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
NotFoundPati:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
    Else
        MsgBox "δ�ҵ���ǰ���ĳ��в���,����!", vbOKOnly + vbInformation, gstrSysName
    End If
    txtCard.Tag = "": cmdReadIC.Tag = ""
End Function
Private Function IsDesinMode() As Boolean
      '���˺� ȷ����ǰģʽΪ���ģʽ
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function

