VERSION 5.00
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmIdentify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���������֤"
   ClientHeight    =   4008
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   6540
   Icon            =   "frmIdentify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4008
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2325
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1350
      Width           =   3015
   End
   Begin VB.TextBox txtCard 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2325
      TabIndex        =   1
      Top             =   1912
      Width           =   3015
   End
   Begin VB.TextBox txtPass 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2340
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2475
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3420
      TabIndex        =   3
      Top             =   3450
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3450
      Width           =   1100
   End
   Begin VB.Frame fraDown 
      Height          =   30
      Left            =   -30
      TabIndex        =   9
      Top             =   3225
      Width           =   7290
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      ScaleHeight     =   1056
      ScaleWidth      =   6540
      TabIndex        =   10
      Top             =   0
      Width           =   6540
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   1140
         X2              =   8715
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   1140
         X2              =   8715
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Label lblFamilyRest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:9999999.00"
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
         Left            =   4260
         TabIndex        =   17
         Tag             =   "�������:"
         Top             =   750
         Width           =   2280
      End
      Begin VB.Label lblRest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:9999999.00"
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
         Left            =   1140
         TabIndex        =   16
         Tag             =   "�������:"
         Top             =   750
         Width           =   2280
      End
      Begin VB.Label lblPatiType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:��ͨ����"
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
         Left            =   1140
         TabIndex        =   15
         Tag             =   "��������:"
         Top             =   420
         Width           =   2040
      End
      Begin VB.Label lblFeeType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�:��ͨ"
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
         Left            =   4740
         TabIndex        =   14
         Tag             =   "�ѱ�:"
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:30��"
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
         Left            =   4740
         TabIndex        =   13
         Tag             =   "����:"
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�:δ֪"
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
         Left            =   3330
         TabIndex        =   12
         Tag             =   "�Ա�:"
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:����༪"
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
         Left            =   1140
         TabIndex        =   11
         Tag             =   "����:"
         Top             =   90
         Width           =   1560
      End
      Begin VB.Image Image1 
         Height          =   576
         Left            =   240
         Picture         =   "frmIdentify.frx":058A
         Top             =   132
         Width           =   576
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -105
         X2              =   7470
         Y1              =   1035
         Y2              =   1035
      End
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   420
      Left            =   1665
      TabIndex        =   7
      Top             =   1905
      Width           =   630
      _ExtentX        =   1101
      _ExtentY        =   741
      Appearance      =   2
      IDKindStr       =   "��|���￨|0|0|0|0|0|;IC|IC����|1|0|0|0|0|"
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
      ShowPropertySet =   -1  'True
      NotContainFastKey=   ""
      AllowAutoICCard =   -1  'True
      AllowAutoIDCard =   -1  'True
      BackColor       =   -2147483633
      SaveRegType     =   4
      ProductName     =   "һ��ͨ����֧��"
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ˢ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1155
      TabIndex        =   5
      Top             =   1425
      Width           =   1140
   End
   Begin VB.Label lblCardNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   6
      Top             =   1980
      Width           =   570
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.4
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1425
      TabIndex        =   8
      Top             =   2580
      Width           =   870
   End
End
Attribute VB_Name = "frmIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mintCount As Integer
Private mstr����IDs As String
Private mlngSys As Long
Private mblnPreCard As Boolean
Private mobjCard As Card '��ǰ����Ŀ�
'--------------------------------------------------
'�����:
Private mobjKeyboard As Object
Private mobjOneCardComLib As zlOneCardComLib.clsOneCardComLib
Private mlngModul As Long
Private mstrPassWord As String
Private mlngDefaultCardTypeID As Long 'ȱʡ��ˢ�����ID
Private mblnBrushCard As Boolean
Private Const VK_RETURN = &HD
Private mblnCheckPassWord As Boolean
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mstrRegSection As String
Private mlngPreBrushCardTypeID As Long '�ϴ�ˢ�����

Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByVal lng����ID As Long, _
    ByVal cur��� As Currency, Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0, _
    Optional lngDefaultCardTypeID As Long = 0, _
    Optional blnCheckPassWord As Boolean = True, _
    Optional blnFamilyMoney As Boolean, _
    Optional strFamilyPatiIDs As String = "", _
    Optional blnˢ����֤ As Boolean = True, _
    Optional bln�����벻�鿨 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤�������
    '���:frmParent-���õ�������
    '       lngSys-ϵͳ��
    '       lng����ID-ָ���Ĳ���ID
    '       lngModul-ģ���
    '       bytOperationType-ҵ������(0-������;1-����;2-סԺ)
    '       mlngDefaultCardTypeID-ȱʡ��ˢ�����ID
    '       blnCheckPassWord-��֤����(true-��֤����,false-ֻˢ��,����������)
    '       blnFamilyMoney-�Ƿ��ȡ����Ԥ�����
    '       strFamilyPatiIDs-���˼����Ĳ���ID
    '       blnˢ����֤-�Ƿ����ˢ����֤����Ҫ���ڲ�ˢ����֤ʱ��ȡ����IDs
    '       bln�����벻�鿨-���˵�����ҽ�ƿ���û����������ʱ�Ƿ��鿨����ΪTrueʱ��ֻҪ��һ�ſ����������붼Ҫ�����鿨,112418
    '����:
    '����:��֤�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-10 16:35:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strSQL As String, intMouse As Integer
    Dim cur������� As Currency, cur������� As Currency
    Dim objPati As clsPatientInfo, objExpenceSvr As clsExpenceSvr
    Dim cllPatiFee As Collection, objPatiFee As clsPatiFeeinfor, i As Long
    
    mblnCheckPassWord = blnCheckPassWord
    mlngSys = lngSys: mlngModul = lngModul: mlngDefaultCardTypeID = lngDefaultCardTypeID
    mblnOk = False: mintCount = 3: mstr����IDs = lng����ID
    
    intMouse = Screen.MousePointer
    strFamilyPatiIDs = ""
    Screen.MousePointer = 0
    
    '��ȡ���￨��Ϣ
    On Error GoTo errH
    If zlGetOneCardComLibObject(Me, mlngModul, mobjOneCardComLib) = False Then Exit Function
    
    If blnFamilyMoney Then '��ȡ������Ϣ
        If mobjOneCardComLib.ZlGetPatiFamilyMember(0, lng����ID, strFamilyPatiIDs) = False Then Exit Function
        If strFamilyPatiIDs <> "" Then mstr����IDs = mstr����IDs & "," & strFamilyPatiIDs
    End If
    
    '����ˢ����ֱ֤�ӷ���
    If Not blnˢ����֤ Then ShowMe = True: Exit Function
    
    '��鲡�˼������Ƿ��п���ֻҪ�����κ�һ���п�����Ҫˢ����79868
    '����:43449���������û�з�����,�������������뼰ˢ������,ֱ�ӽ��пۿ�
    If mobjOneCardComLib.ZlGetPatiCardInfo(mstr����IDs, rsTemp) = False Then Exit Function
    If rsTemp.EOF Then
        '�޼�¼,ֱ�ӷ���true,�����鿨
        ShowMe = True: Exit Function
    Else
        rsTemp.Filter = "����<>'' And ����<>null"
        If rsTemp.RecordCount = 0 And bln�����벻�鿨 Then
            '���п���������,ֱ�ӷ���true,�����鿨
            ShowMe = True: Exit Function
        End If
    End If
    
    '��ȡ������Ϣ
    If mobjOneCardComLib.zlGetPatiInforFromPatiID(lng����ID, objPati) = False Then
        MsgBox "��ȡ������Ϣʧ�ܣ�����!", vbOKOnly, gstrSysName
        Exit Function
    End If
    lblName.Caption = lblName.Tag & objPati.����
    lblSex.Caption = lblSex.Tag & objPati.�Ա�
    lblAge.Caption = lblAge.Tag & objPati.����
    lblPatiType.Caption = lblPatiType.Tag & objPati.��������
    lblFeeType.Caption = lblFeeType.Tag & objPati.�ѱ�
    
    '��ȡ���˼��������
    cur������� = 0: cur������� = 0
    Set objExpenceSvr = New clsExpenceSvr
    If objExpenceSvr.zlInitCommon(glngSys, mlngModul, gcnOracle, gstrDBUser) = False Then Exit Function
    If objExpenceSvr.zlExseSvr_GetRemainMoneyByBatch(mstr����IDs, bytOperationType, cllPatiFee) = False Then Exit Function
    For i = 1 To cllPatiFee.Count
        Set objPatiFee = cllPatiFee(i)
        If objPatiFee.����ID = lng����ID Then
            cur������� = cur������� + objPatiFee.ʣ���
        Else
            cur������� = cur������� + objPatiFee.ʣ���
        End If
    Next
    lblRest.Caption = lblRest.Tag & Format(cur�������, "0.00")
    lblFamilyRest.Caption = lblFamilyRest.Tag & Format(cur�������, "0.00")
    
    txtMoney.Text = Format(cur���, "0.00")
    Me.Show 1, frmParent
    ShowMe = mblnOk
    
    Screen.MousePointer = intMouse
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ������Ч��
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2012-03-19 17:04:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPassWord As String
    Dim str���� As String
    On Error GoTo errHandle
    
    If mobjCard Is Nothing Then Exit Function
    If mobjCard.���� Like "*����" Then
        str���� = mobjCard.����
    ElseIf mobjCard.���� Like "*���֤" Then
        str���� = "���֤��"
    ElseIf mobjCard.���� Like "*��" Then
        str���� = mobjCard.���� & "����"
    Else
        str���� = mobjCard.���� & "������"
    End If

    If UCase(Trim(txtCard.Text)) = "" Then Exit Function
    If Not InStr("," & mstr����IDs & ",", "," & Val(lblPass.Tag) & ",") > 0 Or Val(lblPass.Tag) = 0 Then
        MsgBox "��ǰ" & str���� & "�벡�˵�" & str���� & "�������", vbExclamation, gstrSysName
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Function
    End If
    
    If Not mblnCheckPassWord Then isValied = True: Exit Function
    strPassWord = gobjComlib.zlCommFun.zlStringEncode(txtPass.Text)
    If strPassWord <> mstrPassWord Then
        If mintCount = 1 Then
            MsgBox "���������������,���������룡", vbExclamation, gstrSysName
        Else
            MsgBox "�����������", vbExclamation, gstrSysName
        End If
        txtPass.Text = "": mintCount = mintCount - 1
        If mintCount = 0 Then
            Unload Me '������󣬿�����2��
        ElseIf txtPass.Enabled Then
            txtPass.SetFocus
        End If
        Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IDKind.ActiveFastKey
End Sub
Private Sub Form_Load()
    Dim intIdKind As Integer
    
    mstrRegSection = "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name & Me.Name
    mlngPreBrushCardTypeID = GetSetting("ZLSOFT", mstrRegSection, "ȱʡ�����ID", 0)
    
    Call CreateObjectKeyboard
    Call IDKind.zlInit(Me, mlngSys, mlngModul, gcnOracle, gstrDBUser, mobjOneCardComLib, "", txtCard)
    If mlngPreBrushCardTypeID <> 0 Then
       intIdKind = IDKind.GetKindIndex(mlngPreBrushCardTypeID)
       If intIdKind <> 0 Then
           IDKind.IDKind = intIdKind
       End If
    End If
    
    Call SetCtrlVisible
    HookDefend txtPass.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not IDKind.GetCurCard Is Nothing Then
         SaveSetting "ZLSOFT", mstrRegSection, "ȱʡ�����ID", IDKind.GetCurCard.�ӿ����
    End If
    
    On Error Resume Next
    Set mobjKeyboard = Nothing
    Set mobjCard = Nothing
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As Card)
    txtCard.PasswordChar = ""
    '85565,���ϴ�,2015/7/10:��������
    mblnBrushCard = objCard.�Ƿ�ˢ�� Or objCard.�Ƿ�ɨ��
    If txtCard.Text <> "" Then txtCard.Text = ""
    txtCard.Locked = Not mblnBrushCard
    If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As Card, objPatiInfor As clsPatientInfo, blnCancel As Boolean)
    txtCard.Text = objPatiInfor.����
    
    If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComlib.zlControl.TxtSelAll txtCard
            Exit Sub
    End If
    
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
    If txtCard.Text = "" Then
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Sub
    End If
    If mblnCheckPassWord Then txtPass.SetFocus: Exit Sub
    Call cmdOK_Click
End Sub

Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub

Private Sub txtCard_Change()
    lblPass.Tag = "": txtCard.Tag = ""
    txtPass.Enabled = txtCard.Text <> ""
    If Not txtPass.Enabled Then txtPass.Text = ""
End Sub

Private Sub txtCard_GotFocus()
    Call gobjComlib.zlControl.TxtSelAll(txtCard)
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    mblnPreCard = False

    '�Ƿ�ˢ�����
    blnCard = KeyAscii <> 8 And Len(txtCard.Text) = IDKind.GetCurCard.���ų��� - 1 And txtCard.SelLength <> Len(txtCard.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txtCard.Text = txtCard.Text & Chr(KeyAscii)
            txtCard.SelStart = Len(txtCard.Text)
        End If
        KeyAscii = 0
        If GetPatient(IDKind.GetCurCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComlib.zlControl.TxtSelAll txtCard
            Exit Sub
        End If
        mblnPreCard = blnCard
        If mblnCheckPassWord Then
            If txtPass.Enabled Then txtPass.SetFocus
        Else
            Call cmdOK_Click: Exit Sub
        End If
    Else
        If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If

        '��ȫˢ�����
        If KeyAscii <> 0 And KeyAscii > 32 And IDKind.GetCurCard.�Ƿ�ֿ����� = True Then
            sngNow = Timer
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

Private Sub txtPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    glngTXTProc = GetWindowLong(txtPass.hWnd, GWL_WNDPROC)
    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, AddressOf WndMessage)
End Sub

Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button <> 2 Then Exit Sub
    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtPass_GotFocus()
    If txtCard.Text <> "" And mstrPassWord = "" Then Call cmdOK_Click: Exit Sub
    Call gobjComlib.zlControl.TxtSelAll(txtPass)
    OpenPassKeyboard txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If mblnPreCard Then
            '60580
            mblnPreCard = False
             If (GetAsyncKeyState(VK_RETURN) And &H1) <> 0 Then
                txtPass.Text = ""
                Exit Sub
             End If
        End If
        mblnPreCard = False
        Call cmdOK_Click
    ElseIf KeyAscii = 22 Then
        KeyAscii = 0 '������ճ��
    Else
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        End If
    End If
    '60580
    mblnPreCard = False
End Sub

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
    If gobjComlib.ErrCenter() = 1 Then
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
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    Optional blnIDCard As Boolean = False, Optional blnICCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:objCard-��ָ���Ŀ������ж���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long
    
    On Error GoTo errH
    
    mstrPassWord = ""
    Set mobjCard = Nothing
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Function
    If mobjOneCardComLib.zlGetPatiID(lng�����ID, strInput, True, lng����ID, strPassWord, strErrMsg, lng�����ID, Nothing, Me, False, True) = False Then
        '����ģ������:-1:ҽ�ƿ����(���������ǰ�Ŀ��ų��Ȳ����Ļ�,���������)
        If mobjOneCardComLib.zlGetPatiID(-1, strInput, True, lng����ID, strPassWord, strErrMsg, lng�����ID, Nothing, Me, False, True) = False Then
            GoTo NotFoundPati:
        End If
    End If
    If lng����ID <= 0 Then GoTo NotFoundPati:
    If Not InStr("," & mstr����IDs & ",", "," & lng����ID & ",") > 0 Then
        If objCard.���� Like "*����" Then
            MsgBox "��ǰ" & objCard.���� & "�벡�������е�" & objCard.���� & "�����,���飡", vbExclamation, gstrSysName
        ElseIf objCard.���� Like "*���֤" Then
            MsgBox "��ǰ���֤���벡�������е����֤�Ų����,���飡", vbExclamation, gstrSysName
        ElseIf objCard.���� Like "*��" Then
            MsgBox "��ǰ" & objCard.���� & "�����벡�������е�" & objCard.���� & "���Ų����,���飡", vbExclamation, gstrSysName
        Else
            MsgBox "��ǰ" & objCard.���� & "�������벡�������е�" & objCard.���� & "�����Ų����,���飡", vbExclamation, gstrSysName
        End If
        txtCard.Text = ""
        Exit Function '���Ų�ƥ�䣬��׼����
    End If
    txtCard.Tag = strInput
    lblPass.Tag = lng����ID
    mstrPassWord = strPassWord
    Set mobjCard = objCard
    GetPatient = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Exit Function
NotFoundPati:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
    Else
        MsgBox "δ�ҵ���ǰ���ĳ��в���,����!", vbOKOnly + vbInformation, gstrSysName
        txtCard.Text = ""
    End If
    txtCard.Tag = "": lblPass.Tag = ""
End Function

Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���visible����
    '����:���˺�
    '����:2012-03-13 11:28:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    lblFamilyRest.Visible = InStr(mstr����IDs, ",") > 0 'û�м��������ؼ���������ʾ��79868
    lblPass.Visible = mblnCheckPassWord
    txtPass.Visible = mblnCheckPassWord
    If mblnCheckPassWord Then Exit Sub
    With txtCard
        .Top = picTop.Top + picTop.Height + (fraDown.Top - (picTop.Top + picTop.Height) - .Height) \ 2
        IDKind.Top = .Top
        lblCardNO.Top = .Top + (.Height - lblCardNO.Height) \ 2
    End With
    If Err <> 0 Then Err.Clear
End Sub


