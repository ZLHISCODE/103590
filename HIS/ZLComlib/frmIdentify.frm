VERSION 5.00
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmIdentify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���������֤"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "frmIdentify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
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
         Size            =   14.25
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
         Size            =   14.25
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
         Size            =   10.5
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
         Size            =   10.5
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
      ScaleHeight     =   1050
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
         Height          =   720
         Left            =   240
         Picture         =   "frmIdentify.frx":058A
         Top             =   135
         Width           =   720
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
      _ExtentX        =   1111
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
         Size            =   14.25
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
         Size            =   14.25
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
         Size            =   14.25
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
Private mblnOK As Boolean
Private mintCount As Integer
Private mstr����IDs As String
Private mlngSys As Long
Private mblnPreCard As Boolean
Private mobjCard As Card '��ǰ����Ŀ�
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
Private Const VK_RETURN = &HD
Private mblnCheckPassWord As Boolean
Private mblnReadIDCard As Boolean  '��ȡ�������֤
Private mblnReadICCard As Boolean
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard  '����:47945
Attribute mobjICCard.VB_VarHelpID = -1
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mstrRegSection As String
Private mlngPreBrushCardTypeID As Long '�ϴ�ˢ�����
'--------------------------------------------------
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
    mblnCheckPassWord = blnCheckPassWord
    mlngSys = lngSys: mlngModul = lngModul: mlngDefaultCardTypeID = lngDefaultCardTypeID
    mblnOK = False: mintCount = 3: mstr����IDs = lng����ID
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0
    
    '��ȡ���￨��Ϣ
    On Error GoTo ErrH
    '������Ϣ��Ԥ�����
    strSQL = "Select ����id, Nvl(Sum(Ԥ�����), 0) - Nvl(Sum(�������), 0) As ���" & vbNewLine & _
            " From �������" & vbNewLine & _
            " Where ����id = [1] And ���� = 1 And Decode([2],0,0,����)=[2]" & vbNewLine & _
            " Group By ����id"
    '���˲�����Ԥ������¼ʱ��ȡ����ID���ڶ�ȡ������Ϣ
    strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select To_Number([1]) As ����ID, 0 As ��� From Dual"
    If blnFamilyMoney Then
        '���˼�����Ϣ��Ԥ�����
        strSQL = strSQL & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select b.����id, Nvl(Sum(b.Ԥ�����), 0) - Nvl(Sum(b.�������), 0) As ���" & vbNewLine & _
                " From ���˼��� A, ������� B" & vbNewLine & _
                " Where a.����id = b.����id And a.����id = [1] And b.���� = 1 And Decode([2],0,0,b.����)=[2] " & vbNewLine & _
                "       And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & vbNewLine & _
                " Group By b.����id"
        '���˼���������Ԥ������¼ʱ��ȡ���˼���ID���ڶ�ȡ������Ϣ
        strSQL = strSQL & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select ����id, 0 As ��� From ���˼��� Where ����id = [1] " & vbNewLine & _
                "       And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)"
    End If
    strSQL = "Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�ѱ�, a.���￨��, a.����֤��, Nvl(b.���, 0) As ���" & vbNewLine & _
            " From ������Ϣ A, (" & strSQL & ") B" & vbNewLine & _
            " Where a.����id = b.����id And a.ͣ��ʱ�� Is Null"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˺ͼ�������Ϣ��Ԥ�����", lng����ID, bytOperationType)

    '1-������������Ϣ
    Dim cur������� As Currency
    strFamilyPatiIDs = "": cur������� = 0
    rsTmp.Filter = "����id<>" & lng����ID
    Do While Not rsTmp.EOF
        If InStr(strFamilyPatiIDs & ",", "," & gobjComLib.zlCommFun.NVL(rsTmp!����ID) & ",") = 0 Then
            strFamilyPatiIDs = strFamilyPatiIDs & "," & gobjComLib.zlCommFun.NVL(rsTmp!����ID)
            cur������� = cur������� + Val(gobjComLib.zlCommFun.NVL(rsTmp!���))
        End If
        rsTmp.MoveNext
    Loop
    If strFamilyPatiIDs <> "" Then strFamilyPatiIDs = Mid(strFamilyPatiIDs, 2)
    
    '����ˢ����ֱ֤�ӷ���
    If Not blnˢ����֤ Then ShowMe = True: Exit Function
    
    If strFamilyPatiIDs <> "" Then mstr����IDs = mstr����IDs & "," & strFamilyPatiIDs
    '2-���˱�����Ϣ
    rsTmp.Filter = "����id=" & lng����ID
    If rsTmp.EOF Then
        MsgBox "������Ϣ������,����!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '��鲡�˼������Ƿ��п���ֻҪ�����κ�һ���п�����Ҫˢ����79868
'    If gobjComLib.zlCommFun.NVL(rsTmp!���￨��) = "" Then
        '����:43449���������û�з�����,�������������뼰ˢ������,ֱ�ӽ��пۿ�
        strSQL = _
        "Select Count(1) As ���ڿ�, Sum(Decode(����, Null, 0, 1)) As ��������" & vbNewLine & _
        "From ����ҽ�ƿ���Ϣ" & vbNewLine & _
        "Where ״̬ = 0 And ����id In (Select /*+cardinality(a,10)*/ Column_Value From Table(f_Num2list([1])) A)"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��鲡�˻�����Ƿ񷢿�", mstr����IDs)
        If rsTemp.EOF Then
            '�޼�¼,ֱ�ӷ���true,�����鿨
            ShowMe = True: Exit Function
        Else
            If Val(gobjComLib.zlCommFun.NVL(rsTemp!���ڿ�)) = 0 Then
                'δ����,ֱ�ӷ���true,�����鿨
                ShowMe = True: Exit Function
            End If
            If Val(gobjComLib.zlCommFun.NVL(rsTemp!��������)) = 0 And bln�����벻�鿨 Then
                '���п���������,ֱ�ӷ���true,�����鿨
                ShowMe = True: Exit Function
            End If
        End If
'    End If
    
    If Not rsTmp.EOF Then
        lblName.Caption = lblName.Tag & gobjComLib.zlCommFun.NVL(rsTmp!����)
        lblSex.Caption = lblSex.Tag & gobjComLib.zlCommFun.NVL(rsTmp!�Ա�)
        lblAge.Caption = lblAge.Tag & gobjComLib.zlCommFun.NVL(rsTmp!����)
        lblPatiType.Caption = lblPatiType.Tag & gobjComLib.zlCommFun.NVL(rsTmp!��������)
        lblFeeType.Caption = lblFeeType.Tag & gobjComLib.zlCommFun.NVL(rsTmp!�ѱ�)
        
        lblRest.Caption = lblRest.Tag & Format(Val(gobjComLib.zlCommFun.NVL(rsTmp!���)), "0.00")
    End If
    lblFamilyRest.Caption = lblFamilyRest.Tag & Format(cur�������, "0.00")
    
    txtMoney.Text = Format(cur���, "0.00")
'        txtCard.Tag = .NVL(rsTmp!���￨��)
'        txtPass.Tag = .NVL(rsTmp!����֤��)
    On Error GoTo 0
    Me.Show 1, frmParent
    ShowMe = mblnOK
    
    Screen.MousePointer = intMouse
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ������Ч��
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2012-03-19 17:04:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPassWord As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim blnSucces As Boolean '����ɹ�
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
    
    If Not mblnCheckPassWord Then IsValied = True: Exit Function
    strPassWord = gobjComLib.zlCommFun.zlStringEncode(txtPass.Text)
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
    IsValied = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Sub cmdOK_Click()
    If IsValied = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IDKind.ActiveFastKey
End Sub
Private Sub Form_Load()
    mstrRegSection = "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name & Me.Name
    mlngPreBrushCardTypeID = GetSetting("ZLSOFT", mstrRegSection, "ȱʡ�����ID", 0)

    Call CreateObjectKeyboard
    Call zlCardSquareObject
    Call SetCtrlVisible
    Call NewCardObject
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not IDKind.GetCurCard Is Nothing Then
         SaveSetting "ZLSOFT", mstrRegSection, "ȱʡ�����ID", IDKind.GetCurCard.�ӿ����
    End If
    
    Set mobjKeyboard = Nothing
    Set mobjCard = Nothing
    Call zlCardSquareObject(True)
    Call CloseIDCard
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard Is Nothing Then Exit Sub
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        
        If mobjICCard Is Nothing Then Exit Sub
        txtCard.MaxLength = 0
        txtCard.Text = mobjICCard.Read_Card()
        If txtCard.Text = "" Then
            If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
            Exit Sub
        End If
        
            '�����:42948
        If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            txtCard.Text = "": If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
            Exit Sub
        End If
        mblnReadICCard = True
        If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
        If txtCard.Text <> "" Then
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
            Exit Sub
        End If
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus: Exit Sub
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    
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
    If mobjSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtCard.Text = strOutCardNO
    
    '�����:42948
    If txtCard.Text = "" Then
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Sub
    End If
    
    If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
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

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    txtCard.PasswordChar = IIf(objCard.�������Ĺ��� <> "", "*", "")
    '85565,���ϴ�,2015/7/10:��������
    mblnBrushCard = objCard.�Ƿ�ˢ�� Or objCard.�Ƿ�ɨ��
    If txtCard.Text <> "" Then txtCard.Text = ""
    txtCard.Locked = Not mblnBrushCard
    If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)

    txtCard.Text = objPatiInfor.����
    If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
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

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    'IC����ȡ
    
    If strCardNO = "" Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtCard.MaxLength = Len(strCardNO)
    txtCard.Text = strCardNO: mblnReadICCard = True
    If GetPatient(objCard, strCardNO) = False Then
         mblnReadICCard = False: Exit Sub
    End If
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    '��ʾ����Ϣ
    If strID = "" Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtCard.Text = strID: mblnReadICCard = True
    txtCard.MaxLength = Len(strID)
    If GetPatient(objCard, strID) = False Then
         mblnReadICCard = False: Exit Sub
    End If
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub
Private Sub txtCard_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    IDKind.SetAutoReadCard (False)
End Sub

Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub
Private Sub txtCard_Change()
    lblPass.Tag = "": txtCard.Tag = ""
    txtPass.Enabled = txtCard.Text <> ""
    If Not txtPass.Enabled Then txtPass.Text = ""
    mblnReadIDCard = False
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtCard.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtCard.Text = "")
    IDKind.SetAutoReadCard (txtCard.Text = "")
End Sub

Private Sub txtCard_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtCard)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtCard.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtCard.Text = "")
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
            gobjComLib.zlControl.TxtSelAll txtCard
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
    Call gobjComLib.zlControl.TxtSelAll(txtPass)
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
    If gobjComLib.ErrCenter() = 1 Then
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
    If gobjComLib.ErrCenter() = 1 Then Resume
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
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function
 
 
Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String, i As Integer, intIdKind As Integer
    'ֻ��:ִ�л��˷�ʱ,�ſ��ܹܽ��㿨��
    If blnClosed Then
       If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.CloseWindows
            Set mobjSquareCard = Nothing
        End If
        Exit Sub
    End If
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitCompoent (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Call mobjSquareCard.zlInitComponents(Me, mlngModul, mlngSys, gstrDBUser, gcnOracle, False, strExpend)
    mobjSquareCard.mblnYLMgr = True
    Err = 0: On Error GoTo 0
    Call IDKind.zlInit(Me, mlngSys, mlngModul, gcnOracle, gstrDBUser, mobjSquareCard, "", txtCard)
    
    Err = 0: On Error Resume Next
     If mlngPreBrushCardTypeID <> 0 Then
        intIdKind = IDKind.GetKindIndex(mlngPreBrushCardTypeID)
        If intIdKind <> 0 Then
            IDKind.IDKind = intIdKind
        End If
     End If
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    Optional blnIDCard As Boolean = False, Optional blnICCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:objCard-��ָ���Ŀ������ж���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo ErrH
    
    mstrPassWord = ""
    Set mobjCard = Nothing
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Function
    '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
    If mobjSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID, Nothing, Me, False, True) = False Then
        '����ģ������:-1:ҽ�ƿ����(���������ǰ�Ŀ��ų��Ȳ����Ļ�,���������)
        If mobjSquareCard.zlGetPatiID(-1, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID, Nothing, Me, False, True) = False Then
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
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
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

Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ر�������������
    '����:���˺�
    '����:2012-03-09 16:26:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled (False)
        Set mobjICCard = Nothing
    End If
End Sub
Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���µĿ�����
    '����:���˺�
    '����:2012-03-09 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
    End If
    If mobjICCard Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Err = 0: On Error GoTo 0
    End If
End Sub



