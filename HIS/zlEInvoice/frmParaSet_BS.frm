VERSION 5.00
Begin VB.Form frmParaSet_BS 
   BorderStyle     =   0  'None
   Caption         =   "��˼����Ʊ�ݲ�������"
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame frmPaperCode 
      Caption         =   "ֽ��Ʊ�ݴ���"
      Height          =   1005
      Left            =   90
      TabIndex        =   13
      Top             =   1920
      Width           =   6345
      Begin VB.TextBox txtPaperCode 
         Height          =   285
         Index           =   3
         Left            =   4530
         MaxLength       =   30
         TabIndex        =   38
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txtPaperCode 
         Height          =   285
         Index           =   2
         Left            =   1260
         MaxLength       =   30
         TabIndex        =   16
         Top             =   600
         Width           =   1605
      End
      Begin VB.TextBox txtPaperCode 
         Height          =   285
         Index           =   0
         Left            =   1260
         MaxLength       =   30
         TabIndex        =   14
         Top             =   270
         Width           =   1605
      End
      Begin VB.TextBox txtPaperCode 
         Height          =   285
         Index           =   1
         Left            =   4530
         MaxLength       =   100
         TabIndex        =   15
         Top             =   270
         Width           =   1725
      End
      Begin VB.Label lblPaperCode 
         Caption         =   "Ԥ��Ʊ�ݴ���"
         Height          =   285
         Index           =   3
         Left            =   3390
         TabIndex        =   39
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblPaperCode 
         Caption         =   "����Ʊ�ݴ���"
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblPaperCode 
         Caption         =   "�շ�Ʊ�ݴ���"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label lblPaperCode 
         Caption         =   "�Һ�Ʊ�ݴ���"
         Height          =   315
         Index           =   1
         Left            =   3390
         TabIndex        =   35
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "���ѿ���"
      Height          =   675
      Left            =   120
      TabIndex        =   30
      Top             =   4620
      Width           =   6345
      Begin VB.TextBox txt���Ѷ������� 
         Height          =   285
         Left            =   4500
         MaxLength       =   100
         TabIndex        =   34
         Top             =   270
         Width           =   1665
      End
      Begin VB.TextBox txt���Ѷ��ձ��� 
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   32
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label lbl�������� 
         Caption         =   "���Ѷ�������"
         Height          =   315
         Left            =   3180
         TabIndex        =   33
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label lbl���� 
         Caption         =   "���Ѷ��ձ���"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   330
         Width           =   1335
      End
   End
   Begin VB.CheckBox chk����ÿ�Ʊ 
      Caption         =   "����ÿ��ߵ���Ʊ��"
      Height          =   375
      Left            =   3420
      TabIndex        =   29
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CheckBox chk¼����ԭ�� 
      Caption         =   "�Ƿ��ɲ���Ա¼��Ʊ�ݳ��ԭ��"
      Height          =   375
      Left            =   150
      TabIndex        =   26
      Top             =   5280
      Width           =   3105
   End
   Begin VB.Frame fra��������� 
      Caption         =   "���������"
      Height          =   1530
      Left            =   75
      TabIndex        =   17
      Top             =   2955
      Width           =   6375
      Begin VB.TextBox txtȱʡ����� 
         Height          =   285
         Left            =   5145
         TabIndex        =   28
         Text            =   "99998"
         Top             =   290
         Width           =   1050
      End
      Begin VB.TextBox txtCardNO 
         Height          =   300
         Left            =   5145
         TabIndex        =   25
         Text            =   "-"
         Top             =   1020
         Width           =   1050
      End
      Begin VB.TextBox txtNotCardCode 
         Height          =   315
         Left            =   1950
         TabIndex        =   23
         Text            =   "99999"
         Top             =   1050
         Width           =   1605
      End
      Begin VB.TextBox txtIDCardCode 
         Height          =   285
         Left            =   1950
         TabIndex        =   21
         Text            =   "99998"
         Top             =   675
         Width           =   1605
      End
      Begin VB.ComboBox cboȱʡ����� 
         Height          =   300
         Left            =   1350
         TabIndex        =   19
         Text            =   "cboȱʡ�����"
         Top             =   290
         Width           =   2220
      End
      Begin VB.Label lblCard 
         AutoSize        =   -1  'True
         Caption         =   "���֤�����������"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   27
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label lblNotCardNo 
         AutoSize        =   -1  'True
         Caption         =   "�޿����Ź̶���"
         Height          =   180
         Left            =   3840
         TabIndex        =   24
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblCard 
         AutoSize        =   -1  'True
         Caption         =   "�����޿��Ŀ����ͱ��"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   22
         Top             =   1110
         Width           =   1800
      End
      Begin VB.Label lblCard 
         AutoSize        =   -1  'True
         Caption         =   "ȱʡ�������"
         Height          =   180
         Index           =   1
         Left            =   3840
         TabIndex        =   20
         Top             =   330
         Width           =   1260
      End
      Begin VB.Label lblCard 
         Caption         =   "ȱʡ�����(&D)"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   18
         Top             =   350
         Width           =   1320
      End
   End
   Begin VB.ComboBox cboContentType 
      Height          =   300
      ItemData        =   "frmParaSet_BS.frx":0000
      Left            =   4305
      List            =   "frmParaSet_BS.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1200
      Width           =   2070
   End
   Begin VB.ComboBox cboChar 
      Height          =   300
      ItemData        =   "frmParaSet_BS.frx":0011
      Left            =   1200
      List            =   "frmParaSet_BS.frx":0018
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1620
      Width           =   1605
   End
   Begin VB.TextBox txtKey 
      Height          =   300
      Left            =   1200
      TabIndex        =   6
      Top             =   825
      Width           =   5190
   End
   Begin VB.ComboBox cboVersion 
      Height          =   300
      ItemData        =   "frmParaSet_BS.frx":0022
      Left            =   1200
      List            =   "frmParaSet_BS.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1230
      Width           =   1635
   End
   Begin VB.TextBox txtAppID 
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   465
      Width           =   5190
   End
   Begin VB.ComboBox cboURLType 
      Height          =   300
      ItemData        =   "frmParaSet_BS.frx":0033
      Left            =   315
      List            =   "frmParaSet_BS.frx":003A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   75
      Width           =   825
   End
   Begin VB.TextBox txtAddress 
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Text            =   "http://<ip>:<port>/<service>/api/medical"
      Top             =   75
      Width           =   5190
   End
   Begin VB.Label lblContentType 
      AutoSize        =   -1  'True
      Caption         =   "���ݴ��䷽ʽ(&T)"
      Height          =   180
      Left            =   2970
      TabIndex        =   9
      Top             =   1275
      Width           =   1350
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      Caption         =   "�����ַ���(&B)"
      Height          =   180
      Left            =   15
      TabIndex        =   11
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      Caption         =   "֧�ְ汾(&V)"
      Height          =   180
      Left            =   150
      TabIndex        =   7
      Top             =   1290
      Width           =   990
   End
   Begin VB.Label lblKey 
      AutoSize        =   -1  'True
      Caption         =   "ǩ��˽Կ(&K)"
      Height          =   210
      Left            =   150
      TabIndex        =   5
      Top             =   870
      Width           =   990
   End
   Begin VB.Label lblAppID 
      AutoSize        =   -1  'True
      Caption         =   "Ӧ���ʺ�(&I)"
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   525
      Width           =   990
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      Caption         =   "UR&L"
      Height          =   180
      Left            =   30
      TabIndex        =   0
      Top             =   135
      Width           =   270
   End
End
Attribute VB_Name = "frmParaSet_BS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrAddress = "<ip>:<port>/<service>/api/medical"
Private mblnNotChange As Boolean
Private Const mstrInterfanceName = "��˼����Ʊ��ƽ̨" '�ӿ���

Private Enum PInv_Code
    Pc_�շ� = 0
    Pc_�Һ�
    Pc_����
    Pc_Ԥ��
End Enum

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2020-04-08 10:26:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strText As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    With cboURLType
        .Clear
        .AddItem "http": .ListIndex = .NewIndex
        .AddItem "https"
    End With
   
    With cboVersion
        .Clear
        .AddItem "V2.0.3":  .ListIndex = .NewIndex
        .AddItem "V3.1.0"
    End With
    
    With cboChar
        '����
        .Clear
        .AddItem "UTF8":  .ListIndex = .NewIndex
    End With
    
    With cboContentType
        .Clear
        .AddItem "application/json":  .ListIndex = .NewIndex
    End With
    
    strSql = "Select ID,����,���� From ҽ�ƿ���� where �Ƿ�����=1 Order by ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With cboȱʡ�����
        .Clear
        .AddItem "ȱʡ��ҽ�ƿ�"
        .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        Do While Not rsTemp.EOF
            .AddItem rsTemp!���� & "-" & Nvl(rsTemp!����)
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID))
            rsTemp.MoveNext
        Loop
    End With
    
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',1,'URL_Type','HTTP',NULL);
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',2,'URL_Address','','<ip>:<port>/<service>/api/medical/�ӿڷ����ʶ');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',3,'Ӧ���ʺ�','','��Appid');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',4,'ǩ��˽Կ','','��KEYֵ');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',5,'֧�ְ汾','V2.0.3','Ŀǰֻ֧��:V2.0.3��V3.1.0');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',6,'���ݴ��䷽ʽ','','�ύ�ͷ������ݿ���ΪJSON��ʽ��Content-Type: application/json��');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',7,'�ַ�����','UTF-8','ͳһ����UTF-8�ַ�����');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',8,'ȱʡ�����ID','','ȱʡ��ȡ�Ŀ����');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',9,'ҽ�ƿ����ͱ��','','ȱʡ�����ı��');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',10,'���֤�������ͱ��','999998','ʹ�����֤��Ϊ�ϴ��Ŀ����͵ı��');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',11,'�����޿��Ŀ������','999999','�������κο�ʱ�ϴ��Ŀ����ͱ��');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',12,'�����޿��Ŀ���','-','�������κο�ʱ�ϴ��Ŀ���');
    'insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('��˼����Ʊ��ƽ̨',13,'¼����ԭ��','1','Ʊ�ݳ��ʱ�Ƿ��õ����ɲ���Ա¼����ԭ��');
    
    On Error GoTo errHandle
    
    strSql = "Select �ӿ���,������,upper(������) as ������,����ֵ,˵�� From �����ӿ�����  where �ӿ���='" & mstrInterfanceName & "'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            strText = Nvl(rsTemp!����ֵ)
            
            Select Case Nvl(!������)
            Case "ȱʡ�����ID"
                For i = 0 To cboȱʡ�����.ListCount
                    If cboȱʡ�����.ItemData(i) = Val(Nvl(!����ֵ)) Then cboȱʡ�����.ListIndex = i: Exit For
                Next
            Case "ҽ�ƿ����ͱ��"
                txtȱʡ�����.Text = Nvl(!����ֵ)
            Case "���֤�������ͱ��"
                txtIDCardCode.Text = Nvl(!����ֵ)
            Case "�����޿��Ŀ������"
                txtNotCardCode.Text = Nvl(!����ֵ)
            Case "�����޿��Ŀ���"
                txtCardNO.Text = Nvl(!����ֵ)
            Case UCase("URL_Type")
                If strText <> "" Then
                    Call zlControl.CboLocate(cboURLType, strText)
                    If cboURLType.ListIndex < 0 Then cboURLType.ListIndex = 0
                End If
            Case "Ӧ���ʺ�"
                txtAppID.Text = strText
            Case "ǩ��˽Կ"
                txtKey.Text = strText
            Case "֧�ְ汾"
                If strText <> "" Then
                    Call zlControl.CboLocate(cboVersion, strText)
                    If cboVersion.ListIndex < 0 Then cboVersion.ListIndex = 0
                End If
            Case "�ַ�����"
                If strText <> "" Then
                    Call zlControl.CboLocate(cboChar, strText)
                    If cboChar.ListIndex < 0 Then cboChar.ListIndex = 0
                End If
            Case "���ݴ��䷽ʽ"
                If strText <> "" Then
                    Call zlControl.CboLocate(cboContentType, strText)
                    If cboContentType.ListIndex < 0 Then cboContentType.ListIndex = 0
                End If
            Case UCase("URL_Address")
                 mblnNotChange = True
                If strText <> "" Then
                    txtAddress.Text = strText
                    txtAddress.ForeColor = Me.ForeColor
                Else
                    txtAddress.Text = mstrAddress
                    txtAddress.ForeColor = &HC0C0C0
                End If
                 mblnNotChange = False
            Case "¼����ԭ��"
                chk¼����ԭ��.Value = Val(strText)
            Case "���Ѷ��ձ���"
                txt���Ѷ��ձ���.Text = strText
                txt���Ѷ��ձ���.ToolTipText = Nvl(!˵��)
            Case "���Ѷ�������"
                txt���Ѷ�������.Text = strText
                txt���Ѷ��ձ���.ToolTipText = Nvl(!˵��)
            Case "����ÿ��ߵ���Ʊ��"
                chk����ÿ�Ʊ.Value = Val(strText)
                chk����ÿ�Ʊ.ToolTipText = Nvl(!˵��)
            Case "�շ�ֽ��Ʊ�ݴ���"
                txtPaperCode(Pc_�շ�).Text = strText
            Case "�Һ�ֽ��Ʊ�ݴ���"
                txtPaperCode(Pc_�Һ�).Text = strText
            Case "����ֽ��Ʊ�ݴ���"
                txtPaperCode(Pc_����).Text = strText
            Case "Ԥ��ֽ��Ʊ�ݴ���"
                txtPaperCode(Pc_Ԥ��).Text = strText
            End Select
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitData
End Sub

Private Sub txtAddress_Change()
    If mblnNotChange Then Exit Sub
    If txtAddress.Text = "" Or txtAddress.Text = mstrAddress Then
        txtAddress.ForeColor = &HC0C0C0
        Exit Sub
    End If
    txtAddress.ForeColor = Me.ForeColor
End Sub

Private Sub txtAddress_GotFocus()
    If Not (txtAddress.Text = "" Or txtAddress.Text = mstrAddress) Then Exit Sub
    
    mblnNotChange = True
    txtAddress.Text = ""
    txtAddress.ForeColor = Me.ForeColor
    mblnNotChange = False
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If InStr("'[]����������,.'�ۣ�", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtAddress_LostFocus()
    If Not (txtAddress.Text = "" Or txtAddress.Text = mstrAddress) Then Exit Sub
    mblnNotChange = True
    txtAddress.Text = mstrAddress
    txtAddress.ForeColor = &HC0C0C0
    mblnNotChange = False
End Sub
Private Function SaveParaValue(ByVal ������_In As String, ByVal str����ֵ As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ֵ
    '���:������_In-�����ǲ����źͲ�����
    '     str����ֵ-����Ĳ���ֵ
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-04-08 12:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errHandle
    '  Zl_�����ӿ�����_Set
    strSql = "Zl_�����ӿ�����_Set("
    '�ӿ���_In �����ӿ�����.�ӿ���%Type,
    strSql = strSql & "'" & mstrInterfanceName & "',"
    '����_In   �����ӿ�����.������%Type,
    strSql = strSql & "'" & ������_In & "',"
    '����ֵ_In �����ӿ�����.����ֵ%Type
    strSql = strSql & "'" & str����ֵ & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    SaveParaValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSavePara() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2020-04-08 12:04:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call SaveParaValue(UCase("URL_Type"), cboURLType.Text)
    Call SaveParaValue(UCase("URL_Address"), IIf(Trim(txtAddress.Text) = mstrAddress, "", Trim(txtAddress.Text)))
    Call SaveParaValue(UCase("Ӧ���ʺ�"), Trim(txtAppID.Text))
    Call SaveParaValue(UCase("ǩ��˽Կ"), Trim(txtKey.Text))
    Call SaveParaValue(UCase("֧�ְ汾"), Trim(cboVersion.Text))
    Call SaveParaValue(UCase("���ݴ��䷽ʽ"), Trim(cboContentType.Text))
    Call SaveParaValue(UCase("�ַ�����"), Trim(cboChar.Text))
    Call SaveParaValue(UCase("ȱʡ�����ID"), cboȱʡ�����.ItemData(cboȱʡ�����.ListIndex))
    Call SaveParaValue(UCase("ҽ�ƿ����ͱ��"), Trim(txtȱʡ�����.Text))
    Call SaveParaValue(UCase("���֤�������ͱ��"), Trim(txtIDCardCode.Text))
    Call SaveParaValue(UCase("�����޿��Ŀ������"), Trim(txtNotCardCode.Text))
    Call SaveParaValue(UCase("�����޿��Ŀ���"), Trim(txtCardNO.Text))
    Call SaveParaValue(UCase("¼����ԭ��"), chk¼����ԭ��.Value)
    Call SaveParaValue("���Ѷ��ձ���", Trim(txt���Ѷ��ձ���.Text))
    Call SaveParaValue("���Ѷ�������", Trim(txt���Ѷ�������.Text))
    Call SaveParaValue("����ÿ��ߵ���Ʊ��", chk����ÿ�Ʊ.Value)
    Call SaveParaValue("�շ�ֽ��Ʊ�ݴ���", txtPaperCode(Pc_�շ�).Text)
    Call SaveParaValue("�Һ�ֽ��Ʊ�ݴ���", txtPaperCode(Pc_�Һ�).Text)
    Call SaveParaValue("����ֽ��Ʊ�ݴ���", txtPaperCode(Pc_����).Text)
    Call SaveParaValue("Ԥ��ֽ��Ʊ�ݴ���", txtPaperCode(Pc_Ԥ��).Text)
    zlSavePara = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtAppID_KeyPress(KeyAscii As Integer)
    If InStr("'[]����������,.'�ۣ�", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
    If InStr("'[]����������,.'�ۣ�", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
