VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frm���ķ��Ź��� 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Height          =   1380
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11865
      Begin VB.ComboBox cbo���ϲ��� 
         Height          =   300
         Left            =   900
         TabIndex        =   16
         Text            =   "cbo���ϲ���"
         Top             =   975
         Width           =   4575
      End
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   5850
         ScaleHeight     =   825
         ScaleWidth      =   5940
         TabIndex        =   7
         Top             =   120
         Width           =   5940
         Begin VB.CommandButton cmdIC 
            Caption         =   "����"
            Height          =   300
            Left            =   5280
            TabIndex        =   11
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox txtPati 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2820
            TabIndex        =   10
            Top             =   30
            Width           =   2520
         End
         Begin VB.CommandButton cmd���� 
            Caption         =   "��"
            Height          =   255
            Left            =   5040
            TabIndex        =   9
            Top             =   398
            Width           =   285
         End
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   2820
            TabIndex        =   8
            Top             =   375
            Width           =   2520
         End
         Begin MSComctlLib.TabStrip tbsType 
            Height          =   255
            Left            =   780
            TabIndex        =   12
            Top             =   405
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   450
            MultiRow        =   -1  'True
            Style           =   2
            HotTracking     =   -1  'True
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "�ٴ�"
                  Key             =   "T1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "ҽ��"
                  Key             =   "T2"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "����"
                  Key             =   "T3"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin zlIDKind.IDKindNew IDKNType 
            Height          =   375
            Left            =   1800
            TabIndex        =   13
            Top             =   0
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            ShowSortName    =   0   'False
            IDKindStr       =   "ס|סԺ��|0|0|0|0|0|;��|����|0|0|0|0|0|;��|����|0|0|0|0|0|;��|����id|0|0|0|0|0|;��|�����|0|0|0|0|0|;IC|IC����|1|0|0|0|0|"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   9
            FontName        =   "����"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            DefaultCardType =   "0"
            AutoSize        =   -1  'True
            AllowAutoICCard =   -1  'True
            AllowAutoCommCard=   0   'False
            BackColor       =   -2147483644
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϣ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   840
            TabIndex        =   15
            Top             =   90
            Width           =   720
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Left            =   0
            TabIndex        =   14
            Top             =   435
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdˢ�� 
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   10665
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   1100
      End
      Begin VB.TextBox txtEDIT 
         Height          =   300
         Index           =   0
         Left            =   900
         MaxLength       =   8
         TabIndex        =   5
         Top             =   585
         Width           =   2085
      End
      Begin VB.TextBox txtEDIT 
         Height          =   300
         Index           =   1
         Left            =   3405
         MaxLength       =   8
         TabIndex        =   4
         Top             =   585
         Width           =   2085
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   1
         Left            =   5850
         TabIndex        =   1
         Top             =   885
         Width           =   4800
         Begin VB.CheckBox chkType 
            Caption         =   "סԺ"
            Height          =   240
            Index           =   1
            Left            =   960
            TabIndex        =   3
            Top             =   150
            Value           =   1  'Checked
            Width           =   2145
         End
         Begin VB.CheckBox chkType 
            Caption         =   "����"
            Height          =   240
            Index           =   0
            Left            =   105
            TabIndex        =   2
            Top             =   150
            Value           =   1  'Checked
            Width           =   885
         End
      End
      Begin MSComCtl2.DTPicker Dtp��ʼDate 
         Height          =   300
         Left            =   900
         TabIndex        =   17
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   80740355
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker Dtp����Date 
         Height          =   300
         Left            =   3405
         TabIndex        =   18
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   80740355
         CurrentDate     =   37007
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ϲ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   23
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   3120
         TabIndex        =   22
         Top             =   645
         Width           =   180
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ݷ�Χ"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ�䷶Χ"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   19
         Top             =   270
         Width           =   180
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ķ��Ź���.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ķ��Ź���.frx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm���ķ��Ź���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mArrFilter As Variant
Private mintType As Integer
Private mstrPrivs As String
Private mlngModule As Long
Private mblnCard As Boolean     '�Ƿ�ˢ���Ǿ��￨
Private mobjcard As Card
Private mintOld����ģʽ As Integer

Private Enum mFindType
    סԺ�� = 0
    ���� = 1
    ���� = 2
    ����ID = 3
    ����� = 4
    IC���� = 5
End Enum

Private Enum mtxtIdx
    idx_��ʼNO = 0
    idx_����NO = 1
End Enum

Private mblnDrop As Boolean                     '��KeyDown���ж������б��Ƿ񵯳�

Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

'--------------------------------------------------------------------------------------------------------
'ҩƷ��ҩ����
Private mblnTrans As Boolean            'True��ʾ��ҩƷ������ҩ���ڵ���
Private mstrNo  As String               '���ݺţ������ڶ�λ
Private mlng�ⷿid As Long              '��ҩ�ⷿID��һ��ͷ��ϲ���һ��
Private mstrDrugStartDate As String     'ҩƷ���ݿ�ʼʱ��
Private mstrDrugEndDate As String       'ҩƷ���ݽ���ʱ��
Private mlng����id As Long
Private mlngPre����ID As Long
Private mblnNoClick As Boolean

'--------------------------------------------------------------------------------------------------------
Public Event zlRefreshCon(ByVal arrFilter As Variant)
Public Event zlPopupMenus(ByVal x As Long, ByVal Y As Long)

Private Sub InitIDKindNew()
    Dim int����ģʽ As Integer
    Dim strTemp As String

    strTemp = "ס|סԺ��|0;��|����|0;��|����|0;��|����id|0;��|�����|0;IC|IC����|1"
    Me.IDKNType.IDKindStr = strTemp
    Call IDKNType.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquareCard, strTemp, txtPati)
    IDKNType.SetAutoReadCard True
    Me.IDKNType.IDKind = 0
    
End Sub

Private Sub chkType_Click(Index As Integer)
    Call cmdˢ��_Click
End Sub

Private Sub IDKNType_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    Set mobjcard = objCard
    mintType = Index - 1
    mintOld����ģʽ = mintType
    
    txtPati.Text = ""
    txtPati.MaxLength = objCard.���ų���
    If objCard.�������Ĺ��� <> "" Then
        txtPati.PasswordChar = "*"
    Else
        txtPati.PasswordChar = ""
    End If
    
End Sub

Private Sub IDKNType_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPati.Text = objPatiInfor.����
    If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    If Not txtPati.Locked And txtPati.Text = "" And Me.ActiveControl Is txtPati And strNo <> "" Then
        txtPati.Text = strNo

        If txtPati.Text = "" Then
            Call mobjICCard.SetEnabled(False)
        Else
'            Me.PatiTittle = 6

            Call txtPati_KeyDown(vbKeyReturn, 0)
        End If
    End If
End Sub
Private Function CheckDepend() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim lng���ϲ���ID As Long
    
    On Error GoTo ErrHandle
    CheckDepend = False
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.���� || '-' || a.���� As ���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.���� And (a.վ��=[2] or a.վ�� is null) " & _
        "       AND b.���� ='W' " & _
        "       AND a.id = c.����id " & _
        "       AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
        IIf(InStr(mstrPrivs, "���в���") <> 0, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])") & _
        " Order by a.���� || '-' || a.����"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ӧ�Ŀⷿ", UserInfo.Id, gstrNodeNo)
    
    If rsTemp.EOF Then
        rsTemp.Close
        Exit Function
    End If
    
    '�����ҩƷ���ڴ��룬���÷��ϲ�����ҩƷ��ҩ����һ��
    If mblnTrans Then
        If mlng�ⷿid <> UserInfo.����ID Then
            lng���ϲ���ID = mlng�ⷿid
        Else
            lng���ϲ���ID = UserInfo.����ID
        End If
    End If
    'װ�뷢�ϲ�������
    With cbo���ϲ���
        .Clear
        mblnNoClick = True
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = lng���ϲ���ID Then
                .ListIndex = .NewIndex
                mlngPre����ID = lng���ϲ���ID
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0: mlngPre����ID = .ItemData(.ListIndex)
        mblnNoClick = False
        rsTemp.Close
    End With
    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetFilter() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-30 11:52:50
    '-----------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, strReg As String
    Dim int�շѴ��� As Integer
    Dim lng����id As Long
    Dim strCard As String

    
    strReg = Trim(zlDatabase.GetPara("��ѯҵ������", glngSys, mlngModule, ""))
    If strReg = "" Then strReg = "24,25,26"
    
    strCard = Split(Split(IDKNType.IDKindStr, ";")(mintType), "|")(1)
    '������ѯ����
    Set cllFilter = New Collection
    
    int�շѴ��� = Val(zlDatabase.GetPara("�շѴ�����ʾ��ʽ", glngSys, mlngModule, 0))
    
    If int�շѴ��� < 0 Or int�շѴ��� > 2 Then
        int�շѴ��� = 0
    End If
    
    cllFilter.Add int�շѴ���, "�շѴ���"
    
    cllFilter.Add txtPati.Text, "����"
    
    If cbo���ϲ���.ListIndex < 0 Then
        cllFilter.Add 0, "���ϲ���ID"
    Else
        cllFilter.Add cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex), "���ϲ���ID"
    End If
    cllFilter.Add Array(Format(Dtp��ʼDate.Value, "yyyy-mm-dd HH:MM:SS"), Format(Dtp����Date.Value, "yyyy-mm-dd HH:MM:SS")), "���ڷ�Χ"
    cllFilter.Add strReg, "����"
    If Trim(txt����.Tag) = "" Then
        cllFilter.Add "", "��������ID"
    Else
        cllFilter.Add Trim(txt����.Tag), "��������ID"
    End If
    
    If tbsType.SelectedItem Is Nothing Then
        cllFilter.Add 0, "��������"
    Else
        cllFilter.Add tbsType.SelectedItem.Index - 1, "��������"
    End If
   
    cllFilter.Add Array(Trim(txtEDIT(mtxtIdx.idx_��ʼNO)), Trim(txtEDIT(mtxtIdx.idx_����NO))), "���ݺ�"
    If strCard = "סԺ��" Then
        cllFilter.Add Val(txtPati.Text), "סԺ��"
    Else
        cllFilter.Add 0, "סԺ��"
    End If
    If strCard = "����" Then
        If mblnCard = True Then
            cllFilter.Add Trim(txtPati.Text), "���￨��"
            cllFilter.Add "", "����"
        Else
            cllFilter.Add Trim(txtPati.Text), "����"
        End If
    Else
        cllFilter.Add "", "����"
    End If
    
    If strCard = "����" Then
        cllFilter.Add Trim(txtPati.Text), "����"
    Else
        cllFilter.Add "", "����"
    End If
    If strCard = "����id" Then
        cllFilter.Add Val(txtPati.Text), "����ID"
        cllFilter.Add 0, "IC����"
    ElseIf strCard = "IC��" Then
        If Not gobjSquareCard Is Nothing Then
            Call gobjSquareCard.zlGetPatiID("IC��", txtPati.Text, True, lng����id)
        End If
        cllFilter.Add lng����id, "����ID"
        If txtPati.Text <> "" Then
            cllFilter.Add 1, "IC����"
        Else
            cllFilter.Add 0, "IC����"
        End If
    Else
        '���п�
        If Not gobjSquareCard Is Nothing And strCard <> "����" And strCard <> "����" And strCard <> "סԺ��" And strCard <> "�����" Then
            If gobjSquareCard.zlGetPatiID(mobjcard.�ӿ����, txtPati.Text, False, lng����id) = False And txtPati.Text <> "" Then lng����id = -1
        End If
        cllFilter.Add lng����id, "����ID"
        cllFilter.Add 0, "IC����"
    End If
    
    If strCard = "�����" Then
        cllFilter.Add Val(txtPati.Text), "�����"
    Else
        cllFilter.Add "", "�����"
    End If
    
    If Not (strCard = "����" And mblnCard) Then
        cllFilter.Add "", "���￨��"
    End If
    
    If (chkType(0).Value = 1 And chkType(1).Value = 1) Or (chkType(0).Value = 0 And chkType(1).Value = 0) Then
        '0-����
        cllFilter.Add 0, "��������"
    ElseIf chkType(0).Value = 1 Then
        '1-���Ｐ����
        cllFilter.Add 1, "��������"
    ElseIf chkType(1).Value = 1 Then
        '2-סԺ���ʵ�
        cllFilter.Add 2, "��������"
    End If
    
    'zlDatabase.OpenSQLRecord(gstrsql, Me.Caption, _
        Val(mArrFilter("���ϲ���ID")), _
        CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
        CStr("," & mArrFilter("����") & ","), _
        Val(mArrFilter("��������ID")), _
        CStr(mArrFilter("���ݺ�")(0)), CStr(mArrFilter("���ݺ�")(1)), _
        Val(mArrFilter("����ID")), Val(mArrFilter("סԺ��")), _
        CStr(mArrFilter("����")))
        
    Set mArrFilter = cllFilter
    
End Function

Private Sub cbo���ϲ���_Click()
    If cbo���ϲ���.ListIndex < 0 Then Exit Sub
    If mblnNoClick = True Then Exit Sub
    
    If mlngPre����ID <> cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex) Then
        mlngPre����ID = cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex)
        Call cmdˢ��_Click
    End If
End Sub

Private Sub cbo���ϲ���_KeyDown(KeyCode As Integer, Shift As Integer)
'    mblnDrop = False
'    If KeyCode = 13 Then mblnDrop = SendMessage(cbo���ϲ���.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1
'
''    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab

    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If cbo���ϲ���.ListIndex >= 0 Then
        If mlngPre����ID <> cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex) Then
            mlngPre����ID = cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex)
            Call cmdˢ��_Click
        End If
    End If
    
    If Select����ѡ����(Me, cbo���ϲ���, Trim(cbo���ϲ���.Text), "W", False) = False Then
        DoEvents
        
        cbo���ϲ���.SetFocus
        Exit Sub
    End If
    
    If cbo���ϲ���.ListIndex >= 0 Then
        If mlngPre����ID <> cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex) Then
            mlngPre����ID = cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex)
            Call cmdˢ��_Click
        End If
    End If
End Sub

Private Sub cbo���ϲ���_KeyPress(KeyAscii As Integer)
'    Dim i As Long, intIdx As Integer
'    Dim strText As String, strResult As String, strFilter As String
'
'    If KeyAscii = 13 Then
'        strText = UCase(cbo���ϲ���.Text)
'        If cbo���ϲ���.ListIndex <> -1 Then
'            '�����б�ʱ,�����ı�������������
'            If strText <> cbo���ϲ���.List(cbo���ϲ���.ListIndex) Then Call zlControl.CboSetIndex(cbo���ϲ���.hwnd, -1)
'        End If
'        If strText = "" Then
'            cbo���ϲ���.ListIndex = -1
'        ElseIf cbo���ϲ���.ListIndex = -1 Then
'            intIdx = -1
'
'            For i = 1 To cbo���ϲ���.ListCount - 1
'                If Mid(cbo���ϲ���.List(i), 1, InStr(1, cbo���ϲ���.List(i), "-") - 1) = strText _
'                    Or Mid(cbo���ϲ���.List(i), InStr(1, cbo���ϲ���.List(i), "-")) = strText Then
'                    intIdx = i
'                    Exit For
'                End If
'            Next
'
'            If intIdx = -1 Then
'                For i = 1 To cbo���ϲ���.ListCount - 1
'                    If UCase(cbo���ϲ���.List(i)) Like strText & "*" Then
'                        intIdx = i
'                    End If
'                Next
'            End If
'
'            cbo���ϲ���.ListIndex = intIdx
'            SendMessage cbo���ϲ���.hwnd, CB_SHOWDROPDOWN, True, 0
'        ElseIf Not mblnDrop Then
'            '�س���꾭��
'            Call cbo���ϲ���_Click
'            Exit Sub
'        End If
'        If cbo���ϲ���.ListIndex = -1 Then
'            cbo���ϲ���.ListIndex = 0
'        Else
'            If intIdx <> -1 And mblnDrop Then
'                '�����س�-ǿ�м���Click
'                Call cbo���ϲ���_Click
'            ElseIf intIdx <> cbo���ϲ���.ListIndex And intIdx <> -1 Then
'                '������ѡ��-�Զ�����Click
'                cbo���ϲ���.SetFocus
'                Exit Sub
'            ElseIf intIdx <> -1 Then
'                'һ��������-ǿ�м���Click
'                Call cbo���ϲ���_Click
'            End If
'        End If
'    End If
End Sub


Private Sub cbo���ϲ���_Validate(Cancel As Boolean)
'    Dim i As Long
'    Dim blnTmp As Boolean
'
'    If cbo���ϲ���.Text = "" Then
'        cbo���ϲ���.ListIndex = 0
'    Else
'        For i = 0 To cbo���ϲ���.ListCount - 1
'            If cbo���ϲ���.Text = cbo���ϲ���.List(i) Then
'                blnTmp = True
'                Exit For
'            End If
'        Next
'
'        If blnTmp = False Then
'            cbo���ϲ���.ListIndex = 0
'        End If
'    End If
End Sub

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cmdIC_Click()
    Dim strOutXML As String
    Dim strTemp As String
    Dim strCard As String
    
    strCard = Split(Split(IDKNType.IDKindStr, ";")(mintType), "|")(1)
    If strCard = "IC����" Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPati.Text = mobjICCard.Read_Card()
            If txtPati.Text <> "" Then Call cmdˢ��_Click
        End If
    Else
        If Not gobjSquareCard Is Nothing Then
            Call gobjSquareCard.zlReadCard(Me, mlngModule, mobjcard.�ӿ����, True, "", strTemp, strOutXML)
            txtPati.Text = strTemp
            If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
        End If
    End If
End Sub

Private Sub cmd����_Click()
    If Select��������(txt����, "") = False Then Exit Sub
    Call InitData
End Sub

Private Sub cmdˢ��_Click()
    Call GetFilter
    RaiseEvent zlRefreshCon(mArrFilter)
End Sub
Private Sub InitData()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-05-01 21:55:55
    '-----------------------------------------------------------------------------------------------------------
   
    Dtp����Date.MaxDate = Format(sys.Currentdate, "yyyy-mm-dd") & " 23:59:59"
    If mblnTrans Then
        Dtp��ʼDate.Value = CDate(mstrDrugStartDate)
        Dtp����Date.Value = CDate(mstrDrugEndDate)
    Else
         Dtp����Date.Value = Dtp����Date.MaxDate
         Dtp��ʼDate.Value = Format(DateAdd("d", -7, sys.Currentdate), "yyyy-mm-dd") & " 00:00:00"
    End If
    Dtp��ʼDate.MaxDate = Dtp����Date.MaxDate
'    txtEDIT(mtxtIdx.idx_��ʼNO) = mstrNo

    
End Sub


Private Sub Dtp����Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab

End Sub

Private Sub Dtp��ʼDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs:    mlngModule = glngModul: mblnCard = False
    
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
    
    Call InitIDKindNew
    
    Call CheckDepend
    
    cmdIC.Visible = False
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    
    On Error Resume Next
    
    With fra(0)
        .Top = ScaleTop
        .Height = ScaleHeight
        .Left = ScaleLeft
        .Width = ScaleWidth
    End With
    
    With picFilter
        .Width = IIf(ScaleWidth - .Left - 50 < 0, 0, ScaleWidth - .Left - 50)
        cmdˢ��.Left = .Left + .Width - cmdˢ��.Width - 50
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж��һ��ͨ�ӿ�
    gstrCardType = ""
    Set gobjSquareCard = Nothing
    
    'ж��IC��ˢ���ӿ�
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
End Sub


Private Sub picFilter_Resize()
    err = 0: On Error Resume Next
    With txtPati
        .Width = picFilter.ScaleWidth - .Left
        txt����.Width = .Width
        cmd����.Left = .Left + .Width - cmd����.Width - 10
        cmdIC.Left = picFilter.Width - cmdIC.Width - 20
    End With
End Sub
Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then Exit Sub
    RaiseEvent zlPopupMenus(x, Y)
End Sub
Public Property Get PatiTittle() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ����ر���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-01 12:49:27
    '-----------------------------------------------------------------------------------------------------------
    PatiTittle = mintType
End Property

Public Property Get PatiCardID() As Long
    '��ȡ���ѿ������ID
    If mintType > 5 Then
        PatiCardID = mobjcard.�ӿ����
    Else
        PatiCardID = 0
    End If
End Property
'Public Property Let PatiTittle(ByVal vNewValue As Integer)
'
'    mintType = vNewValue
'
'    If mintType <= 5 Then
'        Me.lblPatiInputType.Caption = Decode(mintType, mFindType.סԺ��, "סԺ�š�", mFindType.����, "��  ����", mFindType.����, " ��  �š�", mFindType.����ID, "����ID��", mFindType.�����, "����š�", mFindType.IC����, "IC����")
'    Else
'        '���п�
'        If gstrCardType <> "" Then
'            Me.lblPatiInputType.Caption = Split(Split(gstrCardType, ";")(mintType - 6), "|")(1) & "��"
'        End If
'    End If
'
'    '��ȷΪ���￨�������:
'    '��ǰ�и�����������ģ����￨�к�������ţ�Ҫ���Σ�" :��;��?��"��
'    '�Ժ��֣���������ImeMode����ΪDisable�Ϳ����ˡ�
'    txtPati.IMEMode = 0
'    cmdIC.Visible = False
'    If mintType = mFindType.����ID Or mintType = mFindType.����� Or mintType = mFindType.���� Or mintType = mFindType.סԺ�� Then
'        txtPati.MaxLength = 18
'    ElseIf mintType = mFindType.IC���� Then
'        cmdIC.Visible = True
'    ElseIf mintType > 5 Then
'        '���п�
'        txtPati.Tag = Split(gstrCardType, ";")(mintType - 6)
'        txtPati.MaxLength = Val(Split(txtPati.Tag, "|")(gCardFormat.���ų���))
'        cmdIC.Visible = (Val(Split(txtPati.Tag, "|")(gCardFormat.ˢ����־)) = 1)
'    Else
'        txtPati.MaxLength = 0
'    End If
'
'End Property
Private Function Select��������(ByVal objCtl As Control, ByVal strSearch As String) As Boolean
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strTemp As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    Dim rsCount As ADODB.Recordset
    Dim strSelectSql As String
    Dim strName As String
    Dim int���� As Integer '0-����;1-����;2-סԺ
    
    On Error GoTo ErrHandle
    If (chkType(0).Value = 1 And chkType(1).Value = 1) Or (chkType(0).Value = 0 And chkType(1).Value = 0) Then
        '0-����
        int���� = 0
    ElseIf chkType(0).Value = 1 Then
        '1-���Ｐ����
        int���� = 1
    ElseIf chkType(1).Value = 1 Then
        '2-סԺ���ʵ�
        int���� = 2
    End If
    
    strKey = GetMatchingSting(UCase(strSearch), False)
    
    strTittle = "����ѡ����"
    vRect = zlControl.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    
    If frm���ķ��Ź���_New.tbPage.Selected.Index = 4 Then
        If tbsType.SelectedItem.Index - 1 = 0 Then
            gstrSQL = "" & _
                " Select ID, ����,���� From ���ű� " & _
                " Where ID in (Select ����ID From ��������˵�� Where ��������='�ٴ�')" & _
                "     And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                "     And (վ��=[2] or վ�� is null) "
        ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
            gstrSQL = "" & _
                " Select ID, ����,����,���� From ���ű� " & _
                " Where ID in (Select ����ID From ��������˵�� Where �������� In ('���','����','����','����'))" & _
                "     And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                "     And (վ��=[2] or վ�� is null) "
        Else
            gstrSQL = "" & _
                " Select ID, ����,����,���� From ���ű� " & _
                " Where ID in (Select ����ID From ��������˵�� Where ��������='����')" & _
                "     And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                "     And (վ��=[2] or վ�� is null) "
        End If
        If strSearch <> "" Then
            gstrSQL = gstrSQL & _
                "     And ( ���� like [1] or ���� like [1] or ���� like [1] )"
        End If
        gstrSQL = gstrSQL & vbCrLf & " Order by ����"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ſ���", strKey, gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "û�����ø��ಿ�ţ������Ź���", vbInformation, gstrSysName
                Exit Function
            End If
        End With
        Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, True, strKey, gstrNodeNo)
    Else
        If tbsType.SelectedItem.Index - 1 = 0 Then
            gstrSQL = "" & _
                "Select Distinct A.ID,A.����,A.���� " & _
                "From ���ű� A, ��������˵�� B, δ��ҩƷ��¼ C, ������ü�¼ D " & _
                "Where (A.վ�� = [6] Or A.վ�� Is Null) And B.�������� ='�ٴ�' And A.ID = B.����id " & _
                "   And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.�ⷿid = [2] " & _
                "   And instr([3],','||C.����||',')>0 And C.�������� Between [4] And [5] And C.NO = D.NO And C.�ⷿid = D.ִ�в���id " & _
                "   And A.ID = D.��������id And D.���˿���id = D.��������id " & _
                IIf(strSearch = "", "", " and (A.���� like upper([1]) or A.���� like [1] or A.���� like upper([1]))")
               
            If int���� = 1 Then
                '�����ʹ��������ü�¼
                gstrSQL = gstrSQL & " Order By A.���� "
            ElseIf int���� = 2 Then
                '��סԺ��ʹ��סԺ���ü�¼
                gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                gstrSQL = gstrSQL & " Order By A.���� "
            Else
                '���У��������סԺ���ü�¼
                gstrSQL = gstrSQL & " Union " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                gstrSQL = gstrSQL & " Order By ���� "
            End If
        ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
            gstrSQL = "" & _
                "Select Distinct A.ID,A.����,A.���� " & _
                "From ���ű� A, ��������˵�� B, δ��ҩƷ��¼ C, ������ü�¼ D " & _
                "Where (A.վ�� = [6] Or A.վ�� Is Null) And B.�������� In ('���','����','����','����') " & _
                "   And A.ID = B.����id And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
                "   And C.�ⷿid = [2] And instr([3],','||C.����||',')>0 And C.�������� Between [4] And [5] And C.NO = D.NO " & _
                "   And C.�ⷿid = D.ִ�в���id And A.ID = D.��������id And D.���˿���id <> D.��������id " & _
                IIf(strSearch = "", "", " and (A.���� like upper([1]) or A.���� like [1] or A.���� like upper([1]))")
                
            If int���� = 1 Then
                '�����ʹ��������ü�¼
                gstrSQL = gstrSQL & " Order By A.���� "
            ElseIf int���� = 2 Then
                '��סԺ��ʹ��סԺ���ü�¼
                gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                gstrSQL = gstrSQL & " Order By A.���� "
            Else
                '���У��������סԺ���ü�¼
                gstrSQL = gstrSQL & " Union " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                gstrSQL = gstrSQL & " Order By ���� "
            End If
        Else
            '�Բ���Ϊ����ʱ������������ʱ����ȡ����
            If int���� = 1 Then
                Exit Function
            End If
                
            gstrSQL = "" & _
                "Select Distinct A.ID,A.����,A.���� " & _
                "From ���ű� A, ��������˵�� B, δ��ҩƷ��¼ C, סԺ���ü�¼ D " & _
                "Where (A.վ�� = [6] Or A.վ�� Is Null) And B.�������� = '����' And A.ID = B.����id " & _
                "   And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.�ⷿid = [2] " & _
                "   And instr([3],','||C.����||',')>0 And C.�������� Between [4] And [5] And C.NO = D.NO " & _
                "   And C.�ⷿid = D.ִ�в���id And A.ID = D.���˲���id " & _
                IIf(strSearch = "", "", " and (A.���� like upper([1]) or A.���� like [1] or A.���� like upper([1]))")
                
            If zlDatabase.GetPara("�������Ϸ�ʽ", glngSys, mlngModule) = "" Then
                gstrSQL = gstrSQL & " And D.���˿���id = D.��������id "
            End If
            
            gstrSQL = gstrSQL & " Order By A.���� "
        End If
    
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKey, Val(mArrFilter("���ϲ���ID")), CStr("," & mArrFilter("����") & ","), CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                Exit Function
            End If
           
            Do While Not .EOF
                gstrSQL = "Select Distinct A.ҩƷid " & _
                    " From ҩƷ�շ���¼ A, δ��ҩƷ��¼ B, ������ü�¼ C " & _
                    " Where A.���� = B.���� And A.NO = B.NO And a.�ⷿid = b.�ⷿid And A.����� Is Null And A.NO = C.NO And B.�ⷿid = C.ִ�в���id " & _
                    " And B.�ⷿid = [2] And instr([3],','||B.����||',')>0 And B.�������� Between [4] And [5] "
                    
                If tbsType.SelectedItem.Index - 1 = 0 Then
                    gstrSQL = gstrSQL & " And C.��������id = [1] And C.���˿���id=C.��������id "
                    
                    If int���� = 2 Then
                        '��סԺʱ��ʹ��סԺ���ü�¼
                        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                    ElseIf int���� = 0 Then
                        '����ʱ���������סԺ���ü�¼
                        gstrSQL = gstrSQL & " Union " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                    End If
                ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
                    gstrSQL = gstrSQL & " And C.��������id = [1] And C.���˿���id<>C.��������id "
                    
                    If int���� = 2 Then
                        '��סԺʱ��ʹ��סԺ���ü�¼
                        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                    ElseIf int���� = 0 Then
                        '����ʱ���������סԺ���ü�¼
                        gstrSQL = gstrSQL & " Union " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                    End If
                Else
                    '�Բ���Ϊ����ʱ������������ʱ����ȡ����
                    If int���� = 1 Then
                        Exit Function
                    End If
            
                    If zlDatabase.GetPara("�������Ϸ�ʽ", glngSys, mlngModule) = "" Then
                        gstrSQL = gstrSQL & " And C.���˲���id = [1] And C.���˿���id=C.��������id "
                    Else
                        gstrSQL = gstrSQL & " And C.���˲���id = [1] "
                    End If
                    gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
                End If
                
                gstrSQL = "Select Count(Distinct ҩƷid) As ҩƷ From (" & gstrSQL & ")"

                Set rsCount = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ſ���", CLng(!Id), Val(mArrFilter("���ϲ���ID")), CStr("," & mArrFilter("����") & ","), CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)))
                
                strName = !���� & "(" & rsCount!ҩƷ & "�����Ĵ�����"
                strSelectSql = IIf(strSelectSql = "", "", strSelectSql & " Union All ") & "Select " & !Id & " As ID," & !���� & " As ����," & "'" & strName & "'" & " As ����  From Dual "
                
                .MoveNext
            Loop
        End With
        
        Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSelectSql, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, True)
    End If

    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgBox "û�����������Ŀ���,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    If objCtl.Enabled Then objCtl.SetFocus
    With rsTemp
        objCtl.Tag = ""
        Do While Not .EOF
            strTemp = strTemp & "," & NVL(rsTemp!����)
            objCtl.Tag = objCtl.Tag & "," & NVL(rsTemp!Id)
            .MoveNext
        Loop
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    strKey = objCtl.Tag
    objCtl.Text = strTemp
    objCtl.Tag = strKey
    OS.PressKey vbKeyTab
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tbsType_Click()
    txt����.Text = ""
    txt����.Tag = ""
End Sub

Private Sub tbsType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtEDIT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim intYear As Integer, strYear As String
    Dim strType As String
    Dim intType As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtEDIT(Index)) = "" Then Exit Sub
    '--���������λ,�򰴹������--
    Me.txtEDIT(Index) = UCase(LTrim(Me.txtEDIT(Index)))
    If Len(txtEDIT(Index)) < 8 Then
        strType = Trim(zlDatabase.GetPara("��ѯҵ������", glngSys, mlngModule, ""))
        
        If strType = "" Or strType = "0,0,0" Or InStr(1, strType, "25") > 0 Or InStr(1, strType, "26") > 0 Then
            intType = 14
        Else
            intType = 13
        End If
        
        txtEDIT(Index).Text = zlCommFun.GetFullNO(txtEDIT(Index).Text, intType, cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex))
    End If
    OS.PressKey (vbKeyTab)
End Sub

Private Sub txtPati_Change()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPati.Text = "" And Me.ActiveControl Is txtPati)
End Sub

Private Sub txtPati_GotFocus()
    If Not mobjICCard Is Nothing And txtPati.Text = "" Then
        Call mobjICCard.SetEnabled(True)
    End If
End Sub

Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCard As String
    
    strCard = Split(Split(IDKNType.IDKindStr, ";")(mintType), "|")(1)
    If KeyCode = vbKeyReturn Then
        If strCard = "����" Then
            '�����ʾ,ˢ���￨
            Call cmdˢ��_Click
        ElseIf strCard = "IC����" Then
            'IC��
            Call cmdˢ��_Click
        ElseIf strCard = "סԺ��" Then
            OS.PressKey vbKeyTab
        ElseIf strCard = "����id" Then
            OS.PressKey vbKeyTab
        ElseIf strCard = "����" Then
            OS.PressKey vbKeyTab
        ElseIf strCard = "�����" Then
            OS.PressKey vbKeyTab
        Else
            '���п�
            Call cmdˢ��_Click
        End If
    End If
End Sub
Private Sub txtPati_KeyPress(KeyAscii As Integer)
    Dim strCard As String
    mblnCard = False

    strCard = Split(Split(IDKNType.IDKindStr, ";")(mintType), "|")(1)
    If strCard <> "����" And strCard <> "IC��" And strCard <> "סԺ��" And strCard <> "����id" And strCard <> "����" And strCard <> "�����" Then
        '�������ѿ�
        If Len(txtPati.Text) = txtPati.MaxLength - 1 And KeyAscii <> 8 Then
            txtPati.Text = txtPati.Text & Chr(KeyAscii)
            txtPati.SelStart = Len(txtPati.Text)
            KeyAscii = 0

            cmdˢ��_Click
        End If
    End If
End Sub

Private Sub txtPati_LostFocus()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
End Sub

Private Sub txt����_Change()
    txt����.Tag = ""
End Sub
Public Sub Set���ϴ�������(ByVal blnTran As Boolean, ByVal strNo As String, ByVal strStartDate As String, ByVal strEndDate As String, ByVal lng����id As Long, ByVal lng�ⷿID As Long)
    '-----------------------------------------------------------------------------------------------------------
    '����:������صķ�ҩ���ڴ��������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-01 22:09:07
    '-----------------------------------------------------------------------------------------------------------
    mstrDrugStartDate = strStartDate: mstrDrugEndDate = strEndDate: mblnTrans = blnTran
    mlng�ⷿid = lng�ⷿID: mlng����id = lng����id: mstrNo = strNo
    
    Call InitData
    
    If mlng����id <> 0 Then
        Me.txtPati.Text = mlng����id
        mintType = 4
        Me.IDKNType.IDKind = 4
        
    End If
End Sub
Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txt����.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If Select��������(txt����, Trim(txt����.Text)) = False Then
        DoEvents
        txt����.SetFocus
    Else
        DoEvents
        cmdˢ��.SetFocus
    End If
End Sub

Public Property Get GetFilterCon() As Variant
    Call GetFilter
    Set GetFilterCon = mArrFilter
End Property

Public Property Get CheckDept() As Boolean
    CheckDept = cbo���ϲ���.ListCount <> 0
End Property




