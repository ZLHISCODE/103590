VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPriceGradeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�۸�ȼ�"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   Icon            =   "frmPriceGradeEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picApplyBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Index           =   1
      Left            =   2190
      ScaleHeight     =   3015
      ScaleWidth      =   1500
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1500
      Begin MSComctlLib.ListView lvwApply 
         Height          =   2055
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.PictureBox picApplyBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Index           =   0
      Left            =   1350
      ScaleHeight     =   3015
      ScaleWidth      =   1380
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1290
      Width           =   1380
      Begin MSComctlLib.ListView lvwApply 
         Height          =   2025
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   3572
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin XtremeSuiteControls.TabControl tbPageGradeApply 
      Height          =   3045
      Left            =   60
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   990
      Width           =   6645
      _Version        =   589884
      _ExtentX        =   11721
      _ExtentY        =   5371
      _StockProps     =   64
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   270
      TabIndex        =   14
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5250
      TabIndex        =   13
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   12
      Top             =   4320
      Width           =   1100
   End
   Begin VB.Frame frmPriceGradeBaseInfo 
      Caption         =   "������Ϣ"
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6645
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   2
         Left            =   4950
         TabIndex        =   6
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   1
         Left            =   2250
         TabIndex        =   4
         Top             =   360
         Width           =   1995
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   0
         Left            =   690
         TabIndex        =   2
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   4560
         TabIndex        =   5
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   1860
         TabIndex        =   3
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   420
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmPriceGradeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private Enum TxtIndex
    Txt_���� = 0
    Txt_���� = 1
    Txt_���� = 2
End Enum
Private Enum TabPageIndex
    Pg_NodeList = 0 'Ժ��
    Pg_PatientType = 1 'ҽ�Ƹ��ʽ
End Enum
Private Enum FunType
    Fun_Add = 0 '����
    Fun_Update = 1 '����
    Fun_Delete = 2 'ɾ��
    Fun_View = 3 '�鿴
End Enum
Private mbytFun As FunType '0-����,1-����,2-ɾ��,3-�鿴
Private mstr�۸�ȼ� As String

Private mblnChanged As Boolean
Private mblnFirst As Boolean
Private mblnLoading As Boolean

Public Function ShowMe(frmParent As Form, ByVal bytFun As Byte, _
    Optional ByVal strIn�۸�ȼ� As String, _
    Optional ByRef strOut�۸�ȼ� As String) As Boolean
    '�������
    '��Σ�
    '   frmParent ���ô��ڶ���
    '   bytFun �������ͣ�0-����,1-����,2-ɾ��,3-�鿴
    '   strIn�۸�ȼ� �鿴��������ɾ��ʱ����۸�ȼ�����
    '���Σ�
    '   strOut�۸�ȼ� ����ʱ���ؼ۸�ȼ����ƣ����ڵ����߶�λ
    mbytFun = bytFun
    mstr�۸�ȼ� = IIF(mbytFun = Fun_Add, "-", strIn�۸�ȼ�)
    
    On Error Resume Next
    mblnOk = False
    If CheckDepend() = False Then Exit Function
    Me.Show 1, frmParent
    ShowMe = mblnOk
    strOut�۸�ȼ� = IIF(mbytFun = Fun_Add, mstr�۸�ȼ�, "")
End Function

Private Function CheckDepend() As Boolean
    '����:���ݼ���ǰ���
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    If Not (mbytFun = Fun_Update Or mbytFun = Fun_Delete) Then CheckDepend = True: Exit Function
    
    '�Ѿ�ͣ�õģ����������/ɾ��
    strSQL = "Select 1 From �շѼ۸�ȼ� Where ���� = [1] And Nvl(����ʱ��, To_Date('3000-01-01','yyyy-mm-dd')) < To_Date('3000-01-01','yyyy-mm-dd')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���۸�ȼ��Ƿ�ͣ��", mstr�۸�ȼ�)
    If Not rsTemp.EOF Then
        MsgBox "��ǰ�۸�ȼ���ͣ�ã�������" & IIF(mbytFun = Fun_Update, "����", "ɾ��") & "��", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If mbytFun = Fun_Delete Then
        '�����ǰ�۸�ȼ��Ѿ����ۣ����Ѿ����ڵ��ۼ�¼����������ɾ����
        strSQL = "Select 1 From �շѼ�Ŀ Where �۸�ȼ� = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���۸�ȼ����շѼ�Ŀ���Ƿ�ʹ��", mstr�۸�ȼ�)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ�۸�ȼ������շѼ�Ŀ��ʹ�ã�������ɾ����", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    CheckDepend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Err = 0: On Error GoTo ErrHandler
    If mblnChanged Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    Err = 0: On Error GoTo ErrHandler
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then Unload Me: Exit Sub
    
    cmdOK.Enabled = False
    If IsValied() = False Then cmdOK.Enabled = True: Exit Sub
    If SaveData() = False Then cmdOK.Enabled = True: Exit Sub
    
    If mbytFun = Fun_Add Then
        mstr�۸�ȼ� = Trim(txtEdit(Txt_����).Text)
    End If
    mblnOk = True
    Unload Me
    Exit Sub
ErrHandler:
    cmdOK.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsValied() As Boolean
    '����:���ݼ��
    '����:���ͨ������True,���򷵻�False
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer, k As Integer
    Dim strTemp As String
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then IsValied = True: Exit Function
    If CheckDepend() = False Then Exit Function
    If mbytFun = Fun_Delete Then
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & mstr�۸�ȼ� & "���ļ۸�ȼ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        IsValied = True: Exit Function
    End If

    If zlControl.TxtCheckInput(txtEdit(Txt_����), "����", , False) = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(Txt_����), "����", , False) = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(Txt_����), "����") = False Then Exit Function
    
    If mbytFun = Fun_Update And mstr�۸�ȼ� <> Trim(txtEdit(Txt_����).Text) Then
        '�����ǰ�۸�ȼ��Ѿ����ۣ����Ѿ����ڵ��ۼ�¼����������������ơ�
        strSQL = "Select 1 From �շѼ�Ŀ Where �۸�ȼ� = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���۸�ȼ����շѼ�Ŀ���Ƿ�ʹ��", mstr�۸�ȼ�)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ�۸�ȼ������շѼ�Ŀ��ʹ�ã�������������ƣ�", vbInformation + vbOKOnly, gstrSysName
            If txtEdit(Txt_����).Visible And txtEdit(Txt_����).Enabled Then txtEdit(Txt_����).SetFocus
            Exit Function
        End If
    End If
    
    '����Ψһ
    strSQL = "Select 1 From �շѼ۸�ȼ� Where ���� = [1] And ���� <> [2] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ψһ���", Trim(txtEdit(Txt_����).Text), mstr�۸�ȼ�)
    If Not rsTemp.EOF Then
        MsgBox "����Ϊ��" & Trim(txtEdit(Txt_����).Text) & "���ļ۸�ȼ��Ѵ��ڣ�", vbInformation + vbOKOnly, gstrSysName
        If txtEdit(Txt_����).Visible And txtEdit(Txt_����).Enabled Then txtEdit(Txt_����).SetFocus
        Exit Function
    End If
    
    '����Ψһ
    strSQL = "Select 1 From �շѼ۸�ȼ� Where ���� = [1] And ���� <> [2] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ψһ���", Trim(txtEdit(Txt_����).Text), mstr�۸�ȼ�)
    If Not rsTemp.EOF Then
        MsgBox "����Ϊ��" & Trim(txtEdit(Txt_����).Text) & "���ļ۸�ȼ��Ѵ��ڣ�", vbInformation + vbOKOnly, gstrSysName
        If txtEdit(Txt_����).Visible And txtEdit(Txt_����).Enabled Then txtEdit(Txt_����).SetFocus
        Exit Function
    End If
    
    'һ��վ�㣬�������ö����Ч�ĵȼ�
    strTemp = ""
    For i = 1 To lvwApply(Pg_NodeList).ListItems.Count
        If lvwApply(Pg_NodeList).ListItems(i).Checked Then
            strTemp = strTemp & "|" & lvwApply(Pg_NodeList).ListItems(i).Tag
        End If
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    strSQL = "Select /*+cardinality(B,10)*/c.���� As վ��, a.�۸�ȼ�" & vbNewLine & _
            " From �շѼ۸�ȼ� D, �շѼ۸�ȼ�Ӧ�� A, Table(f_Str2list([1], '|')) B, Zlnodelist C" & vbNewLine & _
            " Where d.���� = a.�۸�ȼ� And a.վ�� = b.Column_Value And a.վ�� = c.��� And a.���� = 0" & vbNewLine & _
            "       And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01','yyyy-mm-dd')) And a.�۸�ȼ� <> [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�۸�ȼ�Ӧ�ü��", strTemp, mstr�۸�ȼ�)
    If Not rsTemp.EOF Then
        strTemp = ""
        Do While Not rsTemp.EOF
            strTemp = strTemp & vbCrLf & Nvl(rsTemp!վ��) & "��" & Nvl(rsTemp!�۸�ȼ�)
            rsTemp.MoveNext
        Loop
        If MsgBox("����һ��Ժ��ֻ������һ����Ч�ļ۸�ȼ������㵱ǰѡ���" & _
            "����Ժ��������������Ч�ļ۸�ȼ�������������������������ЩԺ����������Ч�۸�ȼ���" & _
            "Ȼ��Ӧ�õ�ǰ�۸�ȼ����Ƿ������" & vbCrLf & strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If tbPageGradeApply.Item(Pg_NodeList).Selected = False Then tbPageGradeApply.Item(Pg_NodeList).Selected = True
            If lvwApply(Pg_NodeList).Visible And lvwApply(Pg_NodeList).Enabled Then lvwApply(Pg_NodeList).SetFocus
            Exit Function
        End If
    End If
    
    'һ��ҽ�Ƹ��ʽ���������ö����Ч�ĵȼ�
    strTemp = ""
    For i = 1 To lvwApply(Pg_PatientType).ListItems.Count
        If lvwApply(Pg_PatientType).ListItems(i).Checked = True Then
            strTemp = strTemp & "|" & lvwApply(Pg_PatientType).ListItems(i).Text
        End If
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    strSQL = "Select /*+cardinality(B,10)*/a.ҽ�Ƹ��ʽ, a.�۸�ȼ�" & vbNewLine & _
            " From �շѼ۸�ȼ� D, �շѼ۸�ȼ�Ӧ�� A, Table(f_Str2list([1], '|')) B" & vbNewLine & _
            " Where d.���� = a.�۸�ȼ� And a.ҽ�Ƹ��ʽ = b.Column_Value And a.���� = 1" & vbNewLine & _
            "       And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01','yyyy-mm-dd')) And a.�۸�ȼ� <> [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�۸�ȼ�Ӧ�ü��", strTemp, mstr�۸�ȼ�)
    If Not rsTemp.EOF Then
        strTemp = ""
        Do While Not rsTemp.EOF
            strTemp = strTemp & vbCrLf & Nvl(rsTemp!ҽ�Ƹ��ʽ) & "��" & Nvl(rsTemp!�۸�ȼ�)
            rsTemp.MoveNext
        Loop
        If MsgBox("����һ��ҽ�Ƹ��ʽֻ������һ����Ч�ļ۸�ȼ������㵱ǰѡ���" & _
            "����ҽ�Ƹ��ʽ������������Ч�ļ۸�ȼ�������������������������Щҽ�Ƹ��ʽ��������Ч�۸�ȼ���" & _
            "Ȼ��Ӧ�õ�ǰ�۸�ȼ����Ƿ������" & vbCrLf & strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If tbPageGradeApply.Item(Pg_PatientType).Selected = False Then tbPageGradeApply.Item(Pg_PatientType).Selected = True
            If lvwApply(Pg_PatientType).Visible And lvwApply(Pg_PatientType).Enabled Then lvwApply(Pg_PatientType).SetFocus
            Exit Function
        End If
    End If
    
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '����:��������
    '����:����ɹ�����True,���򷵻�False
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strӦ��վ�� As String, strӦ��ҽ�Ƹ��ʽ As String
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then SaveData = True: Exit Function
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        For i = 1 To lvwApply(Pg_NodeList).ListItems.Count
            If lvwApply(Pg_NodeList).ListItems(i).Checked = True Then
                strӦ��վ�� = strӦ��վ�� & "|" & lvwApply(Pg_NodeList).ListItems(i).Tag
            End If
        Next
        If strӦ��վ�� <> "" Then strӦ��վ�� = Mid(strӦ��վ��, 2)
        
        For i = 1 To lvwApply(Pg_PatientType).ListItems.Count
            If lvwApply(Pg_PatientType).ListItems(i).Checked = True Then
                strӦ��ҽ�Ƹ��ʽ = strӦ��ҽ�Ƹ��ʽ & "|" & lvwApply(Pg_PatientType).ListItems(i).Text
            End If
        Next
        If strӦ��ҽ�Ƹ��ʽ <> "" Then strӦ��ҽ�Ƹ��ʽ = Mid(strӦ��ҽ�Ƹ��ʽ, 2)
    End If
    
    Select Case mbytFun
    Case Fun_Add
        'Zl_�շѼ۸�ȼ�_Insert(
        strSQL = "Zl_�շѼ۸�ȼ�_Insert("
        '  ����_In             In �շѼ۸�ȼ�.����%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_����).Text) & "',"
        '  ����_In             In �շѼ۸�ȼ�.����%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_����).Text) & "',"
        '  ����_In             In �շѼ۸�ȼ�.����%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_����).Text) & "',"
        '  �Ƿ�����ҩƷ_In     In �շѼ۸�ȼ�.�Ƿ�����ҩƷ%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '  �Ƿ���������_In     In �շѼ۸�ȼ�.�Ƿ���������%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '  �Ƿ�������ͨ��Ŀ_In In �շѼ۸�ȼ�.�Ƿ�������ͨ��Ŀ%Type := 1,
        strSQL = strSQL & "" & 1 & ","
        '  Ӧ��վ��_In         In Varchar2, --Ӧ���ڵ�վ���ţ�����õ�����"|"�ָ����磺01|02|...
        strSQL = strSQL & "'" & strӦ��վ�� & "',"
        '  Ӧ��ҽ�Ƹ��ʽ_In In Varchar2 --Ӧ���ڵ�ҽ�Ƹ��ʽ������õ�����"|"�ָ����磺����ҽ��|�Է�ҽ��|...
        strSQL = strSQL & "'" & strӦ��ҽ�Ƹ��ʽ & "')"
    Case Fun_Update
        'Zl_�շѼ۸�ȼ�_Update(
        strSQL = "Zl_�շѼ۸�ȼ�_Update("
        '  ԭ����_In           In �շѼ۸�ȼ�.����%Type,
        strSQL = strSQL & "'" & mstr�۸�ȼ� & "',"
        '  ����_In             In �շѼ۸�ȼ�.����%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_����).Text) & "',"
        '  ����_In             In �շѼ۸�ȼ�.����%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_����).Text) & "',"
        '  ����_In             In �շѼ۸�ȼ�.����%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_����).Text) & "',"
        '  �Ƿ�����ҩƷ_In     In �շѼ۸�ȼ�.�Ƿ�����ҩƷ%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '  �Ƿ���������_In     In �շѼ۸�ȼ�.�Ƿ���������%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '  �Ƿ�������ͨ��Ŀ_In In �շѼ۸�ȼ�.�Ƿ�������ͨ��Ŀ%Type := 1,
        strSQL = strSQL & "" & 1 & ","
        '  Ӧ��վ��_In         In Varchar2, --Ӧ���ڵ�վ���ţ�����õ�����"|"�ָ����磺01|02|...
        strSQL = strSQL & "'" & strӦ��վ�� & "',"
        '  Ӧ��ҽ�Ƹ��ʽ_In In Varchar2 --Ӧ���ڵ�ҽ�Ƹ��ʽ������õ�����"|"�ָ����磺����ҽ��|�Է�ҽ��|...
        strSQL = strSQL & "'" & strӦ��ҽ�Ƹ��ʽ & "')"
    Case Fun_Delete
        'Zl_�շѼ۸�ȼ�_Delete(
        strSQL = "Zl_�շѼ۸�ȼ�_Delete("
        '  ����_In In �շѼ۸�ȼ�.����%Type
        strSQL = strSQL & "'" & mstr�۸�ȼ� & "')"
    End Select
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    Err = 0: On Error GoTo ErrHandler
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If Trim(txtEdit(Txt_����).Text) <> "" Then
        If txtEdit(Txt_����).Visible And txtEdit(Txt_����).Enabled Then txtEdit(Txt_����).SetFocus
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo ErrHandler
    
    mblnFirst = True: mblnLoading = True
    If InitPage() = False Then Unload Me: Exit Sub
    If GetDefineSize() = False Then Unload Me: Exit Sub
    If InitData() = False Then Unload Me: Exit Sub
    If LoadData() = False Then Unload Me: Exit Sub
    
    If Not (mbytFun = Fun_Add Or mbytFun = Fun_Update) Then
        Call ZlSetEnabled(Me.Controls, False)
        Call ZlSetEnabledBackColor(Me.Controls)
    End If
    
    Me.Caption = Choose(mbytFun + 1, "����", "����", "ɾ��", "�鿴") & "�۸�ȼ�"
    If mbytFun = Fun_View Then
        cmdOK.Visible = False
        cmdCancel.Caption = cmdOK.Caption
    End If
    mblnLoading = False
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitPage() As Boolean
    '����:��ʼ��ҳǩ�ؼ�
    Err = 0: On Error GoTo ErrHandler
    With tbPageGradeApply
        .RemoveAll
        .InsertItem Pg_NodeList, "Ժ��", picApplyBack(Pg_NodeList).hwnd, 0
        .InsertItem Pg_PatientType, "ҽ�Ƹ��ʽ", picApplyBack(Pg_PatientType).hwnd, 0

         With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .Layout = xtpTabLayoutAutoSize
            .StaticFrame = True
            .ClientFrame = xtpTabFrameBorder
        End With
        .Item(Pg_NodeList).Selected = True
    End With
    InitPage = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData() As Boolean
    '��ʼ�������������
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim k As Integer, objListItem As ListItem
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        If mbytFun = Fun_Add Then
            txtEdit(Txt_����).Text = zlDatabase.GetMax("�շѼ۸�ȼ�", "����", 2)
        End If
        
        lvwApply(Pg_NodeList).ListItems.Clear
        lvwApply(Pg_PatientType).ListItems.Clear
        strSQL = "Select 1 As ����, ��� As ����, ���� From Zlnodelist" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select 2 As ����, ����, ���� From ҽ�Ƹ��ʽ" & vbNewLine & _
                " Order By ����, ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������")
        If rsTemp.RecordCount = 0 Then
            tbPageGradeApply.Item(Pg_PatientType).Visible = False
        Else
            '1.Ժ��,2.ҽ�Ƹ��ʽ
            For k = 0 To 1
                rsTemp.Filter = "����=" & IIF(k = 0, 1, 2)
                If rsTemp.RecordCount = 0 Then
                    tbPageGradeApply.Item(k).Visible = False
                    If k = Pg_NodeList Then tbPageGradeApply.Item(Pg_PatientType).Selected = True
                Else
                    Do While Not rsTemp.EOF
                        Set objListItem = lvwApply(k).ListItems.Add(, "K" & Nvl(rsTemp!����), Nvl(rsTemp!����))
                        objListItem.Tag = Nvl(rsTemp!����)
                        rsTemp.MoveNext
                    Loop
                End If
            Next
        End If
    End If
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadData() As Boolean
    '��������
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer, k As Integer, blnFind As Boolean
    Dim objListItem As ListItem
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_Add Then LoadData = True: Exit Function
    
    strSQL = "Select ����, ����, ����, �Ƿ�����ҩƷ, �Ƿ���������, �Ƿ�������ͨ��Ŀ" & vbNewLine & _
            " From �շѼ۸�ȼ�" & vbNewLine & _
            " Where ���� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�շѼ۸�ȼ�", mstr�۸�ȼ�)
    If rsTemp.EOF Then
        MsgBox "�۸�ȼ� " & mstr�۸�ȼ� & " �����ڣ������ѱ�����ɾ������ˢ�º�鿴...", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    txtEdit(Txt_����).Text = Nvl(rsTemp!����)
    txtEdit(Txt_����).Text = Nvl(rsTemp!����)
    txtEdit(Txt_����).Text = Nvl(rsTemp!����)
    
    strSQL = "Select Nvl(a.����, 0) As ����, " & vbNewLine & _
            "        Decode(Nvl(a.����, 0), 0, b.���, c.����) As ����," & vbNewLine & _
            "        Decode(Nvl(a.����, 0), 0, b.����, c.����) As ����" & vbNewLine & _
            " From �շѼ۸�ȼ�Ӧ�� A, Zlnodelist B, ҽ�Ƹ��ʽ C" & vbNewLine & _
            " Where a.վ�� = b.���(+) And a.ҽ�Ƹ��ʽ = c.����(+) And a.�۸�ȼ� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�շѼ۸�ȼ�", mstr�۸�ȼ�)
    If rsTemp.EOF Then
        If mbytFun <> Fun_Update Then tbPageGradeApply.Item(Pg_PatientType).Visible = False
    Else
        For k = 0 To 1
            '0-Ժ��,1-ҽ�Ƹ��ʽ
            rsTemp.Filter = "����=" & IIF(k = 0, 0, 1)
            If rsTemp.RecordCount = 0 Then
                If mbytFun <> Fun_Update Then
                    tbPageGradeApply.Item(k).Visible = False
                    If k = Pg_NodeList Then tbPageGradeApply.Item(Pg_PatientType).Selected = True
                End If
            Else
                Do While Not rsTemp.EOF
                    blnFind = False
                    For i = 1 To lvwApply(k).ListItems.Count
                        Set objListItem = lvwApply(k).ListItems(i)
                        If objListItem.Tag = Nvl(rsTemp!����) Then
                            objListItem.Checked = True
                            blnFind = True: Exit For
                        End If
                    Next
                    If blnFind = False Then
                        Set objListItem = lvwApply(k).ListItems.Add(, "K" & Nvl(rsTemp!����), Nvl(rsTemp!����))
                        objListItem.Tag = Nvl(rsTemp!����)
                        objListItem.Checked = True
                    End If
                    rsTemp.MoveNext
                Loop
            End If
        Next
    End If
    LoadData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDefineSize() As Boolean
'���ܣ��õ����ݿ�ı��ֶεĳ���
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select ����, ����, ���� From �շѼ۸�ȼ� Where Rownum < 0"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "�շѼ۸�ȼ��༭")
    
    txtEdit(Txt_����).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(Txt_����).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(Txt_����).MaxLength = rsTemp.Fields("����").DefinedSize
    
    GetDefineSize = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lvwApply_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Err = 0: On Error GoTo ErrHandler
    If mblnLoading Then Exit Sub
    mblnChanged = True
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwApply_KeyPress(Index As Integer, KeyAscii As Integer)
    Err = 0: On Error GoTo ErrHandler
    If KeyAscii = vbKeyReturn Then
        If tbPageGradeApply.Selected.Index < tbPageGradeApply.ItemCount - 1 Then
            tbPageGradeApply(tbPageGradeApply.Selected.Index + 1).Selected = True
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picApplyBack_Resize(Index As Integer)
    On Error Resume Next
    With lvwApply(Index)
        .Left = 0
        .Top = 0
        .Width = picApplyBack(Index).ScaleWidth
        .Height = picApplyBack(Index).ScaleHeight
    End With
End Sub

Private Sub tbPageGradeApply_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    On Error Resume Next
    If lvwApply(Item.Index).Visible And lvwApply(Item.Index).Enabled Then lvwApply(Item.Index).SetFocus
End Sub

Private Sub txtEdit_Change(Index As Integer)
    Err = 0: On Error GoTo ErrHandler
    If mblnLoading Then Exit Sub
    mblnChanged = True
    If Index = Txt_���� Then
        txtEdit(Txt_����).Text = zlStr.GetCodeByVB(txtEdit(Txt_����).Text)
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Err = 0: On Error GoTo ErrHandler
    zlControl.TxtSelAll txtEdit(Index)
    If Index = Txt_���� Then zlCommFun.OpenIme True
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Err = 0: On Error GoTo ErrHandler
    If InStr("'}|,""/", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Err = 0: On Error GoTo ErrHandler
    zlCommFun.OpenIme False
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ZlSetEnabled(ByVal objControls As Object, ByVal blnEnabled As Boolean)
    '���ÿؼ�����״̬
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To objControls.Count - 1
        If UCase(objControls(i).Name) <> UCase("cmdHelp") _
            And UCase(objControls(i).Name) <> UCase("cmdOk") _
            And UCase(objControls(i).Name) <> UCase("cmdCancel") _
            And UCase(TypeName(objControls(i))) <> UCase("Label") _
            And UCase(TypeName(objControls(i))) <> UCase("Frame") _
            And UCase(TypeName(objControls(i))) <> UCase("TabControl") _
            And UCase(TypeName(objControls(i))) <> UCase("PictureBox") _
            And UCase(TypeName(objControls(i))) <> UCase("VSFlexGrid") Then
            objControls(i).Enabled = blnEnabled
        End If
    Next
End Sub

Private Sub ZlSetEnabledBackColor(ByVal objControls As Object)
    '���ÿؼ�����״̬�벻����״̬�ı�����ɫ
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To objControls.Count - 1
        If UCase(TypeName(objControls(i))) = UCase("TextBox") _
            Or UCase(TypeName(objControls(i))) = UCase("ComboBox") Then
            objControls(i).BackColor = IIF(objControls(i).Enabled, vbWindowBackground, vbButtonFace)
        End If
    Next
End Sub

