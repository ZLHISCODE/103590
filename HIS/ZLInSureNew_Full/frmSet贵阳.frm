VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ղ�������"
   ClientHeight    =   5640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8850
   Icon            =   "frmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo������� 
      ForeColor       =   &H80000012&
      Height          =   300
      Left            =   6630
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1050
      Width           =   2175
   End
   Begin VB.Frame fra 
      Caption         =   "���ղ���"
      Height          =   3225
      Left            =   30
      TabIndex        =   8
      Top             =   1740
      Width           =   4365
      Begin VB.ComboBox cbo�������˿ں� 
         Height          =   300
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2610
         Width           =   1995
      End
      Begin VB.TextBox txtServer 
         Height          =   300
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   16
         Top             =   1545
         Width           =   1995
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "֧����������(&T)"
         Height          =   255
         Left            =   990
         TabIndex        =   10
         Top             =   450
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Left            =   3930
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1575
         Width           =   255
      End
      Begin VB.CheckBox chk�α�ǰ��Ժ 
         Caption         =   "��Ժʱѡ��α�ǰ��Ժ(&T)"
         Height          =   255
         Left            =   990
         TabIndex        =   11
         Top             =   690
         Width           =   2385
      End
      Begin VB.CheckBox chk�շ� 
         Caption         =   "������������շ�(&L)"
         Height          =   255
         Left            =   990
         TabIndex        =   9
         Top             =   180
         Width           =   2055
      End
      Begin VB.TextBox txt��Ŀ�� 
         Height          =   300
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   14
         Top             =   1200
         Width           =   1995
      End
      Begin VB.ComboBox cbo��Ժ���� 
         Height          =   300
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1890
         Width           =   1995
      End
      Begin VB.CheckBox chk����ҩƷ��ʾ 
         Caption         =   "����ҩƷ��ʾ(&M)"
         Height          =   255
         Left            =   990
         TabIndex        =   12
         Top             =   930
         Width           =   2385
      End
      Begin VB.TextBox txtʱ�� 
         Height          =   300
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   20
         Text            =   "20"
         Top             =   2250
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�������˿ں�(D)"
         Height          =   180
         Left            =   810
         TabIndex        =   21
         Top             =   2670
         Width           =   1350
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "frmSet����.frx":000C
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ҽ��������(&S)"
         Height          =   180
         Left            =   990
         TabIndex        =   15
         Top             =   1620
         Width           =   1170
      End
      Begin VB.Label lbl��Ŀ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���������Ŀ��(&N)"
         Height          =   180
         Left            =   630
         TabIndex        =   13
         Top             =   1275
         Width           =   1530
      End
      Begin VB.Label lbl��Ժ���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����(&R)"
         Height          =   180
         Left            =   1170
         TabIndex        =   17
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������㴰�ڸ�           ���Զ��ر�"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   19
         Top             =   2295
         Width           =   3150
      End
   End
   Begin VB.CheckBox chk������Ҫ����ҽ������ 
      Caption         =   "������Ҫ����ҽ������"
      Height          =   315
      Left            =   60
      TabIndex        =   33
      Top             =   5220
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -150
      TabIndex        =   32
      Top             =   5040
      Width           =   10395
   End
   Begin VB.Frame Frame3 
      Caption         =   "�������Լ�����ҩ����"
      Height          =   915
      Left            =   4440
      TabIndex        =   23
      Top             =   90
      Width           =   4365
      Begin VB.CheckBox chk������ҩ���� 
         Caption         =   "�����������Լ�����ҩ���ƹ���"
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   2835
      End
      Begin VB.ComboBox cboʱ�� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   570
         Width           =   795
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�ۼ���ҩ����         ���п���"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   2610
      End
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1605
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   4365
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   1635
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   390
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Top             =   1170
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   30
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7470
      TabIndex        =   31
      Top             =   5190
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3225
      Left            =   4440
      TabIndex        =   29
      Top             =   1740
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   5689
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "סԺ����ʱδ�Ǽǲ�����¼"
      Height          =   180
      Left            =   4440
      TabIndex        =   27
      Top             =   1110
      Width           =   2160
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ֻ�������߲��������շ�����뷢Ʊ������Ŀ����Ķ�Ӧ��ϵ
Dim mlng���� As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Enum enum�ı�
    textҽ���û� = 0
    Textҽ������ = 1
    Textҽ�������� = 2
End Enum

Dim mcnTest As New ADODB.Connection

Private Sub Bill_cboClick(ListIndex As Long)
    If Bill.Active = False Then Exit Sub
    Bill.TextMatrix(Bill.Row, 2) = Bill.CboText
End Sub

Private Sub chk������ҩ����_Click()
    cboʱ��.Enabled = (chk������ҩ����.Value = 1): lblʱ��(1).Enabled = cboʱ��.Enabled
End Sub
Private Sub cmdSelect_Click()
    Dim strServer As String
    
    strServer = GetComputer(Me, "ѡ��ҽ��������")
    If strServer <> "" Then
        txtServer.Text = strServer
        mblnChange = True
    End If
End Sub

Private Sub cmdTest_Click()
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, TxtEdit(Textҽ��������).Text, TxtEdit(textҽ���û�).Text, TxtEdit(Textҽ������).Text) = False Then
        Exit Sub
    End If
    
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Me.ActiveControl.Name <> "Bill" Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.cbo��Ժ����
        .Clear
        .AddItem "HIS��ҽ��ͬʱ��Ժ"
        .AddItem "������HIS��Ժ"
        .ListIndex = 0
    End With
    
    With Me.cbo�������˿ں�
        .Clear
        .AddItem "COM1"
        .AddItem "COM2"
        .AddItem "COM3"
        .AddItem "COM4"
        .AddItem "COM5"
        .AddItem "COM6"
        .AddItem "COM7"
        .AddItem "COM8"
        .AddItem "COM9"
        .AddItem "USB"
        .ListIndex = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim colPara As New Collection
    Dim lngCount As Long
    
    If mcnTest.State = adStateOpen Then mcnTest.Close
    If OraDataOpen(mcnTest, TxtEdit(Textҽ��������).Text, TxtEdit(textҽ���û�).Text, TxtEdit(Textҽ������).Text) = False Then
        Exit Sub
    End If
    
    If txtServer.Text = "" Then
        MsgBox "ҽ��������������Ϊ�ա�", vbInformation, gstrSysName
        txtServer.SetFocus
        Exit Sub
    End If
    If IsNumeric(txt��Ŀ��.Text) = False Then
        MsgBox "��������ȷ����Ŀ����", vbInformation, gstrSysName
        txt��Ŀ��.SetFocus
        Exit Sub
    End If
    If zlCommFun.StrIsValid(txtServer.Text, txtServer.MaxLength) = False Then
        txtServer.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    '��һ���ֲ�������������
    colPara.Add "null,'ҽ���û���','" & TxtEdit(textҽ���û�).Text
    colPara.Add "null,'ҽ���û�����','" & TxtEdit(Textҽ������).Text
    colPara.Add "null,'ҽ��������1','" & TxtEdit(Textҽ��������).Text
    
    colPara.Add "null,'���������շ�','" & chk�շ�.Value
    colPara.Add "null,'֧����������','" & chk����.Value
    colPara.Add "null,'��Ժʱѡ��α�ǰ��Ժ','" & chk�α�ǰ��Ժ.Value
    colPara.Add "null,'ҽ��������','" & txtServer.Text
    colPara.Add "null,'���������Ŀ��','" & txt��Ŀ��.Text
    colPara.Add "null,'����ҩƷ��ʾ','" & chk����ҩƷ��ʾ.Value
    colPara.Add "null,'������㴰�ڹر�ʱ��','" & txtʱ��.Text
    colPara.Add "null,'�����������Լ�����ҩ���ƹ���','" & chk������ҩ����.Value
    colPara.Add "null,'�ۼ���ҩ�������׼','" & Split(cboʱ��.Text, "-")(0)
    colPara.Add "null,'�������','" & cbo�������.ListIndex
    
    For lngCount = 1 To colPara.Count
        gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & colPara(lngCount) & "'," & lngCount & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    '������Ŀ����ı���
    For lngCount = 1 To Bill.Rows - 1
        gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'" & Bill.TextMatrix(lngCount, 0) & "','" & Mid(Bill.TextMatrix(lngCount, 2), 1, 2) & "'," & lngCount + 5 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'��Ժ����','" & Me.cbo��Ժ����.ListIndex & "',90)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    SaveSetting "ZLSOFT", "����ȫ��", "����ҽ������", chk������Ҫ����ҽ������.Value
    gcnOracle.CommitTrans
    
    SaveSetting "ZLSOFT", "����ģ��\������ҽ��", "�˿�", cbo�������˿ں�.Text
    
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub txtServer_Change()
    mblnChange = True
End Sub

Private Sub txtServer_GotFocus()
    zlControl.TxtSelAll txtServer
End Sub

Public Function ��������(ByVal lng���� As Long) As Boolean
'���ܣ�������������ҽ������Ҫ�Ĳ���
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    
    mblnOK = False
    mlng���� = lng����
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1] and ���� is null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����)
    cboʱ��.Clear
    cboʱ��.AddItem "1-��"
    cboʱ��.AddItem "2-��"
    cboʱ��.AddItem "3-��"
    cboʱ��.AddItem "4-��"
    cboʱ��.ListIndex = 1
    
    cbo�������.Clear
    cbo�������.AddItem "0.������������"
    cbo�������.AddItem "1.��ʾ�Ƿ����"
    cbo�������.AddItem "2.��ֹ����"
    cbo�������.ListIndex = 0
    
    cbo�������˿ں�.Text = GetSetting("ZLSOFT", "����ģ��\������ҽ��", "�˿�", "COM1")
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                Me.TxtEdit(textҽ���û�).Text = Nvl(rsTemp!����ֵ)
            Case "ҽ���û�����"
                Me.TxtEdit(Textҽ������).Text = Nvl(rsTemp!����ֵ)
            Case "ҽ��������1"
                Me.TxtEdit(Textҽ��������).Text = Nvl(rsTemp!����ֵ)
            Case "���������շ�"
                chk�շ�.Value = IIf(rsTemp("����ֵ") = 1, 1, 0)
            Case "֧����������"
                chk����.Value = IIf(rsTemp("����ֵ") = 1, 1, 0)
            Case "��Ժʱѡ��α�ǰ��Ժ"
                chk�α�ǰ��Ժ.Value = IIf(rsTemp("����ֵ") = 1, 1, 0)
            Case "ҽ��������"
                txtServer.Text = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "���������Ŀ��"
                txt��Ŀ��.Text = IIf(IsNull(rsTemp("����ֵ")), "7", rsTemp("����ֵ"))
            Case "��Ժ����"
                Me.cbo��Ժ����.ListIndex = Nvl(rsTemp!����ֵ, 0)
            Case "����ҩƷ��ʾ"
                Me.chk����ҩƷ��ʾ.Value = Nvl(rsTemp!����ֵ, 0)
            Case "������㴰�ڹر�ʱ��"
                txtʱ��.Text = Nvl(rsTemp!����ֵ, 20)
            Case "�����������Լ�����ҩ���ƹ���"
                chk������ҩ����.Value = Nvl(rsTemp!����ֵ, 0)
            Case "�ۼ���ҩ�������׼"
                cboʱ��.ListIndex = Val(Nvl(rsTemp!����ֵ, 2)) - 1
            Case "�������"
                cbo�������.ListIndex = Nvl(rsTemp!����ֵ, 0)
        End Select
        
        rsTemp.MoveNext
    Loop
    
    '��ȡ�������úõĹ�����Ŀ�����Ӧ��ϵ�������ж������޸�
    '��Ʊ������Ŀ����
    '01����ҩ��02���г�ҩ��03���в�ҩ��04����λ�ѣ�05�����ѣ�06�����ѣ�
    '07�����Ʒѣ�08������ѣ�09�������ѣ�10������ѣ�11������
    gstrSQL = "Select ����,���,'11-����' ������Ŀ����  " & _
             " From �շ���� " & _
             " Where ���� Not IN  " & _
             "     (Select ������ From ���ղ��� Where ����=" & lng���� & " And ���>=6) " & _
             " union   " & _
             " Select B.����,B.���,decode(A.����ֵ,'01','01-��ҩ','02','02-�г�ҩ', " & _
             " '03','03-�в�ҩ','04','04-��λ��','05','05-����','06','06-����','07','07-���Ʒ�', " & _
             " '08','08-�����','09','09-������','10','10-�����','11-����') ������Ŀ����   " & _
             " From ���ղ��� A,�շ���� B " & _
             " Where A.���>=6 And A.����=[1] And A.������=B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ����", lng����)
    '��ʼ�����ݿؼ�
    With Bill
        .Rows = 2
        .Cols = 3
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "�շ����"
        .TextMatrix(0, 2) = "������Ŀ����"
        .ColWidth(0) = 500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1800
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColData(0) = 0
        .ColData(2) = 3

        .AddItem "01-��ҩ"
        .AddItem "02-�г�ҩ"
        .AddItem "03-�в�ҩ"
        .AddItem "04-��λ��"
        .AddItem "05-����"
        .AddItem "06-����"
        .AddItem "07-���Ʒ�"
        .AddItem "08-�����"
        .AddItem "09-������"
        .AddItem "10-�����"
        .AddItem "11-����"
        .ListIndex = 10
        
        .PrimaryCol = 0
        .LocateCol = 2
    End With
    
    With rsTemp
        Do While Not .EOF
            Bill.TextMatrix(.AbsolutePosition, 0) = !����
            Bill.TextMatrix(.AbsolutePosition, 1) = !���
            Bill.TextMatrix(.AbsolutePosition, 2) = !������Ŀ����
            .MoveNext
            Bill.Rows = Bill.Rows + 1
        Loop
        If .RecordCount <> 0 Then Bill.Rows = Bill.Rows - 1
        Bill.Row = 1
    End With
    
    Bill.AllowAddRow = False
    Bill.Active = OwnerUser(gstrDbUser)
    chk������Ҫ����ҽ������.Value = GetSetting(appName:="ZLSOFT", Section:="����ȫ��", Key:="����ҽ������", Default:="1")
    cboʱ��.Enabled = (chk������ҩ����.Value = 1): lblʱ��(1).Enabled = cboʱ��.Enabled
    mblnChange = False
    frmSet����.Show vbModal, frmҽ�����
    �������� = mblnOK
End Function
Private Sub txtʱ��_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtʱ��, KeyAscii, m����ʽ)
End Sub
