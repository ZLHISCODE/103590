VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTendPrintAsk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡѡ��"
   ClientHeight    =   3150
   ClientLeft      =   2550
   ClientTop       =   2625
   ClientWidth     =   4890
   HelpContextID   =   10322
   Icon            =   "frmTendPrintAsk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab SSTab1 
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4154
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "������ӡ"
      TabPicture(0)   =   "frmTendPrintAsk.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNote"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl�����ļ�"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cbo�����ļ�"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "���´�ӡ"
      TabPicture(1)   =   "frmTendPrintAsk.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPrint(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "�����ش�"
      TabPicture(2)   =   "frmTendPrintAsk.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPrint(2)"
      Tab(2).ControlCount=   1
      Begin VB.ComboBox cbo�����ļ� 
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Frame fraPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1785
         Index           =   2
         Left            =   -74940
         TabIndex        =   14
         Tag             =   "����ش�"
         Top             =   360
         Width           =   4635
         Begin VB.TextBox txtBegin 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   8
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "û����Ҫ�ش��ҳ"
            ForeColor       =   &H000000C0&
            Height          =   375
            Index           =   2
            Left            =   210
            TabIndex        =   17
            Top             =   870
            Width           =   4065
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "��ָ������ʼҳ��ʼ���������ش򣬵��Ѵ�ӡҳ�������޸ĺ������д�λ��ʹ�øù��ܡ�"
            ForeColor       =   &H00C00000&
            Height          =   600
            Index           =   1
            Left            =   180
            TabIndex        =   9
            Top             =   1320
            Width           =   4320
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼҳ"
            Height          =   180
            Left            =   1170
            TabIndex        =   7
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame fraPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   1
         Left            =   -74970
         TabIndex        =   13
         Tag             =   "���´�ӡ"
         Top             =   360
         Width           =   4635
         Begin VB.TextBox txtPage 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   2025
            MaxLength       =   3
            TabIndex        =   5
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ӡָ��ҳ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1050
            TabIndex        =   16
            Top             =   345
            Width           =   900
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "û����Ҫ�ش��ҳ"
            ForeColor       =   &H000000C0&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   870
            Width           =   4065
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "���´�ӡָ��ҳ�ŵĻ������ݡ����������޸ġ�ֽ�Ŷ�ʧ�����𣬻����ӡ�����ϵ��´�ӡ���ɹ��������ʹ�øù��ܡ�"
            ForeColor       =   &H00C00000&
            Height          =   780
            Index           =   0
            Left            =   210
            TabIndex        =   6
            Top             =   1320
            Width           =   4320
         End
      End
      Begin VB.Label lbl�����ļ� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ļ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ�����ļ�һֱδ��ӡ��"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   1110
         TabIndex        =   3
         Top             =   900
         Width           =   2970
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "���������ش�����ʱ������ʹ��������ӡ���ܡ�"
         ForeColor       =   &H00C00000&
         Height          =   360
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   4320
      End
   End
   Begin VB.CommandButton cmd���EXCEL 
      Caption         =   "�����&Excel"
      Height          =   350
      Left            =   180
      TabIndex        =   12
      Top             =   2610
      Width           =   1245
   End
   Begin VB.CommandButton cmdԤ�� 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   2370
      TabIndex        =   11
      Top             =   2610
      Width           =   1100
   End
   Begin VB.CommandButton cmd��ӡ 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   3570
      TabIndex        =   10
      Top             =   2610
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTendPrintAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public byRunMode As Byte        'ִ�з�ʽ
Public intPage As Integer       '�ȴ���ӡҳ�룬����ʱ����Ҫ����ֻ���ش���Ҫ��¼
Public intPageRows As Integer

Dim strSQL As String
Dim blnRePrint As Boolean       'ֻ���ش��޸Ĺ�������
Dim blnRePrintAll As Boolean    '�޸Ĺ������ݼ������ӡ�����ݶ���Ҫ�ش�ֻ��ʹ�������ش��ܺ���ʹ��
Dim strRePrint As String
Dim intBeginPage As Integer
Dim intEndPage As Integer
Dim rsFile As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset


Private Sub SetCommandState()
    '�����ش�������������ش�
    cmdԤ��.Enabled = (intPageRows > 0)
    cmd��ӡ.Enabled = (intPageRows > 0)
    cmd���EXCEL.Enabled = (intPageRows > 0)
    Select Case SSTab1.Tab
    Case ����   'ֻҪ�����ش�����������ʹ��������
        cmdԤ��.Enabled = Not blnRePrint And (intPageRows > 0)
        cmd��ӡ.Enabled = Not blnRePrint And (intPageRows > 0)
        cmd���EXCEL.Enabled = Not blnRePrint And (intPageRows > 0)
    Case �ش�   '����кű仯���������ش���Ҫʹ�������ش���
        cmdԤ��.Enabled = Not blnRePrintAll And (intPageRows > 0)
        cmd��ӡ.Enabled = Not blnRePrintAll And (intPageRows > 0)
        cmd���EXCEL.Enabled = Not blnRePrintAll And (intPageRows > 0)
    Case �����ش�   '�����ش����û��Լ�ʹ�ã�������
        cmdԤ��.Enabled = blnRePrintAll And (intPageRows > 0)
        cmd��ӡ.Enabled = blnRePrintAll And (intPageRows > 0)
        cmd���EXCEL.Enabled = blnRePrintAll And (intPageRows > 0)
    End Select
End Sub

Private Sub cbo�����ļ�_Click()
    Dim blnPrint As Boolean
    On Error GoTo errHand
    
    blnRePrint = False
    blnRePrintAll = False
    Call SetCommandState
    
    strSQL = " Select  d.�����ı�" & vbNewLine & _
             " From �����ļ��ṹ d, �����ļ��ṹ p,���˻����ļ� c" & vbNewLine & _
             " Where p.Id = d.��id And p.�ļ�id = c.��ʽID and C.ID=[1] And p.�������� = 1 And p.�����ı� = '�����ʽ' and d.Ҫ������='��Ч������'"
    Set rsTemp = OpenSQLRecord(strSQL, "��ȡ���������", cbo�����ļ�.ItemData(cbo�����ļ�.ListIndex))
    If rsTemp.RecordCount <> 0 Then
        intPageRows = NVL(rsTemp!�����ı�, 0)
    End If
    
    strSQL = " Select 1 From ���˻����ӡ Where �ļ�ID=[1] And ��ӡ�� is Not NULL And Rownum<2"
    Set rsTemp = OpenSQLRecord(strSQL, "��ȡ�����ӡ����", cbo�����ļ�.ItemData(cbo�����ļ�.ListIndex))
    blnPrint = rsTemp.RecordCount
    If blnPrint Then
        '�Ѿ���ӡ���ļ�����ʾ���ӡҳ��
        strSQL = " Select Min(��ӡҳ��) AS ��ʼҳ��,Max(��ӡҳ��) AS ����ҳ�� From ���˻����ӡ Where �ļ�ID=[1] "
        Set rsTemp = OpenSQLRecord(strSQL, "��ȡ��ӡҳ��Χ", cbo�����ļ�.ItemData(cbo�����ļ�.ListIndex))
        intBeginPage = rsTemp!��ʼҳ��
        intEndPage = rsTemp!����ҳ��
        If rsTemp!��ʼҳ�� <> rsTemp!����ҳ�� Then
            lblNote.Caption = "�Ѵ�ӡҳ��Χ��" & rsTemp!��ʼҳ�� & "-" & rsTemp!����ҳ��
        Else
            lblNote.Caption = "�Ѵ�ӡҳ��Χ��" & rsTemp!��ʼҳ��
        End If
    Else
        lblNote.Caption = "���ļ���δ��ӡ����"
    End If
    
    'ȫԺͳһ���,�û�����ô�����ô��,��������Լ��ش�
    '����Ƿ������Ҫ�ش�����ݣ������������������ȡδ��ӡ����С����ʱ������ݣ�Ȼ�����ʱ��֮���Ƿ���ڴ�ӡ�������ݣ��������˵��������Ҫ�ش�����ݣ�
    strSQL = " Select �в�,��ӡҳ�� From ���˻����ӡ Where �ļ�ID=[1] And ��ӡ�� Is NULL And ��ӡҳ�� Is Not NULL Order by ��ӡҳ��"
    Set rsTemp = OpenSQLRecord(strSQL, "����Ƿ�����ش�����", cbo�����ļ�.ItemData(cbo�����ļ�.ListIndex))
    strRePrint = ""
    If rsTemp.RecordCount <> 0 Then
        blnRePrint = True
        lblTip(�ش�).Caption = ""
        Do While Not rsTemp.EOF
            If InStr(1, "," & strRePrint & ",", "," & rsTemp!��ӡҳ�� & ",") = 0 Then
                strRePrint = strRePrint & "," & rsTemp!��ӡҳ��
                lblTip(�ش�).Caption = lblTip(�ش�).Caption & "," & rsTemp!��ӡҳ��
            End If
            rsTemp.MoveNext
        Loop
        strRePrint = Mid(strRePrint, 2)
        lblTip(�ش�).Caption = "����ҳ����Ҫ�ش�" & Mid(lblTip(�ش�).Caption, 2)
        
        rsTemp.Filter = "�в�<>0"
        blnRePrintAll = (rsTemp.RecordCount <> 0)
        If blnRePrintAll Then
            txtBegin.Text = rsTemp!��ӡҳ��
            lblTip(�����ش�).Caption = "���ڵ�" & rsTemp!��ӡҳ�� & "ҳ�޸ĺ���������������˱仯����" & rsTemp!��ӡҳ�� & "ҳ��ʼ��ҳ��ȫ����Ҫ�ش�"
        End If
        rsTemp.Filter = 0
    End If
    
    If Not blnRePrint Then lblTip(�ش�).Caption = "û����Ҫ�ش��ҳ"
    If Not blnRePrintAll Then lblTip(�����ش�).Caption = "û����Ҫ�����ش��ҳ"
    
    Call SetCommandState
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function PrePrint() As Boolean
    On Error GoTo errHand
    gintPrintState = SSTab1.Tab + 1
    If SSTab1.Tab = �����ش� Then
        '��ɴ�ӡ���ݵ������ش���
        If txtBegin.Text = "" Then
            MsgBox "������������ش��ҳ�룡", vbInformation, gstrSysName
            txtBegin.SetFocus
            Exit Function
        End If
        If Not IsNumeric(txtBegin.Text) Then
            MsgBox "�����ҳ�뺬�зǷ��ַ���", vbInformation, gstrSysName
            txtBegin.SetFocus
            Exit Function
        End If
        intPage = txtBegin.Text
    ElseIf SSTab1.Tab = �ش� Then
        If txtPage(0).Text = "" Then
            MsgBox "�������ش��ҳ�룡", vbInformation, gstrSysName
            txtPage(0).SetFocus
            Exit Function
        End If
        If Not IsNumeric(txtPage(0).Text) Then
            MsgBox "�����ҳ�뺬�зǷ��ַ���", vbInformation, gstrSysName
            txtPage(0).SetFocus
            Exit Function
        End If
        If txtPage(0).Text < intBeginPage Then
            MsgBox "�ش�ҳ�벻��С�ڿ�ʼҳ�룡", vbInformation, gstrSysName
            txtPage(0).SetFocus
            Exit Function
        End If
        If txtPage(0).Text > intEndPage Then
            MsgBox "�ش�ҳ�벻�ܴ��ڽ���ҳ�룡", vbInformation, gstrSysName
            txtPage(0).SetFocus
            Exit Function
        End If
        
        intPage = txtPage(0).Text
    Else
        intPage = 0
    End If
    
    PrePrint = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmd��ӡ_Click()
    If Not PrePrint Then Exit Sub
    byRunMode = 1
    Me.Hide
End Sub

Private Sub cmd���EXCEL_Click()
    If Not PrePrint Then Exit Sub
    byRunMode = 3
    Me.Hide
End Sub

Private Sub cmdԤ��_Click()
    If Not PrePrint Then Exit Sub
    byRunMode = 2
    Me.Hide
End Sub

Private Sub Form_Load()
    '��ȡ���л����ļ�
    strSQL = " Select /*+RULE */ A.�ļ����� " & vbNewLine & _
             " From ���˻����ļ� A" & vbNewLine & _
             " Where A.ID=[1]"
    Set rsFile = OpenSQLRecord(strSQL, "��ȡ���л����ļ�", glng�ļ�ID)
    Me.cbo�����ļ�.AddItem rsFile!�ļ�����
    Me.cbo�����ļ�.ItemData(Me.cbo�����ļ�.NewIndex) = glng�ļ�ID
    cbo�����ļ�.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    byRunMode = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call cbo�����ļ�_Click
End Sub
