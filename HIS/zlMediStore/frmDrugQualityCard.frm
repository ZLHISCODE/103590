VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugQualityCard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҩƷ��������"
   ClientHeight    =   5220
   ClientLeft      =   3825
   ClientTop       =   3465
   ClientWidth     =   8160
   Icon            =   "frmDrugQualityCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6960
      TabIndex        =   38
      Top             =   4320
      Width           =   1100
   End
   Begin TabDlg.SSTab sstQuality 
      Height          =   4935
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "������Ϣ(&D)"
      TabPicture(0)   =   "frmDrugQualityCard.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra������Ϣ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "������Ϣ(&V)"
      TabPicture(1)   =   "frmDrugQualityCard.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraExecute"
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra������Ϣ 
         Height          =   4305
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   6195
         Begin VB.TextBox txtԭ���� 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   300
            Left            =   4545
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1200
            Width           =   1320
         End
         Begin VB.TextBox txt���� 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            MaxLength       =   11
            TabIndex        =   9
            Top             =   1590
            Width           =   4665
         End
         Begin VB.TextBox txt���� 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1215
            Width           =   2505
         End
         Begin VB.CommandButton cmdProvider 
            Caption         =   "��"
            Height          =   300
            Left            =   5610
            TabIndex        =   20
            Top             =   3165
            Width           =   270
         End
         Begin VB.TextBox txtProvider 
            Height          =   300
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   19
            Top             =   3145
            Width           =   4425
         End
         Begin VB.ComboBox cbo����˵�� 
            Height          =   300
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2760
            Width           =   2025
         End
         Begin VB.TextBox TxtName 
            Height          =   300
            Left            =   1215
            MaxLength       =   30
            TabIndex        =   3
            Top             =   825
            Width           =   3345
         End
         Begin VB.TextBox TxtNumber 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1200
            TabIndex        =   15
            Top             =   2735
            Width           =   1440
         End
         Begin VB.CommandButton CmdDrugSelect 
            Caption         =   "��"
            Height          =   300
            Left            =   4560
            TabIndex        =   4
            Top             =   825
            Width           =   270
         End
         Begin VB.ComboBox cboStock 
            Height          =   300
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   4665
         End
         Begin MSComCtl2.DTPicker dtp�������� 
            Height          =   285
            Left            =   3840
            TabIndex        =   22
            Top             =   3563
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   170328067
            CurrentDate     =   36489
         End
         Begin VB.Label lblԭ���� 
            AutoSize        =   -1  'True
            Caption         =   "ԭ����"
            Height          =   180
            Left            =   3840
            TabIndex        =   50
            Top             =   1275
            Width           =   540
         End
         Begin VB.Label lbl��λ 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   2730
            TabIndex        =   49
            Top             =   2805
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label lbl��λ 
            AutoSize        =   -1  'True
            Caption         =   "/��"
            Height          =   180
            Index           =   1
            Left            =   2880
            TabIndex        =   48
            Top             =   2370
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label lbl��λ 
            AutoSize        =   -1  'True
            Caption         =   "/��"
            Height          =   180
            Index           =   0
            Left            =   2880
            TabIndex        =   47
            Top             =   2010
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label txtSale 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4200
            TabIndex        =   13
            Top             =   2325
            Width           =   1680
         End
         Begin VB.Label txtCost 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4200
            TabIndex        =   11
            Top             =   1965
            Width           =   1680
         End
         Begin VB.Label txtSalePrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   12
            Top             =   2325
            Width           =   1680
         End
         Begin VB.Label txtCostPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   10
            Top             =   1965
            Width           =   1680
         End
         Begin VB.Label lblSale 
            AutoSize        =   -1  'True
            Caption         =   "���۽��"
            Height          =   180
            Left            =   3435
            TabIndex        =   46
            Top             =   2370
            Width           =   720
         End
         Begin VB.Label lblCost 
            AutoSize        =   -1  'True
            Caption         =   "�ɱ����"
            Height          =   180
            Left            =   3435
            TabIndex        =   45
            Top             =   2010
            Width           =   720
         End
         Begin VB.Label lblSalePrice 
            AutoSize        =   -1  'True
            Caption         =   "���ۼ�"
            Height          =   180
            Left            =   375
            TabIndex        =   44
            Top             =   2370
            Width           =   540
         End
         Begin VB.Label lblCostPrice 
            AutoSize        =   -1  'True
            Caption         =   "�ɱ���"
            Height          =   180
            Left            =   375
            TabIndex        =   43
            Top             =   2010
            Width           =   540
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   375
            TabIndex        =   8
            Top             =   1635
            Width           =   360
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   375
            TabIndex        =   5
            Top             =   1275
            Width           =   540
         End
         Begin VB.Label txt�Ǽ��� 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   37
            Top             =   3555
            Width           =   1440
         End
         Begin VB.Label txt��λ 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5385
            TabIndex        =   36
            Top             =   825
            Width           =   495
         End
         Begin VB.Label Lbl������λ 
            AutoSize        =   -1  'True
            Caption         =   "��λ"
            Height          =   180
            Left            =   4950
            TabIndex        =   35
            Top             =   885
            Width           =   360
         End
         Begin VB.Label Lbldate 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Left            =   3075
            TabIndex        =   21
            Top             =   3615
            Width           =   720
         End
         Begin VB.Label LblҩƷ��Դ 
            AutoSize        =   -1  'True
            Caption         =   "��ҩ��λ"
            Height          =   180
            Left            =   375
            TabIndex        =   18
            Top             =   3210
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "����ҩƷ"
            Height          =   180
            Left            =   375
            TabIndex        =   2
            Top             =   885
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "�Ǽ���"
            Height          =   180
            Left            =   375
            TabIndex        =   34
            Top             =   3615
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "����˵��"
            Height          =   180
            Left            =   3075
            TabIndex        =   16
            Top             =   2805
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Left            =   375
            TabIndex        =   14
            Top             =   2805
            Width           =   720
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�ⷿ"
            Height          =   180
            Left            =   375
            TabIndex        =   0
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.Frame fraExecute 
         Height          =   3585
         Left            =   -74760
         TabIndex        =   32
         Top             =   840
         Width           =   6195
         Begin VB.ComboBox cbo�����λ 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   2700
            Visible         =   0   'False
            Width           =   3795
         End
         Begin VB.ComboBox cboType 
            Height          =   300
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   2160
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.ComboBox cbo������ 
            Height          =   300
            Left            =   2175
            TabIndex        =   28
            Top             =   1635
            Width           =   2535
         End
         Begin VB.ComboBox cbo����취 
            Height          =   300
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   540
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker dtp�������� 
            Height          =   285
            Left            =   2175
            TabIndex        =   26
            Top             =   1095
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   170328067
            CurrentDate     =   36489
         End
         Begin VB.Label lbl�����λ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�����λ(&D)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1080
            TabIndex        =   42
            Top             =   2760
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblType 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������(&T)"
            Height          =   180
            Left            =   1080
            TabIndex        =   40
            Top             =   2220
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "��������(&Q)"
            Height          =   180
            Left            =   1080
            TabIndex        =   25
            Top             =   1140
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "������(&E)"
            Height          =   180
            Left            =   1260
            TabIndex        =   27
            Top             =   1695
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "����취(&M)"
            Height          =   180
            Left            =   1080
            TabIndex        =   23
            Top             =   600
            Width           =   990
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6960
      TabIndex        =   30
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6960
      TabIndex        =   29
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "frmDrugQualityCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint�༭ģʽ As Integer             '1:�Ǽ�;2:�޸�;3:����,4:�鿴
Private mlng��¼ID As Long
Private mblnSuccess As Boolean
Private mblnChange As Boolean
Private mfrmMain As Form
Private mblnHaveRecord As Boolean           '�����Ѿ�ɾ���ļ�¼���ô˱������жϣ����û��ɾ����Ĭ��ΪTRUE������ΪFALSE

Private mlng�ⷿID As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private mint���ʱ����� As Integer         '���ʱ�Ƿ�ͬ�������棨�൱��ͬʱʵ���������⹦�ܣ���0���������棻1��Ҫͬ��������
Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ

'�Ӳ�������ȡҩƷ�۸����������С��λ�������㾫�ȣ�
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��
Private mstrPrivs As String                 '����ԱȨ��

Private Sub CheckDependOn()
    '�����������
    '�����ʱͬ���������ģʽ��Ҫ���������ⵥ�ݵ��������
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If mint���ʱ����� = 0 Then Exit Sub
    
    gstrSQL = "SELECT b.Id,b.���� " _
            & "FROM ҩƷ�������� A, ҩƷ������ B " _
            & "Where A.���id = B.ID AND A.���� = 11 "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ��������������")
    
    If rsTemp.EOF Then
        MsgBox "δ������������������𣬲���ͬ�������棡", vbExclamation, gstrSysName
        mint���ʱ����� = 0
        Exit Sub
    End If
    
    With cboType
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp.Fields(1)
            .ItemData(.NewIndex) = rsTemp.Fields(0)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    lblType.Visible = True
    cboType.Visible = True
    lbl�����λ.Visible = True
    cbo�����λ.Visible = True
    cbo�����λ.Enabled = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function GetProviderNameById(ByVal lngProviderId As Long) As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select ���� From ��Ӧ�� Where ID = [1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��Ӧ������", lngProviderId)
    
    If Not rsTemp.EOF Then
        GetProviderNameById = rsTemp!����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowCard(ByVal FrmMain As Form, ByVal int�༭ģʽ As Integer, _
        ByVal lng��¼id As Long, Optional strPrivs As String) As Boolean
    Dim rsParrel As New Recordset
    
    mblnSuccess = False
    mblnChange = False
    mint�༭ģʽ = int�༭ģʽ
    mblnHaveRecord = True
    mlng��¼ID = lng��¼id
    mstrPrivs = strPrivs
    
    On Error GoTo errHandle
    If int�༭ģʽ > 1 Then
        gstrSQL = "select nvl(������,'0') from ҩƷ������¼ where id=[1]"
        Set rsParrel = zlDataBase.OpenSQLRecord(gstrSQL, "[��ȡ������]", lng��¼id)
        
        If rsParrel.EOF Then
            MsgBox "��ҩƷ������¼�ѱ�������ɾ�������飡", vbOKOnly, gstrSysName
            Exit Function
        ElseIf rsParrel.Fields(0) <> "0" And InStr(1, 23, mint�༭ģʽ) <> 0 Then
            MsgBox "��ҩƷ������¼�ѱ������˴������飡", vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    Set mfrmMain = FrmMain
    
    Select Case int�༭ģʽ
        Case 1, 2
            sstQuality.TabEnabled(1) = False
            fraExecute.Enabled = False
            sstQuality.Tab = 0
            Me.Caption = "ҩƷ����Ǽ�"
        Case 3
            sstQuality.Tab = 1
            Me.Caption = "ҩƷ������"
        Case 4
            Fra������Ϣ.Enabled = False
            fraExecute.Enabled = False
            Me.Caption = "ҩƷ����鿴"
            cmdOk.Enabled = False
    End Select
    mblnChange = False
    Me.Show vbModal, FrmMain
    ShowCard = mblnSuccess
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function VerifyData() As Boolean
    VerifyData = False
    
    If Val(TxtName.Tag) = 0 Then
        MsgBox "����ҩƷ��������", vbInformation, gstrSysName
        Me.TxtName.SetFocus
        Exit Function
    End If
    If TxtNumber = "" Then
        MsgBox "��������Ӧ������!", vbInformation, gstrSysName
        Me.TxtNumber.SetFocus
        Exit Function
    End If
    
    If Val(TxtNumber) = 0 Then
        MsgBox "�����������������!", vbInformation, gstrSysName
        Me.TxtNumber.SetFocus
        Exit Function
    End If
    
    If Val(Me.TxtNumber) >= 10 ^ 11 - 1 Then
        MsgBox "�����������0С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
        Me.TxtNumber.SetFocus
        Exit Function
    End If
    
    If mint�༭ģʽ = 3 Then
        If cboType.Text = "ҩƷ���" Then
            If cbo�����λ.Text = "" Then
                MsgBox "�����λ����Ϊ�գ������������λ!", vbInformation, gstrSysName
                cbo�����λ.SetFocus
                Exit Function
            End If
        End If
        
        If cboType.Text = "ҩƷ����" Then
            If cbo�����λ.Text = "" Then
                MsgBox "������λ����Ϊ�գ�������������λ!", vbInformation, gstrSysName
                cbo�����λ.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If mint�༭ģʽ = 3 Then
        If Val(txtProvider.Tag) = 0 Then
            MsgBox "��ҩ��λ��������", vbInformation, "��ʾ"
            Me.sstQuality.Tab = 0
            Me.txtProvider.SetFocus
            Exit Function
        End If
        If cbo������.Text = "" Then
            MsgBox "�����˱�������", vbInformation, "��ʾ"
            Me.sstQuality.Tab = 1
            Me.cbo������.SetFocus
            Exit Function
        End If
        
    End If
        
    VerifyData = True
End Function

Private Sub cboStock_Click()
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    Dim rsList As ADODB.Recordset
    
    On Error GoTo errHandle
    
    str�ⷿ���� = ""
    gstrSQL = "Select a.�������� From ��������˵�� A Where a.����id =[1]"
    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, "�ж��ǿⷿ����", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsList.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rsList!��������
        rsList.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
    If bln��ҩ�ⷿ Then
        lblԭ����.Visible = True
        txtԭ����.Visible = True
        txt����.Width = 2505
    Else
        lblԭ����.Visible = False
        txtԭ����.Visible = False
        txt����.Width = 4665
    End If

    mlng�ⷿID = 0
    If mlng�ⷿID <> cboStock.ItemData(cboStock.ListIndex) Then
        mlng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
        Call GetDrugDigit(mlng�ⷿID, "ҩƷ��������", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        mint����� = MediWork_GetCheckStockRule(mlng�ⷿID)
        Call ReleaseSelectorRS
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboType_click()
    cbo�����λ.Clear
    If cboType.Text = "ҩƷ���" Or cboType.Text = "ҩƷ����" Then
        cbo�����λ.Enabled = True
    End If
    
    If cboType.Text = "ҩƷ����" Then
        lbl�����λ.Caption = "������λ"
    End If
    
    If cboType.Text = "ҩƷ��������" Then
        cbo�����λ.Enabled = False
    End If
End Sub

Private Sub cbo����취_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    Dim i As Integer, intIdx As Integer
    Dim strText As String
    Dim rsdepart As New Recordset
    
    On Error GoTo errHandle
    With cbo������
        strText = .Text
        If strText = "" Then
            .ListIndex = -1
        Else
            intIdx = -1
            For i = 0 To .ListCount - 1
                If InStr(.List(i), UCase(strText)) > 0 Then
                    If intIdx = -1 Then .ListIndex = i
                    intIdx = i
                End If
            Next
            If intIdx = -1 Then
                gstrSQL = "Select id,���� from ��Ա�� " & _
                          "Where (վ�� = [2] Or վ�� is Null) And (���� like [1] or ��� like [1]) " & _
                          "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
                Set rsdepart = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ������]", UCase(strText) & "%", gstrNodeNo)
                
                If Not rsdepart.EOF Then
                    Do While Not rsdepart.EOF
                        For i = 0 To .ListCount - 1
                            If InStr(.List(i), rsdepart.Fields(1)) > 0 Then
                                If intIdx = -1 Then .ListIndex = i
                                intIdx = i
                            End If
                        Next
                        rsdepart.MoveNext
                    Loop
                End If
            End If
        End If
        If Trim(.Text) = "" Then
            MsgBox "�Բ��𣬱�������һ��������!", vbExclamation + vbOKOnly, gstrSysName
            .SetFocus
            Exit Sub
        End If
        
        If .ListIndex = -1 Then
            MsgBox "�Բ���û���ҵ����������Ա�������䣡", vbExclamation + vbOKOnly, gstrSysName
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        Else
            If intIdx <> .ListIndex Then SendKeys "{F4}": Exit Sub
        End If
    End With
    OS.PressKey (vbKeyTab)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo����˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub cbo�����λ_DropDown()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If cboType.Text = "ҩƷ���" Then
        gstrSQL = "Select ����||'-'||���� AS �����λ From ҩƷ�����λ Order By ����"
    ElseIf cboType.Text = "ҩƷ����" Then
        gstrSQL = "Select ����||'-'||���� AS �����λ From ҩƷ������λ Order By ����"
    End If
    
    'Call zldatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ�����λ")
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-��ȡ�����λ")
    With cbo�����λ
        .Clear
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!�����λ
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDrugSelect_Click()
    Dim RecReturn As Recordset
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "ҩƷ��������", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    
'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), False)
    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , False, , , , False, mstrPrivs)
    
    If RecReturn.RecordCount > 0 Then
        TxtName.Tag = RecReturn!ҩƷid
        If gintҩƷ������ʾ = 1 Then
            TxtName.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
        Else
            TxtName.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
        End If
        txt��λ = Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ)
        txt��λ.Tag = Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ)
        txt����.Text = IIf(IsNull(RecReturn!����), "", RecReturn!����)
        txtԭ����.Text = IIf(IsNull(RecReturn!ԭ����), "", RecReturn!ԭ����)
        txt����.Text = IIf(IsNull(RecReturn!����), "", RecReturn!����)
        txt����.Tag = IIf(IsNull(RecReturn!����), 0, RecReturn!����)
        txtProvider.Tag = RecReturn!�ϴι�Ӧ��ID
        
'        If IsNull(RecReturn!�ɱ���) Then
'            txtCostPrice.Caption = ""
'            txtCost.Caption = ""
'        Else
'            txtCostPrice.Caption = zlStr.FormatEx(RecReturn!�ɱ���, mintCostDigit, , True)
'            txtCost.Caption = zlStr.FormatEx(Val(txtCostPrice.Caption) * Val(Me.TxtNumber) * Val(txt��λ.Tag), mintMoneyDigit, , True)
'        End If
'        If IsNull(RecReturn!�ۼ�) Then
'            txtSalePrice.Caption = ""
'            txtSale.Caption = ""
'        Else
'            txtSalePrice.Caption = zlStr.FormatEx(RecReturn!�ۼ�, mintPriceDigit, , True)
'            txtSale.Caption = zlStr.FormatEx(Val(txtSalePrice.Caption) * Val(Me.TxtNumber) * Val(txt��λ.Tag), mintMoneyDigit, , True)
'        End If
        
        lbl��λ(0).Visible = True
        lbl��λ(1).Visible = True
        lbl��λ(2).Visible = True
        lbl��λ(0).Caption = "/" & txt��λ.Caption
        lbl��λ(1).Caption = "/" & txt��λ.Caption
        lbl��λ(2).Caption = txt��λ.Caption
        
        txtCostPrice.Tag = zlStr.FormatEx(Get�ɱ���(RecReturn!ҩƷid, mlng�ⷿID, IIf(IsNull(RecReturn!����), 0, RecReturn!����)), gtype_UserDrugDigits.Digit_�ɱ���, , True)
        txtCostPrice.Caption = zlStr.FormatEx(Val(txtCostPrice.Tag) * Val(txt��λ.Tag), mintCostDigit, , True)
        txtCost.Caption = zlStr.FormatEx(Val(txtCostPrice.Caption) * Val(Me.TxtNumber), mintMoneyDigit, , True)
        
        If RecReturn!ʱ�� = 1 Then
            txtSalePrice.Tag = zlStr.FormatEx(Get���ۼ�(RecReturn!ҩƷid, mlng�ⷿID, IIf(IsNull(RecReturn!����), 0, RecReturn!����), 1), gtype_UserDrugDigits.Digit_���ۼ�, , True)
        Else
            txtSalePrice.Tag = zlStr.FormatEx(RecReturn!�ۼ�, gtype_UserDrugDigits.Digit_���ۼ�, , True)
        End If
        txtSalePrice.Caption = zlStr.FormatEx(Val(txtSalePrice.Tag) * Val(txt��λ.Tag), mintPriceDigit, , True)
        txtSale.Caption = zlStr.FormatEx(Val(txtSalePrice.Caption) * Val(Me.TxtNumber), mintMoneyDigit, , True)
        
        
        If Val(txtProvider.Tag) <> 0 Then
            txtProvider.Text = GetProviderNameById(Val(txtProvider.Tag))
        End If
        
        TxtNumber.SetFocus
    End If
End Sub


Private Function SaveCard() As Boolean
    Dim dblTmp As Double
    On Error GoTo errHandle
    SaveCard = False
    
    If mint�༭ģʽ = 2 Then
        gstrSQL = "zl_ҩƷ��������_delete(" & mlng��¼ID & ")"
        Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    gstrSQL = "zl_ҩƷ��������_INSERT("
    '�ⷿID
    gstrSQL = gstrSQL & cboStock.ItemData(cboStock.ListIndex)
    'ҩƷID
    gstrSQL = gstrSQL & "," & Val(TxtName.Tag)
    '����ԭ��
    gstrSQL = gstrSQL & ",'" & cbo����˵��.Text & "'"
    '��������
    gstrSQL = gstrSQL & "," & FormatEx(Val(TxtNumber.Text) * Val(txt��λ.Tag), mintNumberDigit)
    '�Ǽ���
    gstrSQL = gstrSQL & ",'" & txt�Ǽ��� & "'"
    '�Ǽ�ʱ��
    gstrSQL = gstrSQL & ",to_date('" & Format(dtp��������.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')"
    '����
    gstrSQL = gstrSQL & ",'" & txt����.Text & "'"
    '����
    gstrSQL = gstrSQL & ",'" & txt����.Text & "'"
    '����
    gstrSQL = gstrSQL & "," & Val(txt����.Tag)
    '��ҩ��λID
    gstrSQL = gstrSQL & "," & IIf(Val(txtProvider.Tag) = 0, "NULL", txtProvider.Tag)
    '�ɱ�����
    gstrSQL = gstrSQL & "," & IIf(Val(txtCostPrice.Tag) = 0, "null", Val(txtCostPrice.Tag))
    '�ɱ����
    dblTmp = zlStr.FormatEx(Val(txtCost.Caption), mintMoneyDigit, , True)
    gstrSQL = gstrSQL & "," & IIf(dblTmp = 0, "null", dblTmp)
    '���۵���
    gstrSQL = gstrSQL & "," & IIf(Val(txtSalePrice.Tag) = 0, "null", Val(txtSalePrice.Tag))
    '���۽��
    dblTmp = zlStr.FormatEx(Val(txtSale.Caption), mintMoneyDigit, , True)
    gstrSQL = gstrSQL & "," & IIf(dblTmp = 0, "null", dblTmp)
    '˵��
    gstrSQL = gstrSQL & ",NULL"
    gstrSQL = gstrSQL & ")"

    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveHandle() As Boolean
    Dim lng������id As Long
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngTypeID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchID As Long
    Dim strProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim dblOutPrice As Double   '�����
    Dim strOutUnit As String    '�����λ
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim arrSql As Variant
    Dim intRow As Integer
    Dim str��׼�ĺ� As String
    Dim blnTran As Boolean
    
    Dim rsTemp As New Recordset
    
    On Error GoTo errHandle
    SaveHandle = False
    
    If mint���ʱ����� = 1 Then
        lng������id = cboType.ItemData(cboType.ListIndex)
        chrNo = Sys.GetNextNo(28, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        lngSerial = 1
        lngStockid = mlng�ⷿID
        lngDrugID = Val(TxtName.Tag)
        lngBatchID = Val(txt����.Tag)
        
'        dblQuantity = zlStr.FormatEx(Val(TxtNumber.Text) * Val(txt��λ.Tag), gtype_UserSaleDigits.Digit_����)
        dblQuantity = Val(TxtNumber.Tag)    '����ʱ��ԭʼ����
        
        '�����
        If CheckDrugStock(lngStockid, lngDrugID, lngBatchID, dblQuantity, Val(txt��λ.Tag)) = False Then
            Exit Function
        End If
        
        gstrSQL = "Select Nvl(A.ʵ������,0) ʵ������, Nvl(A.ʵ�ʽ��,0) ʵ�ʽ��, Nvl(A.ʵ�ʲ��,0) ʵ�ʲ��, A.Ч��, A.��׼�ĺ�, " & _
            " Nvl(B.�Ƿ���, 0) �Ƿ���, C.�ּ�, Nvl(D.����ѱ���, 0) ����,Nvl(A.����,0) As ����,Nvl(A.���ۼ�,0) As ���ۼ� " & _
            " From ҩƷ��� A, �շ���ĿĿ¼ B, �շѼ�Ŀ C, ҩƷ��� D " & _
            " Where A.ҩƷid = B.ID And A.ҩƷid = C.�շ�ϸĿid And A.ҩƷid = D.ҩƷid And A.���� = 1 And " & _
            " (C.��ֹ���� Is Null Or Sysdate Between C.ִ������ And Nvl(C.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd'))) And " & _
            " A.�ⷿid = [1] And A.ҩƷid = [2] And Nvl(A.����, 0) = [3]" & _
            GetPriceClassString("C")
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ�۸���Ϣ", lngStockid, lngDrugID, lngBatchID)
        
        If rsTemp.RecordCount = 0 Then
            gstrSQL = "Select Nvl(�Ƿ���, 0) �Ƿ���, Nvl(d.����ѱ���, 0) ����, d.�ϴ���׼�ĺ� ��׼�ĺ�, '' Ч��" & vbNewLine & _
                "From �շ���ĿĿ¼ B, ҩƷ��� D" & vbNewLine & _
                "Where ID = [2] And b.Id = d.ҩƷid"
            
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ�۸���Ϣ", lngStockid, lngDrugID, lngBatchID)
        End If
        
'        If rsTemp!�Ƿ��� = 0 Then
'            dblSalePrice = rsTemp!�ּ�
'        Else
'            dblSalePrice = IIf(rsTemp!���� = 0, rsTemp!ʵ�ʽ�� / rsTemp!ʵ������, IIf(rsTemp!���ۼ� = 0, rsTemp!ʵ�ʽ�� / rsTemp!ʵ������, rsTemp!���ۼ�))
'        End If
        
        dblSalePrice = Get�ۼ�(rsTemp!�Ƿ��� = 1, lngDrugID, lngStockid, lngBatchID)
        dblSaleMoney = zlStr.FormatEx(dblSalePrice * dblQuantity, mintMoneyDigit)
        
        dblPurchasePrice = Get�ɱ���(lngDrugID, lngStockid, lngBatchID)
        dblPurchaseMoney = zlStr.FormatEx(dblPurchasePrice * dblQuantity, mintMoneyDigit)
        
        dblMistakePrice = zlStr.FormatEx(dblSaleMoney - dblPurchaseMoney, mintMoneyDigit)
        
        If cboType.Text = "ҩƷ���" Then
            dblOutPrice = zlStr.FormatEx((1 + rsTemp!���� / 100) * dblPurchasePrice, gtype_UserSaleDigits.Digit_�ɱ���)
            If Not cbo�����λ.Text = "" Then
                strOutUnit = Mid(cbo�����λ.Text, 1, InStr(1, cbo�����λ.Text, "-") - 1)
            End If
        End If
        
        strBooker = UserInfo.�û�����
        datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        strProducingArea = IIf(IsNull(txt����.Text), "", txt����.Text)
        strBatchNo = IIf(IsNull(txt����.Text), "", txt����.Text)
        datTimeLimit = IIf(IsNull(rsTemp!Ч��), "", rsTemp!Ч��)
        strBrief = cbo����취.Text & "(���������Զ������)"
        str��׼�ĺ� = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
        
        rsTemp.Close
        
        '���۹�������Ƿ���ڲ��������۵�ҩƷ
        If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
            If IsPriceAdjustMod(lngDrugID) = True Then
                If CheckPriceAdjust(lngDrugID, lngStockid, lngBatchID) = False Then
                    MsgBox "��ҩƷ���������۹���������¼���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        gcnOracle.BeginTrans
        blnTran = True
    End If
    
    gstrSQL = "zl_ҩƷ��������_UPDATE("
    '��¼ID
    gstrSQL = gstrSQL & mlng��¼ID
    '��ҩ��λID
    gstrSQL = gstrSQL & "," & Val(txtProvider.Tag)
    '����취
    gstrSQL = gstrSQL & ",'" & cbo����취.Text & "'"
    '������
    gstrSQL = gstrSQL & ",'" & cbo������.Text & "'"
    '����ʱ��
    gstrSQL = gstrSQL & ",to_date('" & Format(dtp��������.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')"
'    '�ɱ�����
'    gstrSQL = gstrSQL & "," & txtCostPrice.Caption
'    '�ɱ����
'    gstrSQL = gstrSQL & "," & Val(txtCostPrice.Caption) * Val(TxtNumber.Text) * Val(txt��λ.Tag)
'    '���۵���
'    gstrSQL = gstrSQL & "," & txtSalePrice.Caption
'    '���۽��
'    gstrSQL = gstrSQL & "," & Val(txtSalePrice.Caption) * Val(TxtNumber.Text) * Val(txt��λ.Tag)
    '���ⵥNO
    gstrSQL = gstrSQL & "," & IIf(chrNo = "", "Null", "'" & chrNo & "'")
    gstrSQL = gstrSQL & ")"

    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
   
    If mint���ʱ����� = 1 Then
        gstrSQL = "zl_ҩƷ��������_INSERT("
        '������ID
        gstrSQL = gstrSQL & lng������id
        'NO
        gstrSQL = gstrSQL & ",'" & chrNo & "'"
        '���
        gstrSQL = gstrSQL & "," & lngSerial
        '�ⷿID
        gstrSQL = gstrSQL & "," & lngStockid
        'ҩƷID
        gstrSQL = gstrSQL & "," & lngDrugID
        '����
        gstrSQL = gstrSQL & "," & lngBatchID
        '��д����
        gstrSQL = gstrSQL & "," & dblQuantity
        '�ɱ���
        gstrSQL = gstrSQL & "," & dblPurchasePrice
        '�ɱ����
        gstrSQL = gstrSQL & "," & dblPurchaseMoney
        '�ۼ�
        gstrSQL = gstrSQL & "," & dblSalePrice
        '�ۼ۽��
        gstrSQL = gstrSQL & "," & dblSaleMoney
        '���
        gstrSQL = gstrSQL & "," & dblMistakePrice
        '�����
        gstrSQL = gstrSQL & "," & dblOutPrice
        '�����λ
        gstrSQL = gstrSQL & ",'" & strOutUnit & "'"
        '������
        gstrSQL = gstrSQL & ",'" & strBooker & "'"
        '��������
        gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
        '����
        gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
        '����
        gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
        'Ч��
        gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
        'ժҪ
        gstrSQL = gstrSQL & ",'" & strBrief & "'"
        '��׼�ĺ�
        gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
        '��ֵ˰��
        gstrSQL = gstrSQL & ",''"
        'ԭ����
        gstrSQL = gstrSQL & ",''"
        '�޸���
        gstrSQL = gstrSQL & ",''"
        '�޸�����
        gstrSQL = gstrSQL & ",''"
        gstrSQL = gstrSQL & ")"

        Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        gstrSQL = "zl_ҩƷ��������_Verify("
        '���
        gstrSQL = gstrSQL & lngSerial
        'NO
        gstrSQL = gstrSQL & ",'" & chrNo & "'"
        '�ⷿID
        gstrSQL = gstrSQL & "," & lngStockid
        'ҩƷID
        gstrSQL = gstrSQL & "," & lngDrugID
        '����
        gstrSQL = gstrSQL & "," & lngBatchID
        'ʵ������
        gstrSQL = gstrSQL & "," & dblQuantity
        '�ɱ���
        gstrSQL = gstrSQL & "," & dblPurchasePrice
        '�ɱ����
        gstrSQL = gstrSQL & "," & dblPurchaseMoney
        '���۽��
        gstrSQL = gstrSQL & "," & dblSaleMoney
        '���
        gstrSQL = gstrSQL & "," & dblMistakePrice
        '�����
        gstrSQL = gstrSQL & ",'" & strBooker & "'"
        '�������
        gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
        gstrSQL = gstrSQL & ")"
 
        Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        gcnOracle.CommitTrans
        blnTran = False
    End If
    
    mblnSuccess = True
    mblnChange = False
    SaveHandle = True
    Exit Function
errHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckDrugStock(ByVal lng�ⷿID As Long, ByVal lngҩƷID As Long, ByVal lng���� As Long, ByVal DblҩƷ���� As Double, ByVal dbl��װϵ�� As Double) As Boolean
    Dim blnMsg As Boolean
    Dim rsCheck As New ADODB.Recordset
    Dim Dbl���� As Double
    
    On Error GoTo errHandle
    
    If mint����� = 0 Then CheckDrugStock = True: Exit Function
       
    gstrSQL = "Select Nvl(��������,0) ��������,Nvl(ʵ������,0) ʵ������ " & _
              "From ҩƷ��� Where �ⷿID=[1] And Nvl(����,0)=[3] And ����=1 And ҩƷID=[2] "
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[�����]", lng�ⷿID, lngҩƷID, lng����)
    
    If mint�༭ģʽ = 1 Or mint�༭ģʽ = 2 Then
        '����޸�ʱ�ÿ����������
        Dbl���� = rsCheck!��������
    ElseIf mint�༭ģʽ = 3 Then
        '���ʱȡʵ�����������
        Dbl���� = rsCheck!ʵ������
    End If
       
    If DblҩƷ���� > Dbl���� Then
        blnMsg = True
    End If
    
    If blnMsg Then
        If mint����� = 1 Then        '��������
            If MsgBox("���������������еĿ������" & Dbl���� / dbl��װϵ�� & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else                            '�����ֹ
            MsgBox "���������������еĿ������" & Dbl���� / dbl��װϵ�� & "�����ܳ��⣡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    CheckDrugStock = True

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOK_Click()
    
    If Not VerifyData Then Exit Sub
    
    Select Case mint�༭ģʽ
        Case 1
            mblnSuccess = SaveCard
            If mblnSuccess = True Then
                TxtName.Text = ""
                TxtName.Tag = ""
                txt��λ = ""
                TxtNumber.Text = ""
                txtProvider.Text = ""
                txtProvider.Tag = ""
                If cboStock.Enabled = True Then
                    cboStock.SetFocus
                Else
                    TxtName.SetFocus
                End If
            End If
                
        Case 2
            mblnSuccess = SaveCard
            If mblnSuccess = True Then
                Unload Me
                Exit Sub
            End If
            
            
        Case 3
            mblnSuccess = SaveHandle
            If mblnSuccess = True Then
                Unload Me
                Exit Sub
            End If
    End Select
    
End Sub

Private Sub cmdProvider_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txtProvider.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ��Ӧ�� " & _
              "Where (վ�� = [1] Or վ�� is Null) And To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' " & _
              "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
              "Start with �ϼ�ID is null connect by prior ID =�ϼ�ID order by level,ID"
'    Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-��Ӧ��", gstrNodeNo)
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "��Ӧ��", True, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider.State = 0 Then
        Exit Sub
    End If
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    Me.txtProvider.Tag = rsProvider!id
    Me.txtProvider = rsProvider!����
    
    If mint�༭ģʽ = 3 Then
        cmdOk.SetFocus
    Else
        dtp��������.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtp��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub dtp��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdOk.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    mblnChange = True
End Sub

Private Sub Form_Load()
    Dim rsTmp As New Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    With cbo����˵��
        .Clear
        gstrSQL = "select ���� from ������ԭ�� "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-������ԭ��")
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        If .ListCount > 1 Then
            .ListIndex = 0
        End If
        
    End With
    
    With cbo����취
        gstrSQL = "select ���� from �������취"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-�������취")
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        If .ListCount > 1 Then
            .ListIndex = 0
        End If
    End With
    
    With cbo������
'        gstrSQL = "select id, ���� from ��Ա�� " & _
'                  "Where (վ�� = [1] Or վ�� is Null) And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        gstrSQL = "Select distinct a.Id, a.���� " & vbNewLine & _
                  "From ��Ա�� A, ������Ա B, ������Ա C " & vbNewLine & _
                  "Where a.Id = b.��Աid And b.����id = c.����id And c.��Աid = [2] And (a.վ�� = [1] Or a.վ�� Is Null) And " & vbNewLine & _
                  "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & vbNewLine & _
                  "Order By a.���� "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-��Ա��Ϣ", gstrNodeNo, UserInfo.�û�ID)
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            .ItemData(.NewIndex) = rsTmp!id
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        
        .Text = UserInfo.�û�����
        
    End With
    
    With mfrmMain.cboStock
        cboStock.Clear
        
        For intIndex = 0 To .ListCount - 1
            cboStock.AddItem .List(intIndex)
            cboStock.ItemData(cboStock.NewIndex) = .ItemData(intIndex)
        Next
        cboStock.ListIndex = .ListIndex
        cboStock.Enabled = .Enabled
    End With
    
    mlng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    Call GetDrugDigit(mlng�ⷿID, "ҩƷ��������", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    dtp��������.Value = Format(Sys.Currentdate, "yyyy-mm-dd")
    dtp��������.Value = Format(Sys.Currentdate, "yyyy-mm-dd")
    txt�Ǽ��� = UserInfo.�û�����
    
    mint���ʱ����� = Val(zlDataBase.GetPara("���ʱ���ٿ��", glngSys, ģ���.��������))
    Call CheckDependOn
    
    If mint�༭ģʽ > 1 Then
        initCard
    End If
    
    
    If cboType.Text = "ҩƷ��������" Then
        cbo�����λ.Enabled = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim intIndex As Integer
    Dim intBit As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    Dim rsList As ADODB.Recordset
    
    intBit = IIf(gintҩƷ������ʾ = 2, 1, 0)
    '���ǵ��ۡ��������ľ��ȣ�����ȡ����
    On Error GoTo errHandle
    str�ⷿ���� = ""
    gstrSQL = "Select a.�������� From ��������˵�� A Where a.����id =[1]"
    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, "�ж��ǿⷿ����", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsList.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rsList!��������
        rsList.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True 'bln��ҩ�ⷿΪ��ʱ��8�м���֮����кŶ�+1
    
    strsql = "select �ɱ�����, �ɱ����, ���۵���, ���۽��, �������� " _
           & "from ҩƷ������¼ where id=[1]"
    Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, mlng��¼ID)
    If rsTmp.EOF Then Exit Sub
    With frmDrugQualityList.vsfList
        Me.Tag = .TextMatrix(.Row, 0)
        TxtName.Text = .TextMatrix(.Row, 1)
        TxtName.Tag = .TextMatrix(.Row, 2 + intBit)
        txt����.Tag = IIf(IsNull(.TextMatrix(.Row, 3 + intBit)), 0, .TextMatrix(.Row, 3 + intBit))
        txtProvider.Tag = IIf(IsNull(.TextMatrix(.Row, 4 + intBit)), 0, .TextMatrix(.Row, 4 + intBit))
        txtProvider.Text = .TextMatrix(.Row, 16 + intBit + IIf(bln��ҩ�ⷿ, 1, 0))
        txt����.Text = IIf(IsNull(.TextMatrix(.Row, 6 + intBit)), "", .TextMatrix(.Row, 6 + intBit))
        
        txt����.Text = IIf(IsNull(.TextMatrix(.Row, 7 + intBit)), "", .TextMatrix(.Row, 7 + intBit))
    
        If bln��ҩ�ⷿ Then intBit = intBit + 1: txtԭ����.Text = IIf(IsNull(.TextMatrix(.Row, 6 + intBit)), "", .TextMatrix(.Row, 6 + intBit))
    
        txt��λ = .TextMatrix(.Row, 12 + intBit)
        TxtNumber.Text = zlStr.FormatEx(rsTmp!�������� / .TextMatrix(.Row, 14 + intBit), IIf(mint�༭ģʽ = 4, mintNumberDigit, mintNumberDigit), , True)
        TxtNumber.Tag = Nvl(rsTmp!��������, 0)  '��¼ԭʼ����
        txt��λ.Tag = .TextMatrix(.Row, 14 + intBit)        '����ϵ��
        
        lbl��λ(0).Visible = True
        lbl��λ(1).Visible = True
        lbl��λ(2).Visible = True
        lbl��λ(0).Caption = "/" & txt��λ.Caption
        lbl��λ(1).Caption = "/" & txt��λ.Caption
        lbl��λ(2).Caption = txt��λ.Caption
        
        txtCostPrice.Tag = zlStr.FormatEx(rsTmp!�ɱ�����, gtype_UserDrugDigits.Digit_�ɱ���, , True)
        txtCostPrice.Caption = zlStr.FormatEx(Val(txtCostPrice.Tag) * Val(txt��λ.Tag), mintCostDigit, , True)
        If IsNull(rsTmp!�ɱ����) Then
            txtCost.Caption = ""
        Else
            txtCost.Caption = zlStr.FormatEx(rsTmp!�ɱ����, mintMoneyDigit, , True)
        End If
        
        txtSalePrice.Tag = zlStr.FormatEx(rsTmp!���۵���, gtype_UserDrugDigits.Digit_���ۼ�, , True)
        txtSalePrice.Caption = zlStr.FormatEx(Val(txtSalePrice.Tag) * Val(txt��λ.Tag), mintPriceDigit, , True)
        If IsNull(rsTmp!���۽��) Then
            txtSale.Caption = ""
        Else
            txtSale.Caption = zlStr.FormatEx(rsTmp!���۽��, mintMoneyDigit, , True)
        End If
        
        For intIndex = 0 To cbo����˵��.ListCount - 1
            If cbo����˵��.List(intIndex) = .TextMatrix(.Row, 15 + intBit) Then
                cbo����˵��.ListIndex = intIndex
                Exit For
            End If
        Next
        
        txt�Ǽ��� = .TextMatrix(.Row, 17 + intBit)
        dtp��������.Value = .TextMatrix(.Row, 18 + intBit)
        If IIf(IsNull(.TextMatrix(.Row, 19 + intBit)), "", .TextMatrix(.Row, 19 + intBit)) <> "" Then
            For intIndex = 0 To cbo����취.ListCount - 1
                If cbo����취.List(intIndex) = .TextMatrix(.Row, 19 + intBit) Then
                    cbo����취.ListIndex = intIndex
                    Exit For
                End If
            Next
            cbo������.Text = .TextMatrix(.Row, 20 + intBit)
            dtp��������.Value = .TextMatrix(.Row, 21 + intBit)
        End If
        
    End With
    
    If mint�༭ģʽ = 3 Then
        TxtName.Enabled = False
        CmdDrugSelect.Enabled = False
        txt����.Enabled = False
        txtԭ����.Enabled = False
        
        txt����.Enabled = False
        TxtNumber.Enabled = False
        cbo����˵��.Enabled = False
        dtp��������.Enabled = False
        cboStock.Enabled = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If mint�༭ģʽ = 4 Then Call ReleaseSelectorRS: Exit Sub
    If mblnChange = True Then
        If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Call ReleaseSelectorRS
End Sub


Private Sub txtName_GotFocus()
    Me.TxtName.SelStart = 0
    Me.TxtName.SelLength = Len(Me.TxtName.Text)
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Dim strTmp As String
    
    Me.TxtName.Text = Trim(Me.TxtName.Text)
    If Len(LTrim(RTrim(TxtName))) = 0 Then Exit Sub
    strTmp = UCase(TxtName)
    Dim RecReturn As Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
            
    If Mid(strTmp, 1, 1) = "[" Then
        If InStr(2, strTmp, "]") <> 0 Then
            strTmp = Mid(strTmp, 2, InStr(2, strTmp, "]") - 2)
        Else
            strTmp = Mid(strTmp, 2)
        End If
    End If
        
    sngLeft = Me.Left + sstQuality.Left + Fra������Ϣ.Left + TxtName.Left
    sngTop = Me.Top + Me.Height - Me.ScaleHeight + sstQuality.Top + Fra������Ϣ.Top + TxtName.Top + TxtName.Height
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - TxtName.Height - 4530
    End If
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "ҩƷ��������", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    
'    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), strTmp, sngLeft, sngTop, False)
    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strTmp, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , False, , , , False, mstrPrivs)
    
    If RecReturn.RecordCount = 1 Then
        TxtName.Tag = RecReturn!ҩƷid
        If gintҩƷ������ʾ = 1 Then
            TxtName.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
        Else
            TxtName.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
        End If
        txt��λ = Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ)
        txt��λ.Tag = Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ)
        txt����.Text = IIf(IsNull(RecReturn!����), "", RecReturn!����)
        txtԭ����.Text = IIf(IsNull(RecReturn!ԭ����), "", RecReturn!ԭ����)
        txt����.Text = IIf(IsNull(RecReturn!����), "", RecReturn!����)
        txt����.Tag = IIf(IsNull(RecReturn!����), 0, RecReturn!����)
        txtProvider.Tag = RecReturn!�ϴι�Ӧ��ID
'        If IsNull(RecReturn!�ɱ���) Then
'            txtCostPrice.Caption = ""
'            txtCost.Caption = ""
'        Else
'            txtCostPrice.Caption = zlStr.FormatEx(RecReturn!�ɱ���, mintCostDigit, , True)
'            txtCost.Caption = zlStr.FormatEx(Val(txtCostPrice.Caption) * Val(Me.TxtNumber) * Val(txt��λ.Tag), mintMoneyDigit, , True)
'        End If
'        If IsNull(RecReturn!�ۼ�) Then
'            txtSalePrice.Caption = ""
'            txtSale.Caption = ""
'        Else
'            txtSalePrice.Caption = zlStr.FormatEx(RecReturn!�ۼ�, mintPriceDigit, , True)
'            txtSale.Caption = zlStr.FormatEx(Val(txtSalePrice.Caption) * Val(Me.TxtNumber) * Val(txt��λ.Tag), mintMoneyDigit, , True)
'        End If
        
        lbl��λ(0).Visible = True
        lbl��λ(1).Visible = True
        lbl��λ(2).Visible = True
        lbl��λ(0).Caption = "/" & txt��λ.Caption
        lbl��λ(1).Caption = "/" & txt��λ.Caption
        lbl��λ(2).Caption = txt��λ.Caption
        
        txtCostPrice.Tag = zlStr.FormatEx(Get�ɱ���(RecReturn!ҩƷid, mlng�ⷿID, IIf(IsNull(RecReturn!����), 0, RecReturn!����)), gtype_UserDrugDigits.Digit_�ɱ���, , True)
        txtCostPrice.Caption = zlStr.FormatEx(Val(txtCostPrice.Tag) * Val(txt��λ.Tag), mintCostDigit, , True)
        txtCost.Caption = zlStr.FormatEx(Val(txtCostPrice.Caption) * Val(Me.TxtNumber), mintMoneyDigit, , True)
        
        If RecReturn!ʱ�� = 1 Then
            txtSalePrice.Tag = zlStr.FormatEx(Get���ۼ�(RecReturn!ҩƷid, mlng�ⷿID, IIf(IsNull(RecReturn!����), 0, RecReturn!����), 1), gtype_UserDrugDigits.Digit_���ۼ�, , True)
        Else
            txtSalePrice.Tag = zlStr.FormatEx(RecReturn!�ۼ�, gtype_UserDrugDigits.Digit_���ۼ�, , True)
        End If
        txtSalePrice.Caption = zlStr.FormatEx(Val(txtSalePrice.Tag) * Val(txt��λ.Tag), mintPriceDigit, , True)
        txtSale.Caption = zlStr.FormatEx(Val(txtSalePrice.Caption) * Val(Me.TxtNumber), mintMoneyDigit, , True)
        
        If Val(txtProvider.Tag) <> 0 Then
            txtProvider.Text = GetProviderNameById(Val(txtProvider.Tag))
        End If
        
        TxtNumber.SetFocus
        
        
    End If
End Sub


Private Sub TxtNumber_GotFocus()
    TxtNumber.SelStart = 0
    TxtNumber.SelLength = Len(TxtNumber)
End Sub

Private Sub TxtNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(1, "1234567890." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If TxtNumber.SelLength = Len(TxtNumber.Text) Then Exit Sub
            If Len(Mid(TxtNumber, InStr(1, TxtNumber.Text, ".") + 1)) >= mintNumberDigit And TxtNumber.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End If
End Sub


Private Sub TxtNumber_Validate(Cancel As Boolean)
    Dim lng�ⷿID As Long
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim lng���� As Long
    
    If Trim(TxtNumber.Text) <> "" Then
        If Not IsNumeric(TxtNumber.Text) Then
            MsgBox "�Բ��𣬶������������������ͣ����飡", vbExclamation, gstrSysName
            Cancel = True
        Else
            If txtCostPrice.Caption <> "" Then
                txtCost.Caption = zlStr.FormatEx(zlStr.FormatEx(Val(txtCostPrice.Caption), mintCostDigit) * zlStr.FormatEx(Val(TxtNumber.Text), mintNumberDigit), mintMoneyDigit, , True)
            End If
            If txtSalePrice.Caption <> "" Then
                txtSale.Caption = zlStr.FormatEx(zlStr.FormatEx(txtSalePrice.Caption, mintPriceDigit) * zlStr.FormatEx(Val(TxtNumber.Text), mintNumberDigit), mintMoneyDigit, , True)
            End If
        End If
        
        TxtNumber.Text = zlStr.FormatEx(TxtNumber.Text, mintNumberDigit, , True)
        lng���� = Val(TxtNumber.Text) * Val(txt��λ.Tag)
        lng���� = txt����.Tag
        lngҩƷID = TxtName.Tag
        lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
        If CheckDrugStock(lng�ⷿID, lngҩƷID, lng����, lng����, Val(txt��λ.Tag)) = False Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtProvider.hWnd)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint�༭ģʽ = 4 Then Exit Sub
     
    On Error GoTo errHandle
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        
        gstrSQL = "Select id,����,����,���� From ��Ӧ�� " & _
                  "Where (վ�� = [2] Or վ�� is Null) " & _
                  "  And To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' And ĩ��=1 And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
                  "  And (���� like [1] Or ���� like [1] or ���� like [1]) "
'        Set adoProvider = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        Set adoProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ��Ӧ��", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        
        If blnCancel Then txtProvider.SetFocus: Exit Sub
         
        If adoProvider.State = 0 Then
            MsgBox "û��������Ĺ�ҩ��λ�������䣡", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        
        .Text = adoProvider!����
        .Tag = adoProvider!id

        adoProvider.Close
    
    End With
    If mint�༭ģʽ = 3 Then
        cmdOk.SetFocus
    Else
        dtp��������.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt����_GotFocus()
    OS.OpenIme True
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txt����_LostFocus()
    OS.OpenIme
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If Trim(txt����.Text) <> "" Then
        If LenB(StrConv(txt����.Text, vbFromUnicode)) > 30 Then
            MsgBox "�����̳��������������15�����ֻ�30���ַ�!", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub
Private Sub txtԭ����_GotFocus()
    OS.OpenIme True
    txtԭ����.SelStart = 0
    txtԭ����.SelLength = Len(txtԭ����.Text)
End Sub

Private Sub txtԭ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtԭ����_LostFocus()
    OS.OpenIme
End Sub

Private Sub txtԭ����_Validate(Cancel As Boolean)
    If Trim(txtԭ����.Text) <> "" Then
        If LenB(StrConv(txtԭ����.Text, vbFromUnicode)) > 30 Then
            MsgBox "ԭ���س��������������15�����ֻ�30���ַ�!", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
