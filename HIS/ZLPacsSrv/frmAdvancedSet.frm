VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdvancedSet 
   Caption         =   "�߼�����"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   300
      Left            =   5520
      TabIndex        =   2
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   5880
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "�洢�߼�����"
      TabPicture(0)   =   "frmAdvancedSet.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmAutoRoutSet"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Worklist�߼�����"
      TabPicture(1)   =   "frmAdvancedSet.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdResetWLResult"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkForceResult"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkModel"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtDayInterval"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtDayInterval 
         Height          =   300
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   43
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkModel 
         Caption         =   "������豸����"
         Height          =   225
         Left            =   120
         TabIndex        =   42
         Top             =   525
         Width           =   1755
      End
      Begin VB.CheckBox chkForceResult 
         Caption         =   "ʹ��ǿ�ƽ��"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   900
         Width           =   1515
      End
      Begin VB.CommandButton cmdResetWLResult 
         Caption         =   "�ָ�Ĭ�Ͻ��"
         Height          =   350
         Left            =   2760
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame Frame8 
         Caption         =   "���������"
         Height          =   4215
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   8175
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   2055
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   3625
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
         Begin VB.Frame frmSetResult 
            Height          =   1575
            Left            =   120
            TabIndex        =   32
            Top             =   2520
            Width           =   7935
            Begin VB.CheckBox chkMWLItem 
               Caption         =   "ѡ��ʹ�øý��"
               Height          =   180
               Left            =   120
               TabIndex        =   45
               Top             =   0
               Width           =   1575
            End
            Begin VB.TextBox txtResult 
               Height          =   300
               Index           =   0
               Left            =   1200
               TabIndex        =   36
               Top             =   720
               Width           =   5775
            End
            Begin VB.CheckBox chkResult 
               Caption         =   "�Ƿ����"
               Height          =   255
               Left            =   6360
               TabIndex        =   35
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtResult 
               Height          =   300
               Index           =   1
               Left            =   1200
               TabIndex        =   34
               Top             =   1080
               Width           =   6135
            End
            Begin VB.CommandButton cmdBuildResult 
               Appearance      =   0  'Flat
               Caption         =   "��"
               Height          =   235
               Index           =   0
               Left            =   6990
               MaskColor       =   &H80000000&
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   765
               Width           =   315
            End
            Begin VB.Label lblResult 
               Caption         =   "�������"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   360
               Width           =   7215
            End
            Begin VB.Label Label11 
               Caption         =   "����ֵ"
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Top             =   743
               Width           =   735
            End
            Begin VB.Label Label12 
               Caption         =   "ǿ�ƽ��ֵ"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   1110
               Width           =   975
            End
         End
         Begin VB.CheckBox chkUseResult 
            Caption         =   "ѡ��ʹ�øý����"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   3000
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "��������"
         Height          =   855
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   8175
         Begin VB.ComboBox cboStoreDevice 
            Height          =   300
            ItemData        =   "frmAdvancedSet.frx":0038
            Left            =   1275
            List            =   "frmAdvancedSet.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox cboEncode 
            Height          =   300
            ItemData        =   "frmAdvancedSet.frx":0054
            Left            =   5040
            List            =   "frmAdvancedSet.frx":0061
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   360
            Width           =   2835
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "�洢�豸(&F)"
            Height          =   180
            Index           =   8
            Left            =   240
            TabIndex        =   29
            Top             =   405
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ѹ����ʽ(&Y)"
            Height          =   180
            Index           =   0
            Left            =   3960
            TabIndex        =   28
            Top             =   405
            Width           =   990
         End
      End
      Begin VB.Frame frmAutoRoutSet 
         Caption         =   "�Զ�·������"
         Height          =   2145
         Left            =   -74880
         TabIndex        =   14
         Top             =   3360
         Width           =   8175
         Begin VB.ComboBox cboDestination 
            Height          =   300
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1305
            Width           =   1605
         End
         Begin VB.ComboBox cboCondition 
            Height          =   300
            Index           =   1
            Left            =   6450
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   360
            Width           =   1365
         End
         Begin VB.ComboBox cboCondition 
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   6450
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   825
            Width           =   1365
         End
         Begin VB.OptionButton optType 
            Caption         =   "����豸(&R)"
            Height          =   255
            Index           =   2
            Left            =   5040
            TabIndex        =   19
            Top             =   855
            Width           =   1335
         End
         Begin VB.OptionButton optType 
            Caption         =   "Ӱ�����(&S)"
            Height          =   255
            Index           =   1
            Left            =   5040
            TabIndex        =   18
            Top             =   375
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelAutoRouting 
            Caption         =   "ɾ��(&D)"
            Height          =   350
            Left            =   6960
            TabIndex        =   17
            Top             =   1680
            Width           =   1100
         End
         Begin VB.CommandButton cmdModiAutoRouting 
            Caption         =   "�޸�(&M)"
            Height          =   350
            Left            =   5880
            TabIndex        =   16
            Top             =   1680
            Width           =   1100
         End
         Begin VB.CommandButton cmdAddAutoRouting 
            Caption         =   "���(&A)"
            Height          =   350
            Left            =   4800
            TabIndex        =   15
            Top             =   1680
            Width           =   1100
         End
         Begin MSFlexGridLib.MSFlexGrid MSFAutoRout 
            Height          =   1845
            Left            =   150
            TabIndex        =   23
            Top             =   270
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   3254
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ŀ���豸(&B)"
            Height          =   180
            Left            =   5070
            TabIndex        =   24
            Top             =   1365
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "�Զ�ƥ������"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   3
         Top             =   1440
         Width           =   8175
         Begin VB.Frame Frame4 
            Caption         =   "ͼ����Ŀ"
            Height          =   1455
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2685
            Begin VB.OptionButton optImgMatch 
               Caption         =   "Patient ID"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   13
               Top             =   360
               Width           =   1335
            End
            Begin VB.OptionButton optImgMatch 
               Caption         =   "Accession Number"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   12
               Top             =   720
               Width           =   1815
            End
            Begin VB.OptionButton optImgMatch 
               Caption         =   "Patient Name"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   11
               Top             =   1080
               Width           =   1455
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "���ݿ���Ŀ"
            Height          =   1455
            Left            =   2880
            TabIndex        =   6
            Top             =   240
            Width           =   2805
            Begin VB.OptionButton optMatch 
               Caption         =   "����"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   9
               ToolTipText     =   "�����Ž����˺ͽ��յ�Ӱ�����ƥ��"
               Top             =   360
               Width           =   1065
            End
            Begin VB.OptionButton optMatch 
               Caption         =   "���˱�ʶ�ţ�����/סԺ�ţ�"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   8
               ToolTipText     =   "�����˱�ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
               Top             =   720
               Width           =   2655
            End
            Begin VB.OptionButton optMatch 
               Caption         =   "����ʶ��"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   7
               ToolTipText     =   "������ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
               Top             =   1080
               Width           =   1335
            End
         End
         Begin VB.CheckBox chkMatchStudyUID 
            Caption         =   "���� ""���UID"" ƥ��"
            Height          =   350
            Left            =   5880
            TabIndex        =   5
            Top             =   600
            Width           =   2055
         End
         Begin VB.CheckBox chkImageType 
            Caption         =   "����ͼ�����Ͳ������"
            Height          =   350
            Left            =   5880
            TabIndex        =   4
            Top             =   1200
            Width           =   2175
         End
      End
      Begin VB.Label Label9 
         Caption         =   "�������        �������"
         Height          =   195
         Left            =   2730
         TabIndex        =   44
         Top             =   540
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmAdvancedSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlngServiceID As Long           '����ID
Private mstrServiceType As String   '��������

Private aDevices() As Variant       '�洢�豸�б�


Public Sub ShowMe(parent As Object, strServiceType As String, lngServiceID As Long)
    mlngServiceID = lngServiceID
    mstrServiceType = strServiceType
    If mstrServiceType = ZLPACS_�洢���� Then
        Me.SSTab1.TabVisible(0) = True
        Me.SSTab1.TabVisible(1) = False
    ElseIf UCase(mstrServiceType) = UCase(ZLPACS_Worklist����) Then
        Me.SSTab1.TabVisible(0) = False
        Me.SSTab1.TabVisible(1) = True
    End If
    
    Me.Show vbModal, parent
End Sub

Private Sub cmdAddAutoRouting_Click()
    Dim iType As Integer
    
    '��������Ƿ���Ч
    iType = IIf(optType(1).value = True, 1, 2)
    If cboDestination.Text = "" Then MsgBox "�������Զ�·�ɵ�Ŀ���豸��": Exit Sub
    If cboCondition(iType).Text = "" Then MsgBox IIf(iType = 1, "������Ӱ�����", "���������豸"): Exit Sub
    
    On Error GoTo errHand
    '��������
    gstrSQL = "Zl_Ӱ���Զ�·������_INSERT(" & mlngServiceID & ",'" & iType & "','" & cboCondition(iType).Text & "','" & _
                    GetDeviceNameNum(aDevices, cboDestination.Text, 1) & "')"
                        
    ExecuteProcedure "�����Զ�·������"
    'ˢ���б�
    Call subFillMSFAutoRouting(Me.MSFAutoRout.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelAutoRouting_Click()
    
    On Error GoTo errHand
    'ɾ������
    gstrSQL = "Zl_Ӱ���Զ�·������_DELETE(" & Me.MSFAutoRout.TextMatrix(Me.MSFAutoRout.RowSel, 3) & ")"
                        
    ExecuteProcedure "ɾ���Զ�·������"
    'ˢ���б�
    Call subFillMSFAutoRouting
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdModiAutoRouting_Click()
    Dim iType As Integer
    
    '��������Ƿ���Ч
    iType = IIf(optType(1).value = True, 1, 2)
    If cboDestination.Text = "" Then MsgBox "�������Զ�·�ɵ�Ŀ���豸��": Exit Sub
    If cboCondition(iType).Text = "" Then MsgBox IIf(iType = 1, "������Ӱ�����", "���������豸"): Exit Sub
    
    On Error GoTo errHand
    '�޸�����
    gstrSQL = "Zl_Ӱ���Զ�·������_UPDATE(" & Me.MSFAutoRout.TextMatrix(Me.MSFAutoRout.RowSel, 3) & "," & mlngServiceID & ",'" & iType & "','" & cboCondition(iType).Text & "','" & _
                    GetDeviceNameNum(aDevices, cboDestination.Text, 1) & "')"
                        
    ExecuteProcedure "�޸��Զ�·������"
    'ˢ���б�
    Call subFillMSFAutoRouting(Me.MSFAutoRout.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub CmdOK_Click()
    Call subSaveServiceParas
    If mstrServiceType = ZLPACS_�洢���� Then
        '�����Զ�ƥ������
        Call subSaveMatch
    End If
    Unload Me
End Sub

Private Sub subSaveServiceParas()
    '�����������
    Dim strValue As String
    If mstrServiceType = ZLPACS_�洢���� Then
        '�洢�豸��
        strValue = aDevices(0, cboStoreDevice.ListIndex)
        subSaveServicePara ZLPACS_�洢�豸��, strValue
        'ѹ����ʽ
        subSaveServicePara ZLPACS_ѹ����ʽ, cboEncode.ListIndex
        '���ü��UIDƥ��
        subSaveServicePara ZLPACS_���ü��UIDƥ��, chkMatchStudyUID.value
        '��ͼ�����Ͳ������
        subSaveServicePara ZLPACS_��ͼ�����Ͳ������, chkImageType.value
    Else
        '������豸����
        subSaveServicePara ZLPACS_MWL���豸����, chkModel.value
        '��������
        subSaveServicePara ZLPACS_MWL��������, txtDayInterval.Text
        'ʹ��ǿ�ƽ��
        subSaveServicePara ZLPACS_MWL��ǿ�ƽ��, chkForceResult.value
    End If
End Sub

Private Sub subSaveServicePara(strParaName As String, strParaValue As String)
    On Error GoTo errHand
    '��������
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngServiceID & ",'" & strParaName & "','" & strParaValue & "')"
                        
    ExecuteProcedure "����DICOM�������"
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub subSaveMatch()
    Dim intDBItem As Integer
    Dim intImageItem As Integer
    
    For intDBItem = 0 To optMatch.count - 1
        If optMatch(intDBItem).value Then Exit For
    Next
    If intDBItem > optMatch.count - 1 Then intDBItem = 0
    
    For intImageItem = 0 To optImgMatch.count - 1
        If optImgMatch(intImageItem).value Then Exit For
    Next
    If intImageItem > optImgMatch.count - 1 Then intImageItem = 0
    
    On Error GoTo errHand
    '��������
    gstrSQL = "Zl_Ӱ���Զ�ƥ������_SAVE(" & mlngServiceID & ",'" & intImageItem & "','" & intDBItem & "')"
                        
    ExecuteProcedure "�����Զ�ƥ������"
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Load()
    If mstrServiceType = ZLPACS_�洢���� Then
        '���� �洢����
        '���ش洢�豸
        Call subFillcboStoreDevice
        '���ػ�������
        Call subReadPara(1)
        '����ƥ�䷽ʽ
        Call subFillMatch
        '�����Զ�·��
        Call subFillAutoRoutDevice
        Call subFillMSFAutoRouting
    Else
        '����WORKLIST����
        '���ػ�������
        Call subReadPara(2)
        '���ؽ��������
        
    End If
End Sub

Private Sub subFillMSFAutoRouting(Optional iRow As Integer = 1)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRowPos As Long
    
    strSQL = "Select �Զ�·��ID,����ID,��������,����ֵ, Ŀ���豸�� From Ӱ���Զ�·������ Where ����ID =[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�Զ�·������", mlngServiceID)
    
    With MSFAutoRout
        .Clear
        .Rows = 1
        .Cols = 4
        .ColWidth(1) = 2500
        .ColWidth(3) = 0
        .TextMatrix(0, 0) = "��������"
        .TextMatrix(0, 1) = "��������"
        .TextMatrix(0, 2) = "Ŀ���豸"
        .TextMatrix(0, 3) = "ID"
        lngRowPos = 1
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(lngRowPos, 0) = IIf(rsTmp!�������� = 1, "Ӱ�����", "����豸")
            .TextMatrix(lngRowPos, 1) = rsTmp!����ֵ
            .TextMatrix(lngRowPos, 2) = GetDeviceNameNum(aDevices, rsTmp!Ŀ���豸��, 0)
            .TextMatrix(lngRowPos, 3) = rsTmp!�Զ�·��ID
            lngRowPos = .Rows
            rsTmp.MoveNext
        Wend
    End With
    
    Call subClickMSFAutoRouting(iRow)
End Sub

Private Sub subClickMSFAutoRouting(Optional iRow As Integer = 1)

    If iRow > Me.MSFAutoRout.Rows Then iRow = 1

    If Me.MSFAutoRout.Rows > 1 Then
        Me.MSFAutoRout.Row = iRow - 1
        Me.MSFAutoRout.RowSel = iRow
        Call MSFAutoRout_Click
    End If
End Sub

Private Function GetDeviceNameNum(aSource() As Variant, ByVal SeekString As String, iType As Integer) As String
    '��ȡ�豸�����ƻ��豸��
    'iType=0---����SeekStringΪ�豸�ţ������豸����
    'iType=1---����SeekStringΪ�豸���������豸�š�
    Dim i As Long
    For i = 0 To UBound(aSource, 2)
        If aSource(iType, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then GetDeviceNameNum = "": Exit Function
    GetDeviceNameNum = IIf(iType = 1, aSource(0, i), aSource(1, i))
End Function

Private Sub subFillMatch()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "Select ����ID,ͼ����,���ݿ��� From Ӱ���Զ�ƥ������ Where ����ID = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�Զ�ƥ������", mlngServiceID)
    
    If Not rsTmp.EOF Then
        optImgMatch(Val(rsTmp!ͼ����)).value = True
        optMatch(Val(rsTmp!���ݿ���)).value = True
    Else
        optImgMatch(0).value = True
        optMatch(0).value = True
    End If
End Sub

Private Sub subFillcboStoreDevice()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����= [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�洢�豸", 1)
    If rsTmp.EOF Then
        MsgBox "δ����Ӱ��洢�豸���뵽Ӱ���豸Ŀ¼�����ã�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    aDevices = rsTmp.GetRows
    rsTmp.MoveFirst
    
    Me.cboStoreDevice.Clear
    Do While Not rsTmp.EOF
        cboStoreDevice.AddItem Nvl(rsTmp(1))
        '����Զ�·�������е�Ŀ���豸�����б�
        cboDestination.AddItem Nvl(rsTmp(1))
        rsTmp.MoveNext
    Loop
End Sub

Private Sub subReadPara(intType As Integer)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "Select  a.�������ID,a.����ID,a.��������,a.����ֵ From Ӱ��dicom������� a Where a.����ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�������", mlngServiceID)
    
    If intType = 1 Then     '�洢����
        While Not rsTmp.EOF
            Select Case rsTmp!��������
            Case ZLPACS_�洢�豸��
                cboStoreDevice.ListIndex = GetComboxIndex(aDevices, Nvl(rsTmp!����ֵ))
                cboStoreDevice.Tag = 1
            Case ZLPACS_ѹ����ʽ
                cboEncode.ListIndex = Nvl(rsTmp!����ֵ, 0)
                cboEncode.Tag = 1
            Case ZLPACS_���ü��UIDƥ��
                chkMatchStudyUID.value = Nvl(rsTmp!����ֵ, 0)
                chkMatchStudyUID.Tag = 1
            Case ZLPACS_��ͼ�����Ͳ������
                chkImageType.value = Nvl(rsTmp!����ֵ, 0)
                chkImageType.Tag = 1
            End Select
            rsTmp.MoveNext
        Wend
        '����û�в������õ���Ŀ�����ó�Ĭ��ֵ
        If cboStoreDevice.Tag = "" Then cboStoreDevice.ListIndex = 0
        If cboEncode.Tag = "" Then cboEncode.ListIndex = 0
        If chkMatchStudyUID.Tag = "" Then chkMatchStudyUID.value = 0
        If chkImageType.Tag = "" Then chkImageType.value = 0
    ElseIf intType = 2 Then 'worklist����
        While Not rsTmp.EOF
            Select Case rsTmp!��������
            Case ZLPACS_MWL��������
                txtDayInterval.Text = Nvl(rsTmp!����ֵ, 3)
                txtDayInterval.Tag = 1
            Case ZLPACS_MWL���豸����
                chkModel.value = Nvl(rsTmp!����ֵ, 0)
                chkModel.Tag = 1
            Case ZLPACS_MWL��ǿ�ƽ��
                chkForceResult.value = Nvl(rsTmp!����ֵ, 0)
                chkForceResult.Tag = 1
            End Select
            rsTmp.MoveNext
        Wend
        '����û�в������õ���Ŀ�����ó�Ĭ��ֵ
        If txtDayInterval.Tag = "" Then txtDayInterval.Text = "3"
        If chkModel.Tag = "" Then chkModel.value = 0
        If chkForceResult.Tag = "" Then chkForceResult.value = 0
    End If
End Sub

Private Function GetComboxIndex(aSource() As Variant, ByVal SeekString As String) As Long
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetComboxIndex = i
End Function

Private Sub MSFAutoRout_Click()
    Dim iSelected As Integer
    
    With MSFAutoRout
        iSelected = .RowSel
        '��д��������
        Me.optType(IIf(.TextMatrix(iSelected, 0) = "Ӱ�����", 1, 2)).value = True
        '��д����ֵ
        Me.cboCondition(IIf(.TextMatrix(iSelected, 0) = "Ӱ�����", 1, 2)).Text = .TextMatrix(iSelected, 1)
        '��дĿ���豸��
        Me.cboDestination = .TextMatrix(iSelected, 2)
    End With
End Sub

Private Sub subFillAutoRoutDevice()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    '����Զ�·�������У�Ӱ����𣬺ͼ���豸�б�
    strSQL = "Select ���� From Ӱ�������"
    Set rsTmp = OpenSQLRecord(strSQL, "�߼�����")
    Do While Not rsTmp.EOF
        cboCondition(1).AddItem rsTmp(0)
        rsTmp.MoveNext
    Loop
    
    strSQL = "Select distinct ����豸 From Ӱ�����¼"
    Set rsTmp = OpenSQLRecord(strSQL, "�߼�����")
    Do While Not rsTmp.EOF
        cboCondition(2).AddItem Nvl(rsTmp(0))
        rsTmp.MoveNext
    Loop
End Sub

Private Sub optType_Click(Index As Integer)
    Me.cboCondition(Index).Enabled = True
    Me.cboCondition(IIf(Index = 1, 2, 1)).Enabled = False
End Sub

Private Sub txtDayInterval_KeyPress(KeyAscii As Integer)
    If (Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
