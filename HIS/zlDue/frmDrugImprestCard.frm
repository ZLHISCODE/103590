VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrugImprestCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ԥ���"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmDrugImprestCard.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.TreeView tvwProvider 
      Height          =   3585
      Left            =   1560
      TabIndex        =   23
      Top             =   1665
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   6324
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgTree"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   5400
      TabIndex        =   28
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   6120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugImprestCard.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugImprestCard.frx":1D18
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugImprestCard.frx":3A22
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
      Height          =   3375
      Left            =   0
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5953
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fraImprest 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton Cmd��Ӧ�� 
         Caption         =   "��"
         Height          =   300
         Left            =   4680
         TabIndex        =   27
         Top             =   1230
         Width           =   255
      End
      Begin VB.TextBox Txt��ҩ��λ 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   1230
         Width           =   3225
      End
      Begin VB.TextBox TxtNo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3405
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   10
         Top             =   810
         Width           =   1515
      End
      Begin VB.TextBox Txt����˵�� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   3
         Top             =   3960
         Width           =   3585
      End
      Begin VB.TextBox Txt�������� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   4380
         Width           =   1875
      End
      Begin VB.TextBox Txt������� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3045
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   4770
         Width           =   1875
      End
      Begin VB.TextBox Txt����� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4770
         Width           =   1005
      End
      Begin VB.TextBox Txt������ 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4380
         Width           =   1005
      End
      Begin ZL9BillEdit.BillEdit mshImprest 
         Height          =   1485
         Left            =   495
         TabIndex        =   2
         Top             =   2385
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2619
         Appearance      =   0
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
      Begin VB.Label txt˰��� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1440
         TabIndex        =   26
         Top             =   2070
         Width           =   3450
      End
      Begin VB.Label txt������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1440
         TabIndex        =   25
         Top             =   1830
         Width           =   3450
      End
      Begin VB.Label txt�绰��ַ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1440
         TabIndex        =   24
         Top             =   1560
         Width           =   3450
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ����֪ͨ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1410
         TabIndex        =   21
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
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
         Left            =   3075
         TabIndex        =   20
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Lbl��λ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   19
         Top             =   1290
         Width           =   810
      End
      Begin VB.Label Lbl�绰��ַ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ַ�绰:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   18
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Lbl������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   17
         Top             =   1830
         Width           =   810
      End
      Begin VB.Label Lbl����˵�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����˵��:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   495
         TabIndex        =   16
         Top             =   4020
         Width           =   810
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   2265
         TabIndex        =   15
         Top             =   4845
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   525
         TabIndex        =   14
         Top             =   4830
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2265
         TabIndex        =   13
         Top             =   4440
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   525
         TabIndex        =   12
         Top             =   4440
         Width           =   540
      End
      Begin VB.Label lbl˰��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "˰��ǼǺ�:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   2070
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmDrugImprestCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSuccess As Boolean
Private mstr���ݺ� As String
Private mblnSave As Boolean
Private mint�༭״̬ As Integer
Private mint��¼״̬ As Integer
Private mblnChange As Boolean
Private mfrmMain As Object
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Dim mstrPrivs As String                     'Ȩ��
Private Const mlngModule = 1323

Private Function GetDepend() As Boolean
'--------------------------------------------------------------
'���ܣ���ȡ��Ӧ�̡����㷽ʽ������
'������
'���أ�ȡ�������Ƿ�ɹ�
'˵����
'--------------------------------------------------------------
    Dim rsDepend As New Recordset
    Dim rs���㷽ʽ As New Recordset
    Dim intLop As Integer
    Dim strȨ�� As String
    GetDepend = False
    
    strȨ�� = " and (ĩ��<>1 or (ĩ��=1 " & zl_��ȡվ������ & " and " & Get����Ȩ��(gstrPrivs) & "))"
    
    gstrSQL = "" & _
        "   Select ID,�ϼ�ID,����,����,����,ĩ��,��ַ||�绰 as �绰��ַ,��������,�ʺ�,˰��ǼǺ� " & _
        "   From ��Ӧ�� " & _
        "   Where (����ʱ�� is null or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01') " & _
        "        " & strȨ�� & _
        "   Start with �ϼ�ID is Null Connect by prior ID=�ϼ�ID"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rsDepend, gstrSQL, "Ԥ����֪ͨ��"
    With rsDepend
        '��ȡ��Ӧ����Ϣ
        If .EOF Then
            '�޹�Ӧ����Ϣ�˳�
            ShowMsgbox "��Ӧ�̵���Ϣ��ȫ��"
            Exit Function
        End If
    End With
    'by lesfeng 2009-12-2 �����Ż�
    gstrSQL = "Select Ӧ�ó���,���㷽ʽ,ȱʡ��־ From ���㷽ʽӦ�� Where Ӧ�ó���='������' Order by ȱʡ��־ desc"
    zlDatabase.OpenRecordset rs���㷽ʽ, gstrSQL, "Ԥ����֪ͨ��"
    With rs���㷽ʽ
        '��ȡ���㷽ʽ
        If .EOF Then
            '�޽��㷽ʽ�����˳�
            ShowMsgbox "���㷽ʽӦ����Ϣ��ȫ��"
            Exit Function
        End If
        '�����㷽ʽ�б�
        mshImprest.Clear
        For intLop = 1 To .RecordCount
            mshImprest.AddItem !���㷽ʽ
            .MoveNext
        Next
        mshImprest.ListIndex = 0
        
        .Close
    End With
    
    With rsDepend
        '��乩Ӧ�����ݵ�TreeView��
        tvwProvider.Nodes.Clear
        tvwProvider.Nodes.Add , , "R", "���й�Ӧ��", 1, 1
        tvwProvider.Nodes("R").Tag = 0
        .MoveFirst
        
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                If !ĩ�� = 1 Then
                    tvwProvider.Nodes.Add "R", 4, "K_" & !ID, !����, 3, 3
                Else
                    tvwProvider.Nodes.Add "R", 4, "K_" & !ID, !����, 2, 2
                End If
            Else
                If !ĩ�� = 1 Then
                    tvwProvider.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !ID, !����, 3, 3
                Else
                    tvwProvider.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !ID, !����, 2, 2
                End If
            End If
            tvwProvider.Nodes("K_" & !ID).Tag = !ĩ��
            .MoveNext
        Loop
        tvwProvider.Nodes("R").Selected = True
        tvwProvider.Nodes("R").Expanded = True
        
    End With
    GetDepend = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
        Optional int��¼״̬ As Integer = 1, Optional blnSuccess As Boolean = False)
'--------------------------------------------------------------
'���ܣ�Ԥ����¼�뼰�༭
'������FrmMain---------���ô���
'      str���ݺ�-------Ԥ����ݺ�,����Ԥ����ʱΪ""
'      int�༭״̬-----1������Ԥ���2���༭��3����ˣ�4���鿴��ӡ
'      int��¼״̬-----1��������¼��2����������¼��3��������¼
'      BlnSuccess------�����Ƿ�ɹ�����
'���أ�
'˵����
'--------------------------------------------------------------
    mblnSave = False
    mblnSuccess = False
    Set mfrmMain = frmMain
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1323)
    
    If Not GetDepend Then Exit Sub
    
    If mint�༭״̬ = 1 Then            '����Ԥ����
        mstr���ݺ� = NextNo(31)
        TxtNo = mstr���ݺ�
        
    ElseIf mint�༭״̬ = 2 Then        '�༭Ԥ����
'        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then        '���Ԥ����
        'mblnEdit = False
        cmdOk.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then        '�鿴����ӡ
        'mblnEdit = False
        cmdOk.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "Ԥ����֪ͨ����ӡ") = 0 Then
            cmdOk.Visible = False
        Else
            cmdOk.Visible = True
        End If
    End If
    
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cmdCancel_Click()
    '�˳�
    Unload Me
End Sub

Private Function ValidData() As Boolean
'--------------------------------------------------------------
'���ܣ��������������Ƿ����Ҫ��
'������
'���أ�True----���ϣ�False----������
'˵����
'--------------------------------------------------------------
    Dim intRow As Integer
    
    ValidData = False
    '��鹩Ӧ��
    If Txt��ҩ��λ.Text = "" Then
         ShowMsgbox "�Բ���û�й�ҩ��λ!"
         Txt��ҩ��λ.SetFocus
         Exit Function
    End If
    If IIf(Txt��ҩ��λ.Tag = "", 0, Txt��ҩ��λ.Tag) = 0 Then
        ShowMsgbox "�Բ���û����ȷѡ��ҩ��λ��������ѡ��!"
         Txt��ҩ��λ.SetFocus
         Exit Function
    End If
    
    With mshImprest
        '���ÿ�����ݵ�������
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                If IIf(.TextMatrix(intRow, 1) = "", 0, .TextMatrix(intRow, 1)) = 0 And intRow <> .Rows - 1 Then
                    ShowMsgbox "�Բ��𣬽��������룬�Ҳ�Ϊ��"
                    .SetFocus
                    .Row = intRow
                    .MsfObj.TopRow = intRow
                    .Col = 1
                    Exit Function
                End If
                If Val(.TextMatrix(intRow, 1)) > 9999999999999# Then
                    ShowMsgbox "��" & intRow & "�н����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                    .SetFocus
                    .Row = intRow
                    .MsfObj.TopRow = intRow
                    .Col = 1
                    Exit Function
                End If
                If Val(.TextMatrix(intRow, 1)) < -1 * 9999999999999# Then
                    ShowMsgbox "��" & intRow & "�н����������ݿ��ܹ������" & vbCrLf & "���Χ-1*9999999999999�����飡"
                    .SetFocus
                    .Row = intRow
                    .MsfObj.TopRow = intRow
                    .Col = 1
                    Exit Function
                End If
            End If
        Next
    End With
    '��鸶��˵��
    If LenB(StrConv(Txt����˵��.Text, vbFromUnicode)) > 50 Then
        ShowMsgbox "����˵���ĳ��ȳ���!(���Ϊ50���ַ���25������)"
        Txt����˵��.SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Function SaveCard() As Boolean
'--------------------------------------------------------------
'���ܣ�����¼���Ӧ��������
'������
'���أ������Ƿ�ɹ�
'˵����
'--------------------------------------------------------------
    Dim intRow As Integer
    Dim strNO_IN As String
    Dim int���_IN As Integer
    Dim intԤ����_IN As Integer
    Dim lng��λID_IN As Long
    Dim dbl���_IN As Double
    Dim str���㷽ʽ_IN As String
    Dim str�������_IN As String
    Dim str������_IN As String
    Dim str��������_IN As String
    Dim lng�������_IN As Long
    Dim strժҪ_IN As String
    
    SaveCard = False
    '׼������
    strNO_IN = TxtNo
    intԤ����_IN = 1
    lng��λID_IN = Txt��ҩ��λ.Tag
    str������_IN = UserInfo.����
    str��������_IN = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    strժҪ_IN = Txt����˵��
    
    
    On Error GoTo errHandle:
    
    '��ʼ����
    gcnOracle.BeginTrans
    
    If mint�༭״̬ = 2 Then
            gstrSQL = "zl_�����¼_DELETE('" & TxtNo & "')"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End If
        
    'ѭ������ÿ������
    With mshImprest
        'zl_�������_INSERT( /*strNO_IN*/, /*int���_IN*/, /*intԤ����_IN*/, /*lng��λID_IN*/,
            '/*dbl���_IN*/, /*str���㷽ʽ_IN*/, /*str�������_IN*/, /*str������_IN*/, /*str��������_IN*/,
            '/*lng�������_IN*/, /*strժҪ_IN*/ );
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And IIf(.TextMatrix(intRow, 1) = "", 0, .TextMatrix(intRow, 1)) <> 0 Then
                int���_IN = intRow
                dbl���_IN = .TextMatrix(intRow, 1)
                str���㷽ʽ_IN = .TextMatrix(intRow, 0)
                str�������_IN = .TextMatrix(intRow, 2)
                
                gstrSQL = "zl_�������_INSERT('" & strNO_IN & "'," & int���_IN & "," & intԤ����_IN & "," & lng��λID_IN _
                    & "," & dbl���_IN & ",'" & str���㷽ʽ_IN & "','" & str�������_IN & "','" & str������_IN & "',to_date('" _
                    & str��������_IN & "','yyyy-mm-dd HH24:MI:SS'),NULL,'" & strժҪ_IN & "')"
               zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With
    
    '�ύ����
    gcnOracle.CommitTrans
    SaveCard = True
    Exit Function

errHandle:
    gcnOracle.RollbackTrans

    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub cmdHelp_Click()
    '����
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    'ȷ��
    Dim blnSuccess As Boolean
    Dim strReg As String
    Select Case mint�༭״̬
        Case 1, 2           '������༭
            With mshImprest
                If .TextMatrix(1, 0) = "" Then Exit Sub
                If Not ValidData() Then Exit Sub        '���¼�����ݲ����Ϲ淶���˳�ģ��
                blnSuccess = SaveCard                   '��������
                If blnSuccess = False Then Exit Sub     '����ʧ�����˳�ģ��
                mblnChange = False
                mblnSave = False
                mblnSuccess = True
                If IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                     '��ӡ
                    If InStr(mstrPrivs, "Ԥ����֪ͨ����ӡ") <> 0 Then
                        ReportOpen gcnOracle, glngSys, "zl1_bill_1323_2", Me, "���ݱ��=" & TxtNo.Text, "��¼״̬=" & mint��¼״̬, 2
                    End If
                End If
                
                If mint�༭״̬ = 1 Then                '����ʱ���¿�ʼ��һ�����ݵ�����
                    .ClearBill
                    TxtNo = NextNo(31)
                    Txt��ҩ��λ.Text = ""
                    Txt��ҩ��λ.Tag = 0
                    txt�绰��ַ = ""
                    Txt����˵�� = ""
                    txt������ = ""
                    txt˰��� = ""
                    Txt��ҩ��λ.SetFocus
                Else                                    '�༭ʱ�˳�
                    Unload Me
                End If
                mblnChange = False
                Exit Sub
            End With
        Case 3                                          '���
            With mshImprest
                If .TextMatrix(1, 0) = "" Then Exit Sub
                If Not ValidData() Then Exit Sub        '�������
                blnSuccess = SaveVerify                 '���
                If blnSuccess = False Then Exit Sub
                
                If IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                     '��ӡ
                    If InStr(mstrPrivs, "Ԥ����֪ͨ����ӡ") <> 0 Then
                        ReportOpen gcnOracle, glngSys, "zl1_bill_1323_2", Me, "���ݱ��=" & TxtNo.Text, "��¼״̬=" & mint��¼״̬, 2
                    End If
                End If
                
                mblnChange = False
                mblnSave = False
                mblnSuccess = True
                Unload Me
                Exit Sub
            End With
        Case 4
            '��ӡ
            ReportOpen gcnOracle, glngSys, "zl1_bill_1323_2", Me, "���ݱ��=" & TxtNo.Text, "��¼״̬=" & mint��¼״̬, 2
    End Select
End Sub

Private Function SaveVerify() As Boolean
'--------------------------------------------------------------
'���ܣ����Ԥ��������
'������
'���أ�����Ƿ�ɹ�
'˵����
'--------------------------------------------------------------
    Dim intRow As Integer
    Dim strNO_IN As String
    Dim dbl������_IN As Double
    Dim lng��λID_IN As Long
    Dim str�����_IN As String
    
    SaveVerify = False
    '׼������
    strNO_IN = TxtNo
    lng��λID_IN = Txt��ҩ��λ.Tag
    str�����_IN = UserInfo.����
    dbl������_IN = 0
    On Error GoTo errHandle:
    
    With mshImprest
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And IIf(.TextMatrix(intRow, 1) = "", 0, .TextMatrix(intRow, 1)) <> 0 Then
                dbl������_IN = dbl������_IN + Val(.TextMatrix(intRow, 1))
            End If
        Next
    End With
    '��д�������
    
    'zl_�������_VERIFY( /*strNO_IN*/, /*lng��λID_IN*/, /*dbl������_IN*/, /*str�����_IN*/ );
    gstrSQL = "zl_�������_VERIFY('" & strNO_IN & "'" & "" & "" & "" _
        & "" & "" & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    
    
    SaveVerify = True
    Exit Function
errHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            '�����ѱ�ɾ��
            ShowMsgbox "�õ����ѱ�ɾ�������飡"
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            ShowMsgbox "�õ����ѱ���������ˣ����飡"
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim intLop As Integer
    
    TxtNo = mstr���ݺ�
    '��ʼ������
    With mshImprest
        .Clear
        .Cols = 3
        .Rows = 2
        
        .TextMatrix(0, 0) = "���ʽ"
        .TextMatrix(0, 1) = "������"
        .TextMatrix(0, 2) = "�������"
        
        If Not RestoreFlexState(mshImprest, Me.Caption) Then
            .ColWidth(0) = 1000
            .ColWidth(1) = 1500
            .ColWidth(2) = 1800
        End If
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignLeftCenter
        
        .ColData(0) = 3
        .ColData(1) = 4
        .ColData(2) = 4
        
        .PrimaryCol = 0
        
        If mint�༭״̬ > 2 Then
            .Active = False
            Txt��ҩ��λ.Enabled = False
            Cmd��Ӧ��.Enabled = False
            Txt����˵��.Enabled = False
            If mint�༭״̬ = 3 Then
                cmdOk.Caption = "���(&V)"
            Else
                cmdOk.Caption = "��ӡ(&P)"
            End If
        Else
            .Active = True
        End If
    End With
    
    On Error GoTo errHandle
    If mint�༭״̬ = 1 Then                '����ʱ��д
        Txt������ = UserInfo.����
        Txt�������� = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    Else
        '�����ݿ��л�ȡ����
        Dim rsImprest As New Recordset
        Dim intRecord As Integer
        
        gstrSQL = "" & _
            "   Select  a.���,a.���,a.���㷽ʽ,a.�������,a.ժҪ,a.������,a.��������,a.�����," & _
            "           a.�������,b.����,��ַ || �绰 as �绰��ַ,��������,˰��ǼǺ�,b.id " & _
            "   From �����¼ a,��Ӧ�� b " & _
            "   Where a.��λid=b.id " & _
            "       and no=[1] and ��¼״̬=[2]"
        
        Set rsImprest = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr���ݺ�, mint��¼״̬)
        
        If rsImprest.EOF Then
            '������
            mintParallelRecord = 2
            Exit Sub
        End If
        
        '��д���ݵ�����
        intRecord = rsImprest.RecordCount
        Txt��ҩ��λ.Text = rsImprest!����
        Txt��ҩ��λ.Tag = rsImprest!ID
        txt�绰��ַ = IIf(IsNull(rsImprest!�绰��ַ), "", rsImprest!�绰��ַ)
        txt������ = IIf(IsNull(rsImprest!��������), "", rsImprest!��������)
        txt˰��� = IIf(IsNull(rsImprest!˰��ǼǺ�), "", rsImprest!˰��ǼǺ�)
        Txt����˵��.Text = IIf(IsNull(rsImprest!ժҪ), "", rsImprest!ժҪ)
        Txt������ = rsImprest!������
        If mint�༭״̬ = 2 Then
            Txt������ = UserInfo.����
        End If
        Txt�������� = Format(rsImprest!��������, "yyyy-mm-dd hh:mm:ss")
        Txt����� = IIf(IsNull(rsImprest!�����), "", rsImprest!�����)
        Txt������� = IIf(IsNull(rsImprest!�������), "", Format(rsImprest!�������, "yyyy-mm-dd hh:mm:ss"))
        
        If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
            mintParallelRecord = 3
            Exit Sub
        End If
        
        With mshImprest
            For intLop = 1 To intRecord
                .TextMatrix(intLop, 0) = rsImprest!���㷽ʽ
                .TextMatrix(intLop, 1) = GetFormat(rsImprest!���, 2)
                .TextMatrix(intLop, 2) = IIf(IsNull(rsImprest!�������), "", rsImprest!�������)
                If intLop = .Rows - 1 Then .Rows = .Rows + 1
                rsImprest.MoveNext
            Next
        End With
                
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        Txt��ҩ��λ.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If tvwProvider.Visible = True Then
        tvwProvider.Visible = False
        Txt��ҩ��λ.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        
        SaveFlexState mshImprest, Me.Caption
        Exit Sub
    End If
    Dim blnYes As Boolean
    
    ShowMsgbox "���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", True, blnYes
    If blnYes = False Then
        Cancel = 1
        Exit Sub
    End If
    SaveFlexState mshImprest, Me.Caption
End Sub

Private Sub mshImprest_cboClick(ListIndex As Long)
        mshImprest.TextMatrix(mshImprest.Row, mshImprest.Col) = mshImprest.CboText
End Sub

Private Sub mshImprest_EditChange(curText As String)
    With mshImprest
        If .Col <> 0 Then
            .Text = UCase(curText)
            .SelStart = Len(curText)
        End If
    End With
    mblnChange = True
End Sub

Private Sub mshImprest_EnterCell(Row As Long, Col As Long)
    With mshImprest
    Select Case Col
        Case 1
            .TxtCheck = True
            .MaxLength = 16
            .TextMask = ".1234567890-"
        Case 2
            .TxtCheck = True
            .MaxLength = 10
            .ColData(Col) = 4
    End Select
    End With
End Sub

Private Sub mshImprest_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> 13 Then Exit Sub
    
    With mshImprest
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case 1
                If .Row = .Rows - 1 And KeyCode = vbKeyReturn And strKey = "" Then
                    Txt����˵��.SetFocus
                    Cancel = True
                    Exit Sub
                End If
                
                
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    ShowMsgbox "Ԥ�����������룡"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    ShowMsgbox "Ԥ��������Ϊ������,�����䣡"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        ShowMsgbox "Ԥ������Ϊ��,�����䣡"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) >= 10 ^ 12 - 1 Then
                        ShowMsgbox "Ԥ��������С��" & (10 ^ 12 - 1)
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) < -10 ^ 12 - 1 Then
                        ShowMsgbox "Ԥ�����������" & (-1 * 10 ^ 12 - 1)
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = GetFormat(strKey, 2)
                End If
                    
            Case 2
               
                If KeyCode <> vbKeyReturn Then
                    .ColData(2) = 4
                    .TxtCheck = False
                Else
                    .ColData(2) = 0
                    .TxtCheck = True
                    .TextLen = 10
                End If
                
        End Select
    End With
End Sub

Private Sub mshProvider_DblClick()
    mshProvider_KeyPress 13
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intTemp As Integer
    Dim sngWidth As Single
    
    With mshProvider
        If KeyCode = vbKeyRight Then
            If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
                
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If .LeftCol <> 0 Then
                .LeftCol = .LeftCol - 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyHome Then
            If .LeftCol <> 0 Then
                .LeftCol = 0
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyEnd Then
            For intTemp = .Cols - 1 To 0 Step -1
                sngWidth = sngWidth + .ColWidth(intTemp)
                If sngWidth > .Width Then
                    .LeftCol = intTemp + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub mshProvider_KeyPress(KeyAscii As Integer)
    With mshProvider
        If KeyAscii = 13 Then
            Txt��ҩ��λ.Text = .TextMatrix(.Row, 2)
            Txt��ҩ��λ.Tag = .TextMatrix(.Row, 0)
            
            txt�绰��ַ = IIf(IsNull(.TextMatrix(.Row, 4)), "", .TextMatrix(.Row, 4))
            
            txt������ = IIf(IsNull(.TextMatrix(.Row, 5)), "", .TextMatrix(.Row, 5))
            txt˰��� = IIf(IsNull(.TextMatrix(.Row, 7)), "", .TextMatrix(.Row, 7))
            
            .Visible = False
            
            mshImprest.SetFocus
        End If
    End With
End Sub

Private Sub mshProvider_LostFocus()
    SaveFlexState mshProvider, Me.Caption
    If mshProvider.Visible Then mshProvider.Visible = False
End Sub
'���ù�Ӧ��ѡ�����Ŀ�ȼ��������
Private Sub SetProviderWidth(ByVal sngLeft As Single, ByVal sngTop As Single)
    
    With mshProvider
        .Visible = True
        .Left = sngLeft
        .Top = sngTop
'        If RestoreFlexState(mshProvider, Me.Caption) = False Then
            'Select ID,����,����,����,��ַ||�绰 as �绰��ַ,��������,�ʺ�,˰��ǼǺ�
            
            .ColWidth(0) = 0
            .ColWidth(1) = 1000
            .ColWidth(2) = 2500
            .ColWidth(3) = 1000
            
            .ColWidth(4) = 1500
            .ColWidth(5) = 1500
            .ColWidth(6) = 1000
            .ColWidth(7) = 1000
            
'        End If
        
        .SetFocus
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub tvwProvider_DblClick()
    Dim rsProvider As New Recordset
    
    If tvwProvider.SelectedItem.Children <> 0 Then Exit Sub
    If tvwProvider.SelectedItem.Tag = 0 Then Exit Sub
    
    Txt��ҩ��λ = tvwProvider.SelectedItem
    Txt��ҩ��λ.Tag = Mid(tvwProvider.SelectedItem.Key, 3)
    tvwProvider.Tag = "1"
    tvwProvider.Visible = False
    On Error GoTo errHandle
    gstrSQL = "" & _
        "   Select ����,��ַ||�绰 as �绰��ַ,��������,�ʺ�,˰��ǼǺ� " & _
        "   From ��Ӧ��  " & _
        "   Where id=[1]"
    Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Txt��ҩ��λ.Tag))
    
    With rsProvider
        
        If .EOF Then Exit Sub
        
        Txt��ҩ��λ = !����
        txt�绰��ַ = IIf(IsNull(!�绰��ַ), "", !�绰��ַ)
        txt������ = IIf(IsNull(!��������), "", !��������)
        txt˰��� = IIf(IsNull(!˰��ǼǺ�), "", !˰��ǼǺ�)
        mshImprest.SetFocus
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Txt����˵��_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt����˵��, KeyAscii, m�ı�ʽ
End Sub

Private Sub Txt��ҩ��λ_GotFocus()
    tvwProvider.Visible = False
End Sub

Private Sub txt��ҩ��λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(Txt��ҩ��λ)) = "" Then Exit Sub
    Dim strKey As String
    
    Dim rs��Ӧ�� As New ADODB.Recordset
    
    Dim strȨ�� As String
    strȨ�� = " and " & Get����Ȩ��(gstrPrivs)
    strKey = GetMatchingSting(Txt��ҩ��λ.Text, False)
    gstrSQL = "" & _
        "   Select ID,����,����,����,��ַ||�绰 as �绰��ַ,��������,�ʺ�,˰��ǼǺ�  " & _
        "   From ��Ӧ�� " & _
        "   Where " & zl_��ȡվ������(False) & "   " & _
        "       and (���� like upper([1]) Or ���� like [1] Or ���� like upper([1])) " & _
        "       And (����ʱ�� is null or ����ʱ��>=to_date('3000-01-01','yyyy-mm-dd'))  And ĩ��=1 " & strȨ��
    On Error GoTo errHandle
    Set rs��Ӧ�� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKey)
    
    With rs��Ӧ��
    
        If .EOF Then
            ShowMsgbox "�����ڸ������Ĺ�Ӧ�̣�"
            KeyCode = 0
            Txt��ҩ��λ = ""
            tvwProvider.Tag = "0"
            Exit Sub
        End If
        If .RecordCount > 1 Then
            Set mshProvider.Recordset = rs��Ӧ��
            SetProviderWidth Txt��ҩ��λ.Left + fraImprest.Left, Txt��ҩ��λ.Top + Txt��ҩ��λ.Height + fraImprest.Top
            Exit Sub
        Else
            Txt��ҩ��λ = !����
            Txt��ҩ��λ.Tag = !ID
            tvwProvider.Tag = "1"
        End If
    End With
    
    Txt��ҩ��λ = rs��Ӧ��!����
    txt�绰��ַ = IIf(IsNull(rs��Ӧ��!�绰��ַ), "", rs��Ӧ��!�绰��ַ)
    txt������ = IIf(IsNull(rs��Ӧ��!��������), "", rs��Ӧ��!��������)
    txt˰��� = IIf(IsNull(rs��Ӧ��!˰��ǼǺ�), "", rs��Ӧ��!˰��ǼǺ�)
    zlCommFun.PressKey (vbKeyTab)
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Cmd��Ӧ��_Click()
    tvwProvider.Visible = tvwProvider.Visible Xor True
    If tvwProvider.Visible Then
        tvwProvider.Top = Txt��ҩ��λ.Top + Txt��ҩ��λ.Height + fraImprest.Top
        tvwProvider.SetFocus
    End If
End Sub

Private Sub txt����˵��_Change()
    mblnChange = True
End Sub

Private Sub txt����˵��_GotFocus()
    zlCommFun.OpenIme (True)
    With Txt����˵��
        .SelStart = 0
        .SelLength = Len(Txt����˵��.Text)
    End With
End Sub

Private Sub txt����˵��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt����˵��_LostFocus()
    zlCommFun.OpenIme (False)
End Sub

Private Sub Txt��ҩ��λ_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt��ҩ��λ, KeyAscii, m�ı�ʽ
End Sub
