VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm��Ʊ�ݺ�������ҩ 
   Caption         =   "��Ʊ�ݺŷ�ҩ"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   Icon            =   "Frm��Ʊ�ݺ�������ҩ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9705
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtҽ���� 
      Height          =   300
      Left            =   2520
      TabIndex        =   13
      Top             =   180
      Width           =   1725
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   7
      Top             =   5790
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   2580
      TabIndex        =   6
      Top             =   5790
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrintSet 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   1350
      TabIndex        =   5
      Top             =   5790
      Visible         =   0   'False
      Width           =   1100
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   3165
      Left            =   30
      TabIndex        =   2
      Top             =   2400
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5583
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "������ϸ(&D)"
      TabPicture(0)   =   "Frm��Ʊ�ݺ�������ҩ.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Msf������ϸ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ҩƷ����(&T)"
      TabPicture(1)   =   "Frm��Ʊ�ݺ�������ҩ.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf��������"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf������ϸ 
         Height          =   2745
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4842
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�������� 
         Height          =   2745
         Left            =   -74940
         TabIndex        =   9
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4842
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.TextBox TxtNo 
      Height          =   300
      Left            =   660
      TabIndex        =   0
      Top             =   180
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8220
      TabIndex        =   4
      Top             =   5790
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6930
      TabIndex        =   3
      Top             =   5790
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�����б� 
      Height          =   1755
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3096
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblҽ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1920
      TabIndex        =   12
      Top             =   240
      Width           =   540
   End
   Begin VB.Label LblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "δ�����κδ���"
      ForeColor       =   &H80000002&
      Height          =   180
      Left            =   4440
      TabIndex        =   11
      Top             =   240
      Width           =   3630
   End
   Begin VB.Label LblNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ʊ�ݺ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "Frm��Ʊ�ݺ�������ҩ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--�ⲿ���ݲ���--
Private mblnModify As Boolean
Private strUnit As String
Private strPrivs As String
Private mint������� As Integer                     'ҩ���ķ������1-���ﲡ��;2-סԺ����;3-�����סԺ
Private lngҩ��ID As Long                           'ҩ��
Private IntSendAfterDosage As Integer               '����δ��ҩ��ҩ
Private Int����δ��˴�����ҩ As Integer            '����δ��˴�����ҩ
Private mint����δ�շѴ�����ҩ As Integer           '����δ�շѴ�����ҩ
Private IntCheckStock As Integer                    '�����
Private IntУ�鴦�� As Integer                      'У�鴦��
Private Str���� As String                           '��ҩ����
Private int����λ�� As Integer                  '���ý���λ��
Private int��˻��۵� As Integer                    'ִ�к��Զ���˻��۵�
Private mint�����ʾ As Integer                     '�����ʾ��ʽ��0-��ʾӦ�ս��,1-��ʾʵ�ս��,2-��ʾӦ�պ�ʵ�ս��
Private mstrOpr As String
Private mblnConPacker As Boolean
Private mblnLoadDrug As Boolean
Private mbln������ As Boolean

'--������ʹ�ñ���--
Private RecBill As New ADODB.Recordset              '���ݼ�¼
Private RecTotal As New ADODB.Recordset             '��������
Private BlnStartUp As Boolean
Private LngListRow As Long                          '�����б�
Private LngDetailRow As Long                        '������ϸ
Private LngTotalRow As Long                         '��������
Private StrBillNo As String                         '���ܵ��ݺ�
Private strID As String                             '����ID

Private LngBillCount As Long
Public str��ҩ�� As String

Private rs��� As ADODB.Recordset
Private mobjDrugMAC As Object
Private mobjPlugIn As Object             '��ҽӿڶ���
Private mstrDeptNode As String

Public Property Get In_DrugMAC() As Object
    Set In_DrugMAC = mobjDrugMAC
End Property
Public Property Set In_DrugMAC(ByVal objVal As Object)
    Set mobjDrugMAC = objVal
End Property
Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Public Property Get In_���÷�ҩ() As Boolean
    In_���÷�ҩ = mblnLoadDrug
End Property

Public Property Let In_���÷�ҩ(ByVal vNewValue As Boolean)
    mblnLoadDrug = vNewValue
End Property

Public Property Get In_�Զ���ҩ() As Boolean
    In_�Զ���ҩ = mblnConPacker
End Property

Public Property Let In_�Զ���ҩ(ByVal vNewValue As Boolean)
    mblnConPacker = vNewValue
End Property
Private Sub GetRecipe(ByVal intType As Integer, ByVal txtInput As TextBox)
    'intType��1��Ʊ�ݺţ�2��ҽ����
    Dim blnAdd As Boolean
    Dim strNo As String, IntBill As Integer
    Dim rstemp As New ADODB.Recordset
    Dim strInput As String
    Dim strsql As String
    
    If Trim(txtInput.Text) = "" Then Exit Sub
    strInput = Trim(UCase(txtInput.Text))
    
    If intType = 1 Then
        '���������Ʊ�ݺ���ȡ����
        gstrSQL = "Select Distinct A.No " & _
                 " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B " & _
                 " Where A.ID=B.��ӡID And A.��������=1 " & _
                 " And B.Ʊ��=1 And B.����=[1]"
    Else
        '���������ҽ������ȡ����
        gstrSQL = "Select Distinct B.NO " & _
                " From ������Ϣ A, δ��ҩƷ��¼ B " & _
                " Where A.����id = B.����id And B.���� = 8 And A.ҽ���� = [1] And B.�ⷿid = [2]"
    End If
    On Error GoTo errHandle
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[���������Ʊ�ݺ���ȡ����]", strInput, lngҩ��ID)
    
    If rstemp.RecordCount = 0 Then
        MsgBox "û���ҵ��κδ�����", vbInformation, gstrSysName
        GoTo ExitSub
        Exit Sub
    End If
    
    With rstemp
        Do While Not .EOF
            gstrSQL = " Select /*+ Rule*/ Distinct Decode(C.����,8,'�շ�',9,'����') ����,C.No,C.����,A.���շ�,Decode(A.��ҩ��,Null,'','���ŷ�ҩ','',A.��ҩ��) ��ҩ��,P.���� ����,B.����,B.��ʶ�� סԺ��,'' ����," & _
                " B.������ ����ҽ��,B.����Ա���� ������,To_Char(C.��������,'yyyy-MM-dd') ��������,B.��¼����,B.�����־, d.�������� " & _
                " From δ��ҩƷ��¼ A,������ü�¼ B,ҩƷ�շ���¼ C,���ű� P,���ű� S, ������Ϣ D " & IIf(Str���� = "", "", ",Table(Cast(f_Str2list([3]) As zlTools.t_Strlist)) E ") & IIf(mbln������, ",��������¼ Q,���������ϸ K ", "") & _
                " Where C.����ID=B.ID And B.��������ID+0=P.ID And Nvl(C.�ⷿID,0)+0=S.ID and Nvl(A.�ⷿID,0)=Nvl(C.�ⷿID,0) And Mod(C.��¼״̬,3)=1 And A.No=C.No " & IIf(mbln������, " and b.ҽ�����=k.ҽ��id(+) and Q.id(+)=K.��id and K.����ύ(+)=1 And (b.ҽ����� is null or nvl(q.�����,0) = 1)", "") & _
                " And (C.�ⷿID+0=[2] OR C.�ⷿID IS NULL)" & IIf(Str���� = "", "", " And (C.��ҩ����=E.Column_Value Or C.��ҩ���� Is NULL)") & _
                " and Not Exists(select 1 from ҩƷ�շ���¼ F where F.����=C.���� and F.�ⷿid=C.�ⷿid and F.no=C.no and ��ҩ��ʽ=-1) " & _
                " And C.���� =8 And C.����� Is Null And C.����=A.���� And C.No=[1] and nvl(C.��ҩ��ʽ,-999)<>-1 And A.����id=D.����id(+) "     '����һ���������ų��ѱ��Ϊ����ҩ�ļ�¼  by lyq 20050416
            
            If mstrDeptNode <> "" Then
                gstrSQL = gstrSQL & " And (P.վ�� = [4] Or P.վ�� Is Null)"
            End If
            
            If mint������� = 3 Then
                strsql = Replace(gstrSQL, "'' ����", "B.����")
                strsql = Replace(strsql, "������ü�¼", "סԺ���ü�¼")
                gstrSQL = gstrSQL & " Union All " & strsql
            ElseIf mint������� = 2 Then
                gstrSQL = Replace(gstrSQL, "'' ����", "B.����")
                gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
            End If
            On Error GoTo errHandle
            Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(!NO), lngҩ��ID, Str����, mstrDeptNode)
            
            blnAdd = (RecBill.RecordCount <> 0)
            If blnAdd Then     '�ҵ�ָ������
                strNo = RecBill!NO
                IntBill = RecBill!����
                txtInput.Tag = IntBill
                
                '����Ѵ��ڸõ��ݣ����˳�
                blnAdd = Not SetLocateBill(strNo, False)
                
                '���Ϸ���
                If blnAdd Then blnAdd = Not (CheckBill(IntBill, strNo, Val(RecBill!��¼����), Val(RecBill!�����־)) <> 0)
                If blnAdd Then blnAdd = WriteSendListData()
                If blnAdd Then
                    LngBillCount = LngBillCount + 1
                    LblNote.Caption = IIf(LngBillCount = 0, "δ�����κδ���", "������" & LngBillCount & "�Ŵ���")
                End If
            End If
            .MoveNext
        Loop
    End With
    
    '��λ���ղ�����Ĵ�����
    Call SetLocateBill(strNo, True)
    With Msf�����б�
        CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
    End With
    
    mblnModify = True
    If TabShow.Tab = 1 Then Call RefreshData
    Exit Sub
ExitSub:
    With txtInput
        .SelStart = 0
        .SelLength = Len(txtInput.Text)
        .SetFocus
    End With
Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Property Get In_����() As String
    In_���� = mstrOpr
End Property

Public Property Let In_����(ByVal vNewValue As String)
    mstrOpr = vNewValue
End Property

Public Property Get In_Ȩ��() As String
    In_Ȩ�� = strPrivs
End Property

Public Property Let In_Ȩ��(ByVal vNewValue As String)
    strPrivs = vNewValue
End Property

Public Property Get In_�������() As Integer
    In_������� = mint�������
End Property

Public Property Let In_�������(ByVal vNewValue As Integer)
    mint������� = vNewValue
End Property

Public Property Get In_У�鴦��() As Integer
    In_У�鴦�� = IntУ�鴦��
End Property

Public Property Let In_У�鴦��(ByVal vNewValue As Integer)
    IntУ�鴦�� = vNewValue
End Property

Public Property Get In_�����() As Integer
    In_����� = IntCheckStock
End Property

Public Property Let In_�����(ByVal vNewValue As Integer)
    IntCheckStock = vNewValue
End Property

Public Property Get In_ҩ��ID() As Long
    In_ҩ��ID = lngҩ��ID
End Property

Public Property Let In_ҩ��ID(ByVal vNewValue As Long)
    lngҩ��ID = vNewValue
    mstrDeptNode = GetDeptStationNode(lngҩ��ID)
End Property

Public Property Get In_��ҩ����() As String
    In_��ҩ���� = Str����
End Property

Public Property Let In_��ҩ����(ByVal vNewValue As String)
    Str���� = vNewValue
End Property

Public Property Get In_����δ��ҩ��ҩ() As Integer
    In_����δ��ҩ��ҩ = IntSendAfterDosage
End Property

Public Property Let In_����δ��ҩ��ҩ(ByVal vNewValue As Integer)
    IntSendAfterDosage = vNewValue
End Property

Public Property Get IN_����δ��˷�ҩ() As Integer
    IN_����δ��˷�ҩ = Int����δ��˴�����ҩ
End Property

Public Property Let IN_����δ��˷�ҩ(ByVal vNewValue As Integer)
    Int����δ��˴�����ҩ = vNewValue
End Property

Public Property Get IN_����δ�շѷ�ҩ() As Integer
    IN_����δ�շѷ�ҩ = mint����δ�շѴ�����ҩ
End Property

Public Property Let IN_����δ�շѷ�ҩ(ByVal vNewValue As Integer)
    mint����δ�շѴ�����ҩ = vNewValue
End Property

Public Property Get In_����λ��() As Integer
    In_����λ�� = int����λ��
End Property

Public Property Let In_����λ��(ByVal vNewValue As Integer)
    int����λ�� = vNewValue
End Property

Public Property Get IN_��˻��۵�() As Integer
    IN_��˻��۵� = int��˻��۵�
End Property

Public Property Let IN_��˻��۵�(ByVal vNewValue As Integer)
    int��˻��۵� = vNewValue
End Property

Private Sub SetFormat(Optional ByVal IntStyle As Integer = 1)
    Dim intCol As Integer
    '���ø��б�ؼ��ĸ�ʽ

    Select Case IntStyle
    Case 1
        With Msf�����б�
            .rows = 2
            .Cols = 11
    
            .TextMatrix(0, 0) = "����"
            .TextMatrix(0, 1) = "NO"
            .TextMatrix(0, 2) = "����"
            .TextMatrix(0, 3) = "����"
            .TextMatrix(0, 4) = "סԺ��"
            .TextMatrix(0, 5) = "����"
            .TextMatrix(0, 6) = "�շ�Ա"
            .TextMatrix(0, 7) = "����ҽ��"
            .TextMatrix(0, 8) = "��������"
            .TextMatrix(0, 9) = "��¼����"
            .TextMatrix(0, 10) = "�����־"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(0) = 500
                .ColWidth(1) = 1000
                .ColWidth(2) = 1200
                .ColWidth(3) = 1000
                .ColWidth(4) = 1000
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 1000
                .ColWidth(8) = 1200
                .ColWidth(9) = 0
                .ColWidth(10) = 0
                
                .Row = 1
                Call RestoreFlexState(Msf�����б�, Me.Name)
                If glngSys \ 100 <> 1 Then
                    .ColWidth(2) = 0
                    .ColWidth(4) = 0
                    .ColWidth(5) = 0
                End If
                .ColWidth(7) = IIf(IntУ�鴦�� = 1, 0, 1000)
            End If
        End With
    Case 2
        With Msf������ϸ
            .rows = 2
            .Cols = 8
    
            .TextMatrix(0, 0) = "ҩƷ����"
            .TextMatrix(0, 1) = "��Ʒ��"
            .TextMatrix(0, 2) = "���"
            .TextMatrix(0, 3) = "��λ"
            .TextMatrix(0, 4) = "����"
            .TextMatrix(0, 5) = "����"
            .TextMatrix(0, 6) = "Ӧ�ս��"
            .TextMatrix(0, 7) = "ʵ�ս��"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
    
            If BlnStartUp = False Then
                .ColWidth(0) = 2000
                .ColWidth(2) = 1500
                .ColWidth(3) = 500
                .ColWidth(4) = 800
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 1000
                
                .Row = 1
                Call RestoreFlexState(Msf������ϸ, Me.Name)
                If gintҩƷ������ʾ = 2 Then
                    If .ColWidth(1) = 0 Then .ColWidth(1) = 2000
                Else
                    .ColWidth(1) = 0
                End If
                
                If mint�����ʾ = 0 Then
                    .ColWidth(7) = 0
                    If .ColWidth(6) <= 0 Then .ColWidth(6) = 1000
                ElseIf mint�����ʾ = 1 Then
                    .ColWidth(6) = 0
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                Else
                    If .ColWidth(6) <= 0 Then .ColWidth(6) = 1000
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                End If
            End If
        End With
    Case 3
        With Msf��������
            .rows = 2
            .Cols = 9
    
            .TextMatrix(0, 0) = "���"
            .TextMatrix(0, 1) = "ҩƷ����"
            .TextMatrix(0, 2) = "��Ʒ��"
            .TextMatrix(0, 3) = "���"
            .TextMatrix(0, 4) = "��λ"
            .TextMatrix(0, 5) = "����"
            .TextMatrix(0, 6) = "����"
            .TextMatrix(0, 7) = "Ӧ�ս��"
            .TextMatrix(0, 8) = "ʵ�ս��"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(0) = 500
                .ColWidth(1) = 2000
                .ColWidth(3) = 1500
                .ColWidth(4) = 500
                .ColWidth(5) = 800
                .ColWidth(6) = 800
                .ColWidth(7) = 1000
                .ColWidth(8) = 1000
                
                .Row = 1
                Call RestoreFlexState(Msf��������, Me.Name)
                If gintҩƷ������ʾ = 2 Then
                    If .ColWidth(2) = 0 Then .ColWidth(2) = 2000
                Else
                    .ColWidth(2) = 0
                End If
                
                If mint�����ʾ = 0 Then
                    .ColWidth(8) = 0
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                ElseIf mint�����ʾ = 1 Then
                    .ColWidth(7) = 0
                    If .ColWidth(8) <= 0 Then .ColWidth(8) = 1000
                Else
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                    If .ColWidth(8) <= 0 Then .ColWidth(8) = 1000
                End If
            End If
        End With
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    '���õ���ǩ��ʱ����û��Ƿ�ע��
    If gblnESign������ҩ = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Sub
        End If
    End If
    
    Call RefreshData
    If CheckStock = False Then Exit Sub
    If Not CheckCorrelation Then Exit Sub
    If Not CheckBillOperate Then Exit Sub
    If SendBill = False Then Exit Sub
    
    LngBillCount = 0
    LblNote.Caption = IIf(LngBillCount = 0, "δ�����κδ���", "������" & LngBillCount & "�Ŵ���")
    
    '��ʼ��
    strID = ""
    StrBillNo = ""
    TxtNo.Text = ""
    txtҽ����.Text = ""
    
    With Msf��������
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    With Msf�����б�
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    With Msf������ϸ
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    CmdOK.Enabled = False
    TxtNo.SetFocus
End Sub

Private Sub cmdPrint_Click()
    Dim HisPrint As New zlPrint1Grd
    Dim HisRow As New zlTabAppRow
    Dim ArrayNo, IntArray As Integer
    Dim LngSelectRow As Long, intCol As Integer
    
    On Error Resume Next
    'ȡ������ѡ��״̬
    With Msf��������
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngTotalRow > 0 And LngTotalRow < .rows Then
            .Row = LngTotalRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
    End With
    
    HisPrint.Title = "ҩƷ����"
    Set HisRow = New zlTabAppRow
    HisRow.Add "����:" & Format(zldatabase.Currentdate, "yyyy��MM��dd��")
    HisPrint.UnderAppRows.Add HisRow
    
    ArrayNo = Split(StrBillNo, ";")
    
    Set HisRow = New zlTabAppRow
    HisRow.Add "���ݺ�:"
    HisPrint.BelowAppRows.Add HisRow
    For IntArray = 0 To UBound(ArrayNo)
        Set HisRow = New zlTabAppRow
        HisRow.Add Space(10) & ArrayNo(IntArray)
        HisPrint.BelowAppRows.Add HisRow
    Next
    
    Set HisPrint.Body = Msf��������
    Select Case zlPrintAsk(HisPrint)
    Case 1
        zlPrintOrView1Grd HisPrint, 1
    Case 2
        zlPrintOrView1Grd HisPrint, 2
    Case 3
        zlPrintOrView1Grd HisPrint, 3
    End Select
    
    '�ָ�����ѡ��״̬
    With Msf��������
        
        LngTotalRow = LngSelectRow
        .Row = LngTotalRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub CmdPrintSet_Click()
    zlPrintSet
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    BlnStartUp = False
    LngBillCount = 0
    
    mint�����ʾ = Val(zldatabase.GetPara("�����ʾ��ʽ", glngSys, 1341, 0))
    mbln������ = ((gtype_UserSysParms.P240_ҩ��������� = 1 Or gtype_UserSysParms.P240_ҩ��������� = 3) And gtype_UserSysParms.P241_�������ʱ�� = 2)
    
    strID = ""
    StrBillNo = ""
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
   
    BlnStartUp = True
End Sub

Private Function CheckBillOperate() As Boolean
    Dim n, i As Integer
    Dim Dbl��� As Double
    
    For n = 1 To Msf�����б�.rows - 1
        If Msf�����б�.TextMatrix(n, 1) <> "" Then
            Msf�����б�.Row = n
            Call Msf�����б�_EnterCell
            DoEvents
            
            Dbl��� = 0
            
            For i = 1 To Msf������ϸ.rows - 2
                Dbl��� = Dbl��� + Val(Msf������ϸ.TextMatrix(i, 7))
            Next
            
            If CheckBillControl(3, Val(Msf�����б�.RowData(n)), Msf�����б�.TextMatrix(n, 1), Dbl���) = False Then
                Exit Function
            End If
        End If
    Next
    
    CheckBillOperate = True
End Function
Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 8505 Then Me.Width = 8505
    If Me.Height < 6165 Then Me.Height = 6165
    
    With LblNote
        .Left = Me.ScaleWidth - .Width - 100
    End With
    
    With CmdHelp
        .Top = Me.ScaleHeight - .Height - 100
    End With
    With CmdPrintSet
        .Top = CmdHelp.Top
        .Left = CmdHelp.Left + CmdHelp.Width + 100
    End With
    With CmdPrint
        .Top = CmdHelp.Top
        .Left = CmdPrintSet.Left + CmdPrintSet.Width + 100
    End With
    
    With CmdCancel
        .Top = CmdHelp.Top
        .Left = Me.ScaleWidth - .Width - 100
    End With
    With CmdOK
        .Top = CmdHelp.Top
        .Left = CmdCancel.Left - .Width - 100
    End With
    
    With Msf�����б�
        .Height = (CmdOK.Top - 200 - .Top) / 2
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With TabShow
        .Top = Msf�����б�.Top + Msf�����б�.Height + 100
        .Height = CmdOK.Top - 100 - .Top
        .Width = Msf�����б�.Width
    End With
    With Msf��������
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
    With Msf������ϸ
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(Msf��������, Me.Name)
    Call SaveFlexState(Msf�����б�, Me.Name)
    Call SaveFlexState(Msf������ϸ, Me.Name)
End Sub

Private Sub Msf��������_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf��������
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngTotalRow > 0 And LngTotalRow < .rows Then
            .Row = LngTotalRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngTotalRow = LngSelectRow
        .Row = LngTotalRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf��������_GotFocus()
    With Msf��������
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf��������_LostFocus()
    With Msf��������
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf�����б�_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf�����б�
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngListRow > 0 And LngListRow < .rows Then
            .Row = LngListRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H80000005
                If intCol <> 3 Then
                    .CellForeColor = &H80000008
                End If
            Next
            .Col = 0
        End If
        
        LngListRow = LngSelectRow
        .Row = LngListRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellBackColor = &H8000000D
            If intCol <> 3 Then
                .CellForeColor = &H80000005
            End If
        Next
        .Col = 0
        .Redraw = True
        
        If Trim(.TextMatrix(.Row, 1)) = "" Then
            With Msf������ϸ
                .Clear
                .rows = 2
                Call SetFormat(2)
            End With
            Exit Sub
        End If
        
        '��ʾ������ϸ
        Call ReadBillData(.RowData(.Row), .TextMatrix(.Row, 1), Val(.TextMatrix(.Row, 9)), Val(.TextMatrix(.Row, 10)))
    End With
End Sub

Private Sub Msf�����б�_GotFocus()
    With Msf�����б�
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf�����б�_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNo As String, lng���� As Long
    
    If KeyCode = vbKeyDelete Then
        With Msf�����б�
            lng���� = Val(.TextMatrix(.Row, 0))
            strNo = .TextMatrix(.Row, 1)
            If .rows - 1 = 1 Then
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
                .TextMatrix(1, 6) = ""
                .TextMatrix(1, 7) = ""
                .TextMatrix(1, 8) = ""
                .TextMatrix(1, 9) = ""
                .TextMatrix(1, 10) = ""
                .RowData(1) = 0
            Else
                If Trim(.TextMatrix(.Row, 1)) <> "" Then .RemoveItem .Row: LngBillCount = LngBillCount - 1
            End If
            
            CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
            LblNote.Caption = IIf(LngBillCount = 0, "δ�����κδ���", "������" & LngBillCount & "�Ŵ���")
        
            'ɾ���õ���
            With rs���
                If .RecordCount <> 0 Then .MoveFirst
                .Find "���ݱ�ʶ='" & strNo & "|" & lng���� & "'"
                If Not .EOF Then .Delete
            End With
        End With
        
        Msf�����б�_EnterCell
        mblnModify = True
        If TabShow.Tab = 1 Then Call RefreshData
    End If
End Sub

Private Sub Msf�����б�_LostFocus()
    With Msf�����б�
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf������ϸ_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf������ϸ
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngDetailRow > 0 And LngDetailRow < .rows Then
            .Row = LngDetailRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngDetailRow = LngSelectRow
        .Row = LngDetailRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf������ϸ_GotFocus()
    With Msf������ϸ
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf������ϸ_LostFocus()
    With Msf������ϸ
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case TabShow.Tab
    Case 0
        Msf������ϸ.ZOrder
        Msf������ϸ_EnterCell
    Case 1
        Call RefreshData
        Msf��������.ZOrder
        Msf��������_EnterCell
    End Select
End Sub

Private Sub TxtNo_GotFocus()
    GetFocus TxtNo
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call GetRecipe(1, TxtNo)
End Sub

Private Function ReadData(ByVal StrQuery As String) As Boolean
    '--��ȡ����--

'    On Error Resume Next
'    err = 0
    ReadData = False
    On Error GoTo errHandle
    gstrSQL = StrQuery
    With RecBill
        If .State = 1 Then .Close

        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, "ReadData")
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
    End With

    If err <> 0 Then
        MsgBox "��ȡ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    ReadData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadBillData(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer) As Boolean
    Dim IntStyle As Integer
    Dim str��� As String
    Dim str��ϸ��λ�� As String
    '--��ȡ��������--
    'BillStyle-��������;BIllNO-���ݺ�
    '��λ��ʾ���ݷ����������������ﵥλ��סԺ��סԺ���סԺ��λ���������ۼ۵�λ��
    On Error Resume Next
    err = 0
    ReadBillData = False
    
    strUnit = GetUnit(lngҩ��ID, BillStyle, BillNo, int�����־)
    Select Case strUnit
    Case "�ۼ۵�λ"
        str��ϸ��λ�� = "C.���㵥λ ��λ,B.���ۼ� ����,B.ʵ������*Nvl(B.����,1) ����"
    Case "���ﵥλ"
        str��ϸ��λ�� = "D.���ﵥλ ��λ,B.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ) ����,B.ʵ������/Decode(D.�����װ,Null,1,0,1,D.�����װ)*Nvl(B.����,1) ����"
    Case "סԺ��λ"
        str��ϸ��λ�� = "D.סԺ��λ ��λ,B.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ) ����,B.ʵ������/Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ)*Nvl(B.����,1) ����"
    Case "ҩ�ⵥλ"
        str��ϸ��λ�� = "D.ҩ�ⵥλ ��λ,B.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ) ����,B.ʵ������/Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ)*Nvl(B.����,1) ����"
    End Select
    str��ϸ��λ�� = str��ϸ��λ�� & ",B.���۽�� ���,Nvl(B.����, 1) * B.ʵ������ / (Nvl(F.����, 1) * F.����) * F.ʵ�ս�� As ʵ�ս�� "
    
    gstrSQL = " SELECT DISTINCT F.���,'['||C.����||']'|| C.���� As Ʒ��,A.���� As ��Ʒ��, " & _
            " DECODE(C.���,NULL,B.����,DECODE(B.����,NULL,C.���,C.���||'|'||B.����)) ���," & _
            str��ϸ��λ�� & _
            " FROM ҩƷ�շ���¼ B,ҩƷ��� D,�շ���ĿĿ¼ C,�շ���Ŀ���� A,������ü�¼ F" & _
            " WHERE B.ҩƷID=D.ҩƷID AND D.ҩƷID=C.ID" & _
            " AND d.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 And B.����ID=F.ID" & _
            " AND MOD(B.��¼״̬,3)=1 AND B.NO=[1] AND B.����=[2] " & _
            " AND (B.�ⷿID+0=[3] OR B.�ⷿID IS NULL) " & _
            " And ����� Is Null And Nvl(F.����״̬,0)<>1 " & _
            " Order by F.���"
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        gstrSQL = Replace(gstrSQL, "And Nvl(F.����״̬,0)<>1", "")
    End If
    On Error GoTo errHandle
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, BillNo, BillStyle, lngҩ��ID)
        
    With RecBill
        str��� = ""
        Do While Not .EOF
            str��� = str��� & "," & !���
            .MoveNext
        Loop
        If str��� <> "" Then str��� = Mid(str���, 2)
        .MoveFirst
    End With
    
    '��������Ϣ����ϸ���д���ڲ�ӳ���¼����
    With rs���
        If .RecordCount <> 0 Then .MoveFirst
        .Find "���ݱ�ʶ='" & BillNo & "|" & BillStyle & "'"
        If str��� <> "" Then
            If .EOF Then
                .AddNew
                !���ݱ�ʶ = BillNo & "|" & BillStyle
                !��� = str���
                !��¼���� = int��¼����
                !�����־ = int�����־
                .Update
            End If
        End If
    End With
    
    If WriteDataToBill() = False Then Exit Function

    If err <> 0 Then
        MsgBox "��ȡ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    ReadBillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBill(ByVal IntBillStyle As Integer, ByVal strNo As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer) As Integer
    Dim RecCheck As New ADODB.Recordset

    '--���ݽ�Ҫִ�еĲ������ж��Ƿ�����--
    '����:
    '0-�������
    '1-δ��ҩ
    '2-����ҩ
    '3-�ѷ�ҩ
    '4-��ɾ��
    '5-δ��ҩ
    On Error GoTo errHandle
    gstrSQL = " Select A.��ҩ��,A.�����,nvl(B.���շ�,0) ���շ�, C.����Ա���� ������ " & _
            " From ҩƷ�շ���¼ A,δ��ҩƷ��¼ B, ������ü�¼ C " & _
            " Where A.No=B.No And A.����=B.���� And A.����id = C.ID And A.����� IS Null And mod(A.��¼״̬,3)=1 And Rownum=1 " & _
            " And A.No=[1] And A.����=[2]  And (A.�ⷿID+0=[3] Or A.�ⷿID Is NULL)"
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngҩ��ID)
        
    With RecCheck
        If .EOF Then CheckBill = 4: MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!�����) Then
            CheckBill = 3: MsgBox "�ô����ѱ���������Ա��ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
        End If
        If !���շ� = 0 Then
            CheckBill = 3: MsgBox "�ô�����δ�շѣ�����������ֹ��", vbInformation, gstrSysName: Exit Function
        End If
    End With
     
    If mint����δ�շѴ�����ҩ = 0 And IntBillStyle = 8 Then
        If RecCheck!���շ� = 0 Then
            MsgBox "�ô�����δ�շѣ�����������ֹ��", vbInformation, gstrSysName
            CheckBill = 5
            Exit Function
        End If
    End If
    
    If Int����δ��˴�����ҩ = 0 And IntBillStyle <> 8 Then
        If IsNull(RecCheck!������) Then
            MsgBox "�ô�����δ��ˣ�����������ֹ��", vbInformation, gstrSysName
            CheckBill = 5
            Exit Function
        End If
    End If

    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteSendListData() As Boolean
    Dim RecCheck As New ADODB.Recordset
    
    WriteSendListData = False
    
    If IntSendAfterDosage = 0 Then
        If IsNull(RecBill!��ҩ��) Then
            MsgBox "����" & RecBill!NO & "��δ��ҩ����������뷢ҩ�б�", vbInformation, gstrSysName
            Exit Function
        End If
        If Trim(RecBill!��ҩ��) = "" Then
            MsgBox "����" & RecBill!NO & "��δ��ҩ����������뷢ҩ�б�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mint����δ�շѴ�����ҩ = 0 And RecBill!���� = 8 Then
        If RecBill!���շ� = 0 Then
            MsgBox "����" & RecBill!NO & "��δ�շѣ���������뷢ҩ�б�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If Int����δ��˴�����ҩ = 0 And RecBill!���� <> 8 Then
        If IsNull(RecBill!������) Then
            MsgBox "����" & RecBill!NO & "��δ��ˣ���������뷢ҩ�б�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With Msf�����б�
        .Redraw = False
        .TextMatrix(.rows - 1, 0) = RecBill!����
        .TextMatrix(.rows - 1, 1) = RecBill!NO
        .TextMatrix(.rows - 1, 2) = IIf(IsNull(RecBill!����), "", RecBill!����)
        .TextMatrix(.rows - 1, 3) = IIf(IsNull(RecBill!����), "", RecBill!����)
        .TextMatrix(.rows - 1, 4) = IIf(IsNull(RecBill!סԺ��), "", RecBill!סԺ��)
        .TextMatrix(.rows - 1, 5) = IIf(IsNull(RecBill!����), "", RecBill!����)
        .TextMatrix(.rows - 1, 6) = IIf(IsNull(RecBill!������), "", RecBill!������)
        .TextMatrix(.rows - 1, 7) = IIf(IsNull(RecBill!����ҽ��), "", RecBill!����ҽ��)
        .TextMatrix(.rows - 1, 8) = IIf(IsNull(RecBill!��������), "", RecBill!��������)
        .TextMatrix(.rows - 1, 9) = RecBill!��¼����
        .TextMatrix(.rows - 1, 10) = RecBill!�����־
        .RowData(.rows - 1) = RecBill!����
        
        .Row = .rows - 1
        .Col = 3
        .CellForeColor = zldatabase.GetPatiColor(IIf(IsNull(RecBill!��������), "", RecBill!��������))

        .rows = .rows + 1
        .RowData(.rows - 1) = 0
        .Redraw = True
    End With
    WriteSendListData = True
End Function

Private Function RefreshData() As Boolean
    Dim intRow As Integer, intRows As Integer
    Dim arrID
    Dim StrNoThis As String, IntBillThis As Integer
    Dim str���ܵ�λ�� As String
    If mblnModify = False Then Exit Function
    RefreshData = False
    
    '��ջ��ܱ��
    On Error GoTo errHandle
    With Msf��������
        .Clear
        .rows = 2
        SetFormat (3)
    End With
    
    strID = ""
    StrBillNo = ""
    With Msf�����б�
    
        '���NO��
        For intRow = 1 To .rows - 1
            If intRow = 1 Then
                StrBillNo = StrBillNo & .TextMatrix(intRow, 1)
            Else
                If Trim(.TextMatrix(intRow, 1)) <> "" Then
                    If intRow Mod 8 = 0 Then StrBillNo = StrBillNo & ";"
                    StrBillNo = StrBillNo & "," & .TextMatrix(intRow, 1)
                End If
            End If
        Next
        
        '���ID
        For intRow = 1 To .rows - 1
            StrNoThis = .TextMatrix(intRow, 1)
            IntBillThis = .RowData(intRow)
            
            gstrSQL = " Select ID From ҩƷ�շ���¼ Where No=[1] And ����=[2] " & _
                " And Mod(��¼״̬,3)=1 And ����� Is Null And (�ⷿID+0=[3] Or �ⷿID Is NULL)"
            Set RecTotal = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, StrNoThis, IntBillThis, lngҩ��ID)
            
            With RecTotal
                Do While Not .EOF
                    strID = strID & IIf(strID = "", "", ",") & !Id
                    .MoveNext
                Loop
            End With
        Next
    End With
    If strID = "" Then Exit Function
    
    '��ʾ��������
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("ҩ������", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(lngҩ��ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(lngҩ��ID, gint����ҩ��)
    Else
        strUnit = GetSpecUnit(lngҩ��ID, gintסԺҩ��)
    End If
    Select Case strUnit
    Case "�ۼ۵�λ"
        str���ܵ�λ�� = "C.���㵥λ ��λ,B.���ۼ� ����,Sum(B.ʵ������*Nvl(B.����,1)) ����"
    Case "���ﵥλ"
        str���ܵ�λ�� = "D.���ﵥλ ��λ,B.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ) ����,Sum(B.ʵ������/Decode(D.�����װ,Null,1,0,1,D.�����װ)*Nvl(B.����,1)) ����"
    Case "סԺ��λ"
        str���ܵ�λ�� = "D.סԺ��λ ��λ,B.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ) ����,Sum(B.ʵ������/Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ)*Nvl(B.����,1)) ����"
    Case "ҩ�ⵥλ"
        str���ܵ�λ�� = "D.ҩ�ⵥλ ��λ,B.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ) ����,Sum(B.ʵ������/Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ)*Nvl(B.����,1)) ����"
    End Select
    str���ܵ�λ�� = str���ܵ�λ�� & ",Sum(B.���۽��) ���,Sum(Nvl(B.����, 1) * B.ʵ������ / (Nvl(B.���ø���, 1) * B.����) * B.ʵ�ս��) As ʵ�ս��  "
    
    gstrSQL = " Select A.No, A.ҩƷid, A.����, A.���ۼ�, A.ʵ������, A.����, A.���۽��, B.���� As ���ø���,B.����, B.ʵ�ս��, A.���� " & _
        " From ҩƷ�շ���¼ A, ������ü�¼ B , Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) C " & _
        " Where A.����id = B.Id And A.Id =C.Column_Value "
    
    gstrSQL = "Select Distinct D.*,'['||D.����||']'|| D.ͨ������ As Ʒ��,A.���� As ��Ʒ�� " & _
             " From " & _
             "     (SELECT D.ҩƷID,C.����,C.���� ͨ������,NVL(B.����,0) ����," & _
             "     DECODE(C.���,NULL,B.����,DECODE(B.����,NULL,C.���,C.���||'|'||B.����)) ���," & str���ܵ�λ�� & _
             "     FROM (" & gstrSQL & ") B," & _
             "           ҩƷ��� D,�շ���ĿĿ¼ C " & _
             "     WHERE B.ҩƷID+0=D.ҩƷID AND D.ҩƷID=C.ID" & _
             "     GROUP BY D.ҩƷID,C.����,C.����,NVL(B.����,0)," & _
             "     DECODE(C.���,NULL,B.����,DECODE(B.����,NULL,C.���,C.���||'|'||B.����)),"
    Select Case strUnit
    Case "�ۼ۵�λ"
        gstrSQL = gstrSQL & "C.���㵥λ,B.���ۼ�"
    Case "���ﵥλ"
        gstrSQL = gstrSQL & "D.���ﵥλ,B.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ)"
    Case "סԺ��λ"
        gstrSQL = gstrSQL & "D.סԺ��λ,B.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ)"
    Case "ҩ�ⵥλ"
        gstrSQL = gstrSQL & "D.ҩ�ⵥλ,B.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ)"
    End Select
    gstrSQL = gstrSQL & ") D,�շ���Ŀ���� A" & _
            " Where D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3"
    gstrSQL = gstrSQL & " Order By D.����"
    
    Set RecTotal = zldatabase.OpenSQLRecord(gstrSQL, "RefreshData", strID)
    
    Call WriteTotalDataToBill
    
    If err <> 0 Then
        MsgBox "��ʾ��������ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnModify = False
    RefreshData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteTotalDataToBill() As Boolean
    Dim dblӦ�ս�� As Double
    Dim dblʵ�ս�� As Double
    Dim str�����ʾ As String
    
    '����������װ��
    On Error Resume Next
    err = 0
    
    WriteTotalDataToBill = False
    With Msf��������
        .Clear
        .rows = 2
        Call SetFormat(3)
    End With
    
    '��䵥������
    With RecTotal
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Msf��������.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Msf��������.TextMatrix(.AbsolutePosition, 1) = !Ʒ��
            Msf��������.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
            Msf��������.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!���), "", !���)
            Msf��������.TextMatrix(.AbsolutePosition, 4) = IIf(IsNull(!��λ), "", !��λ)
            Msf��������.TextMatrix(.AbsolutePosition, 5) = Format(!����, "#####0.00000;-#####0.00000; ;")
            Msf��������.TextMatrix(.AbsolutePosition, 6) = Format(!����, "#####0.00000;-#####0.00000; ;")
            Msf��������.TextMatrix(.AbsolutePosition, 7) = Format(!���, "#####0.00;-#####0.00; ;")
            Msf��������.TextMatrix(.AbsolutePosition, 8) = Format(!ʵ�ս��, "#####0.00;-#####0.00; ;")
            Msf��������.MergeRow(.AbsolutePosition) = False
            dblӦ�ս�� = dblӦ�ս�� + !���
            dblʵ�ս�� = dblʵ�ս�� + !ʵ�ս��
            
            If .AbsolutePosition >= Msf��������.rows - 1 Then Msf��������.rows = Msf��������.rows + 1
            .MoveNext
        Loop
        
        '��ʾ�ϼ�
        Msf��������.TextMatrix(Msf��������.rows - 1, 0) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.rows - 1, 1) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.rows - 1, 2) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.rows - 1, 3) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.rows - 1, 4) = "�ϼ�"
        
        If mint�����ʾ = 1 Then
            str�����ʾ = "ʵ�ս�" & Format(dblʵ�ս��, "#####0.00;-#####0.00; ;")
        ElseIf mint�����ʾ = 2 Then
            str�����ʾ = "Ӧ�ս�" & Format(dblӦ�ս��, "#####0.00;-#####0.00; ;") & "    ʵ�ս�" & Format(dblʵ�ս��, "#####0.00;-#####0.00; ;")
        Else
            str�����ʾ = "Ӧ�ս�" & Format(dblӦ�ս��, "#####0.00;-#####0.00; ;")
        End If
        
        Msf��������.TextMatrix(Msf��������.rows - 1, 5) = str�����ʾ
        Msf��������.TextMatrix(Msf��������.rows - 1, 6) = str�����ʾ
        Msf��������.TextMatrix(Msf��������.rows - 1, 7) = str�����ʾ
        Msf��������.TextMatrix(Msf��������.rows - 1, 8) = str�����ʾ
        Msf��������.MergeCells = flexMergeFree
        Msf��������.MergeRow(Msf��������.rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "��ʾ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    WriteTotalDataToBill = True
End Function

Private Function WriteDataToBill() As Boolean
    Dim dblӦ�ս�� As Double
    Dim dblʵ�ս�� As Double
    Dim str�����ʾ As String
    
    '--��ʾָ����������ϸ--
    On Error Resume Next
    err = 0
    
    WriteDataToBill = False
    With Msf������ϸ
        .Clear
        .rows = 2
        Call SetFormat(2)
    End With
    
    '��䵥������
    With RecBill
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Msf������ϸ.MergeRow(.AbsolutePosition) = False
            Msf������ϸ.TextMatrix(.AbsolutePosition, 0) = !Ʒ��
            Msf������ϸ.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!���), "", !���)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!��λ), "", !��λ)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 4) = Format(!����, "#####0.00000;-#####0.00000; ;")
            Msf������ϸ.TextMatrix(.AbsolutePosition, 5) = Format(!����, "#####0.00000;-#####0.00000; ;")
            Msf������ϸ.TextMatrix(.AbsolutePosition, 6) = Format(!���, "#####0.00;-#####0.00; ;")
            Msf������ϸ.TextMatrix(.AbsolutePosition, 7) = Format(!ʵ�ս��, "#####0.00;-#####0.00; ;")
            dblӦ�ս�� = dblӦ�ս�� + Val(!���)
            dblʵ�ս�� = dblʵ�ս�� + Val(!ʵ�ս��)
            
            If .AbsolutePosition >= Msf������ϸ.rows - 1 Then Msf������ϸ.rows = Msf������ϸ.rows + 1
            .MoveNext
        Loop
    End With
    With Msf������ϸ
        .TextMatrix(.rows - 1, 0) = "�ϼ�"
        .TextMatrix(.rows - 1, 1) = "�ϼ�"
        .TextMatrix(.rows - 1, 2) = "�ϼ�"
        .TextMatrix(.rows - 1, 3) = "�ϼ�"
        
        If mint�����ʾ = 1 Then
            str�����ʾ = "ʵ�ս�" & Format(dblʵ�ս��, "#####0.00;-#####0.00; ;")
        ElseIf mint�����ʾ = 2 Then
            str�����ʾ = "Ӧ�ս�" & Format(dblӦ�ս��, "#####0.00;-#####0.00; ;") & "    ʵ�ս�" & Format(dblʵ�ս��, "#####0.00;-#####0.00; ;")
        Else
            str�����ʾ = "Ӧ�ս�" & Format(dblӦ�ս��, "#####0.00;-#####0.00; ;")
        End If
        
        .TextMatrix(.rows - 1, 4) = str�����ʾ
        .TextMatrix(.rows - 1, 5) = str�����ʾ
        .TextMatrix(.rows - 1, 6) = str�����ʾ
        .TextMatrix(.rows - 1, 7) = str�����ʾ
        
        .MergeCells = flexMergeFree
        .MergeRow(.rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "��ʾ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    WriteDataToBill = True
End Function

Private Function SetLocateBill(Optional ByVal strNo As String = "", _
    Optional ByVal BlnEnterCell As Boolean = True) As Boolean
    Dim intRow As Integer
    
    SetLocateBill = False
    With Msf�����б�
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 1) = strNo And .RowData(intRow) = 8 Then
                .Row = intRow
                .TopRow = intRow
                SetLocateBill = True
                Exit For
            End If
        Next
    End With
    
    If BlnEnterCell Then Msf�����б�_EnterCell
End Function

Private Function CheckStock() As Boolean
    Dim RecCheckStock As New ADODB.Recordset
    Dim dblStock As Double
    Dim strSubSql As String
    '�����
    If IntCheckStock = 0 Then CheckStock = True: Exit Function
    
    '���������ת��Ϊ��Ӧ��λ��ʵ������
    Dim intUnit As Integer
    On Error GoTo errHandle
    intUnit = Val(zldatabase.GetPara("ҩ������", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(lngҩ��ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(lngҩ��ID, gint����ҩ��)
    Else
        strUnit = GetSpecUnit(lngҩ��ID, gintסԺҩ��)
    End If
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubSql = "/1"
    Case "���ﵥλ"
        strSubSql = "/Decode(B.�����װ,Null,1,0,1,B.�����װ)"
    Case "סԺ��λ"
        strSubSql = "/Decode(B.סԺ��װ,Null,1,0,1,B.סԺ��װ)"
    Case "ҩ�ⵥλ"
        strSubSql = "/Decode(B.ҩ���װ,Null,1,0,1,B.ҩ���װ)"
    End Select
    
    CheckStock = False
    With RecTotal
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
           gstrSQL = " Select nvl(ʵ������,0)" & strSubSql & " AS ����" & _
                 " From ҩƷ��� A,ҩƷ��� B" & _
                 " Where B.ҩƷID=A.ҩƷID And A.����=1 And A.�ⷿID=[1] And A.ҩƷID=[2] And Nvl(A.����,0)=[3]"
           Set RecCheckStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID, CLng(RecTotal!ҩƷid), CLng(RecTotal!����))
           
           With RecCheckStock
                If .EOF Then
                    dblStock = 0
                Else
                    dblStock = !����
                End If
                
                If dblStock < RecTotal!���� Then
                    If RecTotal!���� <> 0 Then
                        MsgBox RecTotal!Ʒ�� & "�����ο�������������ܼ�����ҩ��", vbInformation, gstrSysName: Exit Function
                    Else
                        Select Case IntCheckStock
                        Case 1
                            If MsgBox(RecTotal!Ʒ�� & "�Ŀ�����������Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox RecTotal!Ʒ�� & "�Ŀ�������������ܼ�����ҩ��", vbInformation, gstrSysName: Exit Function
                        End Select
                    End If
                End If
            End With
            .MoveNext
        Loop
    End With
    
    CheckStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendBill() As Boolean
    Dim intRow As Integer
    Dim StrDate As String
    Dim rsSendRecipeByNo As ADODB.Recordset
    Dim int���� As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnInTrans As Boolean
    Dim strǩ����¼ As String
    Dim strReturn As String
    Dim strNo As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    
    arrSql = Array()

    SendBill = False
    
    Set rsSendRecipeByNo = New ADODB.Recordset
    With rsSendRecipeByNo
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 2, adFldIsNullable
        .Fields.Append "�����־", adDouble, 1, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    StrDate = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With Msf�����б�
        For intRow = 1 To .rows - 1
            If .RowData(intRow) <> 0 Then
                '���۹���
                If CheckPriceAdjustByNO(Val(.RowData(intRow)), lngҩ��ID, .TextMatrix(intRow, 1)) = False Then
                    Exit Function
                End If
                
                With rsSendRecipeByNo
                    .AddNew
                    !NO = Msf�����б�.TextMatrix(intRow, 1)
                    !���� = Msf�����б�.RowData(intRow)
                    !��¼���� = Val(Msf�����б�.TextMatrix(intRow, 9))
                    !�����־ = Val(Msf�����б�.TextMatrix(intRow, 10))
                    .Update
                End With
            End If
        Next
    End With
    
    '�������������������ҩ
    rsSendRecipeByNo.Sort = "NO"
    rsSendRecipeByNo.MoveFirst
    For intRow = 1 To rsSendRecipeByNo.RecordCount
        '�ȼ���ִ��Ԥ����
        Call AutoAdjustPrice_ByNO(rsSendRecipeByNo!����, rsSendRecipeByNo!NO)
        
        If Val(rsSendRecipeByNo!��¼����) = 1 Or (Val(rsSendRecipeByNo!��¼����) = 2 And (Val(rsSendRecipeByNo!�����־) = 1 Or Val(rsSendRecipeByNo!�����־) = 4)) Then
            int���� = 1
        Else
            int���� = 2
        End If
        
        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
        '�ⷿID
        gstrSQL = gstrSQL & lngҩ��ID
        '����
        gstrSQL = gstrSQL & "," & rsSendRecipeByNo!����
        'NO
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
        '�����
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '��ҩ��
        gstrSQL = gstrSQL & ",'" & str��ҩ�� & "'"
        'У����
        gstrSQL = gstrSQL & ",NULL"
        '��ҩ��ʽ
        gstrSQL = gstrSQL & ",2"
        '��ҩʱ��
        gstrSQL = gstrSQL & ",to_date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')"
        '����Ա���
        gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
        '����Ա����
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '����λ��
        gstrSQL = gstrSQL & "," & int����λ��
        '��˻��۵�
        gstrSQL = gstrSQL & "," & int��˻��۵�
        '�Ƿ�����
        gstrSQL = gstrSQL & "," & int����
        gstrSQL = gstrSQL & ")"

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        
        strNo = strNo & rsSendRecipeByNo!���� & "," & rsSendRecipeByNo!NO & "|"
        rsSendRecipeByNo.MoveNext
    Next
    
    '���÷�ҩǰ����ҽӿ�
    err.Clear
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        If mobjPlugIn.DrugBeforeSendByRecipe(lngҩ��ID, strNo, strReserve) = False Then
            If err.Number <> 0 Then
                err.Clear: On Error GoTo 0
            Else
                Exit Function
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo ErrHand
    
    '�ȴ���ҩ����
    gcnOracle.BeginTrans
    blnInTrans = True
    
    '�����������ǩ�����ŵ�ҵ����ǰ�棬��ֹ���������������
    '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
    If gblnESign������ҩ = True And gblnESignUserStoped = False Then
        rsSendRecipeByNo.MoveFirst
        For intRow = 1 To rsSendRecipeByNo.RecordCount
            strǩ����¼ = ""
            If GetSignatureRecored(EsignTache.send, rsSendRecipeByNo!����, rsSendRecipeByNo!NO, lngҩ��ID, strǩ����¼, 0, CDate(StrDate), gstrUserName) = False Then
                gcnOracle.RollbackTrans
                blnInTrans = False
                Exit Function
            End If
            
            If strǩ����¼ <> "" Then
                gstrSQL = "Zl_ҩƷǩ����¼_Insert(" & strǩ����¼ & ")"
               
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            Else
                gcnOracle.RollbackTrans
                blnInTrans = False
                MsgBox "�Է�ҩ�˵���ǩ��ʧ�ܣ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            rsSendRecipeByNo.MoveNext
        Next
    End If
    
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_Abolish")
    Next
    
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        If mblnConPacker And strNo <> "" And mblnLoadDrug Then
            Call mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.�û�����, UserInfo.�û�����, lngҩ��ID, Mid(strNo, 1, Len(strNo) - 1), strReturn)
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        If mblnConPacker Then
            If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
            mobjDrugMAC.Operation gstrDbUser, Val("22-��ʼ��ҩ"), "1|" & Replace(strNo, "|", ";"), strReturn
        End If
    End If
    
    If MsgBox("����Ҫ��ӡ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_2", "ZL8_BILL_1341_2"), Me, "�ⷿ=" & lngҩ��ID, "��ҩ��ʽ=������ҩ|2", "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "��ҩʱ��=" & StrDate, 2)
    End If
    
    '���÷�ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        mobjPlugIn.DrugSendByRecipe lngҩ��ID, strNo, CDate(StrDate), strReserve
        err.Clear: On Error GoTo 0
    End If
    
    SendBill = True
    Exit Function
ErrHand:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng���� As Long, str��� As String
    '��鴦���Ƿ��ѽ��ʡ����ò����Ƿ��ѳ�Ժ������Ȩ�޽��м��
    With rs���
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !���ݱ�ʶ
            lng���� = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            str��� = nvl(!���)
            If Not IsReceiptBalance_Charge(0, strPrivs, lng����, strNo, str���, Val(!��¼����), Val(!�����־)) Then Exit Function
            If Not IsOutPatient(strPrivs, lng����, strNo, Val(!��¼����), Val(!�����־)) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function

Private Sub InitRec()
    Set rs��� = New ADODB.Recordset
    With rs���
        If .State = 1 Then .Close
        .Fields.Append "���ݱ�ʶ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "�����־", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub txtҽ����_GotFocus()
    GetFocus txtҽ����
End Sub


Private Sub txtҽ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call GetRecipe(2, txtҽ����)
End Sub


