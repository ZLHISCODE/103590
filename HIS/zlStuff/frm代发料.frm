VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm������ 
   Caption         =   "�������ⷿ����"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9705
   Icon            =   "frm������.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9705
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   8520
      TabIndex        =   13
      Top             =   5880
      Width           =   990
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   7080
      TabIndex        =   12
      Top             =   5880
      Width           =   990
   End
   Begin VB.TextBox Txt���� 
      Height          =   300
      Left            =   7320
      MaxLength       =   12
      TabIndex        =   11
      Top             =   570
      Width           =   2150
   End
   Begin VB.TextBox txtҽ���� 
      Height          =   300
      Left            =   4080
      TabIndex        =   7
      Top             =   570
      Width           =   2085
   End
   Begin VB.TextBox TxtNo 
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   570
      Width           =   1845
   End
   Begin VB.ComboBox cbo���ϲ��� 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
   Begin MSComCtl2.DTPicker Dtp��ʼDate 
      Height          =   300
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   120127491
      CurrentDate     =   37007
   End
   Begin MSComCtl2.DTPicker Dtp����Date 
      Height          =   300
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   120127491
      CurrentDate     =   37007
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�����б� 
      Height          =   1755
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   9375
      _ExtentX        =   16536
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
   Begin TabDlg.SSTab TabShow 
      Height          =   2685
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4736
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "������ϸ(&D)"
      TabPicture(0)   =   "frm������.frx":74F2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Msf������ϸ"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "���ϻ���(&T)"
      TabPicture(1)   =   "frm������.frx":750E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Msf��������"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf������ϸ 
         Height          =   2265
         Left            =   -74940
         TabIndex        =   16
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3995
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
         Height          =   2265
         Left            =   60
         TabIndex        =   17
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3995
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
   Begin VB.Label lblҽ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3420
      TabIndex        =   10
      Top             =   630
      Width           =   540
   End
   Begin VB.Label Lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6840
      TabIndex        =   9
      Top             =   630
      Width           =   360
   End
   Begin VB.Label LblNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   8
      Top             =   630
      Width           =   540
   End
   Begin VB.Label Lbl����Date 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6480
      TabIndex        =   5
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Lbl��ʼDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   4
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lbl���ϲ��� 
      Caption         =   "���ϲ���"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frm������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public mstr������ As String

'--�ⲿ���ݲ���--
Private mblnModify As Boolean
Private mstrPrivs As String
Private mint���� As Integer                          '����
Private mlng���ϲ���ID As Long                           '�Ϸ�
Private mintSendAfterDosage As Integer               '����δ���Ϸ���
Private mint����δ��˴������� As Integer            '����δ��˴�������
Private mIntCheckStock As Integer                    '�����
Private mintУ�鴦�� As Integer                      'У�鴦��
Private mintUnit  As Integer                        '��λ
'--������ʹ�ñ���--
Private mrsBill As New ADODB.Recordset              '���ݼ�¼
Private mrsTotal As New ADODB.Recordset             '��������
Private mrs��� As ADODB.Recordset
Private mrs������Դ���� As ADODB.Recordset            '��¼���д����ϴ�������Դ����
Private mrs����������ϸ As ADODB.Recordset            '��¼�������ܵļ�¼��ʵ���ǰ����ݺŵ���ϸ��¼


Private mblnStartUp As Boolean
Private mlngListRow As Long                          '�����б�
Private mlngDetailRow As Long                        '������ϸ
Private mlngTotalRow As Long                         '��������
Private mstrBillNo As String                         '���ܵ��ݺ�
Private mstrID As String                             '����ID
Private mlngBillCount As Long
Private mstr���ݺ� As String
Private mint�������� As String
Private mstr����IN  As String
Private mbln��Ʊ�ݺŷ��� As Boolean
Private Const mlngModule = 1723
Private mobjPlugIn As Object             '��ҽӿڶ���
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Public Property Get ��Ʊ�ݺŷ���() As Boolean
   ��Ʊ�ݺŷ��� = mbln��Ʊ�ݺŷ���
End Property

Public Property Let ��Ʊ�ݺŷ���(ByVal vNewValue As Boolean)
    mbln��Ʊ�ݺŷ��� = vNewValue
End Property

Public Property Get In_Ȩ��() As String
    In_Ȩ�� = mstrPrivs
End Property

Public Property Let In_Ȩ��(ByVal vNewValue As String)
    mstrPrivs = vNewValue
End Property
Public Property Get In_����IN() As String
    In_����IN = mstr����IN
End Property

Public Property Let In_����IN(ByVal vNewValue As String)
    mstr����IN = vNewValue
End Property


Public Property Get In_����() As Integer
    In_���� = mint����
End Property

Public Property Let In_����(ByVal vNewValue As Integer)
    mint���� = vNewValue
End Property

Public Property Get In_У�鴦��() As Integer
    In_У�鴦�� = mintУ�鴦��
End Property

Public Property Let In_У�鴦��(ByVal vNewValue As Integer)
    mintУ�鴦�� = vNewValue
End Property

Public Property Get In_�����() As Integer
    In_����� = mIntCheckStock
End Property

Public Property Let In_�����(ByVal vNewValue As Integer)
    mIntCheckStock = vNewValue
End Property

Public Property Get In_���ϲ���id() As Long
    In_���ϲ���id = mlng���ϲ���ID
End Property

Public Property Let In_���ϲ���id(ByVal vNewValue As Long)
    mlng���ϲ���ID = vNewValue
End Property

Public Property Get In_����δ���Ϸ���() As Integer
    In_����δ���Ϸ��� = mintSendAfterDosage
End Property

Public Property Let In_����δ���Ϸ���(ByVal vNewValue As Integer)
    mintSendAfterDosage = vNewValue
End Property

Public Property Get IN_����δ��˷���() As Integer
    IN_����δ��˷��� = mint����δ��˴�������
End Property

Public Property Let IN_����δ��˷���(ByVal vNewValue As Integer)
    mint����δ��˴������� = vNewValue
End Property

Private Sub SetFormat(Optional ByVal IntStyle As Integer = 1)
    Dim intCol As Integer
    '���ø��б�ؼ��ĸ�ʽ

    Select Case IntStyle
    Case 1
        With Msf�����б�
            .Rows = 2
            .Cols = 10
    
            .TextMatrix(0, 0) = "����"
            .TextMatrix(0, 1) = "NO"
            .TextMatrix(0, 2) = "����"
            .TextMatrix(0, 3) = "����"
            .TextMatrix(0, 4) = "סԺ��"
            .TextMatrix(0, 5) = "����"
            .TextMatrix(0, 6) = "�շ�Ա"
            .TextMatrix(0, 7) = "����ҽ��"
            .TextMatrix(0, 8) = "��������"
            .TextMatrix(0, 9) = "�����־"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If mblnStartUp = False Then
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
                
                .Row = 1
                Call RestoreFlexState(Msf�����б�, Me.Name)
                If glngSys \ 100 <> 1 Then
                    .ColWidth(2) = 0
                    .ColWidth(4) = 0
                    .ColWidth(5) = 0
                End If
                .ColWidth(7) = IIf(mintУ�鴦�� = 1, 0, 1000)
            End If
        End With
    Case 2
        With Msf������ϸ
            .Rows = 2
            .Cols = 6
    
            .TextMatrix(0, 0) = "��������"
            .TextMatrix(0, 1) = "���"
            .TextMatrix(0, 2) = "��λ"
            .TextMatrix(0, 3) = "����"
            .TextMatrix(0, 4) = "����"
            .TextMatrix(0, 5) = "���"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
                If intCol < 2 Then .ColAlignment(intCol) = 1
                If intCol > 2 Then .ColAlignment(intCol) = 7
            Next
    
            If mblnStartUp = False Then
                .ColWidth(0) = 2000
                .ColWidth(1) = 1500
                .ColWidth(2) = 500
                .ColWidth(3) = 800
                .ColWidth(4) = 800
                .ColWidth(5) = 1000
                
                .Row = 1
                Call RestoreFlexState(Msf������ϸ, Me.Name)
            End If
        End With
    Case 3
        With Msf��������
            .Rows = 2
            .Cols = 9
    
            .TextMatrix(0, 0) = "���"
            .TextMatrix(0, 1) = "��������"
            .TextMatrix(0, 2) = "���"
            .TextMatrix(0, 3) = "��λ"
            .TextMatrix(0, 4) = "����"
            .TextMatrix(0, 5) = "����"
            .TextMatrix(0, 6) = "���"
            .TextMatrix(0, 7) = "����ID"
            .TextMatrix(0, 8) = "����"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
                If intCol < 3 Then .ColAlignment(intCol) = 1
                If intCol > 3 Then .ColAlignment(intCol) = 7
                
            Next
            
            If mblnStartUp = False Then
                .ColWidth(0) = 500
                .ColWidth(1) = 2000
                .ColWidth(2) = 1500
                .ColWidth(3) = 500
                .ColWidth(4) = 800
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 0
                .ColWidth(8) = 0
                .Row = 1
                Call RestoreFlexState(Msf��������, Me.Name)
            End If
        End With
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If CheckStock = False Then Exit Sub
    If Not CheckCorrelation Then Exit Sub
    If SendBill = False Then Exit Sub
    
    mlngBillCount = 0
'    lblNote.Caption = IIf(mlngBillCount = 0, "δ�����κδ���", "������" & mlngBillCount & "�Ŵ���")
    
    '��ʼ��
    mstrID = ""
    mstrBillNo = ""
    TxtNo = ""
    
    With Msf��������
        .Clear
        .Rows = 2
        .RowData(1) = 0
    End With
    With Msf�����б�
        .Clear
        .Rows = 2
        .RowData(1) = 0
    End With
    With Msf������ϸ
        .Clear
        .Rows = 2
        .RowData(1) = 0
    End With
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    cmdOk.Enabled = False
    TxtNo.SetFocus
End Sub

Private Sub CmdPrint_Click()
    Dim HisPrint As New zlPrint1Grd
    Dim HisRow As New zlTabAppRow
    Dim ArrayNo, IntArray As Integer
    Dim LngSelectRow As Long, intCol As Integer
    
    On Error Resume Next
    'ȡ������ѡ��״̬
    With Msf��������
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If mlngTotalRow > 0 And mlngTotalRow < .Rows Then
            .Row = mlngTotalRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
    End With
    
    HisPrint.Title = "���ϻ���"
    Set HisRow = New zlTabAppRow
    HisRow.Add "����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    HisPrint.UnderAppRows.Add HisRow
    
    ArrayNo = Split(mstrBillNo, ";")
    
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
        
        mlngTotalRow = LngSelectRow
        .Row = mlngTotalRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub cmdPrintSet_Click()
    zlPrintSet
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strReg As String
    
    mblnStartUp = False
    mlngBillCount = 0
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    
    Dtp����Date.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59"
    Dtp��ʼDate.Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00"
    
    Dtp����Date.MinDate = zlDatabase.Currentdate - 30
    Dtp��ʼDate.MinDate = zlDatabase.Currentdate - 30
    
    CheckDepend
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
   ' mbln��Ʊ�ݺŷ��� = False
    If mbln��Ʊ�ݺŷ��� = True Then LblNo.Caption = "Ʊ�ݺ�": Me.Caption = "��Ʊ�ݺŷ���"
    mstrID = ""
    mstrBillNo = ""
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    mblnStartUp = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 7730 Then Me.Width = 7730
    If Me.Height < 6000 Then Me.Height = 6000
    
'    With lblNote
'        .Left = Me.ScaleWidth - .Width - 100
'    End With
    
'    With CmdHelp
'        .Top = Me.ScaleHeight - .Height - 100
'    End With
'    With cmdPrintSet
'        .Top = CmdHelp.Top
'        .Left = CmdHelp.Left + CmdHelp.Width + 100
'    End With
'    With CmdPrint
'        .Top = CmdHelp.Top
'        .Left = cmdPrintSet.Left + cmdPrintSet.Width + 100
'    End With
    
    With cmdCancle
         .Top = Me.ScaleHeight - .Height - 200
        .Left = Me.ScaleWidth - .Width - 100
    End With
    With cmdOk
        .Top = Me.ScaleHeight - .Height - 200
        .Left = cmdCancle.Left - .Width - 300
    End With
    
    With Msf�����б�
        .Height = (cmdOk.Top - 200 - .Top) / 2
        .Width = Me.ScaleWidth - .Left - 150
    End With
    
    With TabShow
        .Top = Msf�����б�.Top + Msf�����б�.Height + 100
        .Height = cmdOk.Top - 100 - .Top
        .Width = Msf�����б�.Width
    End With
    With Msf��������
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 150
    End With
    With Msf������ϸ
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 150
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
        If mlngTotalRow > 0 And mlngTotalRow < .Rows Then
            .Row = mlngTotalRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngTotalRow = LngSelectRow
        .Row = mlngTotalRow       '���õ�ǰѡ����
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
        If mlngListRow > 0 And mlngListRow < .Rows Then
            .Row = mlngListRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngListRow = LngSelectRow
        .Row = mlngListRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
        
        If Trim(.TextMatrix(.Row, 1)) = "" Then
            With Msf������ϸ
                .Clear
                .Rows = 2
                Call SetFormat(2)
            End With
            Exit Sub
        End If
        
        '��ʾ������ϸ
        Call ReadBillData(.RowData(.Row), .TextMatrix(.Row, 1), Val(.TextMatrix(.Row, 9)))
    End With
End Sub

Private Sub Msf�����б�_GotFocus()
    With Msf�����б�
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf�����б�_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng���� As Long, strNo As String
    
    If KeyCode = vbKeyDelete Then
        If Msf�����б�.TextMatrix(Msf�����б�.Row, 1) = "" Then Exit Sub
        
        With mrs����������ϸ
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    .Find "���ݺ�='" & Msf�����б�.TextMatrix(Msf�����б�.Row, 1) & "'"
                    If Not .EOF Then .Delete
                    If Not .EOF Then .MoveNext
                Loop
            End If
        End With
        With mrs������Դ����
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "��Դ����='" & Msf�����б�.TextMatrix(Msf�����б�.Row, 2) & "'"
                If Not .EOF Then .Delete
            End If
        End With
        With Msf�����б�
            lng���� = Val(.TextMatrix(.Row, 0))
            strNo = .TextMatrix(.Row, 1)
            If .Rows - 1 = 1 Then
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
                .TextMatrix(1, 6) = ""
                .RowData(1) = 0
            Else
                If Trim(.TextMatrix(.Row, 1)) <> "" Then .RemoveItem .Row: mlngBillCount = mlngBillCount - 1
            End If
            
            cmdOk.Enabled = (.RowData(IIf(.Rows - 1 = 1, 1, .Rows - 2)) <> 0)
'            lblNote.Caption = IIf(mlngBillCount = 0, "δ�����κδ���", "������" & mlngBillCount & "�Ŵ���")
        
            'ɾ���õ���
            With mrs���
                If .RecordCount <> 0 Then .MoveFirst
                .Find "���ݱ�ʶ='" & strNo & "|" & lng���� & "'"
                If Not .EOF Then .Delete
            End With
            
        End With
        
        Msf�����б�_EnterCell
        mblnModify = True
        Call WriteTotalDataToBill
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
        If mlngDetailRow > 0 And mlngDetailRow < .Rows Then
            .Row = mlngDetailRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngDetailRow = LngSelectRow
        .Row = mlngDetailRow       '���õ�ǰѡ����
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
        WriteTotalDataToBill
        Msf��������.ZOrder
        Msf��������_EnterCell
    End Select
End Sub

Private Sub TxtNo_GotFocus()
    zlControl.TxtSelAll TxtNo
End Sub

Private Function Send������(ByVal intType As Integer, ByVal strInput As String) As Boolean
    '--------------------------------------------------------------------------------------------------------------------------------------------
    '--����:�������ŷ���
    '--����:
    '--����:���ϳɹ�,����true,���򷵻�False
    '--------------------------------------------------------------------------------------------------------------------------------------------

    Dim intYear As Integer, strYear As String
    Dim intRow As Integer
    Dim strNo As String, IntBill As Integer, ArrTmp, strTmp As String
    Dim strSQL As String
    Dim int���� As Integer
    Dim strBeginDate As Date
    Dim strEndDate As Date
    
    
    '--���������λ,�򰴹������--
'    Me.TxtNo = UCase(LTrim(Me.TxtNo))
'    Me.TxtNo.Text = GetFullNO(Me.TxtNo.Text, 13)
    On Error GoTo errHandle
    strBeginDate = Format(Dtp��ʼDate.Value, "yyyy-MM-dd hh:mm:ss")
    strEndDate = Format(Dtp����Date.Value, "yyyy-MM-dd hh:mm:ss")
    
    gstrSQL = "" & _
             " Select /*+ Rule*/ Distinct Decode(C.����,24,'�շ�',25,'����',26,'���ʱ�') ����,C.No,C.����,A.���շ�," & _
             "      Decode(A.��ҩ��,Null,'','���ŷ���','',A.��ҩ��) ������,P.���� ����,decode(c.����,26,'',B.����) ����," & _
             "      Decode(c.����,26,'',B.��ʶ��)  סԺ��,decode(c.����,26,'','') ����,B.������ ����ҽ��,B.����Ա���� ������," & _
             "      To_Char(C.��������,'yyyy-MM-dd') ��������,1 ���� " & _
             " From δ��ҩƷ��¼ A,������ü�¼ B,ҩƷ�շ���¼ C,���ű� P,���ű� S" & IIf(intType = 2, ",������Ϣ W ", " ") & _
             "     ,Table(cast(f_Str2List([2]) as zlTools.t_StrList)) D " & _
             " Where C.����ID=B.ID And B.��������ID+0=P.ID(+) And Nvl(C.�ⷿID,0)+0=S.ID(+) " & _
             "     And Nvl(A.�ⷿID,0)=Nvl(C.�ⷿID,0) And Mod(C.��¼״̬,3)=1 And A.No=C.No " & _
             "     And (C.�ⷿID+0=[1] OR C.�ⷿID IS NULL)" & _
             "     And C.����=D.Column_Value And C.����� Is Null " & _
             "     And C.����=A.���� and nvl(C.��ҩ��ʽ,-999)<>-1 And Nvl(B.����״̬,0)<>1 " & _
             "     And A.�������� Between [3] And  [4]"
    If intType = 1 Then
        gstrSQL = gstrSQL & " And C.No=[5]"
    ElseIf intType = 2 Then
        gstrSQL = gstrSQL & " And A.����id=W.����id And W.ҽ����=[5] "
    ElseIf intType = 3 Then
        gstrSQL = gstrSQL & " And B.���� Like [5]"
    End If

    If mstr����IN = "24" Then
    ElseIf mstr����IN = "26" Then
        gstrSQL = Replace(gstrSQL, "1 ����", "2 ����")
        gstrSQL = Replace(gstrSQL, "B.����", "R.����")
        gstrSQL = Replace(gstrSQL, "decode(c.����,26,'','') ����", "decode(c.����,26,'',B.����) ����")
        gstrSQL = Replace(gstrSQL, "������ü�¼ B", "סԺ���ü�¼ B,������ҳ R")
        gstrSQL = Replace(gstrSQL, "And Nvl(B.����״̬,0)<>1", "And B.����id=R.����id And B.��ҳid=R.��ҳid")
    ElseIf InStr(1, mstr����IN, "25") > 0 Or InStr(1, mstr����IN, "26") > 0 Then
        strSQL = Replace(gstrSQL, "1 ����", "2 ����")
        strSQL = Replace(strSQL, "B.����", "R.����")
        strSQL = Replace(strSQL, "decode(c.����,26,'','') ����", "decode(c.����,26,'',B.����) ����")
        strSQL = Replace(strSQL, "������ü�¼ B", "סԺ���ü�¼ B,������ҳ R")
        strSQL = Replace(strSQL, "And Nvl(B.����״̬,0)<>1", "And B.����id=R.����id And B.��ҳid=R.��ҳid")
        gstrSQL = gstrSQL & " Union All " & strSQL
    End If
    
'    err = 0: On Error Resume Next
'    Set mrsBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtNO, mlng���ϲ���ID, mstr����IN)
    Set mrsBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
                    Val(Me.cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex)), _
                    mstr����IN, _
                    CDate(strBeginDate), _
                    CDate(strEndDate), _
                    IIf(intType = 3, "%" & strInput & "%", strInput))
    If err <> 0 Then
        MsgBox "��ȡ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        GoTo ExitSub  '��ȡ����δ���ϼ�¼
    End If
'    If ReadData(gstrSQL) = False Then GoTo ExitSub  '��ȡ����δ���ϼ�¼

    If mrsBill.EOF Then
        MsgBox "δ�ҵ�ָ�����������������룡", vbInformation, gstrSysName
        GoTo ExitSub
    End If
    
    If mrsBill.RecordCount > 1 Then
        strTmp = Frm����ѡ��.ShowMe(Me, mrsBill)
        If strTmp = "" Then GoTo ExitSub
        
        ArrTmp = Split(strTmp, ";")
        strNo = ArrTmp(0)
        IntBill = ArrTmp(1)
        
        mrsBill.MoveFirst
        mrsBill.Filter = "����=" & IntBill & " And NO='" & strNo & "'"
        int���� = mrsBill!����
    Else
        strNo = mrsBill!NO
        IntBill = mrsBill!����
        int���� = mrsBill!����
    End If
    Me.TxtNo.Tag = IntBill
    
    '����Ѵ��ڸõ��ݣ����˳�
    If SetLocateBill(TxtNo.Text, IntBill, False) Then
        MsgBox "�ô����Ѿ����룬�����䣡", vbInformation, gstrSysName
        GoTo ExitSub
    End If
    
    '���Ϸ���
    If CheckBill(IntBill, strNo) <> 0 Then GoTo ExitSub
    '�����ǰ���봦���Ŀ�������¼��Ĵ����Ŀ��Ҳ�ͬ���������ʾ
    If CheckSource(IntBill, strNo) = False Then Exit Function
    If WriteSendListData(IntBill, strNo, int����) = False Then GoTo ExitSub
    
    mlngBillCount = mlngBillCount + 1
'    lblNote.Caption = IIf(mlngBillCount = 0, "δ�����κδ���", "������" & mlngBillCount & "�Ŵ���")
    
    '��λ���ղ�����Ĵ�����
    Call SetLocateBill(TxtNo.Text, Val(TxtNo.Tag))
    
    With Msf�����б�
        cmdOk.Enabled = (.RowData(IIf(.Rows - 1 = 1, 1, .Rows - 2)) <> 0)
    End With
    
    mblnModify = True
    Call RefreshData
    With TxtNo
        .SelStart = 0
        .SelLength = Len(TxtNo)
    End With
    Send������ = True
    Exit Function
ExitSub:
    With TxtNo
        .SelStart = 0
        .SelLength = Len(TxtNo)
        .SetFocus
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtNo) = "" Then Exit Sub
    Me.TxtNo.Text = zlCommFun.GetFullNO(Me.TxtNo.Text, 13)
    
    If Send������(1, TxtNo.Text) = False Then Exit Sub
    
End Sub

Private Function CheckSource(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim bln�ظ����� As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "" & _
        "   Select B.���� as ����,B.���� as ��Դ���� " & _
        "   From ҩƷ�շ���¼ A,���ű� B " & _
        "   Where A.�Է�����id=B.id and No=[1] And ����=[2]" & _
        "           And Mod(��¼״̬,3)=1 And ����� Is Null And (�ⷿID+0=[3] Or �ⷿID Is NULL) And Rownum<2"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "���", strNo, int����, Val(Me.cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex)))
    
    
    If rs.RecordCount = 0 Then
        CheckSource = False
        Exit Function
    End If
    
    With mrs������Դ����
        If .RecordCount = 0 Then
            .AddNew
            !���� = rs!����
            !��Դ���� = rs!��Դ����
            CheckSource = True
        Else
            .MoveFirst
            For n = 1 To .RecordCount
                If !���� = rs!���� Then
                    bln�ظ����� = True
                    Exit For
                End If
                .MoveNext
            Next
            If Not bln�ظ����� Then
                If MsgBox("��ǰ�����Ŀ���������[" & rs!���� & "]" & rs!��Դ���� & "����ȷ��Ҫ����ô�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                Else
                    .AddNew
                    !���� = rs!����
                    !��Դ���� = rs!��Դ����
                    CheckSource = True
                End If
            Else
                CheckSource = True
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadData(ByVal StrQuery As String) As Boolean
    '--��ȡ����--

'    On Error Resume Next
'    err = 0
    ReadData = False
    On Error GoTo errHandle
    
    gstrSQL = StrQuery
    Call zlDatabase.OpenRecordset(mrsBill, gstrSQL, Me.Caption)
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

Private Function ReadBillData(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int���� As Integer) As Boolean
    Dim IntStyle As Integer
    Dim str��� As String
    Dim str��ϸ��λ�� As String, str���ܵ�λ�� As String
    '--��ȡ��������--
    'BillStyle-��������;BIllNO-���ݺ�
    '��λ��ʾ���ݷ����������������ﵥλ��סԺ��סԺ���סԺ��λ���������ۼ۵�λ��
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    ReadBillData = False
    
    
    Select Case mintUnit
    Case 0
        str��ϸ��λ�� = "C.���㵥λ ��λ,B.���ۼ� ����,B.ʵ������*Nvl(B.����,1) ����"
        str���ܵ�λ�� = "C.���㵥λ ��λ,B.���ۼ� ����,Sum(B.ʵ������*Nvl(B.����,1)) ����"
    Case Else
        str��ϸ��λ�� = "D.��װ��λ ��λ,B.���ۼ�*nvl(D.����ϵ��,1) ����,B.ʵ������/nvl(D.����ϵ��,1)*Nvl(B.����,1) ����"
        str���ܵ�λ�� = "D.��װ��λ ��λ,B.���ۼ�*nvl(D.����ϵ��,1) ����,Sum(B.ʵ������/nvl(D.����ϵ��,1)*Nvl(B.����,1)) ����"
    End Select
    
    str��ϸ��λ�� = str��ϸ��λ�� & ",B.���۽�� ��� "
    str���ܵ�λ�� = str���ܵ�λ�� & ",Sum(B.���۽��) ��� "

    gstrSQL = "" & _
        "   SELECT DISTINCT F.���,F.����ID,'['||C.����||']'||C.����  As Ʒ��,DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)) ���," & _
                str��ϸ��λ�� & _
        " FROM ҩƷ�շ���¼ B,�������� D,�շ���ĿĿ¼ C,������ü�¼ F" & _
        " WHERE B.ҩƷID=D.����ID AND D.����ID=C.ID And B.����ID=F.ID" & _
        "       AND MOD(B.��¼״̬,3)=1 AND B.NO=[1] AND B.����=[2]" & _
        "       AND (B.�ⷿID+0=[3] OR B.�ⷿID IS NULL)"
    gstrSQL = gstrSQL & " And b.����� Is Null"
    
    If int���� = 2 Then
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    
    gstrSQL = gstrSQL & " Order by ���"

    Set mrsBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, BillNo, BillStyle, Val(Me.cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex)))

    With mrsBill
        str��� = ""
        Do While Not .EOF
            str��� = str��� & "," & !���
            .MoveNext
        Loop
        If str��� <> "" Then str��� = Mid(str���, 2)
        .MoveFirst
    End With
    
    
    '��������Ϣ����ϸ���д���ڲ�ӳ���¼����
    With mrs���
        If .RecordCount <> 0 Then .MoveFirst
        .Find "���ݱ�ʶ='" & BillNo & "|" & BillStyle & "'"
        If str��� <> "" Then
            If .EOF Then
                .AddNew
                !���ݱ�ʶ = BillNo & "|" & BillStyle
                !��� = str���
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

Private Function CheckBill(ByVal intBillStyle As Integer, ByVal strNo As String) As Integer
    Dim rsCheck As New ADODB.Recordset

    '--���ݽ�Ҫִ�еĲ������ж��Ƿ�����--
    '����:
    '0-�������
    '1-δ����
    '2-������
    '3-�ѷ���
    '4-��ɾ��
    '5-δ����
    
    On Error GoTo errHandle
    gstrSQL = "" & _
        "   Select A.��ҩ�� ������,A.�����,nvl(B.���շ�,0) ���շ� " & _
        "   From ҩƷ�շ���¼ A,δ��ҩƷ��¼ B " & _
        "  Where A.No=B.No And A.����=B.���� And A.No=[1] And A.����=[2]" & _
        "           And mod(A.��¼״̬,3)=1 And Rownum=1 And (A.�ⷿID+0=[3] Or A.�ⷿID Is NULL)"
    gstrSQL = gstrSQL & " And A.����� IS Null"
    
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, intBillStyle, Val(Me.cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex)))
    
    With rsCheck
        If .EOF Then CheckBill = 4: MsgBox "δ�ҵ�����[" & strNo & "],�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!�����) Then
            CheckBill = 3: MsgBox "�ô���[" & strNo & "]�ѱ���������Ա���ϣ����ϲ�����ֹ��", vbInformation, gstrSysName: Exit Function
        End If
'        If frm���ķ��Ź���.mint����δ��˴������� = 0 Then
'            If !���շ� = 0 Then
'                CheckBill = 3: MsgBox "�ô���[" & strNo & "]��δ�շѣ����ϲ�����ֹ��", vbInformation, gstrSysName: Exit Function
'            End If
'        End If
    End With

    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteSendListData(ByVal int���� As Integer, ByVal strNo As String, ByVal int���� As Integer) As Boolean
    Dim rsCheck As New ADODB.Recordset
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    WriteSendListData = False
    
    If mintSendAfterDosage = 0 Then
        If IsNull(mrsBill!������) Then
            MsgBox "�ô�����δ���ϣ�����ִ�з��ϲ�����", vbInformation, gstrSysName
            Exit Function
        End If
        If Trim(mrsBill!������) = "" Then
            MsgBox "�ô�����δ���ϣ�����ִ�з��ϲ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mint����δ��˴������� = 0 Then
        If mrsBill!���շ� = 0 Then
            MsgBox "�ô�����δ�շѻ���ʣ�����ִ�з��ϲ�����", vbInformation, gstrSysName
            Exit Function
        End If
        
        gstrSQL = "Select ����Ա���� " & _
            "   From ������ü�¼ " & _
            "   Where ID =( Select ����ID From ҩƷ�շ���¼ Where ����� Is Null And Mod(��¼״̬,3)=1  And No=[1] And ����=[2] And Rownum=1)"
        
        If int���� = 2 Then
            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        End If
        
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, int����)
        
        With rsCheck
            If IsNull(!����Ա����) Then
                MsgBox "�ô�����δ��ˣ�����ִ�з��ϲ�����", vbInformation, gstrSysName
                Exit Function
            End If
        End With
    End If
    
    With Msf�����б�
        .Redraw = False
        .TextMatrix(.Rows - 1, 0) = mrsBill!����
        .TextMatrix(.Rows - 1, 1) = mrsBill!NO
        .TextMatrix(.Rows - 1, 2) = IIf(IsNull(mrsBill!����), "", mrsBill!����)
        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(mrsBill!����), "", mrsBill!����)
        .TextMatrix(.Rows - 1, 4) = IIf(IsNull(mrsBill!סԺ��), "", mrsBill!סԺ��)
        .TextMatrix(.Rows - 1, 5) = IIf(IsNull(mrsBill!����), "", mrsBill!����)
        .TextMatrix(.Rows - 1, 6) = IIf(IsNull(mrsBill!������), "", mrsBill!������)
        .TextMatrix(.Rows - 1, 7) = IIf(IsNull(mrsBill!����ҽ��), "", mrsBill!����ҽ��)
        .TextMatrix(.Rows - 1, 8) = IIf(IsNull(mrsBill!��������), "", mrsBill!��������)
        .TextMatrix(.Rows - 1, 9) = mrsBill!����
        .RowData(.Rows - 1) = mrsBill!����
        mstr���ݺ� = mrsBill!NO
        mint�������� = mrsBill!����
    End With
    
    If err <> 0 Then
        MsgBox "д�����б�ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        With Msf�����б�
            If .Rows - 1 >= 2 Then
                .Rows = .Rows - 1
            Else
                .TextMatrix(.Rows - 1, 0) = ""
                .TextMatrix(.Rows - 1, 1) = ""
                .TextMatrix(.Rows - 1, 2) = ""
                .TextMatrix(.Rows - 1, 3) = ""
                .TextMatrix(.Rows - 1, 4) = ""
                .TextMatrix(.Rows - 1, 5) = ""
                .TextMatrix(.Rows - 1, 6) = ""
                .TextMatrix(.Rows - 1, 7) = ""
                .TextMatrix(.Rows - 1, 8) = ""
                .TextMatrix(.Rows - 1, 9) = ""
                .RowData(.Rows - 1) = 0
            End If
            .Redraw = True
        End With
        Exit Function
    End If
    
    With Msf�����б�
        .Rows = .Rows + 1
        .RowData(.Rows - 1) = 0
        .Redraw = True
    End With
    
    WriteSendListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function RefreshData() As Boolean
    Dim intRow As Integer, intRows As Integer
    Dim arrID
    Dim strNoThis As String, intBillThis As Integer
    Dim str��ϸ��λ�� As String, str���ܵ�λ�� As String
    
    On Error GoTo errHandle
    If mblnModify = False Then Exit Function
    RefreshData = False
    
    '��ջ��ܱ��
    With Msf��������
        .Clear
        .Rows = 2
        SetFormat (3)
    End With
    
    gstrSQL = ""
    If mbln��Ʊ�ݺŷ��� Then
            With Msf�����б�
                    '.TextMatrix(0, 0) = "����"
                    '.TextMatrix(0, 1) = "NO"
                    '.TextMatrix(0, 2) = "����"
                    '.TextMatrix(0, 3) = "����"
                    '.TextMatrix(0, 4) = "סԺ��"
                    '.TextMatrix(0, 5) = "����"
                    '.TextMatrix(0, 6) = "�շ�Ա"
                    '.TextMatrix(0, 7) = "����ҽ��"
                    '.TextMatrix(0, 8) = "��������"
                For intRow = 1 To .Rows - 1
                            
                    If Trim(.TextMatrix(intRow, 1)) <> "" Then
                            '����SQL���
                            gstrSQL = gstrSQL & " UNION ALL  SELECT " & .RowData(intRow) & " as ����,'" & Trim(.TextMatrix(intRow, 1)) & "' as NO From DUAL" & vbCrLf
                    End If
                Next
            End With
            
        If gstrSQL = "" Then Exit Function
        gstrSQL = Mid(gstrSQL, Len(" UNION ALL "))
        gstrSQL = "" & _
            "   Select NO,ҩƷID,����,���ۼ�,ʵ������,����,���۽�� " & _
            "   From ҩƷ�շ���¼ " & _
            "   Where (����,No) in (" & gstrSQL & ") And Mod(��¼״̬,3)=1 And ����� Is Null And (�ⷿID+0=[3] Or �ⷿID Is NULL)"
    Else
        gstrSQL = "" & _
        "   Select NO,ҩƷID,����,���ۼ�,ʵ������,����,���۽�� " & _
        "   From ҩƷ�շ���¼ " & _
        "   Where No=[1] And ����=[2]" & _
        "            And Mod(��¼״̬,3)=1 And ����� Is Null And (�ⷿID+0=[3] Or �ⷿID Is NULL)"
    End If
    
    '��ʾ��������
    Select Case mintUnit
    Case 0
        str��ϸ��λ�� = "C.���㵥λ ��λ,B.���ۼ� ����,B.ʵ������*Nvl(B.����,1) ����"
        str���ܵ�λ�� = "C.���㵥λ ��λ,B.���ۼ� ����,Sum(B.ʵ������*Nvl(B.����,1)) ����"
    Case Else
        str��ϸ��λ�� = "D.��װ��λ ��λ,B.���ۼ�*nvl(D.����ϵ��,1) ����,B.ʵ������/nvl(D.����ϵ��,1)*Nvl(B.����,1) ����"
        str���ܵ�λ�� = "D.��װ��λ ��λ,B.���ۼ�*nvl(D.����ϵ��,1) ����,Sum(B.ʵ������/nvl(D.����ϵ��,1)*Nvl(B.����,1)) ����"
    End Select
    
    str��ϸ��λ�� = str��ϸ��λ�� & ",B.���۽�� ��� "
    str���ܵ�λ�� = str���ܵ�λ�� & ",Sum(B.���۽��) ��� "
    
    
    gstrSQL = "Select Distinct D.*,'['||D.����||']'||D.ͨ������  As Ʒ��" & _
             " From (   SELECT B.NO,D.����ID,C.����,C.���� ͨ������,NVL(B.����,0) ����," & _
             "                  DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)) ���," & str���ܵ�λ�� & _
             "          FROM (" & gstrSQL & ") B, �������� D,�շ���ĿĿ¼ C " & _
             "          WHERE B.ҩƷID+0=D.����ID AND D.����ID=C.ID" & _
             "          GROUP BY B.NO,D.����ID,C.����,C.����,NVL(B.����,0)," & _
             "                 DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)),"
    
    Select Case mintUnit
    Case 0
        gstrSQL = gstrSQL & "C.���㵥λ,B.���ۼ�"
    Case Else
        gstrSQL = gstrSQL & "D.��װ��λ,B.���ۼ�*nvl(D.����ϵ��,1)"
    End Select
    gstrSQL = gstrSQL & ") D"
    gstrSQL = gstrSQL & " Order By D.����"
    
'    err = 0: On Error Resume Next
    Set mrsTotal = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr���ݺ�, mint��������, Val(Me.cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex)))
    
    If mbln��Ʊ�ݺŷ��� Then
        'ɾ����ǰ�ĵ���
        With mrs����������ϸ
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                If Not .EOF Then .Delete
                If Not .EOF Then .MoveNext
            Loop
        End With
    End If
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
    Dim Dbl��� As Double
    '����������װ��
    On Error Resume Next
    err = 0
    WriteTotalDataToBill = False
    With Msf��������
        .Clear
        .Rows = 2
        Call SetFormat(3)
    End With
    
    '��䵥������
    Dbl��� = 0
    
    If mrsTotal.State = 0 Then Exit Function
    
    If mrsTotal.RecordCount > 0 Then
        Do While Not mrsTotal.EOF
            With mrs����������ϸ
                .AddNew
                !���ݺ� = mrsTotal!NO
                !�������� = mrsTotal!Ʒ��
                !���� = mrsTotal!����
                !��� = IIf(IsNull(mrsTotal!���), "", mrsTotal!���)
                !��λ = IIf(IsNull(mrsTotal!��λ), "", mrsTotal!��λ)
                !���� = mrsTotal!����
                !���� = mrsTotal!����
                !��� = mrsTotal!���
                !����ID = mrsTotal!����ID
                !���� = mrsTotal!����
            End With
            mrsTotal.MoveNext
        Loop
    End If
    
    With mrs����������ϸ
        If .RecordCount <> 0 Then
            .Sort = "����,����"
            .MoveFirst
        End If
        Do While Not .EOF
            If Msf��������.Rows = 2 And Msf��������.TextMatrix(1, 1) = "" Then
                Msf��������.TextMatrix(Msf��������.Rows - 1, 0) = Msf��������.Rows - 1
                Msf��������.TextMatrix(Msf��������.Rows - 1, 1) = !��������
                Msf��������.TextMatrix(Msf��������.Rows - 1, 2) = IIf(IsNull(!���), "", !���)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 3) = IIf(IsNull(!��λ), "", !��λ)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 4) = Format(!����, mFMT.FM_���ۼ�)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 5) = Format(!����, mFMT.FM_����)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 6) = Format(!���, mFMT.FM_���)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 7) = !����ID
                Msf��������.TextMatrix(Msf��������.Rows - 1, 8) = !����
                Msf��������.MergeRow(Msf��������.Rows - 1) = False
            ElseIf Msf��������.TextMatrix(Msf��������.Rows - 1, 7) <> !����ID Then
                Msf��������.Rows = Msf��������.Rows + 1
                Msf��������.TextMatrix(Msf��������.Rows - 1, 0) = Msf��������.Rows - 1
                Msf��������.TextMatrix(Msf��������.Rows - 1, 1) = !��������
                Msf��������.TextMatrix(Msf��������.Rows - 1, 2) = IIf(IsNull(!���), "", !���)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 3) = IIf(IsNull(!��λ), "", !��λ)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 4) = Format(!����, mFMT.FM_���ۼ�)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 5) = Format(!����, mFMT.FM_����)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 6) = Format(!���, mFMT.FM_���)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 7) = !����ID
                Msf��������.TextMatrix(Msf��������.Rows - 1, 8) = !����
                Msf��������.MergeRow(Msf��������.Rows - 1) = False
            ElseIf Msf��������.TextMatrix(Msf��������.Rows - 1, 8) <> !���� Then
                Msf��������.Rows = Msf��������.Rows + 1
                Msf��������.TextMatrix(Msf��������.Rows - 1, 0) = Msf��������.Rows - 1
                Msf��������.TextMatrix(Msf��������.Rows - 1, 1) = !��������
                Msf��������.TextMatrix(Msf��������.Rows - 1, 2) = IIf(IsNull(!���), "", !���)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 3) = IIf(IsNull(!��λ), "", !��λ)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 4) = Format(!����, mFMT.FM_���ۼ�)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 5) = Format(!����, mFMT.FM_����)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 6) = Format(!���, mFMT.FM_���)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 7) = !����ID
                Msf��������.TextMatrix(Msf��������.Rows - 1, 8) = !����
                Msf��������.MergeRow(Msf��������.Rows - 1) = False
            Else
                Msf��������.TextMatrix(Msf��������.Rows - 1, 5) = Format(CDbl(IIf(Msf��������.TextMatrix(Msf��������.Rows - 1, 5) = "", 0, Msf��������.TextMatrix(Msf��������.Rows - 1, 5))) + !����, mFMT.FM_����)
                Msf��������.TextMatrix(Msf��������.Rows - 1, 6) = Format(CDbl(IIf(Msf��������.TextMatrix(Msf��������.Rows - 1, 6) = "", 0, Msf��������.TextMatrix(Msf��������.Rows - 1, 6))) + !���, mFMT.FM_���)
            End If
            Dbl��� = Dbl��� + !���
            .MoveNext
        Loop
        
        '��ʾ�ϼ�
        Msf��������.Rows = Msf��������.Rows + 1
        Msf��������.TextMatrix(Msf��������.Rows - 1, 0) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.Rows - 1, 1) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.Rows - 1, 2) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.Rows - 1, 3) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.Rows - 1, 4) = "�ϼ�"
        Msf��������.TextMatrix(Msf��������.Rows - 1, 5) = Format(Dbl���, mFMT.FM_���)
        Msf��������.TextMatrix(Msf��������.Rows - 1, 6) = Format(Dbl���, mFMT.FM_���)
        Msf��������.MergeCells = flexMergeFree
        Msf��������.MergeRow(Msf��������.Rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "��ʾ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    WriteTotalDataToBill = True
End Function

Private Function WriteDataToBill() As Boolean
    Dim dbl�ϼƽ�� As Double
    '--��ʾָ����������ϸ--
    On Error Resume Next
    err = 0
    
    WriteDataToBill = False
    With Msf������ϸ
        .Clear
        .Rows = 2
        Call SetFormat(2)
    End With
    dbl�ϼƽ�� = 0
    
    '��䵥������
    With mrsBill
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Msf������ϸ.MergeRow(.AbsolutePosition) = False
            Msf������ϸ.TextMatrix(.AbsolutePosition, 0) = !Ʒ��
            Msf������ϸ.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!���), "", !���)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!��λ), "", !��λ)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 3) = Format(!����, mFMT.FM_���ۼ�)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 4) = Format(!����, mFMT.FM_����)
            Msf������ϸ.TextMatrix(.AbsolutePosition, 5) = Format(!���, mFMT.FM_���)
            dbl�ϼƽ�� = dbl�ϼƽ�� + Val(!���)
            
            If .AbsolutePosition >= Msf������ϸ.Rows - 1 Then Msf������ϸ.Rows = Msf������ϸ.Rows + 1
            .MoveNext
        Loop
    End With
    With Msf������ϸ
        .TextMatrix(.Rows - 1, 0) = "�ϼ�"
        .TextMatrix(.Rows - 1, 1) = "�ϼ�"
        .TextMatrix(.Rows - 1, 2) = "�ϼ�"
        .TextMatrix(.Rows - 1, 3) = "�ϼ�"
        .TextMatrix(.Rows - 1, 4) = Format(dbl�ϼƽ��, mFMT.FM_���)
        .TextMatrix(.Rows - 1, 5) = Format(dbl�ϼƽ��, mFMT.FM_���)
        .MergeCells = flexMergeFree
        .MergeRow(.Rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "��ʾ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    WriteDataToBill = True
End Function

Private Function SetLocateBill(ByVal strNo As String, ByVal intBillType As Integer, Optional ByVal BlnEnterCell As Boolean = True) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ҵ�ָ�������Ƿ����
    '����:strNo-���ݺ�
    '     intBillType-��������
    '     BlnEnterCell-�Ƿ��������б�
    '����:�ҵ��˷���true,���򷵻�false
    Dim intRow As Integer
    
    SetLocateBill = False
    With Msf�����б�
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) = strNo And intBillType = .RowData(intRow) Then
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
    Dim rsCheckStock As New ADODB.Recordset
    Dim dblStock As Double
    Dim strSubSql As String
    Dim n As Integer
    On Error GoTo errHandle
    
    '�����
    If mIntCheckStock = 0 Then CheckStock = True: Exit Function
    
    '���������ת��Ϊ��Ӧ��λ��ʵ������
    Select Case mintUnit
    Case 0
        strSubSql = "/1"
    Case Else
        strSubSql = "/Decode(B.����ϵ��,0,1,null,1,b.����ϵ��)"
    End Select
    
    CheckStock = False
    If Msf�����б�.TextMatrix(1, 1) <> "" Then
        For n = 1 To Msf��������.Rows - 2
            
            gstrSQL = "" & _
                "   Select nvl(ʵ������,0)" & strSubSql & " AS ����" & _
                "   From ҩƷ��� A,�������� B" & _
                "   Where A.ҩƷID=B.����ID And A.����=1 And A.�ⷿID=[3]" & _
                "           And A.ҩƷID=[1] And Nvl(A.����,0)=[2]"
        
            Set rsCheckStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Msf��������.TextMatrix(n, 7)), Val(Msf��������.TextMatrix(n, 8)), mlng���ϲ���ID)
            With rsCheckStock
                If .EOF Then
                    dblStock = 0
                Else
                    dblStock = !����
                End If
                
                If dblStock < Msf��������.TextMatrix(n, 5) Then
                    If Msf��������.TextMatrix(n, 8) <> 0 Then
                        MsgBox Msf��������.TextMatrix(n, 1) & "�����ο�������������ܼ������ϣ�", vbInformation, gstrSysName: Exit Function
                    Else
                        Select Case mIntCheckStock
                        Case 1
                            If MsgBox(Msf��������.TextMatrix(n, 1) & "�Ŀ�����������Ƿ�������ϣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox Msf��������.TextMatrix(n, 1) & "�Ŀ�������������ܼ������ϣ�", vbInformation, gstrSysName: Exit Function
                        End Select
                    End If
                End If
            End With
        Next
    End If
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
    Dim strDate As String
    Dim strNo As String
    Dim str���� As String
    Dim arrSql As Variant
    Dim strNos As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    err = 0
    SendBill = False
    arrSql = Array()
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If CheckDate = False Then Exit Function
    
    
    With Msf�����б�
        For intRow = 1 To .Rows - 1
            If .RowData(intRow) <> 0 Then
                '��鴦��
                If CheckBill(.RowData(intRow), .TextMatrix(intRow, 1)) <> 0 Then
                    Exit Function
                End If
                strNo = Trim(.TextMatrix(intRow, 1))
                str���� = .RowData(intRow)
                
''                �޸ĵ��ݵĿⷿ
                gstrSQL = "Zl_ҩƷ�շ���¼_���Ŀⷿ("
                '�ֿⷿID
                gstrSQL = gstrSQL & mlng���ϲ���ID
                '����
                gstrSQL = gstrSQL & "," & .RowData(intRow)
                'NO
                gstrSQL = gstrSQL & ",'" & .TextMatrix(intRow, 1) & "'"
                'ԭ�ⷿID
                gstrSQL = gstrSQL & "," & Val(Me.cbo���ϲ���.ItemData(cbo���ϲ���.ListIndex))
                '����
                gstrSQL = gstrSQL & "," & .TextMatrix(intRow, 9)
                '��������
                gstrSQL = gstrSQL & ",to_date('" & .TextMatrix(intRow, 8) & "','yyyy-MM-dd')"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL

                '----:���Ϸ�ʽ_IN�����ݷ�ҩΪ1��������ҩΪ2�����ŷ�ҩΪ3
                '���̲���:�ⷿID_IN,����_IN,NO_IN,�����_IN,������_IN,У����_IN,���Ϸ�ʽ_IN,�������_IN
                gstrSQL = "zl_�����շ���¼_��������(" & _
                    mlng���ϲ���ID & "," & _
                    .RowData(intRow) & ",'" & _
                    .TextMatrix(intRow, 1) & "','" & _
                    gstrUserName & "','" & _
                    gstrUserName & "','NULL'," & _
                    2 & ",to_date('" & _
                    strDate & "','yyyy-MM-dd hh24:mi:ss'))"
                
                strNos = IIf(strNos = "", "", strNos & "|") & .RowData(intRow) & "," & .TextMatrix(intRow, 1)
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        Next
    End With
    
    gcnOracle.BeginTrans
        For intRow = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(intRow)), "cmdOk_Click")
        Next
    gcnOracle.CommitTrans
    
    '���÷��Ϻ����ҽӿ�
    If Not mobjPlugIn Is Nothing And strNos <> "" Then
        mobjPlugIn.StuffSendByRecipe mlng���ϲ���ID, strNos, CDate(strDate), strReserve
    End If
    
    SendBill = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckDate() As Boolean
'���ڷ�����ҩ������ʱ������Ƿ��ǵ���ĵ���
    Dim i As Integer
    Dim dateCur As Date
    
    dateCur = zlDatabase.Currentdate
    With Msf�����б�
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" Then
                If Format(.TextMatrix(i, 8), "YYYY-MM-DD") < Format(dateCur, "YYYY-MM-DD") Then
                    If MsgBox("        �����ǵ��쵥�ݣ���ɾ�������������»��ܣ�" & vbCrLf & "����Ѿ����˱���Ŀ�����Ҫ���³������Ƿ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        CheckDate = False
                    Else
                        CheckDate = True
                    End If
                    Exit Function
                End If
            End If
        Next
    End With
    
    CheckDate = True
End Function

Private Sub BillListPrint(Optional int���Ϸ�ʽ As Integer = 1, Optional strDate As String = "", Optional strNo As String = "", Optional str���� As String = "0")
    '���ݻ�����ӡ
    '���Ϸ�ʽ:1-��������;2-��������;3-���ŷ���
    ' intStyle:0-�����Ϸ�ʽ��ӡ,1-���ݴ�ӡ
    Dim bln���ϵ� As Boolean
    Dim bln�ѷ����嵥 As Boolean
    Dim bln���ݴ�ӡ As Boolean
    Dim strReg As String
    Dim intPrint As Integer '0-��ʾ��ӡ,1-�Զ���ӡ,<>0��1:����ӡ
    
    bln���ϵ� = InStr(1, mstrPrivs, "����֪ͨ��") <> 0
    bln�ѷ����嵥 = InStr(1, gstrPrivs, "��ӡ�ѷ����嵥") <> 0
    bln���ݴ�ӡ = InStr(1, gstrPrivs, "���ݴ�ӡ") <> 0
    
    If bln���ݴ�ӡ = False Then Exit Sub
    
    intPrint = Val(zlDatabase.GetPara("���ϴ�ӡ���ѷ�ʽ", glngSys, mlngModule, "0"))
    
    If intPrint = 0 Then
        '��ʾ��ӡ
        If MsgBox("����Ҫ��ӡ��ص�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    ElseIf intPrint = 1 Then
        '�Զ���ӡ
    Else
        Exit Sub
    End If
    Select Case int���Ϸ�ʽ
    Case 1  '������ӡ
        If strNo <> "" Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723", Me, "�ⷿ==" & mlng���ϲ���ID, "NO=" & strNo, "����=" & str����, "�����=����� is not null", 1)
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "�ⷿ=" & mlng���ϲ���ID, "���Ϸ�ʽ=���ݷ���|1", "���Ϻ�=" & strDate, 1)
        End If
    Case 2
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "�ⷿ=" & mlng���ϲ���ID, "���Ϸ�ʽ=��������|2", "���Ϻ�=" & strDate, 1)
    Case 3
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "�ⷿ=" & mlng���ϲ���ID, "���Ϸ�ʽ=���ŷ���|3", "���Ϻ�=" & strDate, 1)
    End Select
    
End Sub

Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng���� As Long, str��� As String
    '��鴦���Ƿ��ѽ��ʡ����ò����Ƿ��ѳ�Ժ������Ȩ�޽��м��
    '���޴˷���ļ��
'    With mrs���
'        If .RecordCount <> 0 Then .MoveFirst
'        Do While Not .EOF
'            StrNo = !���ݱ�ʶ
'            lng���� = Split(StrNo, "|")(1)
'            StrNo = Split(StrNo, "|")(0)
'            str��� = NVL(!���)
'            '���ޡ����˽��ʡ�������Ȩ�ޣ��������
'            'If Not IsReceiptBalance_Charge(mstrPrivs, lng����, StrNo, str���) Then Exit Function
'            '����Ժ����
'            If Not IsOutPatient(mstrPrivs, lng����, StrNo) Then Exit Function
'            .MoveNext
'        Loop
'    End With
'
    CheckCorrelation = True
End Function

Private Sub InitRec()
    Set mrs��� = New ADODB.Recordset
    With mrs���
        If .State = 1 Then .Close
        .Fields.Append "���ݱ�ʶ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 500, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set mrs������Դ���� = New ADODB.Recordset
    With mrs������Դ����
        If .State = 1 Then .Close
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��Դ����", adLongVarChar, 100, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set mrs����������ϸ = New ADODB.Recordset
    With mrs����������ϸ
        If .State = 1 Then .Close
        .Fields.Append "���ݺ�", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
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

    CheckDepend = False

    On Error GoTo ErrHand
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.���� || '-' || a.���� As ���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.���� And (a.վ��=[2] or a.վ�� is null) " & _
        "       AND b.���� ='W' " & _
        "       AND a.id = c.����id " & _
        "       AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
        " Order by a.���� || '-' || a.����"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ӧ�Ŀⷿ", UserInfo.Id, gstrNodeNo)

    If rsTemp.EOF Then
        rsTemp.Close
        Exit Function
    End If

    'װ�뷢�ϲ�������
    With cbo���ϲ���
        .Clear
'        mblnNoClick = True
        Do While Not rsTemp.EOF
            If rsTemp!Id <> mlng���ϲ���ID Then
                .AddItem rsTemp!����
                .ItemData(.NewIndex) = rsTemp!Id
            End If
            
            rsTemp.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
        rsTemp.Close
    End With
    CheckDepend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt����) = "" Then Exit Sub
    
    If Send������(3, Txt����.Text) = False Then Exit Sub
End Sub

Private Sub txtҽ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtҽ����) = "" Then Exit Sub
    
    If Send������(2, txtҽ����.Text) = False Then Exit Sub
End Sub
