VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm�����ŷ��� 
   Caption         =   "��������"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "Frm�����ŷ���.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7605
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   7
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   2580
      TabIndex        =   6
      Top             =   5190
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrintSet 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   1350
      TabIndex        =   5
      Top             =   5190
      Visible         =   0   'False
      Width           =   1100
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   2685
      Left            =   30
      TabIndex        =   2
      Top             =   2400
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4736
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "������ϸ(&D)"
      TabPicture(0)   =   "Frm�����ŷ���.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Msf������ϸ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "���ϻ���(&T)"
      TabPicture(1)   =   "Frm�����ŷ���.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf��������"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf������ϸ 
         Height          =   2265
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
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
         Left            =   -74940
         TabIndex        =   9
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
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
      Left            =   6420
      TabIndex        =   4
      Top             =   5190
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�����б� 
      Height          =   1755
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   7545
      _ExtentX        =   13309
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
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5130
      TabIndex        =   3
      Top             =   5190
      Width           =   1100
   End
   Begin VB.Label LblNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "δ�����κδ���"
      ForeColor       =   &H80000002&
      Height          =   180
      Left            =   3390
      TabIndex        =   11
      Top             =   240
      Width           =   4110
   End
   Begin VB.Label LblNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "Frm�����ŷ���"
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

Private Sub cmdOk_Click()
    If CheckStock = False Then Exit Sub
    If Not CheckCorrelation Then Exit Sub
    If SendBill = False Then Exit Sub
    
    mlngBillCount = 0
    lblNote.Caption = IIf(mlngBillCount = 0, "δ�����κδ���", "������" & mlngBillCount & "�Ŵ���")
    
    '��ʼ��
    mstrID = ""
    mstrBillNo = ""
    txtNO = ""
    
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
    cmdOK.Enabled = False
    txtNO.SetFocus
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
    HisRow.Add "����:" & Format(Sys.Currentdate, "yyyy��MM��dd��")
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
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
   ' mbln��Ʊ�ݺŷ��� = False
    If mbln��Ʊ�ݺŷ��� = True Then lblNO.Caption = "Ʊ�ݺ�": Me.Caption = "��Ʊ�ݺŷ���"
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
    
    With lblNote
        .Left = Me.ScaleWidth - .Width - 100
    End With
    
    With cmdHelp
        .Top = Me.ScaleHeight - .Height - 100
    End With
    With cmdPrintSet
        .Top = cmdHelp.Top
        .Left = cmdHelp.Left + cmdHelp.Width + 100
    End With
    With cmdPrint
        .Top = cmdHelp.Top
        .Left = cmdPrintSet.Left + cmdPrintSet.Width + 100
    End With
    
    With CmdCancel
        .Top = cmdHelp.Top
        .Left = Me.ScaleWidth - .Width - 100
    End With
    With cmdOK
        .Top = cmdHelp.Top
        .Left = CmdCancel.Left - .Width - 100
    End With
    
    With Msf�����б�
        .Height = (cmdOK.Top - 200 - .Top) / 2
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With TabShow
        .Top = Msf�����б�.Top + Msf�����б�.Height + 100
        .Height = cmdOK.Top - 100 - .Top
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
            
            cmdOK.Enabled = (.RowData(IIf(.Rows - 1 = 1, 1, .Rows - 2)) <> 0)
            lblNote.Caption = IIf(mlngBillCount = 0, "δ�����κδ���", "������" & mlngBillCount & "�Ŵ���")
        
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
    zlControl.TxtSelAll txtNO
End Sub

Private Function Send������() As Boolean
    '--------------------------------------------------------------------------------------------------------------------------------------------
    '--����:�������ŷ���
    '--����:
    '--����:���ϳɹ�,����true,���򷵻�False
    '--------------------------------------------------------------------------------------------------------------------------------------------

    Dim intYear As Integer, strYear As String
    Dim intRow As Integer
    Dim strNo As String, IntBill As Integer, ArrTmp, strTmp As String
    Dim strSql As String
    Dim int���� As Integer
    
    '--���������λ,�򰴹������--
    Me.txtNO = UCase(LTrim(Me.txtNO))
    Me.txtNO.Text = zlCommFun.GetFullNO(Me.txtNO.Text, 13)
    On Error GoTo ErrHandle
    gstrSQL = "" & _
             " Select /*+ Rule*/ Distinct Decode(C.����,24,'�շ�',25,'����',26,'���ʱ�') ����,C.No,C.����,A.���շ�," & _
             "      Decode(A.��ҩ��,Null,'','���ŷ���','',A.��ҩ��) ������,P.���� ����,decode(c.����,26,'',B.����) ����," & _
             "      Decode(c.����,26,'',B.��ʶ��)  סԺ��,decode(c.����,26,'','') ����,B.������ ����ҽ��,B.����Ա���� ������," & _
             "      To_Char(C.��������,'yyyy-MM-dd') ��������,0 ���� " & _
             " From δ��ҩƷ��¼ A,������ü�¼ B,ҩƷ�շ���¼ C,���ű� P,���ű� S " & _
             "     ,Table(cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
             " Where C.����ID=B.ID And B.��������ID+0=P.ID(+) And Nvl(C.�ⷿID,0)+0=S.ID(+) " & _
             "     And Nvl(A.�ⷿID,0)=Nvl(C.�ⷿID,0) And Mod(C.��¼״̬,3)=1 And A.No=C.No " & _
             "     And (C.�ⷿID+0=[2] OR C.�ⷿID IS NULL)" & _
             "     And C.����=D.Column_Value And C.����� Is Null " & _
             "     And C.����=A.���� And C.No=[1] and nvl(C.��ҩ��ʽ,-999)<>-1 And Nvl(B.����״̬,0)<>1 "
     
    If mstr����IN = "24" Then
    ElseIf mstr����IN = "26" Then
        gstrSQL = Replace(gstrSQL, "0 ����", "1 ����")
        gstrSQL = Replace(gstrSQL, "B.����", "nvl(R.����,B.����)")
        gstrSQL = Replace(gstrSQL, "decode(c.����,26,'','') ����", "decode(c.����,26,'',B.����) ����")
        gstrSQL = Replace(gstrSQL, "������ü�¼ B", "סԺ���ü�¼ B,������ҳ R")
        gstrSQL = Replace(gstrSQL, "And Nvl(B.����״̬,0)<>1", "And B.����id=R.����id And B.��ҳid=R.��ҳid")
    ElseIf InStr(1, mstr����IN, "25") > 0 Or InStr(1, mstr����IN, "26") > 0 Then
        strSql = Replace(gstrSQL, "0 ����", "1 ����")
        strSql = Replace(strSql, "B.����", "nvl(R.����,B.����)")
        strSql = Replace(strSql, "decode(c.����,26,'','') ����", "decode(c.����,26,'',B.����) ����")
        strSql = Replace(strSql, "������ü�¼ B", "סԺ���ü�¼ B,������ҳ R")
        strSql = Replace(strSql, "And Nvl(B.����״̬,0)<>1", "And B.����id=R.����id And B.��ҳid=R.��ҳid")
        gstrSQL = gstrSQL & " Union All " & strSql
    End If
    
'    err = 0: On Error Resume Next
    Set mrsBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtNO, mlng���ϲ���ID, mstr����IN)
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
        mrsBill.Find "����=" & IntBill & " And No=" & strNo
        int���� = mrsBill!����
    Else
        strNo = mrsBill!NO
        IntBill = mrsBill!����
        int���� = mrsBill!����
    End If
    Me.txtNO.Tag = IntBill
    
    '����Ѵ��ڸõ��ݣ����˳�
    If SetLocateBill(txtNO.Text, IntBill, False) Then
        MsgBox "�ô����Ѿ����룬�����䣡", vbInformation, gstrSysName
        GoTo ExitSub
    End If
    
    '���Ϸ���
    If CheckBill(IntBill, strNo) <> 0 Then GoTo ExitSub
    '�����ǰ���봦���Ŀ�������¼��Ĵ����Ŀ��Ҳ�ͬ���������ʾ
    If CheckSource(IntBill, strNo) = False Then Exit Function
    If WriteSendListData(IntBill, strNo, int����) = False Then GoTo ExitSub
    
    mlngBillCount = mlngBillCount + 1
    lblNote.Caption = IIf(mlngBillCount = 0, "δ�����κδ���", "������" & mlngBillCount & "�Ŵ���")
    
    '��λ���ղ�����Ĵ�����
    Call SetLocateBill(txtNO.Text, Val(txtNO.Tag))
    
    With Msf�����б�
        cmdOK.Enabled = (.RowData(IIf(.Rows - 1 = 1, 1, .Rows - 2)) <> 0)
    End With
    
    mblnModify = True
    Call RefreshData
    With txtNO
        .SelStart = 0
        .SelLength = Len(txtNO)
    End With
    Send������ = True
    Exit Function
ExitSub:
    With txtNO
        .SelStart = 0
        .SelLength = Len(txtNO)
        .SetFocus
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtNO) = "" Then Exit Sub
    If ��Ʊ�ݺŷ��� = True Then
        If Send��Ʊ�� = False Then Exit Sub
    Else
        If Send������() = False Then Exit Sub
    End If
    
End Sub
Private Function Send��Ʊ��() As Boolean
    Dim blnAdd As Boolean
    
    Dim strNo As String, IntBill As Integer
    Dim rsƱ�� As New ADODB.Recordset
    txtNO.Text = Trim(UCase(txtNO.Text))
    
    On Error GoTo ErrHandle
    '���������Ʊ�ݺ���ȡ����
    gstrSQL = "Select Distinct A.No " & _
             " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B " & _
             " Where A.ID=B.��ӡID And A.��������=1 " & _
             " And B.Ʊ��=1 And B.����=[1]"
    Set rsƱ�� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[���������Ʊ�ݺ���ȡ����]", txtNO.Text)
    
    If rsƱ��.RecordCount = 0 Then
        MsgBox "û���ҵ��κδ�����", vbInformation, gstrSysName
        GoTo ExitSub
        Exit Function
    End If
    
    With rsƱ��
        Do While Not .EOF
            gstrSQL = " Select Distinct Decode(C.����,24,'�շ�','����') ����,C.No,C.����,A.���շ�,Decode(A.��ҩ��,Null,'','���ŷ�ҩ','',A.��ҩ��) ��ҩ��,P.���� ����,B.����,B.��ʶ�� סԺ��,'' ����,B.������ ����ҽ��,B.����Ա���� ������,To_Char(C.��������,'yyyy-MM-dd') ��������,0 ���� " & _
                      " From δ��ҩƷ��¼ A,������ü�¼ B,ҩƷ�շ���¼ C,���ű� P,���ű� S " & _
                      " Where C.����ID=B.ID And B.��������ID+0=P.ID(+) And Nvl(C.�ⷿID,0)+0=S.ID(+) and Nvl(A.�ⷿID,0)=Nvl(C.�ⷿID,0) And Mod(C.��¼״̬,3)=1 And A.No=C.No " & _
                      "     And (C.�ⷿID+0=[2] OR C.�ⷿID IS NULL)" & _
                      "     And C.���� =24 And C.����� Is Null And C.����=A.���� And C.No=[1] and nvl(C.��ҩ��ʽ,-999)<>-1 "
                  
            Set mrsBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(Nvl(!NO)), mlng���ϲ���ID)
            blnAdd = (mrsBill.RecordCount <> 0)
            
            If blnAdd Then     '�ҵ�ָ������
                strNo = mrsBill!NO
                IntBill = mrsBill!����
                txtNO.Tag = IntBill
                
                '����Ѵ��ڸõ��ݣ����˳�
                blnAdd = Not SetLocateBill(strNo, IntBill, False)
                '���Ϸ���
                If blnAdd Then blnAdd = Not (CheckBill(IntBill, strNo) <> 0)
                If blnAdd Then blnAdd = WriteSendListData(IntBill, strNo, 0)
                If blnAdd Then
                    mlngBillCount = mlngBillCount + 1
                    lblNote.Caption = IIf(mlngBillCount = 0, "δ�����κδ���", "������" & mlngBillCount & "�Ŵ���")
                End If
            End If
            .MoveNext
        Loop
    End With
    
    '��λ���ղ�����Ĵ�����
    Call SetLocateBill(strNo, True)
    
    With Msf�����б�
        cmdOK.Enabled = (.RowData(IIf(.Rows - 1 = 1, 1, .Rows - 2)) <> 0)
    End With
    mblnModify = True
    Call RefreshData
    Send��Ʊ�� = True
    Exit Function
ExitSub:
    With txtNO
        .SelStart = 0
        .SelLength = Len(txtNO)
        .SetFocus
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckSource(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim bln�ظ����� As Boolean
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select B.���� as ����,B.���� as ��Դ���� " & _
        "   From ҩƷ�շ���¼ A,���ű� B " & _
        "   Where A.�Է�����id=B.id and No=[1] And ����=[2]" & _
        "           And Mod(��¼״̬,3)=1 And ����� Is Null And (�ⷿID+0=[3] Or �ⷿID Is NULL) And Rownum<2"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "���", strNo, int����, mlng���ϲ���ID)
    
    
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadData(ByVal StrQuery As String) As Boolean
    '--��ȡ����--

'    On Error Resume Next
'    err = 0
    On Error GoTo ErrHandle
    ReadData = False

    gstrSQL = StrQuery
    Call zlDatabase.OpenRecordset(mrsBill, gstrSQL, Me.Caption)
    If err <> 0 Then
        MsgBox "��ȡ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    ReadData = True
    Exit Function
ErrHandle:
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
    On Error GoTo ErrHandle
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
    
    If int���� = 1 Then
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    
    gstrSQL = gstrSQL & " Order by ���"

    Set mrsBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, BillNo, BillStyle, mlng���ϲ���ID)

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
ErrHandle:
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
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select A.��ҩ�� ������,A.�����,nvl(B.���շ�,0) ���շ� " & _
        "   From ҩƷ�շ���¼ A,δ��ҩƷ��¼ B " & _
        "  Where A.No=B.No And A.����=B.���� And A.No=[1] And A.����=[2]" & _
        "           And mod(A.��¼״̬,3)=1 And Rownum=1 And (A.�ⷿID+0=[3] Or A.�ⷿID Is NULL)"
    gstrSQL = gstrSQL & " And A.����� IS Null"
    
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, intBillStyle, mlng���ϲ���ID)
    
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteSendListData(ByVal int���� As Integer, ByVal strNo As String, ByVal int���� As Integer) As Boolean
    Dim rsCheck As New ADODB.Recordset
'    On Error Resume Next
'    err = 0
    On Error GoTo ErrHandle
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
        
        If int���� = 1 Then
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
ErrHandle:
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
    
    On Error GoTo ErrHandle
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
    
    err = 0: On Error Resume Next
    Set mrsTotal = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr���ݺ�, mint��������, mlng���ϲ���ID)
    
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
ErrHandle:
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
    
    '�����
    On Error GoTo ErrHandle
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
ErrHandle:
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
    Dim lngCount As Long
    Dim strNos As String
    Dim strReserve As String
    
    On Error GoTo ErrHand
    err = 0
    SendBill = False
    
    strDate = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
    lngCount = 0
    gcnOracle.BeginTrans
    With Msf�����б�
        For intRow = 1 To .Rows - 1
            If .RowData(intRow) <> 0 Then
                '��鴦��
                If CheckBill(.RowData(intRow), .TextMatrix(intRow, 1)) <> 0 Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
                strNo = Trim(.TextMatrix(intRow, 1))
                str���� = .RowData(intRow)
                '----:���Ϸ�ʽ_IN�����ݷ�ҩΪ1��������ҩΪ2�����ŷ�ҩΪ3
                '���̲���:�ⷿID_IN,����_IN,NO_IN,�����_IN,������_IN,У����_IN,���Ϸ�ʽ_IN,�������_IN
                gstrSQL = "zl_�����շ���¼_��������(" & _
                    mlng���ϲ���ID & "," & _
                    .RowData(intRow) & ",'" & _
                    .TextMatrix(intRow, 1) & "','" & _
                    gstrUserName & "','" & _
                    gstrUserName & "',NULL," & _
                    2 & ",to_date('" & _
                    strDate & "','yyyy-MM-dd hh24:mi:ss'))"
                
                strNos = IIf(strNos = "", "", strNos & "|") & .RowData(intRow) & "," & .TextMatrix(intRow, 1)
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���Ϸ���")
               lngCount = lngCount + 1
            End If
        Next
    End With
    gcnOracle.CommitTrans
    
    If lngCount = 0 Then
    Else
        If lngCount = 1 Then
            Call BillListPrint(1, strDate, strNo, str����)
        Else
            Call BillListPrint(2, strDate)
        End If
    End If
    
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
