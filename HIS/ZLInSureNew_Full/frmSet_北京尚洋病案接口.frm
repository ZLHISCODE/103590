VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSet_�������󲡰��ӿ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "frmSet_�������󲡰��ӿ�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5370
      TabIndex        =   2
      Top             =   1080
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5370
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   630
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit BillEdit1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5953
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
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������վ���Ŀ���գ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   210
      Width           =   1800
   End
End
Attribute VB_Name = "frmSet_�������󲡰��ӿ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr������Ŀ As String

Private Sub BillEdit1_cboClick(ListIndex As Long)
    BillEdit1.TextMatrix(BillEdit1.Row, 1) = BillEdit1.CboText
End Sub

Private Sub CMD����_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim strSave As String
    Dim intDO As Integer, intCOUNT As Integer
    
    '���ÿ����Ŀ�Ƿ�������
    intCOUNT = BillEdit1.Rows - 1
    For intDO = 1 To intCOUNT
        If Trim(BillEdit1.TextMatrix(intDO, 1)) = "" Then
            MsgBox "����Ŀδ������ҽ���Ķ��չ�ϵ�����飡", vbInformation, gstrSysName
            Exit Sub
        End If
    Next
    
    '��������'�в�ҩ,A-�в�ҩ|����ҩ...
    For intDO = 1 To intCOUNT
        If Trim(BillEdit1.TextMatrix(intDO, 1)) <> "" Then
            strSave = strSave & "|" & BillEdit1.TextMatrix(intDO, 0) & "," & BillEdit1.TextMatrix(intDO, 1)
        End If
    Next
    If strSave <> "" Then mstr������Ŀ = Mid(strSave, 2)
    
    '���浽ע�����
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & Me.Name, "������Ŀ", mstr������Ŀ)
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim arr������Ŀ
    Dim strSQL As String
    Dim intDO As Integer, intCOUNT As Integer
    Dim rsTemp As New ADODB.Recordset
    
    mstr������Ŀ = ""
    With BillEdit1
        .Rows = 2: .Cols = 2
        .Active = True
        .PrimaryCol = 0
        .LocateCol = 1
        
        .TextMatrix(0, 0) = "HIS����"
        .TextMatrix(0, 1) = "��������"
        .ColData(0) = 5
        .ColData(1) = 3
        .ColWidth(0) = 1800
        .ColWidth(1) = 1800
        
        .AddItem "A-��λ��"
        .AddItem "B-��ҩ��"
        .AddItem "C-�г�ҩ"
        .AddItem "D-�в�ҩ"
        .AddItem "E-������"
        .AddItem "F-���鲡���"
        .AddItem "G-�����"
        .AddItem "H-����"
        .AddItem "I-���Ʒ�"
        .AddItem "J-���Ʒ�"
        .AddItem "K-�����"
        .AddItem "M-������"
        .AddItem "N-������"
        .AddItem "O-Ӥ����"
        .AddItem "P-���̷�"
        .AddItem "R-Ѫ  ��"
        .AddItem "Y-�����"
        .AddItem "Z-��  ��"
    End With
    
    '��ȡ������Ŀ
    strSQL = " Select ����,���� From ������Ŀ"
    Call OpenRecordset(rsTemp, "��ȡ������Ŀ", strSQL)
    With rsTemp
        Do While Not .EOF
            BillEdit1.TextMatrix(.AbsolutePosition, 0) = !����
            .MoveNext
            
            If Not .EOF Then BillEdit1.Rows = BillEdit1.Rows + 1
        Loop
    End With
    
    '�в�ҩ,A-�в�ҩ|����ҩ...
    mstr������Ŀ = GetSetting("ZLSOFT", "˽��ģ��\" & Me.Name, "������Ŀ", "")
    arr������Ŀ = Split(mstr������Ŀ, "|")
    intCOUNT = UBound(arr������Ŀ)
    For intDO = 0 To intCOUNT
        rsTemp.MoveFirst
        rsTemp.Find "����='" & Split(arr������Ŀ(intDO), ",")(0) & "'"
        If rsTemp.EOF = False Then
            BillEdit1.TextMatrix(rsTemp.AbsolutePosition, 1) = Split(arr������Ŀ(intDO), ",")(1)
        End If
    Next
End Sub
