VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.2#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm���ñ���_�ֵ���ϸ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ֵ���ϸ"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "frm���ñ���_�ֵ���ϸ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2970
      TabIndex        =   1
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4230
      TabIndex        =   2
      Top             =   3210
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   5371
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
End
Attribute VB_Name = "frm���ñ���_�ֵ���ϸ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln���� As Boolean
Private mrsData As New ADODB.Recordset
Private Const strFormat_��� As String = "#####0.00;-#####0.00; ;"

Public Sub ShowME(Optional bln���� As Boolean = False, Optional ByVal rsData As ADODB.Recordset = Nothing)
    On Error Resume Next
    mbln���� = bln����
    Set mrsData = rsData
    Me.Show 1
End Sub

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strInput As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With Bill
        If .TxtVisible = False Then Exit Sub
        strInput = Val(.Text)
        If Not IsNumeric(strInput) Then
            MsgBox "�����к��зǷ��ַ���", vbInformation, gstrSysName
            Cancel = True
            .TxtSetFocus
            Exit Sub
        End If
        If Val(strInput) < 0 Then
            MsgBox "������������㣡", vbInformation, gstrSysName
            Cancel = True
            .TxtSetFocus
            Exit Sub
        End If
        If Val(strInput) > 100 Then
            MsgBox "�������ܴ���100%��", vbInformation, gstrSysName
            Cancel = True
            .TxtSetFocus
            Exit Sub
        End If
        
        .Text = Format(.Text, strFormat_���)
        .TextMatrix(.Row, 3) = Format(Val(.TextMatrix(.Row, 2)) * Val(.Text) / 100, strFormat_���)
    End With
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lngRow As Long
    Dim curͳ��֧�� As Currency
    '������д�빫����¼��
    
    If Not mbln���� Then
        Set rs�ֵ�֧��_���� = New ADODB.Recordset
        With rs�ֵ�֧��_����
            If .State = 1 Then .Close
            .Fields.Append "����", adDouble, 10  '0:��ʾ����
            .Fields.Append "����", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
            .Fields.Append "����ͳ��", adDouble, 18, adFldIsNullable
            .Fields.Append "ͳ�ﱨ��", adDouble, 18, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
        
        curͳ��֧�� = 0
        With Bill
            For lngRow = 1 To .Rows - 1
                rs�ֵ�֧��_����.AddNew
                rs�ֵ�֧��_����!���� = lngRow
                rs�ֵ�֧��_����!���� = Val(.TextMatrix(lngRow, 1))
                rs�ֵ�֧��_����!���� = .TextMatrix(lngRow, 0)
                rs�ֵ�֧��_����!����ͳ�� = Val(.TextMatrix(lngRow, 2))
                rs�ֵ�֧��_����!ͳ�ﱨ�� = Val(.TextMatrix(lngRow, 3))
                rs�ֵ�֧��_����.Update
                curͳ��֧�� = curͳ��֧�� + Val(.TextMatrix(lngRow, 3))
            Next
        End With
        gComInfo_üɽ.ͳ��֧�� = curͳ��֧�� + gComInfo_üɽ.ʵ�ʱ���
        gComInfo_üɽ.ͳ���Ը� = gComInfo_üɽ.����ͳ�� - curͳ��֧��
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsObj As New ADODB.Recordset
    On Error Resume Next
    
    If mbln���� Then
        Set rsObj = mrsData
    Else
        Set rsObj = rs�ֵ�֧��_����
    End If
    
    With Bill
        .ClearBill
        .Active = True
        .Rows = 1 + rsObj.RecordCount
        .Cols = 4
        .msfObj.FixedCols = 1
        
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����ͳ��"
        .TextMatrix(0, 3) = "�������"
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 800
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .msfObj.ColAlignmentFixed = 1
        .ColData(0) = 5
        .ColData(1) = 4
        .ColData(2) = 5
        .ColData(3) = 5
        
        .PrimaryCol = 1
        .LocateCol = 1
    End With
    
    With rsObj
        If .RecordCount <> 0 Then
            .Sort = "���� asc"
            .MoveFirst
        Else
            Unload Me
            Exit Sub
        End If
        
        Do While Not .EOF
            Bill.TextMatrix(.AbsolutePosition, 0) = !����
            Bill.TextMatrix(.AbsolutePosition, 1) = Format(!����, strFormat_���)
            Bill.TextMatrix(.AbsolutePosition, 2) = Format(!����ͳ��, strFormat_���)
            Bill.TextMatrix(.AbsolutePosition, 3) = Format(!ͳ�ﱨ��, strFormat_���)
            .MoveNext
        Loop
    End With
    
    Bill.AllowAddRow = False
    If mbln���� Then
        Bill.Active = False
        cmdȷ��.Visible = False
        cmdȡ��.Caption = "ȷ��(&O)"
    End If
End Sub
