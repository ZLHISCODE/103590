VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form Frm���������� 
   Caption         =   "�����ݽ�������"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   Icon            =   "Frm����������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9690
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   1320
      TabIndex        =   17
      Top             =   4980
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   12
      TabIndex        =   16
      Top             =   4980
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8472
      TabIndex        =   13
      Top             =   5088
      Width           =   1100
   End
   Begin VB.Frame fraTop 
      Height          =   672
      Left            =   -12
      TabIndex        =   14
      Top             =   24
      Width           =   9660
      Begin VB.ComboBox cbo�շѵ��� 
         Height          =   300
         Left            =   852
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   288
         Width           =   1116
      End
      Begin VB.TextBox TxtNo 
         Height          =   300
         Left            =   2700
         TabIndex        =   3
         Top             =   276
         Width           =   1668
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   3
         Left            =   96
         TabIndex        =   0
         Top             =   336
         Width           =   720
      End
      Begin VB.Label LblNote 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����κδ���"
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   5412
         TabIndex        =   4
         Top             =   324
         Width           =   4116
      End
      Begin VB.Label LblNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2136
         TabIndex        =   2
         Top             =   336
         Width           =   540
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3540
      Left            =   -12
      TabIndex        =   5
      Top             =   744
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   6244
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      TxtCheck        =   -1  'True
      TxtCheck        =   -1  'True
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   360
      RowHeightMin    =   360
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
   Begin VB.Frame fraBillDown 
      Height          =   684
      Left            =   0
      TabIndex        =   15
      Top             =   4236
      Width           =   9624
      Begin VB.TextBox txtDown 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   7620
         TabIndex        =   11
         Text            =   "2009-09-09 23:59:59"
         Top             =   204
         Width           =   1944
      End
      Begin VB.TextBox txtDown 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   5124
         TabIndex        =   9
         Text            =   "������Ա"
         Top             =   204
         Width           =   1308
      End
      Begin VB.TextBox txtDown 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   888
         TabIndex        =   7
         Text            =   "[01]һ����"
         Top             =   204
         Width           =   3336
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ʱ��"
         Height          =   180
         Index           =   2
         Left            =   7140
         TabIndex        =   10
         Top             =   264
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   4668
         TabIndex        =   8
         Top             =   264
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   0
         Left            =   132
         TabIndex        =   6
         Top             =   264
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "����(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7188
      TabIndex        =   12
      Top             =   5088
      Width           =   1100
   End
End
Attribute VB_Name = "Frm����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mblnFirst As Boolean
Private mlng���ϲ��� As Long
Private mfrmMain As Form
Private mblnSucces As Boolean
Private mintUnit As Integer
Private Const mstrCaption As String = "�����ݽ�������"
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

Private mblnChange As Boolean
Private Enum mCol
    c_��� = 0
    c_ִ�п���
    c_����
    c_����
    c_����
    c_��Ŀ
    c_���
    c_����
    c_����
    c_����ϵ��
    c_��λ
    c_����
    c_����
    c_ԭʼ����
    c_������
    c_׼����
    c_������
    c_����
    c_���
End Enum
Private mintPreBillType  As Integer
Private Const mCols = 19
Private Const mlngModule = 1723

Private mobjPlugIn As Object             '��ҽӿڶ���

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Public Function ShowCard(ByVal frmMain As Form, ByVal lng���ϲ���ID As Long, ByVal strPrivs As String) As Boolean
    '------------------------------------------------------------------------------------------------------
    '����:��ָ���Ĵ������в������ϵ����
    '���:mfrmMain-������
    '     lng���ϲ���ID-���ϲ���ID
    '     strPrivs-Ȩ�޴�
    '����:
    '����:���ϳɹ�,����true,���򷵻�false
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/1
    '------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    mstrPrivs = strPrivs
    mlng���ϲ��� = lng���ϲ���ID
    Me.Show 1, frmMain
    ShowCard = mblnSucces
End Function

 

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
        Cancel = True
End Sub

Private Sub Bill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    With Bill
        Select Case .Col
            Case mCol.c_������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
        End Select
    End With
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Bill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        Select Case .Col
            
            Case mCol.c_������
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "������������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        MsgBox "�����������������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(Format(Val(strKey), mFMT.FM_����)) > Val(Format(Val(.TextMatrix(.Row, c_׼����)), mFMT.FM_����)) Then
                        MsgBox "�����������ܴ���׼������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If

                    If Abs(Val(strKey)) >= 10 ^ 11 - 1 Then
                        MsgBox "��������������(-" & (10 ^ 11 - 1) & " �� " & (10 ^ 11 - 1) & ") ֮��", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    .TextMatrix(.Row, .Col) = .Text
                End If
        End Select
    End With
End Sub

Private Sub cbo�շѵ���_Click()
    If mintPreBillType = cbo�շѵ���.ItemData(cbo�շѵ���.ListIndex) Then Exit Sub
    mintPreBillType = cbo�շѵ���.ItemData(cbo�շѵ���.ListIndex)
    With Bill
        .Rows = 2
        .ClearBill
        txtDown(0).Text = ""
        txtDown(1).Text = ""
        txtDown(2).Text = ""
    End With
End Sub

Private Sub cbo�շѵ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    With Bill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, mCol.c_��Ŀ) <> "" Then
                .TextMatrix(intRow, mCol.c_������) = ""
            End If
        Next
    End With
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    With Bill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, mCol.c_��Ŀ) <> "" Then
                .TextMatrix(intRow, mCol.c_������) = .TextMatrix(intRow, mCol.c_׼����)
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    '����
    Dim strDate As String
    
    Dim bln���ϵ� As Boolean
    If ISValied = False Then Exit Sub
    strDate = Format(Sys.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If Save����(strDate) = False Then Exit Sub
    bln���ϵ� = InStr(1, mstrPrivs, "����֪ͨ��") <> 0
    
    If bln���ϵ� Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "����ʱ��=" & strDate, "��λ=" & mintUnit + 1, 2)
    End If
    mblnSucces = True
    mblnChange = False
    MsgBox "���ϳɹ�", vbInformation + vbDefaultButton1, gstrSysName
    Bill.ClearBill
End Sub
Private Function Save����(ByVal strDate As String) As Boolean
    '������ϵ��������
    Dim strNo As String
    Dim lngId As Long
    Dim lngRow As Long
    Dim ����ID As Long
    Dim int�Զ����� As Integer
    Dim strReg As String
    Dim cllTemp As New Collection
    Dim dbl���� As Double
    Dim strReturnInfo As String
    Dim strReserve As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln�������� As Boolean
    Dim int�Զ�����_ԭʼֵ As Integer
    
    int�Զ�����_ԭʼֵ = IIf(Val(zlDatabase.GetPara("�Զ�����", glngSys, mlngModule)) = 1, 1, 0)
    
    Save���� = False
    err = 0
    
    With Bill
        For lngRow = 1 To .Rows - 1
                
                 
                If Trim(.TextMatrix(lngRow, mCol.c_��Ŀ)) <> "" And Val(.TextMatrix(lngRow, mCol.c_������)) <> 0 Then
                    int�Զ����� = int�Զ�����_ԭʼֵ
                    
                    If int�Զ����� <> 1 Then
                        '�ж��Ƿ񱸻�����
                        gstrSQL = " Select 1 From ҩƷ�շ���¼ Where ���� = 21 And ������� Is Not Null And ����id = (select ����id from ҩƷ�շ���¼ where id=[1]) And Rownum < 2 "
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ񱸻�����", Val(.RowData(lngRow)))
                        bln�������� = Not rsTemp.EOF
                    
                        '����Ǹ�ֵ����Ҳ�����Զ�����
                        If bln�������� Then int�Զ����� = 1
                    End If
                 
                 '   ���̲���:ID_IN,�����_IN,�������_IN,����_IN,Ч��_IN,����_IN,��������_IN,�Զ�����_IN(1-�Զ�����,0-���Զ�����)
                   gstrSQL = "zl_�����շ���¼_��������("
                   gstrSQL = gstrSQL & .RowData(lngRow) & ","
                   gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                   gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
                   gstrSQL = gstrSQL & "'" & Replace(.TextMatrix(lngRow, mCol.c_����), "(" & .TextMatrix(lngRow, mCol.c_����) & ")", "") & "',"
                   gstrSQL = gstrSQL & "NULL" & ","
                   gstrSQL = gstrSQL & "NULL" & ","
                   dbl���� = 0
                   If Val(.TextMatrix(lngRow, mCol.c_׼����)) = Val(.TextMatrix(lngRow, mCol.c_������)) Then
                            dbl���� = Val(.TextMatrix(lngRow, mCol.c_ԭʼ����))
                   Else
                    If mintUnit = 0 Then
                         dbl���� = Val(.TextMatrix(lngRow, mCol.c_������))
                    Else
                         dbl���� = Round(Val(.TextMatrix(lngRow, mCol.c_������)) * Val(.TextMatrix(lngRow, mCol.c_����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                    End If
                   End If
                   gstrSQL = gstrSQL & dbl����
                   gstrSQL = gstrSQL & "," & int�Զ����� & ")"
                   Call AddArray(cllTemp, gstrSQL)
                   
                   strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & NVL(.RowData(lngRow)) & "," & dbl����
                End If
        Next
    End With
    
    On Error GoTo ErrHand:
    
    Call ExecuteProcedureArrAy(cllTemp, mstrCaption)
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing And strReturnInfo <> "" Then
        mobjPlugIn.DrugReturnByID mlng���ϲ���, strReturnInfo, CDate(strDate), strReserve
    End If
    
    Save���� = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ISValied() As Boolean
    '------------------------------------------------------------------------------------------------------
    '����:������ݵĺϷ���
    '���:
    '����:
    '����:���ݺϷ�,����ture,���򷵻�False
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/2
    '------------------------------------------------------------------------------------------------------
    Dim intRow As Integer
    Dim blnHave As Boolean
    Dim dblTemp As Double
    blnHave = False
    With Bill
        For intRow = 1 To .Rows - 1
            If Trim(.TextMatrix(intRow, mCol.c_��Ŀ)) <> "" Then
                dblTemp = Val(.TextMatrix(intRow, mCol.c_������))
                If dblTemp > Val(.TextMatrix(intRow, mCol.c_׼����)) Then
                    ShowMsgBox "��������(��:" & Format(dblTemp, mFMT.FM_����) & ") ������׼������(��:" & Format(Val(.TextMatrix(intRow, mCol.c_׼����)), mFMT.FM_����) & "),����!"
                    .Row = intRow
                    .Col = c_������
                    .SetFocus
                    Exit Function
                End If
                If Abs(dblTemp) >= 10 ^ 11 - 1 Then
                    MsgBox "��������������(-" & (10 ^ 11 - 1) & " �� " & (10 ^ 11 - 1) & ") ֮��", vbInformation + vbOKOnly, gstrSysName
                    .Row = intRow
                    .Col = c_������
                    .TxtSetFocus
                    Exit Function
                End If
                If dblTemp <> 0 Then
                    blnHave = True
                End If
            End If
        Next
    End With
    If blnHave = False Then
        ShowMsgBox "�㻹δ������������,����!"
        Bill.Row = 1
        Bill.Col = c_������
        Bill.SetFocus
        Exit Function
    End If
    ISValied = True
End Function
Private Sub Form_Load()
    Dim strReg As String
    mblnFirst = True
    
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
  
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    Call initGrid

    RestoreWinState Me, App.ProductName, mstrCaption
    
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mintPreBillType = 0
    
    '��ʼ
    Call InitData
    cmdOK.Enabled = True
    mblnChange = False
    cmdAllCls.Visible = True
    cmdAllSel.Visible = True
    txtDown(0).Text = ""
    txtDown(1).Text = ""
    txtDown(2).Text = ""
    
End Sub
Private Sub InitData()
    '------------------------------------------------------------------------------------------------------
    '����:��ʼһЩ��Ҫ��������
    '���:
    '����:
    '����:
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/21
    '------------------------------------------------------------------------------------------------------
    Dim strReg As String
    
    strReg = Trim(zlDatabase.GetPara("������ϵ�������", glngSys, mlngModule, "25", Array(lbl(3), cbo�շѵ���), zlStr.IsHavePrivs(mstrPrivs, "��������")))
    If Val(strReg) = 0 Then strReg = "25"
    With cbo�շѵ���
        .Clear
        .AddItem "0-�շ�"
        .ItemData(.NewIndex) = 24
        If Val(strReg) = 24 Then .ListIndex = .NewIndex
        .AddItem "1-����"
        .ItemData(.NewIndex) = 25
        If Val(strReg) = 25 Then .ListIndex = .NewIndex
        .AddItem "2-���ʱ�"
        .ItemData(.NewIndex) = 26
        If Val(strReg) = 26 Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = 0
    End With

End Sub

Private Sub initGrid()
    '------------------------------------------------------------------------------------------------------
    '����:��ʼ����ؼ�
    '���:
    '����:
    '����:
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/1
    '------------------------------------------------------------------------------------------------------
     With Bill
        .Cols = mCols
      '  .MsfObj.FixedCols = 1
        .AllowAddRow = False
        .TextMatrix(0, c_���) = "���"
        .TextMatrix(0, c_��Ŀ) = "��Ŀ"
        .TextMatrix(0, c_���) = "���"
        .TextMatrix(0, c_����) = "����"
        .TextMatrix(0, c_����) = "����"
        .TextMatrix(0, c_����ϵ��) = "����ϵ��"
        .TextMatrix(0, c_��λ) = "��λ"
        .TextMatrix(0, c_����) = "����"
        
        .TextMatrix(0, c_����) = "����"
        .TextMatrix(0, c_ԭʼ����) = "ԭʼ����"
        .TextMatrix(0, c_������) = "������"
        .TextMatrix(0, c_׼����) = "׼����"
        .TextMatrix(0, c_������) = "������"
        .TextMatrix(0, c_����) = "����"
        .TextMatrix(0, c_���) = "���"
        .TextMatrix(0, c_ִ�п���) = "ִ�п���"
        .TextMatrix(0, c_����) = "����"
        .TextMatrix(0, c_����) = "����"
        .TextMatrix(0, c_����) = "����"
 
        .ColWidth(c_���) = 600
        .ColWidth(c_��Ŀ) = 2000
        .ColWidth(c_���) = 1000
        .ColWidth(c_����) = 0
        .ColWidth(c_����ϵ��) = 0
        .ColWidth(c_ԭʼ����) = 0
        .ColWidth(c_����) = 1000
        
        .ColWidth(c_��λ) = 1000
        .ColWidth(c_����) = 1000
        .ColWidth(c_����) = 1000
        .ColWidth(c_������) = 1000
        .ColWidth(c_׼����) = 1000
        
        .ColWidth(c_������) = 1000
        .ColWidth(c_����) = 1000
        .ColWidth(c_���) = 1000
        .ColWidth(c_ִ�п���) = 0
        .ColWidth(c_����) = 1000
        .ColWidth(c_����) = 1000
        .ColWidth(c_����) = 1000
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��
        .ColData(c_���) = 5
        .ColData(c_��Ŀ) = 5
        .ColData(c_���) = 5
        .ColData(c_����) = 5
        .ColData(c_����) = 5
        .ColData(c_����ϵ��) = 5
        
        .ColData(c_��λ) = 5
        .ColData(c_����) = 5
        .ColData(c_����) = 5
        .ColData(c_ԭʼ����) = 5
        .ColData(c_������) = 5
        .ColData(c_׼����) = 5
        .ColData(c_������) = 4
        .ColData(c_����) = 5
        .ColData(c_���) = 5
        .ColData(c_ִ�п���) = 5
        .ColData(c_����) = 5
        .ColData(c_����) = 5
        .ColData(c_����) = 5
            
        .PrimaryCol = c_��Ŀ
        .LocateCol = c_������
        .Active = True
    End With
End Sub
Private Function InitBill(ByVal strNo As String, ByVal IntBill As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------
    '����:��ʼ��������
    '���:strNO-������
    '    intBill-��������(24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ�������)
    '����:
    '����:��ʼ���ݳɹ�,����true,���򷵻�false
    '�޸���:���˺�
    '�޸�ʱ��:2007/1/25
    '------------------------------------------------------------------------------------------------------
    Dim rsBill As New ADODB.Recordset
    Dim strFields As String
    Dim lngRow As Long
    Dim str���� As String

    err = 0: On Error GoTo ErrHand:
    '24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ������ϣ�
    Select Case mintUnit
    Case 0  'ɢװ��λ
         strFields = "S.���㵥λ ��λ,D.����ϵ��,ltrim(to_char(S.����,'9999999999')) ����,S.�ѷ����� as ԭʼ����,ltrim(to_char(S.ʵ������," & mOraFMT.FM_���� & ")) ����,ltrim(to_char(S.��������," & mOraFMT.FM_���� & ")) as ������,ltrim(to_char(S.�ѷ�����," & mOraFMT.FM_���� & ")) as ׼����,'' ������,ltrim(to_char(S.���ۼ�," & mOraFMT.FM_���ۼ� & ")) ����,''  �����, "
    Case Else
         strFields = "D.��װ��λ ��λ,D.����ϵ��,ltrim(to_char(S.����,'9999999999')) ����,S.�ѷ����� as ԭʼ����,ltrim(to_char(S.ʵ������/D.����ϵ��," & mOraFMT.FM_���� & ")) ����,ltrim(to_char(S.��������/D.����ϵ��," & mOraFMT.FM_���� & ")) as ������,ltrim(to_char(S.�ѷ�����/D.����ϵ��," & mOraFMT.FM_���� & ")) as ׼����,'' ������,ltrim(to_char(S.���ۼ�*D.����ϵ��," & mOraFMT.FM_���ۼ� & ")) ����,'' �����, "
    End Select
    
    
    gstrSQL = "" & _
        " SELECT DISTINCT S.id,S.��¼״̬ ,S.����ID,s.����ҽ��,decode(S.����,24,'�շ�',25,'���ʵ�',26,'���ʱ�' ) as ����,s.סԺ��,s.����Ա,s.����Ա, S.ID,S.����,S.ҩƷID as ����id,S.NO,S.����,P.���� ����,s.�����־,to_char(s.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,S.����,s.����,s.����," & _
        "            '['||X.����||']'||X.����  ��������,NVL(D.���÷���,0) ���÷���,DECODE(x.���,NULL,x.����,DECODE(x.����,NULL,x.���,x.���||'|'||x.����)) ���," & strFields & _
        "  DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����,NVL(S.����,0) ����,S.Ч��," & _
        "  S.���۽�� ���,S.ժҪ ˵��,S.�����,TO_CHAR(S.�������,'YYYY-MM-DD HH24:MI:SS') ����ʱ��,s.�ɲ���" & _
        " FROM (    SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,NVL(A.����,0) ����," & _
        "                   NVL(A.����,1) ����,A.ʵ������ ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
        "                   A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����ID,A.�ⷿID,A.����ҽ��,A.���㵥λ,A.סԺ��,A.����Ա,A.����Ա,A.�����־,A.����ʱ�� ,A.����,A.����,A.����,A.�ɲ���" & _
        "           FROM(SELECT A.ID,A.NO,A.ҩƷid,A.���,A.����,A.����ID,A.����,A.����,A.Ч��,nvl(A.����,0) ����,nvl(A.����,0) ����,A.ʵ������,A.��¼״̬," & _
        "                       A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����id,A.�ⷿID," & _
        "                       m.������ as ����ҽ��,M.���㵥λ,m.��ʶ�� as סԺ��,m.����Ա���� as ����Ա,m.������ ����Ա,m.�����־,m.����ʱ��,m.����,'' ����,m.����,1 �ɲ��� " & _
        "                FROM ҩƷ�շ���¼ A,������ü�¼ M" & _
        "                WHERE  A.����� IS NOT NULL and A.����id=M.ID  and nvl(a.��ҩ��ʽ,0)<>-1 AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
        "                       AND A.�ⷿID+0=[1] and a.����=[2] and a.NO=[3] ) A," & _
        "               (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
        "                FROM ҩƷ�շ���¼ A" & _
        "                WHERE A.����� IS NOT NULL and nvl(a.��ҩ��ʽ,0)<>-1 AND A.�ⷿID+0=[1]" & _
        "                        and A.NO=[3] and A.����=[2]" & _
        "               GROUP BY A.NO,A.����,A.ҩƷID,A.��� ) B" & _
        "           WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� AND B.�ѷ�����<>0" & _
        "       ) S,���ű� P,�������� D,�շ���ĿĿ¼ X" & _
        " WHERE S.ҩƷID=D.����ID AND S.�Է�����ID+0=P.ID  AND d.����ID=X.ID" & _
        "       AND (S.��¼״̬=1 OR MOD(S.��¼״̬,3)=0) AND S.ʵ������*S.����>S.�������� " & _
        "       AND S.����� IS NOT NULL AND S.�ⷿID+0=[1] "
    
    If IntBill = 25 Then
        str���� = gstrSQL
        gstrSQL = Replace(gstrSQL, "'' ����", "M.����")
        gstrSQL = Replace(gstrSQL, "m.����", "nvl(R.����,m.����) ����")
        gstrSQL = Replace(gstrSQL, "m.����", "nvl(R.����,m.����) ����")
        gstrSQL = Replace(gstrSQL, "������ü�¼ M", "סԺ���ü�¼ M,������ҳ R")
        gstrSQL = Replace(gstrSQL, "A.����id=M.ID", "A.����id=M.ID And M.����id=R.����id And M.��ҳid=R.��ҳid ")
        gstrSQL = str���� & " Union All " & gstrSQL
    ElseIf IntBill = 26 Then
        gstrSQL = Replace(gstrSQL, "'' ����", "M.����")
        gstrSQL = Replace(gstrSQL, "m.����", "nvl(R.����,m.����) ����")
        gstrSQL = Replace(gstrSQL, "m.����", "nvl(R.����,m.����) ����")
        gstrSQL = Replace(gstrSQL, "A.����id=M.ID", "A.����id=M.ID And M.����id=R.����id And M.��ҳid=R.��ҳid ")
        gstrSQL = Replace(gstrSQL, "������ü�¼ M", "סԺ���ü�¼ M,������ҳ R")
    End If
    
    gstrSQL = gstrSQL & " Order By No,����"
    
    Set rsBill = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mlng���ϲ���, IntBill, strNo)
    With rsBill
        '�����������
        Bill.ClearBill
        If .RecordCount = 0 Then
            Bill.Rows = 2
            Bill.ClearBill
            txtDown(0).Text = ""
            txtDown(1).Text = ""
            txtDown(2).Text = ""
        
            lblNote.Caption = "δ�����κδ���"
            MsgBox "�������ҵĴ�����������,����!", vbInformation + vbDefaultButton1
            Exit Function
        Else
            Bill.Rows = .RecordCount + 1
            lblNote.Caption = "�ҵ�����"
        End If
        txtDown(0).Text = NVL(!����)
        txtDown(1).Text = NVL(!����ҽ��)
        txtDown(2).Text = NVL(!����ʱ��)
           
        lngRow = 1
        Do While Not .EOF
            Bill.TextMatrix(lngRow, c_���) = NVL(!����)
            Bill.TextMatrix(lngRow, c_��Ŀ) = NVL(!��������)
            Bill.TextMatrix(lngRow, c_���) = NVL(!���)
            Bill.TextMatrix(lngRow, c_����) = NVL(!����)
            Bill.TextMatrix(lngRow, c_����) = NVL(!����)
            
            Bill.TextMatrix(lngRow, c_��λ) = NVL(!��λ)
            Bill.TextMatrix(lngRow, c_����) = NVL(!����)
            Bill.TextMatrix(lngRow, c_����ϵ��) = NVL(!����ϵ��)
            Bill.TextMatrix(lngRow, c_����) = Format(Val(NVL(!����)), mFMT.FM_����)
            Bill.TextMatrix(lngRow, c_ԭʼ����) = Val(NVL(!ԭʼ����))
            Bill.TextMatrix(lngRow, c_������) = Format(Val(NVL(!������)), mFMT.FM_����)
            Bill.TextMatrix(lngRow, c_׼����) = Format(Val(NVL(!׼����)), mFMT.FM_����)
            Bill.TextMatrix(lngRow, c_������) = Format(Val(NVL(!׼����)), mFMT.FM_����)
            Bill.TextMatrix(lngRow, c_����) = Format(NVL(!����), mFMT.FM_���ۼ�)
            Bill.TextMatrix(lngRow, c_���) = Format(NVL(!���), mFMT.FM_���)
            Bill.TextMatrix(lngRow, c_ִ�п���) = ""
            Bill.TextMatrix(lngRow, c_����) = NVL(!����)
            Bill.TextMatrix(lngRow, c_����) = NVL(!����)
            Bill.TextMatrix(lngRow, c_����) = NVL(!����)
            Bill.RowData(lngRow) = Val(NVL(!Id))
            lngRow = lngRow + 1
            .MoveNext
        Loop
    End With
    
    InitBill = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 7730 Then Me.Width = 7730
    If Me.Height < 6000 Then Me.Height = 6000
     
    With fraTop
       .Top = ScaleTop + 50
       .Left = ScaleLeft + 50
       .Width = ScaleWidth - .Left
    End With
    
    With CmdCancel
        .Top = Me.ScaleHeight - .Height - 50
        .Left = Me.ScaleWidth - .Width - 100
    End With
    With cmdOK
        .Top = CmdCancel.Top
        .Left = CmdCancel.Left - 50 - .Width
    End With
    With cmdAllSel
        .Top = CmdCancel.Top
        .Left = fraTop.Left
    End With
    With cmdAllCls
        .Top = CmdCancel.Top
        .Left = cmdAllSel.Left + cmdAllSel.Width + 50
    End With
    With fraBillDown
        .Top = CmdCancel.Top - .Height - 50
        .Left = fraTop.Left
        .Width = fraTop.Width
    End With
    With Bill
        .Top = fraTop.Top + fraTop.Height
        .Left = fraTop.Left
        .Width = fraTop.Width
        .Height = fraBillDown.Top - .Top
    End With
    With lblNote
        .Left = fraTop.Width - .Width - 10
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If mblnChange = True Then
        If MsgBox("�������ݿ����Ѹı䣬����δ���ϣ���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    err = 0: On Error Resume Next
    SaveWinState Me, App.ProductName, mstrCaption
    Call zlDatabase.SetPara("������ϵ�������", Me.cbo�շѵ���.ItemData(Me.cbo�շѵ���.ListIndex), glngSys, mlngModule)
 
End Sub

Private Sub TxtNo_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNo As String, IntBill As Integer
    
    err = 0: On Error GoTo ErrHand:
    
    If cbo�շѵ���.ListIndex < 0 Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtNO) = "" Then Exit Sub
    Me.txtNO = UCase(LTrim(Me.txtNO))
    Me.txtNO.Text = zlCommFun.GetFullNo(Me.txtNO.Text, 13)
    strNo = txtNO.Text
    IntBill = cbo�շѵ���.ItemData(cbo�շѵ���.ListIndex)
    If InitBill(Me.txtNO, IntBill) = False Then
       If txtNO.Enabled Then txtNO.SetFocus
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


