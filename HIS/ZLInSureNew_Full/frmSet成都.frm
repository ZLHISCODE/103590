VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmSet�ɶ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�Ʊ��սӿ�����"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ControlBox      =   0   'False
   Icon            =   "frmSet�ɶ�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   1365
      Left            =   900
      TabIndex        =   14
      Top             =   2100
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2408
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
   Begin VB.CheckBox chkhisCharge 
      Caption         =   "HIS�շ�"
      Height          =   200
      Left            =   2640
      TabIndex        =   13
      Top             =   2050
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtInterCode 
      Height          =   300
      Left            =   4500
      MaxLength       =   6
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "713"
      Top             =   1965
      Visible         =   0   'False
      Width           =   960
   End
   Begin MSComCtl2.UpDown UDCard 
      Height          =   315
      Left            =   2190
      TabIndex        =   2
      Top             =   1965
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   30
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtCard"
      BuddyDispid     =   196613
      OrigLeft        =   2415
      OrigTop         =   1965
      OrigRight       =   2655
      OrigBottom      =   2280
      Max             =   30
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtCard 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "30"
      Top             =   1980
      Width           =   480
   End
   Begin VB.CommandButton cmdODBC 
      Caption         =   "����Դ(&D)"
      Height          =   350
      Left            =   225
      TabIndex        =   6
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   5
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "����(&T)"
      Height          =   350
      Left            =   1425
      TabIndex        =   3
      Top             =   2520
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -210
      TabIndex        =   9
      Top             =   2340
      Width           =   5850
   End
   Begin VB.TextBox txt���Ӵ� 
      Height          =   720
      Left            =   915
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1095
      Width           =   4575
   End
   Begin VB.Label Lbl�շ������� 
      AutoSize        =   -1  'True
      Caption         =   "�շ�������"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   930
      TabIndex        =   15
      Top             =   1890
      Width           =   1080
   End
   Begin VB.Label lblInterCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ������"
      Height          =   180
      Left            =   3720
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblCard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ų���"
      Height          =   180
      Left            =   930
      TabIndex        =   10
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label lbl���Ӵ� 
      AutoSize        =   -1  'True
      Caption         =   "���Ӵ�"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   930
      TabIndex        =   8
      Top             =   885
      Width           =   540
   End
   Begin VB.Label lblNote 
      Caption         =   "    ���õ�ҽ�Ʊ������ݷ����������Ӵ���Ϊ��֤������Ч����ʱҽ�Ʊ������ݷ�����������á�"
      Height          =   390
      Left            =   930
      TabIndex        =   7
      Top             =   225
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   210
      Picture         =   "frmSet�ɶ�.frx":030A
      Top             =   180
      Width           =   240
   End
End
Attribute VB_Name = "frmSet�ɶ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mint���� As Integer
Dim mblnOK As Boolean
Dim str������Ŀ As String   '�շ�����Ӧҽ���ķ�����Ŀ

Private Sub Bill_cboClick(ListIndex As Long)
    Bill.TextMatrix(Bill.Row, Bill.COL) = Bill.CboText
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdODBC_Click()
    On Error Resume Next
    Shell "ODBCAD32", vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "���ܽ���ODBC����Դ������������ϵͳ�Ƿ���ȷ��װ��", vbInformation, gstrSysName
    End If
    Err.Clear
End Sub

Private Sub cmdOK_Click()
    Select Case mint����
        Case TYPE_�ɶ���
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("ConnectionStrINg"), Trim(txt���Ӵ�.Text)
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("CardNOLength"), txtCard.Text
        Case TYPE_�ɶ��ϳ�
            If Not CheckItem Then
                If MsgBox("�в����շ����δ���ö�Ӧ��ҽ���շ���Ŀ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            Call Combinate
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("LCConnectionString"), Trim(txt���Ӵ�.Text)
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("LCItem"), str������Ŀ
        Case TYPE_�ɶ�����
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("CardNOLength"), txtCard.Text
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("LHConnectionStrINg"), Trim(txt���Ӵ�.Text)
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("intercode"), txtInterCode.Text
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("HIS�շ�"), chkhisCharge.Value
        Case TYPE_����
            '20050124
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("LHConnectionStrINg"), Trim(txt���Ӵ�.Text)
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("intercode"), txtInterCode.Text
            SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("HIS�շ�"), chkhisCharge.Value
    End Select
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim cnInsure As New ADODB.Connection
    Err = 0
    On Error Resume Next
    With cnInsure
        If .State = adStateOpen Then .Close
        .ConnectionString = Trim(Me.txt���Ӵ�.Text)
        .Open
        If Err <> 0 Then
            MsgBox "���Բ��ɹ�������ҽ�����ݷ������Ƿ���ã��Լ�����Դ�Ƿ���ȷ���ã�", vbExclamation, gstrSysName
            Exit Sub
        End If
        .Close
        If txtInterCode.Visible = True Then
            If txtInterCode.Text = "" Then
                MsgBox "ҽ�����벻��Ϊ�գ������䣡", vbExclamation, gstrSysName
                Exit Sub
            End If
            If IsNumeric(txtInterCode.Text) = False Then
                MsgBox "ҽ���������Ϊ�����ͣ������䣡", vbExclamation, gstrSysName
                txtInterCode.SelStart = 0
                txtInterCode.SelLength = Len(txtInterCode.Text)
                txtInterCode.SetFocus
                Exit Sub
            End If
        End If
        
        MsgBox "���Գɹ�����ҽ�����ݷ������������ӣ�", vbInformation, gstrSysName
        Me.cmdOK.Enabled = True
    End With
End Sub

Private Sub txtCard_Change()
    If txtCard.Locked Then Exit Sub
    Me.cmdOK.Enabled = True
End Sub

Private Sub txtInterCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    End If
End Sub

Private Sub txt���Ӵ�_KeyPress(KeyAscii As Integer)
    Me.cmdOK.Enabled = False
End Sub

Public Function ShowSet(ByVal int���� As Integer) As Boolean
'���ܣ��õ�����������Ϣ
    Dim rsTemp As New ADODB.Recordset
    mblnOK = False
    mint���� = int����
    
    Lbl�շ�������.Visible = False
    Bill.Visible = False
    If int���� <> TYPE_�ɶ��ϳ� Then
        Frame1.Top = txtCard.Top + txtCard.Height + 100
        Me.Height = 3380
    Else
        Frame1.Top = Bill.Top + Bill.Height + 100
        Me.Height = 4560
    End If
    Call AdjustCons
    
    Select Case int����
        Case TYPE_�ɶ���
            txt���Ӵ�.Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ConnectionStrINg"), "dsn=cnnSyb;uID=face;pwd=facepass")
            txtCard.Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("CardNOLength"), 20)
        Case TYPE_�ɶ��ϳ�
            Lbl�շ�������.Visible = True
            Bill.Visible = True
            txtCard.Visible = False
            lblCard.Visible = False
            UDCard.Visible = False
            lblInterCode.Visible = False
            txtInterCode.Visible = False
            chkhisCharge.Visible = False
            txt���Ӵ�.Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LCConnectionStrINg"), "dsn=lcyb;uid=hisuser;pwd=hiscdgk;")
            str������Ŀ = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LCItem"), "")
            
            '��ʼ�����
            Call InitBill
            
            'װ�������Ŀ
            gstrSQL = "Select ��� From �շ���� Order By ����"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ����")
            
            'װ���趨ֵ
            Call LoadSet(rsTemp)
        Case TYPE_�ɶ�����
            lblCard.Caption = "��������"
            UDCard.Visible = False
            lblInterCode.Visible = True
            txtInterCode.Visible = True
            chkhisCharge.Visible = True
            txtCard.Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("CardNOLength"), 10)
            txt���Ӵ�.Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LHConnectionStrINg"), "dsn=lhyb;uid=sa;pwd=;")
            txtInterCode.Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("intercode"), 713)
            chkhisCharge.Value = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("HIS�շ�"), 0)
            txtCard.Locked = False
        Case TYPE_����
            txtCard.Visible = False
            lblCard.Visible = False
            UDCard.Visible = False
            lblInterCode.Visible = True
            txtInterCode.Visible = True
            chkhisCharge.Visible = True
            txt���Ӵ�.Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LHConnectionStrINg"), "dsn=lhyb;uid=sa;pwd=;")
            txtInterCode.Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("intercode"), 713)
            chkhisCharge.Value = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("HIS�շ�"), 0)
    End Select
    frmSet�ɶ�.Show vbModal
    
    ShowSet = mblnOK
End Function

Private Sub AdjustCons()
    With cmdODBC
        .Top = Frame1.Top + 200
    End With
    cmdCancel.Top = cmdODBC.Top
    cmdOK.Top = cmdODBC.Top
    cmdTest.Top = cmdODBC.Top
End Sub

Private Sub InitBill()
    With Bill
        .AllowAddRow = False
        .Active = True
        .ClearBill
        .Cols = 2
        
        .TextMatrix(0, 0) = "�շ����"
        .TextMatrix(0, 1) = "ҽ��������Ŀ"
        
        .ColData(0) = 0
        .ColData(1) = 3
        
        .ColWidth(0) = 1200
        .ColWidth(1) = 2500
        
        .PrimaryCol = 1
        .LocateCol = 1
    End With
End Sub

Private Sub LoadSet(ByVal rsTemp As ADODB.Recordset)
    Dim arrItem, intItem As Integer, strItem As String
    
    Bill.Rows = rsTemp.RecordCount + 1
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
    'װ���շ����
    For intItem = 1 To rsTemp.RecordCount
        Bill.TextMatrix(intItem, 0) = rsTemp!���
        rsTemp.MoveNext
    Next
    
    'װ��ҽ��������Ŀ
    arrItem = Split(gstr������Ŀ, gstrSplit����)
    For intItem = 0 To UBound(arrItem)
        Bill.AddItem arrItem(intItem)
    Next
    
    'װ���趨��ȷ�ķ�����Ŀ
    arrItem = Split(str������Ŀ, gstrSplit����)
    For intItem = 0 To UBound(arrItem)
        strItem = Split(arrItem(intItem), gstrSplitС��)(1)
        '������趨��ҽ��������Ŀ�Ƿ�����ȷ��
        If InStr(1, gstrSplit���� & gstr������Ŀ & gstrSplit����, gstrSplit���� & strItem & gstrSplit����) <> 0 Then
            rsTemp.MoveFirst
            rsTemp.Find "���='" & Split(arrItem(intItem), gstrSplitС��)(0) & "'"  '�ҵ����Ӧ���շ����
            If Not rsTemp.EOF Then Bill.TextMatrix(rsTemp.AbsolutePosition, 1) = strItem
        End If
    Next
End Sub

Private Sub Combinate()
    Dim intItem  As Integer
    '���趨��������ϳɴ�
    str������Ŀ = ""
    For intItem = 1 To Bill.Rows - 1
        str������Ŀ = str������Ŀ & gstrSplit���� & Bill.TextMatrix(intItem, 0) & gstrSplitС�� & Bill.TextMatrix(intItem, 1)
    Next
    str������Ŀ = Mid(str������Ŀ, 2)
End Sub

Private Function CheckItem() As Boolean
    Dim intRow As Integer
    CheckItem = False
    For intRow = 1 To Bill.Rows - 1
        If Trim(Bill.TextMatrix(intRow, 1)) = "" Then Exit Function
    Next
    CheckItem = True
End Function
