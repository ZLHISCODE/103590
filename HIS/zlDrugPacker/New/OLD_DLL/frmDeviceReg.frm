VERSION 5.00
Begin VB.Form frmDeviceReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�豸��Ϣ"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7575
   Icon            =   "frmDeviceReg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7575
   StartUpPosition =   1  '����������
   Begin VB.Frame fraFunc 
      Caption         =   "������ҵ��"
      Height          =   2295
      Left            =   3840
      TabIndex        =   19
      Top             =   960
      Width           =   3615
      Begin VB.ComboBox cboDispense 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkDispensing 
         Caption         =   "����ҩƷ������ҩ"
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ���ܣ�"
         Height          =   180
         Index           =   11
         Left            =   240
         TabIndex        =   20
         Top             =   390
         Width           =   900
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���͹��ܣ�"
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   21
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblDevice 
         BackStyle       =   0  'Transparent
         Caption         =   "  ָ����ҩ������HIS�ĸ�ҵ�����ҩƷ��ϸ�ϴ�"
         Height          =   420
         Index           =   13
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label lblDevice 
         BackStyle       =   0  'Transparent
         Caption         =   "  ָ�����͹����Ƿ��ڴ�����ҩ����"
         Height          =   420
         Index           =   14
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   3000
      End
   End
   Begin VB.Frame fraService 
      Caption         =   "�������"
      Height          =   735
      Left            =   3840
      TabIndex        =   16
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton optObject 
         Caption         =   "סԺ"
         Height          =   180
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optObject 
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   6240
      TabIndex        =   25
      Top             =   3360
      Width           =   1110
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   360
      Left            =   5040
      TabIndex        =   24
      Top             =   3360
      Width           =   1110
   End
   Begin VB.Frame fraDevice 
      Caption         =   "������Ϣ"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox cboLink 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtManufacturer 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtModel 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optState 
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         Top             =   750
         Width           =   855
      End
      Begin VB.OptionButton optState 
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   750
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ��ҩ����"
         Height          =   180
         Index           =   7
         Left            =   360
         TabIndex        =   14
         Top             =   2550
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����̣�"
         Height          =   180
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   2190
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͺţ�"
         Height          =   180
         Index           =   5
         Left            =   360
         TabIndex        =   10
         Top             =   1830
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ƣ�"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   8
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���룺"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�豸״̬��"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   750
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   390
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmDeviceReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long

Public Sub ShowMe(ByVal frmOwner As Form, ByVal bytState As Byte, ByVal lngID As Long)
'���ܣ��������
'������
'  frmOwner�����������
'  bytState������״̬��0-������1-�޸�
'  lngID������״̬Ϊ0��������ʱ����ʾ����ID������״̬Ϊ1���޸ģ�ʱ����ʾ�豸ID
    
    mlngID = lngID
    Me.Tag = bytState
    
    Call Init
    Call FullData(mlngID)
    Call cboLink_Click
    If Val(Me.Tag) = 0 Then Call cboDept_Click
    
    Me.Show vbModal, frmOwner
    
End Sub

Private Sub cboDept_Click()
    cmdSave.Enabled = cboDept.ListIndex >= 0 And cboLink.ListIndex >= 0
    If cboDept.ListIndex < 0 Then
        optObject(0).Value = False
        optObject(1).Value = False
        optObject(0).Enabled = False
        optObject(1).Enabled = False
    Else
        'ҩ���������
        Dim rsTmp As ADODB.Recordset
        
        On Error GoTo errHandle
        gstrSQL = "Select ������� From ��������˵�� " & _
                  "Where ����id = [1] And ������� in (1,2,3) " & _
                  "Order By ������� "
        Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���ŷ������", cboDept.ItemData(cboDept.ListIndex))
        Do While rsTmp.EOF = False
            Select Case gobjComLib.zlCommFun.Nvl(rsTmp!�������, 0)
                Case 1                  '���ﲡ��
                    optObject(0).Value = True
                    optObject(0).Enabled = True
                    optObject(1).Enabled = False
                Case 2                  'סԺ����
                    optObject(1).Value = True
                    optObject(1).Enabled = True
                    optObject(0).Enabled = False
                Case 3                  '���ﲡ����סԺ����
                    optObject(0).Enabled = True
                    optObject(1).Enabled = True
                Case Else               '�ǲ���
                    optObject(0).Value = False
                    optObject(1).Value = False
                    optObject(0).Enabled = False
                    optObject(1).Enabled = False
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        Set rsTmp = Nothing
        
    End If
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub cboDispense_Click()
    chkDispensing.Value = False
    chkDispensing.Enabled = cboDispense.ListIndex = 2
End Sub

Private Sub cboLink_Click()
    cmdSave.Enabled = cboDept.ListIndex >= 0 And cboLink.ListIndex >= 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer

    '���
    If Trim(txtCode.Text) = "" Then
        MsgBox "δ��д�����롱��", vbInformation, GSTR_INTERFACE_NAME
        txtCode.SetFocus
        Exit Sub
    End If
    If Trim(txtName.Text) = "" Then
        MsgBox "δ��д�����ơ���", vbInformation, GSTR_INTERFACE_NAME
        txtName.SetFocus
        Exit Sub
    End If
    If cboLink.ListIndex < 0 Then
        MsgBox "δѡ������������", vbInformation, GSTR_INTERFACE_NAME
        cboLink.SetFocus
        Exit Sub
    End If
    If cboDept.ListIndex < 0 Then
        MsgBox "δѡ��ʹ��ҩ������", vbInformation, GSTR_INTERFACE_NAME
        cboDept.SetFocus
        Exit Sub
    End If
    If optObject(0).Value = False And optObject(1) = False Then
        MsgBox "��������󡱱����ѡһ��", vbInformation, GSTR_INTERFACE_NAME
        optObject(0).SetFocus
        Exit Sub
    End If
    If cboDispense.Enabled Then
        If cboDispense.ListIndex < 0 Then
            MsgBox "δѡ����ҩ���ܡ���Ӧ��ҵ��", vbInformation, GSTR_INTERFACE_NAME
            cboDispense.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(Me.Tag) = 1 Then
        '�޸�
        gstrSQL = "Zl_ҩ��ע���豸_Update("
        gstrSQL = gstrSQL & mlngID & ","
        gstrSQL = gstrSQL & "'" & txtCode.Text & "',"
        gstrSQL = gstrSQL & "'" & txtName.Text & "',"
        gstrSQL = gstrSQL & IIf(Trim(txtModel.Text) = "", "null", "'" & txtModel.Text & "'") & ","
        gstrSQL = gstrSQL & IIf(Trim(txtManufacturer.Text) = "", "null", "'" & txtManufacturer.Text & "'") & ","
        gstrSQL = gstrSQL & cboLink.ItemData(cboLink.ListIndex) & ","
        gstrSQL = gstrSQL & cboDept.ItemData(cboDept.ListIndex) & ","
        gstrSQL = gstrSQL & IIf(optState(0).Value, "1", "null") & ","
        gstrSQL = gstrSQL & IIf(optObject(0).Value, "'1'", "'2'") & ","
        If cboDispense.Enabled Then
            gstrSQL = gstrSQL & "'" & cboDispense.ListIndex + 1 & "',"
        Else
            gstrSQL = gstrSQL & "null,"
        End If
        If chkDispensing.Enabled Then
            gstrSQL = gstrSQL & IIf(chkDispensing.Value, "'1'", "null")
        Else
            gstrSQL = gstrSQL & "null"
        End If
        gstrSQL = gstrSQL & ")"
        
        On Error GoTo errHandle
        Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "ҩ��ע���豸-�޸�")
        
    Else
        '����
        gstrSQL = "Zl_ҩ��ע���豸_Insert("
        gstrSQL = gstrSQL & "'" & txtCode.Text & "',"
        gstrSQL = gstrSQL & "'" & txtName.Text & "',"
        gstrSQL = gstrSQL & IIf(Trim(txtModel.Text) = "", "null", "'" & txtModel.Text & "'") & ","
        gstrSQL = gstrSQL & IIf(Trim(txtManufacturer.Text) = "", "null", "'" & txtManufacturer.Text & "'") & ","
        gstrSQL = gstrSQL & cboLink.ItemData(cboLink.ListIndex) & ","
        gstrSQL = gstrSQL & cboDept.ItemData(cboDept.ListIndex) & ","
        gstrSQL = gstrSQL & IIf(optState(0).Value, "1", "null") & ","
        gstrSQL = gstrSQL & IIf(optObject(0).Value, "'1'", "'2'") & ","
        If cboDispense.Enabled Then
            gstrSQL = gstrSQL & "'" & cboDispense.ListIndex + 1 & "',"
        Else
            gstrSQL = gstrSQL & "null,"
        End If
        If chkDispensing.Enabled Then
            gstrSQL = gstrSQL & IIf(chkDispensing.Value, "'1'", "null")
        Else
            gstrSQL = gstrSQL & "null"
        End If
        gstrSQL = gstrSQL & ")"
        
        On Error GoTo errHandle
        Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "ҩ��ע���豸-����")
        
    End If

    Unload Me
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub Form_Load()
    '
End Sub

Private Sub Init()
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select ID, ���� From ҩ���豸���� Order by ���� "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ���豸����")
    Do While rsTmp.EOF = False
        cboLink.AddItem rsTmp!����
        cboLink.ItemData(cboLink.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    gstrSQL = "Select Distinct a.Id, '��' || a.���� || '��' || a.���� ���� " & _
              "From ���ű� A, ��������˵�� B " & _
              "Where a.Id = b.����id And b.�������� In ('��ҩ��', '��ҩ��', '��ҩ��') And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000/1/1', 'YYYY/MM/DD')) " & _
              "Order By '��' || a.���� || '��' || a.���� "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ��������Ϣ")
    Do While rsTmp.EOF = False
        cboDept.AddItem rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub FullData(ByVal lngID As Long)
'���ܣ����������ؼ�
'������
'  lngID���豸ID

    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    With cboDispense
        .Clear
        .AddItem "�����շ�", 0
        .AddItem "������ҩ-��ҩ", 1
        .AddItem "������ҩ-��ҩ", 2
    End With
    
    '����
    If Val(Me.Tag) = 0 Then
        gstrSQL = "Select ���� From ҩ���豸���� Where ID = [1] "
        Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ���豸����", lngID)
        If rsTmp.EOF = False Then
            cboLink.Text = rsTmp!����
        End If
        optObject(0).Enabled = False
        optObject(1).Enabled = False
        cboDispense.Enabled = False
        chkDispensing.Enabled = False
        Exit Sub
    End If
    
    'ҩ���豸��Ϣ
    gstrSQL = "Select a.*, b.���� ������, '��' || c.���� || '��' || c.���� ҩ�� " & _
              "From ҩ��ע���豸 A, ҩ���豸���� B, ���ű� C " & _
              "Where a.����id = b.Id and a.����id = c.Id and a.ID = [1] "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ��ע���豸", lngID)
    If rsTmp.EOF = False Then
        cboLink.Text = rsTmp!������
        txtCode.Text = rsTmp!����
        txtName.Text = rsTmp!����
        txtModel.Text = gobjComLib.zlCommFun.Nvl(rsTmp!�ͺ�)
        txtManufacturer.Text = gobjComLib.zlCommFun.Nvl(rsTmp!������)
        If gobjComLib.zlCommFun.Nvl(rsTmp!����, 0) = 1 Then
            optState(0).Value = True
        Else
            optState(1).Value = True
        End If
        cboLink.Text = rsTmp!������
        cboDept.Text = rsTmp!ҩ��
    End If
    rsTmp.Close
    
    'ҩ���豸����
    gstrSQL = "Select b.Id, b.����, b.����, b.�ͺ�, b.����, a.������, c.����ֵ " & _
              "From Zlparameters A, ҩ��ע���豸 B, ҩ���豸���� C " & _
              "Where a.Id = c.����id And b.Id = c.�豸id And a.ϵͳ = 100 And a.ģ�� = [1] And b.Id = [2] "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ���豸����", GINT_INTERFACE_MODULENO, lngID)
    
    '�������
    optObject(0).Enabled = True
    optObject(1).Enabled = True
    rsTmp.Filter = "������=1"
    If rsTmp.EOF = False Then
        Select Case Val(rsTmp!����ֵ)
            Case 1      '����
                optObject(0).Value = True
            Case 2      'סԺ
                optObject(1).Value = True
            Case Else   '�쳣
                optObject(0).Value = False
                optObject(1).Value = False
        End Select
    Else
        optObject(0).Value = False
        optObject(1).Value = False
    End If
    
    '��ҩ��Ӧҵ��
    cboDispense.ListIndex = -1
    cboDispense.Enabled = optObject(0).Value
    rsTmp.Filter = "������=2"
    If rsTmp.EOF = False Then
        If optObject(0).Value Then
            '����
            cboDispense.ListIndex = Val(rsTmp!����ֵ) - 1
        End If
    End If
    
    '���Ͷ�Ӧҵ��
    chkDispensing.Value = False
    chkDispensing.Enabled = True
    rsTmp.Filter = "������=3"
    If rsTmp.EOF = False Then
        If cboDispense.ListIndex = 2 Then
            '������ҩ-��ҩ
            chkDispensing.Value = Val(rsTmp!����ֵ)
        Else
            chkDispensing.Enabled = False
        End If
    End If
    
    rsTmp.Close
    Set rsTmp = Nothing
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub optObject_Click(Index As Integer)
    cboDispense.Enabled = Index = 0
End Sub
