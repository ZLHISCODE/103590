VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmServiceSet 
   Caption         =   "��������"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   300
      Left            =   5520
      TabIndex        =   34
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   300
      Left            =   2160
      TabIndex        =   33
      Top             =   6000
      Width           =   1100
   End
   Begin TabDlg.SSTab tabService 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "�豸����"
      TabPicture(0)   =   "frmServiceSet.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(3)=   "MSFDevice"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "txtDeviceName"
      Tab(0).Control(6)=   "cmdAddDevice"
      Tab(0).Control(7)=   "cmdModiDevice"
      Tab(0).Control(8)=   "cmdDelDevice"
      Tab(0).Control(9)=   "txtIPAddr"
      Tab(0).Control(10)=   "cboModality"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "��������"
      TabPicture(1)   =   "frmServiceSet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAddService"
      Tab(1).Control(1)=   "MSFService"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "cmdModiService"
      Tab(1).Control(4)=   "cmdDelService"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "DICOM��������"
      TabPicture(2)   =   "frmServiceSet.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "MSFDicomDevice"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "MSFDicomService"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "DICOM�������"
         Height          =   2415
         Left            =   120
         TabIndex        =   56
         Top             =   3240
         Width           =   8775
      End
      Begin MSFlexGridLib.MSFlexGrid MSFDicomService 
         Height          =   1335
         Left            =   120
         TabIndex        =   55
         Top             =   1800
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFDicomDevice 
         Height          =   1215
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin VB.ComboBox cboModality 
         Height          =   300
         ItemData        =   "frmServiceSet.frx":0054
         Left            =   -71040
         List            =   "frmServiceSet.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmdDelService 
         Caption         =   "ɾ��"
         Height          =   300
         Left            =   -68160
         TabIndex        =   51
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton cmdModiService 
         Caption         =   "�޸�"
         Height          =   300
         Left            =   -71040
         TabIndex        =   50
         Top             =   5280
         Width           =   1100
      End
      Begin VB.Frame Frame2 
         Caption         =   "��������"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   38
         Top             =   3600
         Width           =   8775
         Begin VB.ComboBox cboServiceType 
            Height          =   300
            ItemData        =   "frmServiceSet.frx":006E
            Left            =   3720
            List            =   "frmServiceSet.frx":007B
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdServiceSetup 
            Caption         =   "�߼�����"
            Enabled         =   0   'False
            Height          =   300
            Left            =   7440
            TabIndex        =   43
            Top             =   360
            Width           =   1100
         End
         Begin VB.TextBox txtServiceName 
            Height          =   300
            Left            =   720
            MaxLength       =   20
            TabIndex        =   42
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtServicePort 
            Height          =   300
            Left            =   6840
            MaxLength       =   4
            TabIndex        =   41
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtServiceIP 
            Height          =   300
            Left            =   720
            MaxLength       =   15
            TabIndex        =   40
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtServiceAE 
            Height          =   300
            Left            =   3720
            MaxLength       =   20
            TabIndex        =   39
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "�˿�"
            Height          =   165
            Left            =   6240
            TabIndex        =   48
            Top             =   908
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "������"
            Height          =   165
            Left            =   120
            TabIndex        =   47
            Top             =   435
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "��������"
            Height          =   165
            Left            =   3000
            TabIndex        =   46
            Top             =   435
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "AE����"
            Height          =   165
            Left            =   3000
            TabIndex        =   45
            Top             =   908
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "IP��ַ"
            Height          =   165
            Left            =   120
            TabIndex        =   44
            Top             =   908
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFService 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5318
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.TextBox txtIPAddr 
         Height          =   300
         Left            =   -67920
         MaxLength       =   15
         TabIndex        =   31
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelDevice 
         Caption         =   "ɾ��"
         Height          =   300
         Left            =   -68160
         TabIndex        =   23
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton cmdModiDevice 
         Caption         =   "�޸�"
         Height          =   300
         Left            =   -71040
         TabIndex        =   22
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddDevice 
         Caption         =   "����"
         Height          =   300
         Left            =   -73920
         TabIndex        =   21
         Top             =   5280
         Width           =   1100
      End
      Begin VB.TextBox txtDeviceName 
         Height          =   300
         Left            =   -73920
         MaxLength       =   100
         TabIndex        =   20
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "�豸��������"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   4
         Top             =   3240
         Width           =   8775
         Begin VB.CheckBox chkService 
            Caption         =   "Q/R��������"
            Height          =   255
            Index           =   2
            Left            =   6000
            TabIndex        =   37
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkService 
            Caption         =   "Worklist����"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   36
            Top             =   240
            Width           =   1455
         End
         Begin VB.Frame frmService 
            Enabled         =   0   'False
            Height          =   1455
            Index           =   1
            Left            =   3000
            TabIndex        =   24
            Top             =   240
            Width           =   2800
            Begin VB.TextBox txtDeviceAE 
               Height          =   300
               Index           =   1
               Left            =   840
               MaxLength       =   20
               TabIndex        =   27
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txtDevicePort 
               Height          =   300
               Index           =   1
               Left            =   840
               MaxLength       =   4
               TabIndex        =   26
               Top             =   720
               Width           =   1695
            End
            Begin VB.ComboBox cboService 
               Height          =   300
               Index           =   1
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label Label6 
               Caption         =   "AE����"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   390
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "�˿�"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   750
               Width           =   615
            End
            Begin VB.Label Label8 
               Caption         =   "����"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   1110
               Width           =   615
            End
         End
         Begin VB.CheckBox chkService 
            Caption         =   "ͼ��洢����"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.Frame frmService 
            Enabled         =   0   'False
            Height          =   1455
            Index           =   2
            Left            =   5880
            TabIndex        =   6
            Top             =   240
            Width           =   2800
            Begin VB.ComboBox cboService 
               Height          =   300
               Index           =   2
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   1080
               Width           =   1695
            End
            Begin VB.TextBox txtDevicePort 
               Height          =   300
               Index           =   2
               Left            =   840
               MaxLength       =   4
               TabIndex        =   18
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txtDeviceAE 
               Height          =   300
               Index           =   2
               Left            =   840
               MaxLength       =   20
               TabIndex        =   17
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label11 
               Caption         =   "����"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   1110
               Width           =   615
            End
            Begin VB.Label Label10 
               Caption         =   "�˿�"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   750
               Width           =   615
            End
            Begin VB.Label Label9 
               Caption         =   "AE����"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   390
               Width           =   615
            End
         End
         Begin VB.Frame frmService 
            Enabled         =   0   'False
            Height          =   1455
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2800
            Begin VB.ComboBox cboService 
               Height          =   300
               Index           =   0
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   1080
               Width           =   1695
            End
            Begin VB.TextBox txtDevicePort 
               Height          =   300
               Index           =   0
               Left            =   840
               MaxLength       =   4
               TabIndex        =   11
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txtDeviceAE 
               Height          =   300
               Index           =   0
               Left            =   840
               MaxLength       =   20
               TabIndex        =   10
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label5 
               Caption         =   "����"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   1110
               Width           =   615
            End
            Begin VB.Label Label4 
               Caption         =   "�˿�"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   750
               Width           =   615
            End
            Begin VB.Label Label3 
               Caption         =   "AE����"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   390
               Width           =   615
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFDevice 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3836
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.CommandButton cmdAddService 
         Caption         =   "����"
         Height          =   300
         Left            =   -73920
         TabIndex        =   49
         Top             =   5280
         Width           =   1100
      End
      Begin VB.Label Label17 
         Caption         =   "Ӱ�����"
         Height          =   255
         Left            =   -71880
         TabIndex        =   35
         Top             =   2903
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "�豸IP��ַ"
         Height          =   255
         Left            =   -68880
         TabIndex        =   3
         Top             =   2903
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "�豸������"
         Height          =   255
         Left            =   -74880
         TabIndex        =   2
         Top             =   2903
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmServiceSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngServiceID As Long    '����ID

Private Sub chkService_Click(Index As Integer)
    If Me.chkService(Index).value = 1 Then
        Me.frmService(Index).Enabled = True
    Else
        Me.frmService(Index).Enabled = False
        Me.txtDeviceAE(Index).Text = ""
        Me.txtDevicePort(Index).Text = ""
        Me.cboService(Index).ListIndex = -1
    End If
End Sub

Private Sub cmdAddDevice_Click()
    Dim i As Integer
    Dim blnResult As Boolean    '��Ч�Լ����
    '��������Ƿ���Ч
    blnResult = funValidateDevice
    If blnResult = False Then Exit Sub
    
    On Error GoTo errHand
    '��������
    gstrSQL = "Zl_Ӱ������豸_INSERT('" & Me.txtIPAddr.Text & "','" & Me.txtDeviceName.Text & "','" & _
                    Left(cboModality.Text, InStr(cboModality.Text, "-") - 1) & "'"
    For i = 0 To 2
        If chkService(i).value = 1 Then
            gstrSQL = gstrSQL & ",'" & Me.txtDeviceAE(i).Text & "','" & Me.txtDevicePort(i).Text & "'," & Me.cboService(i).ItemData(Me.cboService(i).ListIndex)
        Else
            gstrSQL = gstrSQL & ",null,null,null"
        End If
    Next i
    gstrSQL = gstrSQL & ")"
                        
    ExecuteProcedure "����DICOM�豸"
    'ˢ���б�
    Call subFillMSFDevice(Me.MSFDevice.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdAddService_Click()
    Dim i As Integer
    Dim blnResult As Boolean    '��Ч�Լ����
    '��������Ƿ���Ч
    blnResult = funValidateService
    If blnResult = False Then Exit Sub
    On Error GoTo errHand
    '��������
    gstrSQL = "Zl_Ӱ��DICOM����_INSERT('" & Me.txtServiceName.Text & "','" & Me.txtServiceIP.Text & "','" & _
                    Me.txtServiceAE.Text & "','" & Me.txtServicePort.Text & "','" & Me.cboServiceType.Text & "')"
                        
    ExecuteProcedure "����DICOM����"
    'ˢ���б�
    Call subFillMSFService(Me.MSFService.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function funValidateService()
    Dim arrIP() As String
    Dim i As Integer
    
    '�ж�������豸��Ϣ�Ƿ���Ч
    '��������Ϊ��
    If Me.txtServiceName = "" Then GoTo inValidate
    '�������Ͳ�Ϊ��
    If Me.cboServiceType.ListIndex = -1 Then GoTo inValidate
    'IP��ַ��Ч
    If Me.txtServiceIP.Text = "" Then
        GoTo inValidate
    Else
        arrIP = Split(Me.txtServiceIP.Text, ".")
        If UBound(arrIP) = 3 Then
            If IsNumeric(arrIP(0)) And IsNumeric(arrIP(1)) And IsNumeric(arrIP(2)) And IsNumeric(arrIP(3)) Then
                '��Ч��������
            Else
                GoTo inValidate
            End If
        Else
            GoTo inValidate
        End If
    End If
    'AE���Ʋ�Ϊ��
    If Me.txtServiceAE.Text = "" Then GoTo inValidate
    '�˿���Ч
    If Me.txtServicePort.Text = "" Then GoTo inValidate
    
    funValidateService = True
    Exit Function
inValidate:
    MsgBox "��������������߲��������������������", vbOKOnly, "������������"
    Exit Function
End Function

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelDevice_Click()
    If MSFDevice.Rows <= 1 Then Exit Sub
    On Error GoTo errHand
    'ɾ������
    gstrSQL = "Zl_Ӱ������豸_DELETE(" & Me.MSFDevice.TextMatrix(Me.MSFDevice.RowSel, 12) & ")"
                        
    ExecuteProcedure "�޸�DICOM�豸"
    'ˢ���б�
    Call subFillMSFDevice
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdDelService_Click()
    If MSFService.Rows <= 1 Then Exit Sub
    On Error GoTo errHand
    'ɾ������
    gstrSQL = "Zl_Ӱ��DICOM����_DELETE(" & Me.MSFService.TextMatrix(Me.MSFService.RowSel, 0) & ")"
                        
    ExecuteProcedure "ɾ����DICOM����"
    'ˢ���б�
    Call subFillMSFService
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdModiDevice_Click()
    Dim i As Integer
    Dim blnResult As Boolean    '��Ч�Լ����
    '��������Ƿ���Ч
    blnResult = funValidateDevice
    If blnResult = False Then Exit Sub
    
    On Error GoTo errHand
    '��������
    gstrSQL = "Zl_Ӱ������豸_UPDATE(" & Me.MSFDevice.TextMatrix(Me.MSFDevice.RowSel, 12) & ",'" & Me.txtIPAddr.Text & "','" & Me.txtDeviceName.Text & "','" & _
                    Left(cboModality.Text, InStr(cboModality.Text, "-") - 1) & "'"
    For i = 0 To 2
        If chkService(i).value = 1 Then
            gstrSQL = gstrSQL & ",'" & Me.txtDeviceAE(i).Text & "','" & Me.txtDevicePort(i).Text & "'," & Me.cboService(i).ItemData(Me.cboService(i).ListIndex)
        Else
            gstrSQL = gstrSQL & ",null,null,null"
        End If
    Next i
    gstrSQL = gstrSQL & ")"
                        
    ExecuteProcedure "�޸�DICOM�豸"
    'ˢ���б�
    Call subFillMSFDevice(Me.MSFDevice.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdModiService_Click()
    Dim i As Integer
    Dim blnResult As Boolean    '��Ч�Լ����
    '��������Ƿ���Ч
    blnResult = funValidateService
    If blnResult = False Then Exit Sub
    On Error GoTo errHand
    '�޸�����
    gstrSQL = "Zl_Ӱ��DICOM����_UPDATE(" & Me.MSFService.TextMatrix(Me.MSFService.RowSel, 0) & ",'" & Me.txtServiceName.Text & "','" & Me.txtServiceIP.Text & "','" & _
                    Me.txtServiceAE.Text & "','" & Me.txtServicePort.Text & "','" & Me.cboServiceType.Text & "')"
                        
    ExecuteProcedure "�޸�DICOM����"
    'ˢ���б�
    Call subFillMSFService(Me.MSFService.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub CmdOK_Click()
    Unload Me
End Sub

Private Sub cmdServiceSetup_Click()
    If Me.MSFService.RowSel <> 0 Then
        frmAdvancedSet.ShowMe Me, Me.cboServiceType.Text, lngServiceID
    End If
End Sub

Private Sub Form_Load()
    '���Ӱ������б�
    Call subFillcboModality
    '����ѡ�����б�
    Call subFillcboService
    '����豸�б�
    Call subFillMSFDevice
    '�������б�
    Call subFillMSFService
End Sub

Private Sub subFillcboModality()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    '�����ݿ��ж�ȡ�豸
    strSQL = "Select ����,���� From Ӱ������� "
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡӰ�����")
    
    cboModality.Clear
    While Not rsTmp.EOF
        cboModality.AddItem rsTmp!���� & "-" & (rsTmp!����)
        rsTmp.MoveNext
    Wend
End Sub

Private Sub subFillMSFService(Optional iRow As Integer = 1)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intRowPos As Integer
    
    '�����ݿ��ж�ȡ����
    strSQL = "Select a.����ID,a.������,a.����IP,a.����AE,a.����˿�,a.������  From Ӱ��DICOM���� a "
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡDICOM����")
    
    With MSFService
        .Clear
        .Rows = 1
        .Cols = 6
        .ColWidth(0) = 800
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .ColWidth(5) = 800
        
        .FixedCols = 0
        .TextMatrix(0, 0) = "����ID"
        .TextMatrix(0, 1) = "������"
        .TextMatrix(0, 2) = "������"
        .TextMatrix(0, 3) = "����IP"
        .TextMatrix(0, 4) = "����AE"
        .TextMatrix(0, 5) = "����˿�"
        
        intRowPos = 1
        
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(intRowPos, 0) = Nvl(rsTmp!����ID)
            .TextMatrix(intRowPos, 1) = Nvl(rsTmp!������)
            .TextMatrix(intRowPos, 2) = Nvl(rsTmp!������)
            .TextMatrix(intRowPos, 3) = Nvl(rsTmp!����IP)
            .TextMatrix(intRowPos, 4) = Nvl(rsTmp!����AE)
            .TextMatrix(intRowPos, 5) = Nvl(rsTmp!����˿�)
            
            rsTmp.MoveNext
            intRowPos = .Rows
        Wend
    End With
    subClickMSFService iRow
End Sub

Private Function funValidateDevice() As Boolean
    Dim arrIP() As String
    Dim i As Integer
    
    '�ж�������豸��Ϣ�Ƿ���Ч
    '�豸����Ϊ��
    If Me.txtDeviceName = "" Then
        GoTo inValidate
    End If
    '�豸IP��ַ��Ч
    If Me.txtIPAddr = "" Then
        GoTo inValidate
    Else
        arrIP = Split(Me.txtIPAddr.Text, ".")
        If UBound(arrIP) = 3 Then
            If IsNumeric(arrIP(0)) And IsNumeric(arrIP(1)) And IsNumeric(arrIP(2)) And IsNumeric(arrIP(3)) Then
                '��Ч��������
            Else
                GoTo inValidate
            End If
        Else
            GoTo inValidate
        End If
    End If
    'Ӱ�����Ϊ��
    If Me.cboModality.Text = "" Then
        GoTo inValidate
    End If
    '����������AE���˿ڣ�PACS����
    For i = 0 To 2
        If chkService(i).value = 1 Then
            'ֻ������õ�����
            If Me.txtDeviceAE(i).Text = "" Then GoTo inValidate
            If Me.txtDevicePort(i).Text = "" Then GoTo inValidate
            If Me.cboService(i).ListIndex = -1 Then GoTo inValidate
        End If
    Next i
    
    funValidateDevice = True
    Exit Function
inValidate:
    MsgBox "��������������߲��������������������", vbOKOnly, "������������"
    Exit Function
End Function

Private Sub subFillcboService()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    '�����ݿ��ж�ȡ�豸
    strSQL = "Select ����ID,������,������ From Ӱ��DICOM���� "
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ��ѡDICOM����")
    
    For i = 0 To 2
        cboService(i).Clear
    Next i
    
    While Not rsTmp.EOF
        Select Case rsTmp!������
        Case ZLPACS_�洢����
            i = 0
        Case ZLPACS_Worklist����
            i = 1
        Case ZLPACS_��������
            i = 2
        End Select

        cboService(i).AddItem (rsTmp!������)
        cboService(i).ItemData(cboService(i).ListCount - 1) = rsTmp!����ID
        rsTmp.MoveNext
    Wend
    
End Sub

Private Sub subFillMSFDevice(Optional iRow As Integer = 1)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intRowPos As Integer
    
    '�����ݿ��ж�ȡ�豸
    strSQL = "Select a.����id, a.ip��ַ, a.�豸����,a.Ӱ�����,a.�洢AE,a.�洢�˿�,a.�洢����ID, b1.������ As �洢������, " & _
             "a.WORKLISTAE,a.WORKLIST�˿�,a.WORKLIST����ID,b2.������ As WORKLIST������,a.����AE,a.�����˿�,a.��������ID, " & _
             "b3.������ As ���������� From Ӱ������豸 a ,Ӱ��DICOM���� b1,Ӱ��DICOM���� b2,Ӱ��DICOM���� b3 Where " & _
             "a.�洢����id=b1.����id(+) And a.worklist����id=b2.����id(+) And a.��������id = b3.����id(+) "
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡDICOM�豸")
    
    With MSFDevice
        .Clear
        .Rows = 1
        .Cols = 13
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .ColWidth(2) = 1400
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 800
        .ColWidth(9) = 800
        .ColWidth(10) = 800
        .ColWidth(11) = 800
        .ColWidth(12) = 800
        
        .FixedCols = 0
        '.FixedRows = 1
        .TextMatrix(0, 0) = "�豸����"
        .TextMatrix(0, 1) = "Ӱ�����"
        .TextMatrix(0, 2) = "IP��ַ"
        .TextMatrix(0, 3) = "�洢AE"
        .TextMatrix(0, 4) = "�洢�˿�"
        .TextMatrix(0, 5) = "�洢������"
        .TextMatrix(0, 6) = "WorklistAE"
        .TextMatrix(0, 7) = "Worklist�˿�"
        .TextMatrix(0, 8) = "Worklist������"
        .TextMatrix(0, 9) = "����AE"
        .TextMatrix(0, 10) = "�����˿�"
        .TextMatrix(0, 11) = "����������"
        .TextMatrix(0, 12) = "����ID"
        intRowPos = 1
        
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(intRowPos, 0) = Nvl(rsTmp!�豸����)
            .TextMatrix(intRowPos, 1) = Nvl(rsTmp!Ӱ�����)
            .TextMatrix(intRowPos, 2) = Nvl(rsTmp!IP��ַ)
            .TextMatrix(intRowPos, 3) = Nvl(rsTmp!�洢AE)
            .TextMatrix(intRowPos, 4) = Nvl(rsTmp!�洢�˿�)
            .TextMatrix(intRowPos, 5) = Nvl(rsTmp!�洢������)
            .TextMatrix(intRowPos, 6) = Nvl(rsTmp!WORKLISTAE)
            .TextMatrix(intRowPos, 7) = Nvl(rsTmp!WORKLIST�˿�)
            .TextMatrix(intRowPos, 8) = Nvl(rsTmp!WORKLIST������)
            .TextMatrix(intRowPos, 9) = Nvl(rsTmp!����AE)
            .TextMatrix(intRowPos, 10) = Nvl(rsTmp!�����˿�)
            .TextMatrix(intRowPos, 11) = Nvl(rsTmp!����������)
            .TextMatrix(intRowPos, 12) = Nvl(rsTmp!����id)
            rsTmp.MoveNext
            intRowPos = .Rows
        Wend
    End With
    
    Call subClickMSFDevice(iRow)
End Sub

Private Sub MSFDevice_Click()
    Dim iSelected As Integer
    If MSFDevice.Rows <= 1 Then Exit Sub
    
    With MSFDevice
        iSelected = .RowSel
        '��д������Ϣ
        Me.txtDeviceName.Text = .TextMatrix(iSelected, 0)
        Me.cboModality.Text = funcGetModalityText(.TextMatrix(iSelected, 1))
        Me.txtIPAddr.Text = .TextMatrix(iSelected, 2)
        '��д�洢����
        If .TextMatrix(iSelected, 5) = "" Then
            Me.chkService(0).value = 0
        Else
            Me.chkService(0).value = 1
            Me.txtDeviceAE(0).Text = .TextMatrix(iSelected, 3)
            Me.txtDevicePort(0).Text = .TextMatrix(iSelected, 4)
            Me.cboService(0).Text = .TextMatrix(iSelected, 5)
        End If
        '��дWORKLIST����
        If .TextMatrix(iSelected, 8) = "" Then
            Me.chkService(1).value = 0
        Else
            Me.chkService(1).value = 1
            Me.txtDeviceAE(1).Text = .TextMatrix(iSelected, 6)
            Me.txtDevicePort(1).Text = .TextMatrix(iSelected, 7)
            Me.cboService(1).Text = .TextMatrix(iSelected, 8)
        End If
        '��д��������
        If .TextMatrix(iSelected, 11) = "" Then
            Me.chkService(2).value = 0
        Else
            Me.chkService(2).value = 1
            Me.txtDeviceAE(2).Text = .TextMatrix(iSelected, 9)
            Me.txtDevicePort(2).Text = .TextMatrix(iSelected, 10)
            Me.cboService(2).Text = .TextMatrix(iSelected, 11)
        End If
    End With
End Sub

Private Function funcGetModalityText(strModality As String) As String
    Dim i As Integer
    For i = 0 To cboModality.ListCount - 1
        If Left(cboModality.list(i), InStr(cboModality.list(i), "-") - 1) = strModality Then
            funcGetModalityText = cboModality.list(i)
            Exit Function
        End If
    Next i
End Function

Private Sub MSFService_Click()
    Dim iSelected As Integer
    If MSFService.Rows <= 1 Then Exit Sub
    With MSFService
        iSelected = .RowSel
        lngServiceID = .TextMatrix(iSelected, 0)
        '��д������
        Me.txtServiceName.Text = .TextMatrix(iSelected, 1)
        '��д��������
        Me.cboServiceType.Text = .TextMatrix(iSelected, 2)
        '����IP��ַ
        Me.txtServiceIP.Text = .TextMatrix(iSelected, 3)
        '����AE
        Me.txtServiceAE.Text = .TextMatrix(iSelected, 4)
        '����˿�
        Me.txtServicePort.Text = .TextMatrix(iSelected, 5)
        If Me.cboServiceType.Text = ZLPACS_�������� Then
            cmdServiceSetup.Enabled = False
        Else
            cmdServiceSetup.Enabled = True
        End If
    End With
End Sub

Private Sub subClickMSFService(Optional iRow As Integer = 1)

    If iRow > Me.MSFService.Rows Or iRow < 1 Then iRow = 1

    If Me.MSFService.Rows > 1 Then
        Me.MSFService.Row = iRow - 1
        Me.MSFService.RowSel = iRow
        Call MSFService_Click
    End If
End Sub

Private Sub subClickMSFDevice(Optional iRow As Integer = 1)

    If iRow > Me.MSFDevice.Rows Or iRow < 1 Then iRow = 1

    If Me.MSFDevice.Rows > 1 Then
        Me.MSFDevice.Row = iRow - 1
        Me.MSFDevice.RowSel = iRow
        Call MSFDevice_Click
    End If
End Sub

Private Sub tabService_Click(PreviousTab As Integer)
    '������豸ҳ����ˢ���豸�еĿ�ѡ�����б�
    If tabService.Tab = 0 Then
        '����ѡ�����б�
        Call subFillcboService
        '����豸�б�
        Call subFillMSFDevice
    ElseIf tabService.Tab = 2 Then
        Call subFillMSFDicomDevice
       ' Call subFillMSFDicomService
    
    End If
End Sub

Private Sub subFillMSFDicomDevice()
    '���DICOM�豸
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intRowPos As Integer
    
    '�����ݿ��ж�ȡ�豸
    strSQL = "Select  �豸��,�豸��,����,IP��ַ From Ӱ���豸Ŀ¼ Where ���� =4 "
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡDICOM�豸")
    
    With MSFDicomDevice
        .Clear
        .Rows = 1
        .Cols = 4
        .ColWidth(0) = 1400
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        
        .FixedCols = 0
        .TextMatrix(0, 0) = "�豸��"
        .TextMatrix(0, 1) = "�豸����"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "IP��ַ"
        intRowPos = 1
        
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(intRowPos, 0) = Nvl(rsTmp!�豸��)
            .TextMatrix(intRowPos, 1) = Nvl(rsTmp!�豸��)
            .TextMatrix(intRowPos, 2) = "DICOMӰ���豸"
            .TextMatrix(intRowPos, 3) = Nvl(rsTmp!IP��ַ)
            rsTmp.MoveNext
            intRowPos = .Rows
        Wend
    End With
    
   ' Call subClickMSFDevice(iRow)
End Sub

Private Sub txtDevicePort_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtServicePort_KeyPress(KeyAscii As Integer)
    If (Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
