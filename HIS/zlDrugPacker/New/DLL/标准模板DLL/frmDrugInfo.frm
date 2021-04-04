VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDrugInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҩƷ��Ϣ�ϴ�"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6030
   Icon            =   "frmDrugInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6030
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdUpload 
      Caption         =   "�ϴ�(&U)"
      Height          =   360
      Left            =   3480
      TabIndex        =   4
      Top             =   5400
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   4680
      TabIndex        =   5
      Top             =   5400
      Width           =   1110
   End
   Begin TabDlg.SSTab sstDrug 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "����"
      TabPicture(0)   =   "frmDrugInfo.frx":0A02
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvwDrugType"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�ϴ���Ϣ"
      TabPicture(1)   =   "frmDrugInfo.frx":0A1E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstMess"
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstMess 
         Height          =   3840
         ItemData        =   "frmDrugInfo.frx":0A3A
         Left            =   -74880
         List            =   "frmDrugInfo.frx":0A3C
         TabIndex        =   6
         Top             =   480
         Width           =   5535
      End
      Begin MSComctlLib.TreeView tvwDrugType 
         Height          =   3975
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   7011
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.ComboBox cboDevice 
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�豸"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmDrugInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const STR_ROOT = "ROOT"

Public Sub ShowMe(ByVal frmOwner As Form)
    
    Call InitDevice
    Call InitDrugType
    
    cmdUpload.Enabled = False
    
    If cboDevice.ListCount = 0 Then
        MsgBox "��δע��ҩ���Զ����豸��", vbInformation, GSTR_INTERFACE_NAME
        Unload Me
        Exit Sub
    End If

    Me.Show vbModal, frmOwner
End Sub

Private Sub cboDevice_Change()
    Dim i As Integer
    Dim blnAll As Boolean
    For i = 1 To tvwDrugType.Nodes.Count
        If tvwDrugType.Nodes(i).Checked Then
            blnAll = True
            Exit For
        End If
    Next
    cmdUpload.Enabled = cboDevice.ListIndex >= 0 And blnAll
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpload_Click()
    Dim rsData As ADODB.Recordset
    Dim strTmp As String, strConnect As String
    Dim i As Integer
    Dim objDevice As clsDevice
    
    '�豸����
    Set objDevice = New clsDevice
    objDevice.ID = cboDevice.ItemData(cboDevice.ListIndex)
    
    If objDevice.Status = False Then
        MsgBox "��" & objDevice.Name & "���豸δ���ӳɹ���", vbInformation, GSTR_INTERFACE_NAME
        Set objDevice = Nothing
        Exit Sub
    End If
    
    'ҩƷ�����ַ���
    If tvwDrugType.Nodes(STR_ROOT).Checked = True Then
        '���м���
        strTmp = "0"
    Else
        'strTmp�գ���ʾδѡ�����
        For i = 1 To tvwDrugType.Nodes.Count
            If tvwDrugType.Nodes(i).Key <> STR_ROOT And tvwDrugType.Nodes(i).Checked = True Then
                strTmp = strTmp & tvwDrugType.Nodes(i).Text & ","
            End If
        Next
    End If
    
    If strTmp = "" Then Exit Sub
    If Right(strTmp, 1) = "," Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    
    cmdUpload.Enabled = False
    Screen.MousePointer = vbHourglass
    
    '�õ�Ҫ�ϴ�������
    Set rsData = mdlProcessData.ProcDrugInfo(strTmp, objDevice)
    
    '�ϴ����豸
    If Not rsData Is Nothing Then
        lstMess.Clear
        With rsData
            If rsData.State <> adStateOpen Then rsData.Open
            If .RecordCount > 0 Then .MoveFirst
            Do While .EOF = False
                If mdlDrugPacker.DrugInfo(objDevice, !Content) = False Then
                    strTmp = "ҩƷ��" & rsData!Drug & "���ϴ�ʧ�ܣ�"
                    lstMess.AddItem strTmp
                End If
                .MoveNext
            Loop
            .Close
        End With
    End If
    
    Set objDevice = Nothing
    Screen.MousePointer = vbDefault
    cmdUpload.Enabled = True
    sstDrug.TabIndex = 1
    
End Sub

Private Sub InitDrugType()
'���ܣ�����ҩƷ����

    Dim rsTmp As ADODB.Recordset
    
    tvwDrugType.Nodes.Add , , STR_ROOT, "ȫ��"
    
    gstrSQL = "Select ����, ���� From ҩƷ���� Order By ���� "
    On Error GoTo errHandle
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ����")
    Do While Not rsTmp.EOF
        tvwDrugType.Nodes.Add STR_ROOT, tvwChild, "N_" & rsTmp!����, rsTmp!����
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    tvwDrugType.Nodes(STR_ROOT).Expanded = True
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub InitDevice()
'���ܣ���������
    
    Dim rsTmp As ADODB.Recordset
        
    'gstrSQL = "Select ID, ���� From ҩ����ҩ�豸 Where �Ƿ����� = 1 Order By ���� "
    
    'ͬһ�����ӵ��豸��ֻȡһ���豸
    gstrSQL = "Select a.Id, b.����, b.���� " & vbCr & _
              "From (Select Max(ID) ID " & vbCr & _
              "       From ҩ����ҩ�豸 " & vbCr & _
              "       Where �������� Is Not Null And �Ƿ����� = 1 " & vbCr & _
              "       Group By ��������, ��������) A, ҩ����ҩ�豸 B " & vbCr & _
              "Where a.Id = b.Id " & vbCr & _
              "Order By b.���� "

    On Error GoTo errHandle
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ����ҩ�豸")
    Do While Not rsTmp.EOF
        cboDevice.AddItem "��" & rsTmp!���� & "��" & rsTmp!����
        cboDevice.ItemData(cboDevice.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    If cboDevice.ListCount > 0 Then cboDevice.ListIndex = 0
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub tvwDrugType_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    Dim blnAll As Boolean
    
    If Node.Key = STR_ROOT Then
        cmdUpload.Enabled = Node.Checked
        For i = 1 To tvwDrugType.Nodes.Count
            tvwDrugType.Nodes(i).Checked = Node.Checked
        Next
    Else
        blnAll = True
        For i = 1 To tvwDrugType.Nodes.Count
            If tvwDrugType.Nodes(i).Checked = False And tvwDrugType.Nodes(i).Key <> STR_ROOT Then
                blnAll = False
                Exit For
            End If
        Next
        tvwDrugType.Nodes(STR_ROOT).Checked = blnAll
        
        For i = 1 To tvwDrugType.Nodes.Count
            If tvwDrugType.Nodes(i).Checked Then
                cmdUpload.Enabled = True And Me.cboDevice.ListIndex >= 0
                Exit Sub
            End If
        Next
        cmdUpload.Enabled = False
    End If
End Sub

